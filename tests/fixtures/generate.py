"""Generate the deterministic QC fixture set with its ground-truth manifest.

Usage: ``conda run -n py311 python -m tests.fixtures.generate [dest]``
(defaults to ``tests/fixtures/generated/``, which is gitignored).

Determinism: document timestamps are pinned and every zip is rewritten with
sorted entries and epoch dates, so consecutive runs are byte-identical —
except ``current_encrypted.xlsx``, whose encryption salt is random by design
(tests assert decrypt-equality instead).
"""

import io
import logging
import re
import sys
import zipfile
from pathlib import Path

from msoffcrypto.format.ooxml import OOXMLFile

from tests.fixtures import domain, excel_builder, ppt_builder
from tests.fixtures.manifest_schema import (
    Artifact,
    DefectClass,
    ExpectedChange,
    FixtureManifest,
    SeededDefect,
)
from tests.fixtures.xlsb_writer import CellValue, write_xlsb

logger = logging.getLogger(__name__)

FIXED_GENERATED_AT = "2026-07-26T00:00:00+00:00"
_ZIP_EPOCH = (1980, 1, 1, 0, 0, 0)
_DCTERMS_RE = re.compile(
    rb"(<dcterms:(?:created|modified)[^>]*>)[^<]*(</dcterms:(?:created|modified)>)"
)
_FIXED_DCTERMS = rb"\g<1>2026-07-01T00:00:00Z\g<2>"


def _rewrite_deterministic(parts: dict[str, bytes]) -> bytes:
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for name in sorted(parts):
            info = zipfile.ZipInfo(name, date_time=_ZIP_EPOCH)
            info.compress_type = zipfile.ZIP_DEFLATED
            zf.writestr(info, parts[name])
    return buffer.getvalue()


def _normalize_embedded_xlsx(data: bytes) -> bytes:
    """Normalize a chart's embedded workbook (XlsxWriter stamps wall-clock times)."""
    with zipfile.ZipFile(io.BytesIO(data)) as zf:
        parts = {info.filename: zf.read(info.filename) for info in zf.infolist()}
    if "docProps/core.xml" in parts:
        parts["docProps/core.xml"] = _DCTERMS_RE.sub(_FIXED_DCTERMS, parts["docProps/core.xml"])
    return _rewrite_deterministic(parts)


def normalize_zip(
    path: Path,
    extra_parts: dict[str, bytes] | None = None,
    content_type_overrides: dict[str, str] | None = None,
) -> None:
    """Rewrite a zip deterministically; optionally inject parts + content types."""
    with zipfile.ZipFile(path) as zf:
        parts = {info.filename: zf.read(info.filename) for info in zf.infolist()}
    parts.update(extra_parts or {})
    if "docProps/core.xml" in parts:
        # openpyxl stamps dcterms:modified with wall-clock time at save.
        parts["docProps/core.xml"] = _DCTERMS_RE.sub(_FIXED_DCTERMS, parts["docProps/core.xml"])
    for name, data in parts.items():
        if name.startswith("ppt/embeddings/") and name.endswith(".xlsx"):
            parts[name] = _normalize_embedded_xlsx(data)
    if content_type_overrides:
        content_types = parts["[Content_Types].xml"].decode("utf-8")
        overrides = "".join(
            f'<Override PartName="{name}" ContentType="{ctype}"/>'
            for name, ctype in content_type_overrides.items()
        )
        parts["[Content_Types].xml"] = content_types.replace(
            "</Types>", overrides + "</Types>"
        ).encode("utf-8")
    path.write_bytes(_rewrite_deterministic(parts))


def encrypt_file(src: Path, dest: Path, password: str) -> None:
    with src.open("rb") as fin:
        office_file = OOXMLFile(fin)
        with dest.open("wb") as fout:
            office_file.encrypt(password, fout)


def _xlsb_rows(*, current: bool) -> list[list[CellValue]]:
    months = domain.CURRENT_MONTHS if current else domain.BASELINE_MONTHS
    revenue_fn = domain.current_monthly_revenue if current else domain.monthly_revenue
    rows: list[list[CellValue]] = [["Period", "Region", "Revenue", "Cost", "Margin"]]
    for month in range(months):
        for region in range(len(domain.REGION_LABELS)):
            revenue = revenue_fn(month, region)
            cost = domain.monthly_cost(month, region)
            rows.append(
                [
                    domain.MONTH_LABELS[month],
                    domain.REGION_LABELS[region],
                    revenue,
                    cost,
                    revenue - cost,
                ]
            )
    return rows


def _build_xlsb_pair(dest: Path) -> tuple[list[SeededDefect], list[ExpectedChange]]:
    write_xlsb(dest / "baseline.xlsb", {"Long_Monthly": _xlsb_rows(current=False)})
    write_xlsb(dest / "current.xlsb", {"Long_Monthly": _xlsb_rows(current=True)})
    defects = [
        SeededDefect(
            defect_id="XB01",
            artifact=Artifact.XLSB,
            classes=[DefectClass.VALUE_CHANGED],
            sheet="Long_Monthly",
            cell="C7",
            baseline=str(domain.monthly_revenue(1, 1)),
            current=str(domain.monthly_revenue(1, 1) + domain.E01_DELTA),
            note="historical revenue edited (xlsb mirror of E01)",
        )
    ]
    expected = [
        ExpectedChange(
            change_id="XBX1",
            artifact=Artifact.XLSB,
            kind="new_rows",
            sheet="Long_Monthly",
            detail="Jun-26 rows appended",
        )
    ]
    return defects, expected


def generate(dest: Path) -> FixtureManifest:
    """Build all fixture artifacts into ``dest`` and write manifest.json."""
    dest.mkdir(parents=True, exist_ok=True)

    excel_defects, excel_expected = excel_builder.build_workbooks(dest)
    normalize_zip(
        dest / "baseline.xlsx",
        extra_parts=excel_builder.pivot_parts(current=False),
        content_type_overrides=excel_builder.PIVOT_CONTENT_TYPES,
    )
    normalize_zip(
        dest / "current.xlsx",
        extra_parts=excel_builder.pivot_parts(current=True),
        content_type_overrides=excel_builder.PIVOT_CONTENT_TYPES,
    )
    encrypt_file(
        dest / "current.xlsx", dest / "current_encrypted.xlsx", domain.FIXTURE_PASSWORD
    )

    xlsb_defects, xlsb_expected = _build_xlsb_pair(dest)

    ppt_defects, ppt_expected, crosscheck = ppt_builder.build_decks(dest)
    normalize_zip(dest / "baseline.pptx")
    normalize_zip(dest / "current.pptx")

    manifest = FixtureManifest(
        seed=0,
        generated_at=FIXED_GENERATED_AT,
        password=domain.FIXTURE_PASSWORD,
        files={
            "baseline_xlsx": "baseline.xlsx",
            "current_xlsx": "current.xlsx",
            "current_xlsx_encrypted": "current_encrypted.xlsx",
            "baseline_xlsb": "baseline.xlsb",
            "current_xlsb": "current.xlsb",
            "baseline_pptx": "baseline.pptx",
            "current_pptx": "current.pptx",
        },
        defects=excel_defects + xlsb_defects + ppt_defects,
        expected=excel_expected + xlsb_expected + ppt_expected,
        crosscheck_map=crosscheck,
    )
    (dest / "manifest.json").write_text(manifest.model_dump_json(indent=2), encoding="utf-8")
    logger.info(
        "generated %d files, %d defects, %d expected changes at %s",
        len(manifest.files) + 1,
        len(manifest.defects),
        len(manifest.expected),
        dest,
    )
    return manifest


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s")
    target = Path(sys.argv[1]) if len(sys.argv) > 1 else Path(__file__).parent / "generated"
    generate(target)
