"""Standalone HTML report rendered from the triaged `QCRunResult`.

Self-contained (inline CSS, vanilla JS, no external assets) and
autoescaped — finding text originates from client files and must never
inject markup.
"""

import datetime as dt
import logging
from pathlib import Path

from jinja2 import Environment, FileSystemLoader, select_autoescape

from qc_tool.engine import QCRunResult
from qc_tool.security import private_directory, private_file

logger = logging.getLogger(__name__)

_ENV = Environment(
    loader=FileSystemLoader(Path(__file__).parent / "templates"),
    autoescape=select_autoescape(enabled_extensions=("j2", "html")),
)


def render_html_report(result: QCRunResult) -> str:
    template = _ENV.get_template("report.html.j2")
    return template.render(
        result=result,
        counts=result.counts,
        generated_at=dt.datetime.now(dt.UTC).isoformat(timespec="seconds"),
    )


def write_html_report(result: QCRunResult, path: Path) -> None:
    private_directory(path.parent)
    path.write_text(render_html_report(result), encoding="utf-8")
    private_file(path)
    logger.info("HTML report written to %s (%d findings)", path, len(result.findings))
