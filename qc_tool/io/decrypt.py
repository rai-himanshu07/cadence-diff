"""Encryption detection and in-memory decryption of Office files.

Sources are only ever opened for reading; decryption output goes to an
in-memory stream, never back to disk.
"""

import io
from pathlib import Path

import msoffcrypto
import msoffcrypto.exceptions

from qc_tool.io.opc import validate_office_package

#: OLE compound-file magic — encrypted OOXML files are CFB containers.
_OLE_MAGIC = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
_ZIP_MAGIC = b"PK\x03\x04"


class PasswordRequiredError(ValueError):
    """The file is encrypted and no password was supplied."""


class InvalidPasswordError(ValueError):
    """The supplied password does not decrypt the file."""


class UnsupportedFileError(ValueError):
    """The file is neither a plain OOXML zip nor an encrypted container."""


def is_encrypted(path: Path) -> bool:
    with path.open("rb") as fh:
        magic = fh.read(8)
    if magic.startswith(_ZIP_MAGIC):
        return False
    if magic == _OLE_MAGIC:
        return True
    raise UnsupportedFileError(
        f"{path.name}: unrecognized file signature; expected an OOXML zip or an "
        "encrypted Office container"
    )


def open_decrypted(path: Path, password: str | None = None) -> io.BytesIO:
    """Return the plaintext package bytes of ``path`` as an in-memory stream."""
    if not is_encrypted(path):
        data = path.read_bytes()
        validate_office_package(data, source_name=path.name)
        return io.BytesIO(data)
    if password is None:
        raise PasswordRequiredError(
            f"{path.name} is password-protected; supply the open password"
        )
    plaintext = io.BytesIO()
    with path.open("rb") as fh:
        office_file = msoffcrypto.OfficeFile(fh)
        try:
            office_file.load_key(password=password)
            office_file.decrypt(plaintext)
        except msoffcrypto.exceptions.InvalidKeyError as exc:
            raise InvalidPasswordError(f"{path.name}: {exc}") from exc
        except msoffcrypto.exceptions.DecryptionError as exc:
            raise InvalidPasswordError(
                f"{path.name}: decryption failed — likely a wrong password ({exc})"
            ) from exc
    plaintext.seek(0)
    validate_office_package(plaintext.getvalue(), source_name=path.name)
    return plaintext
