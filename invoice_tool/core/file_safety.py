from __future__ import annotations

import hashlib
import os
import shutil
from pathlib import Path
from typing import Any, Dict, Mapping
from uuid import uuid4

from ..infra.paths import is_relative_to


FINGERPRINT_ALGORITHM = "sha256"
FINGERPRINT_CHUNK_SIZE = 1024 * 1024


class UnsafeFileOperationError(ValueError):
    """Raised when a recorded file operation cannot be proven safe."""


def resolve_contained_path(path: Path, root: Path, *, allow_root: bool = False) -> Path:
    """Resolve *path* and ensure it remains inside the trusted *root*."""

    resolved_root = root.resolve()
    resolved_path = path.resolve()
    if (not allow_root and resolved_path == resolved_root) or not is_relative_to(resolved_path, resolved_root):
        raise UnsafeFileOperationError(f"路径超出受信目录：{path}")
    return resolved_path


def fingerprint_file(path: Path) -> Dict[str, Any]:
    """Return a stable content fingerprint and reject files changing mid-read."""

    if path.is_symlink():
        raise UnsafeFileOperationError(f"不允许处理符号链接：{path.name}")
    before = path.stat()
    if not path.is_file():
        raise OSError(f"不是普通文件：{path}")

    digest = hashlib.sha256()
    with path.open("rb") as stream:
        while chunk := stream.read(FINGERPRINT_CHUNK_SIZE):
            digest.update(chunk)

    after = path.stat()
    if before.st_size != after.st_size or before.st_mtime_ns != after.st_mtime_ns:
        raise OSError(f"读取期间文件发生变化：{path.name}")
    return {
        "algorithm": FINGERPRINT_ALGORITHM,
        "sha256": digest.hexdigest(),
        "size": after.st_size,
    }


def has_valid_fingerprint(value: object) -> bool:
    if not isinstance(value, Mapping):
        return False
    digest = value.get("sha256")
    size = value.get("size")
    return (
        value.get("algorithm") == FINGERPRINT_ALGORITHM
        and isinstance(digest, str)
        and len(digest) == 64
        and all(character in "0123456789abcdef" for character in digest.lower())
        and isinstance(size, int)
        and size >= 0
    )


def fingerprint_matches(path: Path, expected: object) -> bool:
    if not has_valid_fingerprint(expected):
        return False
    actual = fingerprint_file(path)
    expected_mapping = dict(expected)  # type: ignore[arg-type]
    return (
        actual["algorithm"] == expected_mapping["algorithm"]
        and actual["sha256"] == str(expected_mapping["sha256"]).lower()
        and actual["size"] == expected_mapping["size"]
    )


def copy_file_exclusive(
    source: Path,
    target: Path,
    expected_fingerprint: object,
) -> None:
    """Copy via a verified temporary file and never replace an existing target."""

    target.parent.mkdir(parents=True, exist_ok=True)
    temporary = target.with_name(f".{target.name}.{uuid4().hex}.tmp")
    try:
        shutil.copy2(source, temporary)
        if not fingerprint_matches(temporary, expected_fingerprint):
            raise OSError(f"复制后校验失败：{target.name}")
        with temporary.open("rb+") as stream:
            os.fsync(stream.fileno())

        if os.name == "nt":
            os.rename(temporary, target)
        else:
            os.link(temporary, target)
            temporary.unlink()
    finally:
        try:
            temporary.unlink(missing_ok=True)
        except OSError:
            pass
