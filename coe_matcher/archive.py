"""Safe ZIP extraction and file inventory/classification."""
import os
import zipfile
from dataclasses import dataclass, field
from typing import List

MAX_FILES = 20000
MAX_TOTAL_BYTES = 4 * 1024 ** 3          # 4 GiB uncompressed cap
EXCEL_EXT = (".xlsx", ".xlsm")
SOURCE_EXT = (".csv", ".xlsx", ".xlsm")


@dataclass
class ArchiveReport:
    zip_name: str
    extract_dir: str
    ok: bool = True
    error: str = ""
    n_members: int = 0
    skipped: List[str] = field(default_factory=list)   # unsafe / unsupported members, with reason
    files: List[str] = field(default_factory=list)     # extracted file paths (relative to extract_dir)


def safe_extract(zip_path: str, dest: str) -> ArchiveReport:
    """Extract a ZIP with zip-slip protection and size limits. Never raises: problems are reported."""
    rep = ArchiveReport(os.path.basename(zip_path), dest)
    os.makedirs(dest, exist_ok=True)
    try:
        with zipfile.ZipFile(zip_path) as zf:
            infos = zf.infolist()
            rep.n_members = len(infos)
            if len(infos) > MAX_FILES:
                rep.ok, rep.error = False, f"archive has {len(infos)} members (limit {MAX_FILES})"
                return rep
            if sum(i.file_size for i in infos) > MAX_TOTAL_BYTES:
                rep.ok, rep.error = False, "uncompressed size exceeds limit"
                return rep
            root = os.path.realpath(dest)
            for info in infos:
                if info.is_dir():
                    continue
                name = info.filename.replace("\\", "/")
                target = os.path.realpath(os.path.join(dest, name))
                if not (target == root or target.startswith(root + os.sep)):
                    rep.skipped.append(f"{name}: unsafe path (blocked)")
                    continue
                if info.flag_bits & 0x1:
                    rep.skipped.append(f"{name}: encrypted member")
                    continue
                try:
                    os.makedirs(os.path.dirname(target), exist_ok=True)
                    with zf.open(info) as src, open(target, "wb") as out:
                        while True:
                            chunk = src.read(1 << 20)
                            if not chunk:
                                break
                            out.write(chunk)
                    rep.files.append(os.path.relpath(target, dest).replace(os.sep, "/"))
                except Exception as e:  # corrupt member
                    rep.skipped.append(f"{name}: {e!r}")
    except zipfile.BadZipFile as e:
        rep.ok, rep.error = False, f"not a valid ZIP: {e}"
    except Exception as e:
        rep.ok, rep.error = False, repr(e)
    return rep


def is_lock_file(rel_path: str) -> bool:
    return os.path.basename(rel_path).startswith("~$")


def is_junk(rel_path: str) -> bool:
    b = os.path.basename(rel_path)
    return b.startswith("._") or b.lower() in ("thumbs.db", ".ds_store") or "__MACOSX" in rel_path
