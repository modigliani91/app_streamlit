"""safe_member_path copié identiquement de la cellule 22."""
import zipfile
import stat
from pathlib import PurePosixPath, PureWindowsPath

def safe_member_path(info: zipfile.ZipInfo) -> PurePosixPath:
    raw = info.filename
    normalized = raw.replace("\\", "/")
    path = PurePosixPath(normalized)
    if (not raw or "\x00" in raw or path.is_absolute() or PureWindowsPath(raw).drive
            or ".." in normalized.split("/") or any(":" in p for p in path.parts)):
        raise ValueError(f"Chemin ZIP interdit : {raw!r}")
    mode = info.external_attr >> 16
    kind = stat.S_IFMT(mode)
    if stat.S_ISLNK(mode) or kind not in {0, stat.S_IFREG, stat.S_IFDIR}:
        raise ValueError(f"Lien ou entrée spéciale ZIP interdite : {raw!r}")
    if not path.parts:
        raise ValueError("Entrée ZIP sans nom.")
    return path
