"""Entrées, métadonnées atomiques et verrous POSIX. Aucun import GPU."""
from contextlib import contextmanager
from dataclasses import dataclass
from pathlib import Path, PurePosixPath
import fcntl
import hashlib
import json
import os
import socket
import tempfile
import threading
import zipfile
import zlib

from archive_paths import safe_member_path
from exceptions import ConcurrentRun, InvalidArchive
from progress import utc_now

_PROCESS_LOCK = threading.Lock()

def sha256_file(path):
    digest = hashlib.sha256()
    with Path(path).open("rb") as stream:
        for block in iter(lambda: stream.read(1024**2), b""): digest.update(block)
    return digest.hexdigest()

def atomic_json(path, value):
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    fd, name = tempfile.mkstemp(prefix=".write-", dir=path.parent)
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as stream:
            json.dump(value, stream, ensure_ascii=False, indent=2, allow_nan=False)
            stream.flush(); os.fsync(stream.fileno())
        os.replace(name, path)
    finally:
        Path(name).unlink(missing_ok=True)

class RunLock:
    """Un inode stable; jamais de suppression ni de cassage par délai/PID.

    flock est libéré par le noyau à la fermeture du FD ou à la mort du processus.
    Le fichier JSON peut rester après un crash: son existence n'est pas un verrou.
    """
    def __init__(self, path, run_id):
        self.path, self.run_id, self.fd = Path(path), run_id, None
        self.local_acquired = False

    def acquire(self):
        if not _PROCESS_LOCK.acquire(blocking=False):
            raise ConcurrentRun("Un traitement est déjà en cours.")
        self.local_acquired = True
        try:
            self.path.parent.mkdir(parents=True, exist_ok=True)
            self.fd = os.open(self.path, os.O_RDWR | os.O_CREAT | os.O_NOFOLLOW, 0o600)
            try: fcntl.flock(self.fd, fcntl.LOCK_EX | fcntl.LOCK_NB)
            except BlockingIOError as exc:
                raise ConcurrentRun("Un traitement est déjà en cours.") from exc
            payload = json.dumps({"run_id": self.run_id, "pid": os.getpid(), "hostname": socket.gethostname(), "started_at": utc_now()}).encode()
            os.ftruncate(self.fd, 0); os.lseek(self.fd, 0, os.SEEK_SET)
            os.write(self.fd, payload); os.fsync(self.fd)
            return self
        except BaseException:
            self.release(); raise

    def release(self):
        if self.fd is not None:
            os.close(self.fd); self.fd = None
        if self.local_acquired:
            self.local_acquired = False; _PROCESS_LOCK.release()

    def __enter__(self): return self.acquire()
    def __exit__(self, *args): self.release()

def is_busy(path):
    if _PROCESS_LOCK.locked(): return True
    path = Path(path)
    if not path.exists(): return False
    fd = os.open(path, os.O_RDWR | os.O_NOFOLLOW)
    try:
        try: fcntl.flock(fd, fcntl.LOCK_EX | fcntl.LOCK_NB)
        except BlockingIOError: return True
        return False
    finally: os.close(fd)

@dataclass(frozen=True)
class ArchiveInfo:
    members: int
    customers: int
    uncompressed_bytes: int

def validate_archive(source, settings, *, verify_crc=False):
    """Validation de TOUTES les entrées; l'extraction ciblée reste celle du notebook."""
    from notebook_config import StructuredConfig
    limits = StructuredConfig()
    if hasattr(source, "seek"): source.seek(0)
    try:
        with zipfile.ZipFile(source) as archive:
            infos = archive.infolist()
            if not infos or len(infos) > limits.MAX_ZIP_MEMBERS:
                raise InvalidArchive("L'archive ZIP est vide ou contient trop d'entrées.")
            validated, seen, total = [], set(), 0
            for item in infos:
                path = safe_member_path(item)
                key = path.as_posix()
                if key in seen: raise InvalidArchive("L'archive contient des chemins dupliqués.")
                seen.add(key); validated.append((item, path))
                if item.flag_bits & 1: raise InvalidArchive("Les archives chiffrées ne sont pas prises en charge.")
                if item.file_size > limits.MAX_PDF_BYTES:
                    raise InvalidArchive("Un fichier de l'archive dépasse la taille autorisée.")
                total += item.file_size
                if total > limits.MAX_EXTRACT_BYTES or item.file_size / max(1, item.compress_size) > limits.MAX_COMPRESSION_RATIO:
                    raise InvalidArchive("L'archive dépasse les limites de décompression.")
            # Reject file/directory collisions before any extraction.
            files = {p.as_posix() for i,p in validated if not i.is_dir()}
            for _,path in validated:
                if any(parent.as_posix() in files for parent in path.parents if parent.parts):
                    raise InvalidArchive("Un chemin de fichier est aussi utilisé comme dossier.")
            prefix = ()
            if settings.zip_customer_root:
                root = PurePosixPath(settings.zip_customer_root.replace("\\", "/"))
                if root.is_absolute() or ".." in root.parts:
                    raise InvalidArchive("La racine ZIP configurée est invalide.")
                prefix = root.parts
            else:
                tops = {p.parts[0] for i,p in validated if p.parts[0] != "__MACOSX" and (i.is_dir() or len(p.parts)>=2)}
                stem = Path(settings.input_filename).stem
                if tops == {stem}: prefix = (stem,)
            customers = set()
            for item,path in validated:
                if path.parts[0] == "__MACOSX" or tuple(path.parts[:len(prefix)]) != prefix: continue
                parts = path.parts[len(prefix):]
                if parts and (item.is_dir() or len(parts)>=2): customers.add(parts[0])
            if not customers or not any(not i.is_dir() for i in infos):
                raise InvalidArchive("Aucun dossier client contenant des fichiers n'a été trouvé.")
            if verify_crc:
                # Stream CRC validation: bounded RAM, no unsafe extractall, even for ignored files.
                for item in infos:
                    if item.is_dir(): continue
                    count = 0
                    with archive.open(item) as stream:
                        for block in iter(lambda: stream.read(1024**2), b""):
                            count += len(block)
                            if count > item.file_size: raise InvalidArchive("Taille ZIP incohérente.")
                    if count != item.file_size: raise InvalidArchive("Un fichier ZIP est incomplet.")
            return ArchiveInfo(len(infos), len(customers), total)
    except InvalidArchive: raise
    except (ValueError, OSError, zipfile.BadZipFile, RuntimeError, NotImplementedError, EOFError, zlib.error) as exc:
        raise InvalidArchive("L'archive ZIP fournie n'est pas valide ou contient un chemin interdit.") from exc
    finally:
        if hasattr(source, "seek"): source.seek(0)

def save_upload(upload, settings):
    """Appelé sous RunLock; remplace seulement l'entrée canonique après contrôle de structure."""
    if Path(upload.name).name != settings.input_filename or Path(upload.name).suffix.lower() != ".zip":
        raise InvalidArchive("Le fichier doit être nommé kyc_documents_test.zip.")
    settings.input_dir.mkdir(parents=True, exist_ok=True)
    fd, name = tempfile.mkstemp(prefix=".upload-", suffix=".zip", dir=settings.input_dir)
    try:
        count = 0; upload.seek(0)
        with os.fdopen(fd,"wb") as stream:
            for block in iter(lambda: upload.read(1024**2), b""):
                count += len(block)
                if count > settings.max_upload_mb*1024**2:
                    raise InvalidArchive("L'archive dépasse la taille maximale configurée.")
                stream.write(block)
            stream.flush(); os.fsync(stream.fileno())
        if count == 0: raise InvalidArchive("L'archive ZIP est vide.")
        validate_archive(Path(name), settings)
        os.replace(name, settings.input_path)
        return sha256_file(settings.input_path)
    finally:
        Path(name).unlink(missing_ok=True); upload.seek(0)

def readiness(settings):
    errors=[]
    if not settings.reference_csv.is_file():
        errors.append(f"Le fichier de référence {settings.reference_csv} est introuvable.")
    if not settings.model_path.is_dir() or not (settings.model_path/"config.json").is_file():
        errors.append(f"Le répertoire du modèle ou son config.json est introuvable : {settings.model_path}")
    for directory in (settings.output_dir, settings.input_dir, settings.extract_dir):
        try:
            directory.mkdir(parents=True, exist_ok=True)
            fd,name=tempfile.mkstemp(prefix=".writable-",dir=directory)
            os.close(fd); Path(name).unlink()
        except OSError: errors.append(f"Le répertoire {directory} n'est pas accessible en écriture.")
    return errors
