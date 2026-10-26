import logging
from logging.handlers import RotatingFileHandler
import threading

_LOCK = threading.Lock()

def configure_logging(output_dir):
    directory = output_dir / "logs"
    directory.mkdir(parents=True, exist_ok=True)
    with _LOCK:
        for name, filename in (("kyc_app", "application.log"), ("kyc_structured", "kyc_structured.log")):
            logger = logging.getLogger(name)
            logger.setLevel(logging.INFO)
            logger.propagate = False
            target = str((directory / filename).resolve())
            if not any(getattr(h, "baseFilename", None) == target for h in logger.handlers):
                handler = RotatingFileHandler(target, maxBytes=10*1024**2, backupCount=5, encoding="utf-8")
                handler.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(message)s"))
                logger.addHandler(handler)
    return logging.getLogger("kyc_app")
