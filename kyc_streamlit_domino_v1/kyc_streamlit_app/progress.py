from dataclasses import dataclass, asdict
from datetime import datetime, timezone

def utc_now():
    return datetime.now(timezone.utc).isoformat()

@dataclass(frozen=True)
class ProgressEvent:
    stage: str
    message: str
    current: int | None = None
    total: int | None = None

    def to_dict(self):
        return {**asdict(self), "updated_at": utc_now()}
