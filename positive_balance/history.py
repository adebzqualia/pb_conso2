"""Versioned snapshots. A rerun replaces one complete reporting period."""

from contextlib import contextmanager
import json
import os
from pathlib import Path
import tempfile


def load_history(path: Path) -> dict:
    if not path.exists():
        return {"schema_version": 1, "snapshots": {}}
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except (ValueError, OSError) as exc:
        raise ValueError(f"Historique illisible : {path}: {exc}") from exc
    if data.get("schema_version") != 1 or not isinstance(data.get("snapshots"), dict):
        raise ValueError("Format/version d'historique non pris en charge.")
    for period, snapshot in data["snapshots"].items():
        if snapshot.get("period") != period or not isinstance(snapshot.get("entities"), list):
            raise ValueError(f"Snapshot historique invalide : {period}")
    return data


def merge_snapshot(history: dict, snapshot: dict) -> dict:
    """No append of overlapping cohorts; the reporting period is the snapshot key."""
    return {"schema_version": 1,
            "snapshots": dict(sorted({**history["snapshots"], snapshot["period"]: snapshot}.items()))}


def atomic_json(path: Path, data: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    fd, temporary = tempfile.mkstemp(prefix=path.stem + "_", suffix=".tmp", dir=path.parent)
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as handle:
            json.dump(data, handle, ensure_ascii=False, indent=2, allow_nan=False)
            handle.write("\n")
        os.replace(temporary, path)
    finally:
        Path(temporary).unlink(missing_ok=True)


@contextmanager
def history_lock(path: Path):
    """Fail explicitly on concurrent runs, rather than lose another quarter."""
    path.parent.mkdir(parents=True, exist_ok=True)
    lock = path.with_suffix(path.suffix + ".lock")
    try:
        fd = os.open(lock, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
    except FileExistsError as exc:
        raise ValueError(f"Historique verrouillé : {lock}. Une autre exécution est en cours. "
                         "Après un arrêt anormal, supprimer ce verrou uniquement si aucune exécution ne tourne.") from exc
    try:
        with os.fdopen(fd, "w") as handle:
            handle.write(str(os.getpid()))
        yield
    finally:
        lock.unlink(missing_ok=True)
