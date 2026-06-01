"""
Forensic logging for the ASI QB middleware.

Each operation appends to logs/{operation}_history.jsonl — one JSON line per
call, trimmed to _MAX_JSONL_ENTRIES at app startup.

Invoice errors also write the outbound qbXML to logs/ as a named file and
auto-prune to _MAX_XML_FILES files.

All public functions are wrapped in try/except and never raise — a logging
failure cannot affect the calling route.
"""

import json
from datetime import datetime, timezone
from pathlib import Path

_LOGS_DIR   = Path(__file__).parent.parent / "logs"
_MAX_JSONL  = 1000
_MAX_XML    = 20


def ensure_logs_dir() -> None:
    """Create the logs directory. Called at app startup."""
    try:
        _LOGS_DIR.mkdir(parents=True, exist_ok=True)
    except Exception:
        pass


def trim_all_logs() -> None:
    """Trim every jsonlines file to _MAX_JSONL entries. Called at app startup."""
    try:
        for path in _LOGS_DIR.glob("*_history.jsonl"):
            _trim_jsonl(path)
    except Exception:
        pass


def log_event(operation: str, **fields) -> None:
    """Append one JSON line to logs/{operation}_history.jsonl. Never raises."""
    try:
        entry = {"ts": datetime.now(timezone.utc).isoformat(), **fields}
        path  = _LOGS_DIR / f"{operation}_history.jsonl"
        with open(path, "a", encoding="utf-8") as f:
            f.write(json.dumps(entry) + "\n")
    except Exception:
        pass


def log_invoice_error(order_id: str, company: str, request_xml: str, error_msg: str) -> None:
    """
    Write the outbound qbXML for a failed invoice post. Never raises.
    File is named error_{ts}_{order_id}_request.xml so it can be cross-referenced
    with the matching invoice_post_history.jsonl entry by order_id.
    """
    try:
        ts         = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")
        safe_order = str(order_id).replace("/", "-").replace("\\", "-")
        filename   = f"error_{ts}_{safe_order}_request.xml"
        (_LOGS_DIR / filename).write_text(request_xml, encoding="utf-8")
        _prune_xml_files()
    except Exception:
        pass


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------

def _prune_xml_files() -> None:
    """Delete oldest error XML files, keeping only the most recent _MAX_XML."""
    try:
        files  = sorted(_LOGS_DIR.glob("error_*_request.xml"))
        excess = max(0, len(files) - _MAX_XML)
        for f in files[:excess]:
            f.unlink(missing_ok=True)
    except Exception:
        pass


def _trim_jsonl(path: Path) -> None:
    """Trim a jsonlines file to the last _MAX_JSONL lines."""
    try:
        lines = path.read_text(encoding="utf-8").splitlines(keepends=True)
        if len(lines) > _MAX_JSONL:
            path.write_text("".join(lines[-_MAX_JSONL:]), encoding="utf-8")
    except Exception:
        pass
