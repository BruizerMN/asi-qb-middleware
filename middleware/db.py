import sqlite3
import uuid
from datetime import datetime
from contextlib import contextmanager
from middleware.config import DB_PATH

# Job status values
PENDING   = "pending"    # queued, waiting for QBWC to pick it up
SENT      = "sent"       # QBWC has received the qbXML, processing in QB
DONE      = "done"       # QB responded with success; qb_invoice_id is set
ERROR     = "error"      # QB or QBWC reported a failure


def init_db():
    with _conn() as conn:
        conn.execute("""
            CREATE TABLE IF NOT EXISTS jobs (
                id          TEXT PRIMARY KEY,
                company     TEXT NOT NULL,
                order_id    TEXT NOT NULL,
                payload     TEXT NOT NULL,
                status      TEXT NOT NULL DEFAULT 'pending',
                qb_invoice_id TEXT,
                error_msg   TEXT,
                created_at  TEXT NOT NULL,
                updated_at  TEXT NOT NULL
            )
        """)


@contextmanager
def _conn():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    try:
        yield conn
        conn.commit()
    finally:
        conn.close()


def enqueue_job(company: str, order_id: str, payload: str) -> str:
    """Insert a new invoice job and return its job_id."""
    job_id = str(uuid.uuid4())
    now = datetime.utcnow().isoformat()
    with _conn() as conn:
        conn.execute(
            "INSERT INTO jobs (id, company, order_id, payload, status, created_at, updated_at) "
            "VALUES (?, ?, ?, ?, ?, ?, ?)",
            (job_id, company, order_id, payload, PENDING, now, now),
        )
    return job_id


def next_pending_job(company: str):
    """Return the oldest pending job for a company, or None."""
    with _conn() as conn:
        row = conn.execute(
            "SELECT * FROM jobs WHERE company=? AND status=? ORDER BY created_at LIMIT 1",
            (company, PENDING),
        ).fetchone()
    return dict(row) if row else None


def mark_sent(job_id: str):
    _update_job(job_id, status=SENT)


def mark_done(job_id: str, qb_invoice_id: str):
    _update_job(job_id, status=DONE, qb_invoice_id=qb_invoice_id)


def mark_error(job_id: str, error_msg: str):
    _update_job(job_id, status=ERROR, error_msg=error_msg)


def get_job(job_id: str):
    with _conn() as conn:
        row = conn.execute("SELECT * FROM jobs WHERE id=?", (job_id,)).fetchone()
    return dict(row) if row else None


def _update_job(job_id: str, **fields):
    fields["updated_at"] = datetime.utcnow().isoformat()
    set_clause = ", ".join(f"{k}=?" for k in fields)
    with _conn() as conn:
        conn.execute(
            f"UPDATE jobs SET {set_clause} WHERE id=?",
            (*fields.values(), job_id),
        )
