"""
QuickBooks Desktop COM interface.

Sends qbXML requests directly to the open QB company file via win32com.
This module only works on Windows with QuickBooks Desktop installed.
"""

import xml.etree.ElementTree as ET

try:
    import win32com.client
    import pythoncom
    _COM_AVAILABLE = True
except ImportError:
    _COM_AVAILABLE = False

from .config import QB_COMPANY_FILES
from .qbxml_builder import build_company_query

APP_NAME = "ASI QB Middleware"


def _normalize_path(path: str) -> str:
    """Lowercase and ensure .qbw extension for comparison."""
    p = path.strip().lower()
    if not p.endswith(".qbw"):
        p += ".qbw"
    return p


def _open_session():
    """
    Open a QB COM connection and session. Returns (rp, ticket).
    Caller is responsible for calling _close_session().
    """
    if not _COM_AVAILABLE:
        raise RuntimeError(
            "pywin32 is not installed. Run: pip install pywin32"
        )

    # COM must be initialized on each thread that uses it.
    # Flask handles requests on worker threads where this hasn't been done.
    # MTA (COINIT_MULTITHREADED) is required -- STA needs a Windows message
    # pump to deliver cross-process COM responses, which Flask threads don't have.
    pythoncom.CoInitializeEx(0, pythoncom.COINIT_MULTITHREADED)

    try:
        rp = win32com.client.Dispatch("QBXMLRP2.RequestProcessor")
    except Exception as e:
        pythoncom.CoUninitialize()
        raise RuntimeError(
            "Could not connect to QuickBooks. Make sure QuickBooks Desktop "
            f"is running and a company file is open. ({e})"
        )

    try:
        rp.OpenConnection2("", APP_NAME, 1)
    except Exception as e:
        pythoncom.CoUninitialize()
        raise RuntimeError(f"Could not open QB connection: {e}")

    try:
        # "" = currently open file, 2 = omDontCare (single or multi-user)
        ticket = rp.BeginSession("", 2)
    except Exception as e:
        try:
            rp.CloseConnection()
        except Exception:
            pass
        pythoncom.CoUninitialize()

        raise RuntimeError(
            "Could not begin QB session. Make sure a company file is open "
            f"in QuickBooks. ({e})"
        )

    return rp, ticket


def _close_session(rp, ticket):
    """Close QB session, connection, and uninitialize COM for this thread."""
    try:
        rp.EndSession(ticket)
    except Exception:
        pass
    try:
        rp.CloseConnection()
    except Exception:
        pass
    try:
        pythoncom.CoUninitialize()
    except Exception:
        pass


def _get_company_info(rp, ticket) -> dict:
    """Returns {'name': '...', 'file': '...'} for the currently open company."""
    response = rp.ProcessRequest(ticket, build_company_query())
    root = ET.fromstring(response)

    ret = root.find(".//CompanyRet")
    if ret is None:
        raise RuntimeError("CompanyQueryRq returned no data from QuickBooks.")

    return {
        "name": ret.findtext("CompanyName") or "",
        "file": ret.findtext("CompanyFileName") or "",
    }


def get_open_company() -> dict:
    """
    Returns {'name': '...', 'file': '...'} for the currently open QB company.
    Raises RuntimeError if QB is not running or no file is open.
    """
    rp, ticket = _open_session()
    try:
        return _get_company_info(rp, ticket)
    finally:
        _close_session(rp, ticket)


def submit_invoice(qbxml: str, expected_slug: str) -> str:
    """
    Verify the correct QB company file is open, then submit an InvoiceAdd.
    Returns the QB invoice number (TxnNumber) on success.
    Raises RuntimeError with a user-friendly message on any failure.
    """
    rp, ticket = _open_session()
    try:
        # Verify the correct company file is open before submitting
        info = _get_company_info(rp, ticket)
        expected_file = QB_COMPANY_FILES.get(expected_slug, "")

        if not expected_file:
            raise RuntimeError(
                f"No QB company file configured for '{expected_slug}'. "
                "Check QB_FILE_ACOUSTICAL / QB_FILE_ARCHITECTURAL in .env."
            )

        if _normalize_path(info["file"]) != _normalize_path(expected_file):
            raise RuntimeError(
                f"Wrong QuickBooks company file is open. "
                f"This invoice requires the {expected_slug} company file "
                f"({expected_file}), but '{info['name']}' is currently open. "
                f"Please switch QuickBooks to the correct company file and try again."
            )

        # Submit the invoice
        response = rp.ProcessRequest(ticket, qbxml)

        # Parse the invoice number from the response
        root = ET.fromstring(response)
        rs = root.find(".//InvoiceAddRs")

        if rs is None:
            raise RuntimeError("No InvoiceAddRs in QB response.")

        status_code = rs.get("statusCode", "")
        status_msg  = rs.get("statusMessage", "")

        if status_code != "0":
            raise RuntimeError(
                f"QuickBooks rejected the invoice: {status_msg} (code {status_code})"
            )

        txn_number = rs.findtext(".//TxnNumber")
        if not txn_number:
            raise RuntimeError(
                "Invoice was created in QuickBooks but no invoice number was returned."
            )

        return txn_number

    finally:
        _close_session(rp, ticket)
