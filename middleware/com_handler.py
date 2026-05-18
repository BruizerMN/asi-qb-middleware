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


def _expected_company_name(configured_path: str) -> str:
    """
    Derive the expected QB company name from the configured file path.
    QB returns the company name, not the file path, in CompanyQueryRq.
    We extract it from the filename: 'Q:\\Test Parent 2.23.QBW' -> 'test parent 2.23'
    Comparison is case-insensitive.
    """
    import os
    basename = os.path.basename(configured_path.strip())
    # Strip .qbw extension (case-insensitive)
    if basename.lower().endswith(".qbw"):
        basename = basename[:-4]
    return basename.lower()


def _open_session():
    """
    Open a QB COM connection and session. Returns (rp, ticket).
    Caller is responsible for calling _close_session().
    """
    if not _COM_AVAILABLE:
        raise RuntimeError(
            "pywin32 is not installed. Run: pip install pywin32"
        )

    try:
        rp = win32com.client.Dispatch("QBXMLRP2.RequestProcessor")
    except Exception:
        raise RuntimeError(
            "QuickBooks is not running. Please open QuickBooks and a company file, then try again."
        )

    try:
        rp.OpenConnection2("", APP_NAME, 1)
    except Exception as e:
        raise RuntimeError(
            f"QuickBooks refused the connection. Make sure QuickBooks is open and try again. ({e})"
        )

    try:
        # "" = currently open file, 2 = omDontCare (single or multi-user)
        ticket = rp.BeginSession("", 2)
    except Exception:
        try:
            rp.CloseConnection()
        except Exception:
            pass
        raise RuntimeError(
            "No QuickBooks company file is open. "
            "Please open the correct company file and try again."
        )

    return rp, ticket


def _close_session(rp, ticket):
    """Close QB session and connection."""
    try:
        rp.EndSession(ticket)
    except Exception:
        pass
    try:
        rp.CloseConnection()
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


def verify_company(info: dict, expected_slug: str):
    """
    Raise RuntimeError if the open company doesn't match expected_slug.
    Pass the dict returned by get_open_company() or _get_company_info().
    """
    expected_file = QB_COMPANY_FILES.get(expected_slug, "")

    if not expected_file:
        raise RuntimeError(
            f"No QB company file configured for '{expected_slug}'. "
            "Check QB_FILE_ACOUSTICAL / QB_FILE_ARCHITECTURAL in .env."
        )

    expected_name = _expected_company_name(expected_file)
    actual_name   = info["name"].strip().lower()

    if actual_name != expected_name:
        raise RuntimeError(
            f"Wrong QuickBooks company file is open. "
            f"This invoice requires '{expected_file}', "
            f"but '{info['name']}' is currently open. "
            f"Please switch QuickBooks to the correct company file and try again."
        )


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
        # Verify the correct company file is open before submitting.
        info = _get_company_info(rp, ticket)
        verify_company(info, expected_slug)

        # Log the qbxml being sent — before ProcessRequest so we capture it even on error
        import os
        log_path = os.path.normpath(os.path.join(os.path.dirname(__file__), "..", "qb_last_request.xml"))
        with open(log_path, "w", encoding="utf-8") as f:
            f.write(qbxml)

        # Submit the invoice
        response = rp.ProcessRequest(ticket, qbxml)

        # Log raw QB response
        log_path2 = os.path.normpath(os.path.join(os.path.dirname(__file__), "..", "qb_last_response.xml"))
        with open(log_path2, "w", encoding="utf-8") as f:
            f.write(response)

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
        ref_number = rs.findtext(".//RefNumber")
        txn_id     = rs.findtext(".//TxnID")

        # Return whichever identifier QB provided — prefer RefNumber (user-visible)
        invoice_number = ref_number or txn_number or txn_id
        if not invoice_number:
            raise RuntimeError(
                f"Invoice was created in QuickBooks but no invoice number was returned. "
                f"Response: {ET.tostring(rs, encoding='unicode')}"
            )

        return invoice_number

    finally:
        _close_session(rp, ticket)
