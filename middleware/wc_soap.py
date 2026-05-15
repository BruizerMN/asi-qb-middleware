"""
QuickBooks Web Connector SOAP handler.

QBWC communicates via SOAP over HTTP. It calls these methods in order:
  1. authenticate    — validate credentials, return session ticket
  2. sendRequestXML  — we give QBWC the next qbXML to run in QB
  3. receiveResponseXML — QB's response comes back here; return 100 when done
  4. closeConnection — session ends

We parse the incoming SOAP envelope manually (no WSDL library needed)
and return plain SOAP XML strings.
"""

import uuid
import xml.etree.ElementTree as ET

import middleware.db as db
from middleware.config import WC_USER_TO_COMPANY, WC_CREDENTIALS

# In-memory map of active QBWC sessions: ticket → company slug
_sessions: dict[str, str] = {}


def handle_soap(body: str) -> str:
    """Route an incoming SOAP request to the right handler."""
    root = ET.fromstring(body)
    ns = {"s": "http://schemas.xmlsoap.org/soap/envelope/"}

    soap_body = root.find("s:Body", ns)
    if soap_body is None:
        return _fault("Client", "No SOAP Body found")

    action = list(soap_body)[0]
    method = action.tag.split("}")[-1] if "}" in action.tag else action.tag

    handlers = {
        "authenticate":       _authenticate,
        "sendRequestXML":     _send_request_xml,
        "receiveResponseXML": _receive_response_xml,
        "getLastError":       _get_last_error,
        "connectionError":    _connection_error,
        "closeConnection":    _close_connection,
    }

    handler = handlers.get(method)
    if handler is None:
        return _fault("Client", f"Unknown method: {method}")

    return handler(action)


# ---------------------------------------------------------------------------
# SOAP method handlers
# ---------------------------------------------------------------------------

def _authenticate(el: ET.Element) -> str:
    user = _find_text(el, "strUserName")
    password = _find_text(el, "strPassword")

    company = WC_USER_TO_COMPANY.get(user)
    if company and WC_CREDENTIALS[company]["password"] == password:
        ticket = str(uuid.uuid4())
        _sessions[ticket] = company
        # Empty string for company file path — QB uses whatever file is open
        return _response("authenticate", f"<ticket>{ticket}</ticket><hresult></hresult><message></message>")

    return _response("authenticate", "<ticket></ticket><hresult>0x80040400</hresult><message>Invalid credentials</message>")


def _send_request_xml(el: ET.Element) -> str:
    ticket = _find_text(el, "ticket")
    company = _sessions.get(ticket)
    if not company:
        return _response("sendRequestXML", "<request></request>")

    job = db.next_pending_job(company)
    if not job:
        # Empty string tells QBWC there's nothing to do — it will close the session
        return _response("sendRequestXML", "<request></request>")

    db.mark_sent(job["id"])
    # Embed job_id in the iterator so we can match the response back
    qbxml = job["payload"].replace('requestID="1"', f'requestID="{job["id"]}"')
    return _response("sendRequestXML", f"<request>{_cdata(qbxml)}</request>")


def _receive_response_xml(el: ET.Element) -> str:
    ticket = _find_text(el, "ticket")
    response = _find_text(el, "response")
    hresult = _find_text(el, "hresult")
    message = _find_text(el, "message")

    if hresult:
        # QB reported an error at the transport level
        _handle_error_response(response, hresult, message)
        return _response("receiveResponseXML", "<retCode>-1</retCode>")

    _handle_success_response(response)
    # 100 = done with this session
    return _response("receiveResponseXML", "<retCode>100</retCode>")


def _get_last_error(el: ET.Element) -> str:
    return _response("getLastError", "<message></message>")


def _connection_error(el: ET.Element) -> str:
    ticket = _find_text(el, "ticket")
    hresult = _find_text(el, "hresult")
    message = _find_text(el, "message")
    _sessions.pop(ticket, None)
    return _response("connectionError", "<result>done</result>")


def _close_connection(el: ET.Element) -> str:
    ticket = _find_text(el, "ticket")
    _sessions.pop(ticket, None)
    return _response("closeConnection", "<result>OK</result>")


# ---------------------------------------------------------------------------
# Response parsing helpers
# ---------------------------------------------------------------------------

def _handle_success_response(response_xml: str):
    """Dispatch QB response to the appropriate handler by response type."""
    if not response_xml:
        return
    try:
        root = ET.fromstring(response_xml)
        rs = root.find(".//{*}InvoiceAddRs")
        if rs is not None:
            _handle_invoice_add_rs(rs)
            return
        rs = root.find(".//{*}CompanyQueryRs")
        if rs is not None:
            _handle_company_query_rs(rs)
            return
    except ET.ParseError:
        pass


def _handle_invoice_add_rs(rs: ET.Element):
    job_id = rs.get("requestID")
    status = rs.get("statusCode", "")
    if status == "0":
        txn_id = _find_text(rs, "InvoiceRet/TxnID") or ""
        ref_number = _find_text(rs, "InvoiceRet/RefNumber") or ""
        db.mark_done(job_id, qb_invoice_id=ref_number or txn_id)
    else:
        msg = rs.get("statusMessage", "Unknown QB error")
        db.mark_error(job_id, error_msg=f"statusCode={status}: {msg}")


def _handle_company_query_rs(rs: ET.Element):
    job_id = rs.get("requestID")
    status = rs.get("statusCode", "")
    if status == "0":
        company_name = _find_text(rs, "CompanyRet/CompanyName") or "unknown"
        db.mark_done(job_id, qb_invoice_id=company_name)
    else:
        msg = rs.get("statusMessage", "Unknown QB error")
        db.mark_error(job_id, error_msg=f"statusCode={status}: {msg}")


def _handle_error_response(response_xml: str, hresult: str, message: str):
    """Mark the in-flight sent job as errored when QBWC reports a problem."""
    job_id = None
    if response_xml:
        try:
            root = ET.fromstring(response_xml)
            for el in root.iter():
                req_id = el.get("requestID")
                if req_id:
                    job_id = req_id
                    break
        except ET.ParseError:
            pass
    if job_id:
        db.mark_error(job_id, error_msg=f"QBWC hresult={hresult}: {message}")


# ---------------------------------------------------------------------------
# XML / SOAP helpers
# ---------------------------------------------------------------------------

def _find_text(el: ET.Element, tag: str) -> str:
    # Use local-name matching to handle QBWC's xmlns namespace on elements
    for child in el.iter():
        local = child.tag.split("}")[-1] if "}" in child.tag else child.tag
        if local == tag:
            return (child.text or "").strip()
    return ""


def _cdata(content: str) -> str:
    return f"<![CDATA[{content}]]>"


def _response(method: str, inner: str) -> str:
    return (
        '<?xml version="1.0" encoding="utf-8"?>'
        '<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">'
        "<soap:Body>"
        f'<{method}Response xmlns="http://developer.intuit.com/qbxml/2.0">'
        f"<{method}Result>{inner}</{method}Result>"
        f"</{method}Response>"
        "</soap:Body>"
        "</soap:Envelope>"
    )


def _fault(code: str, message: str) -> str:
    return (
        '<?xml version="1.0" encoding="utf-8"?>'
        '<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">'
        "<soap:Body>"
        "<soap:Fault>"
        f"<faultcode>{code}</faultcode>"
        f"<faultstring>{message}</faultstring>"
        "</soap:Fault>"
        "</soap:Body>"
        "</soap:Envelope>"
    )
