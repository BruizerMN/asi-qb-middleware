"""
Simulates the QuickBooks Web Connector SOAP conversation.
Run this while the middleware is running to test the full flow on Mac.

Usage:
    python tests/simulate_wc.py
"""

import sys
import requests

BASE = "http://localhost:5100"

SOAP_ENVELOPE = """\
<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
  <soap:Body>
    <{method} xmlns="http://developer.intuit.com/qbxml/2.0">
      {body}
    </{method}>
  </soap:Body>
</soap:Envelope>"""


def soap(method, body):
    xml = SOAP_ENVELOPE.format(method=method, body=body)
    r = requests.post(f"{BASE}/wc", data=xml.encode(), headers={"Content-Type": "text/xml"})
    print(f"\n--- {method} ---")
    print(r.text[:800])
    return r.text


def run_simulation(company="acoustical", user="wc_acoustical", password="wc-acoustical-2026"):
    print("=== Simulating QuickBooks Web Connector session ===")
    print(f"Company: {company}")

    import re

    # Step 1: authenticate
    auth_response = soap("authenticate", f"""
        <strUserName>{user}</strUserName>
        <strPassword>{password}</strPassword>
    """)

    ticket_match = re.search(r"<ticket>(.*?)</ticket>", auth_response)
    if not ticket_match or not ticket_match.group(1):
        print("\nAuthentication failed or no pending jobs.")
        return

    ticket = ticket_match.group(1)
    print(f"\nTicket: {ticket}")

    # Step 2: request qbXML
    send_response = soap("sendRequestXML", f"<ticket>{ticket}</ticket>")

    request_match = re.search(r"<request>(.*?)</request>", send_response, re.DOTALL)
    qbxml = request_match.group(1).strip() if request_match else ""

    if not qbxml:
        print("\nNo pending jobs in queue.")
        soap("closeConnection", f"<ticket>{ticket}</ticket>")
        return

    fake_request_id = re.search(r'requestID="([^"]+)"', qbxml)
    req_id = fake_request_id.group(1) if fake_request_id else "1"

    # Step 3: build fake response matching the request type
    if "CompanyQueryRq" in qbxml:
        fake_response = f"""<?xml version="1.0"?><?qbxml version="16.0"?>
<QBXML><QBXMLMsgsRs>
  <CompanyQueryRs requestID="{req_id}" statusCode="0" statusSeverity="Info" statusMessage="Status OK">
    <CompanyRet>
      <CompanyName>Acoustical Surfaces (SIMULATED)</CompanyName>
    </CompanyRet>
  </CompanyQueryRs>
</QBXMLMsgsRs></QBXML>"""
    else:
        fake_response = f"""<?xml version="1.0"?><?qbxml version="16.0"?>
<QBXML><QBXMLMsgsRs>
  <InvoiceAddRs requestID="{req_id}" statusCode="0" statusSeverity="Info" statusMessage="Status OK">
    <InvoiceRet>
      <TxnID>12345-1234567890</TxnID>
      <RefNumber>1042</RefNumber>
    </InvoiceRet>
  </InvoiceAddRs>
</QBXMLMsgsRs></QBXML>"""

    soap("receiveResponseXML", f"""
        <ticket>{ticket}</ticket>
        <response><![CDATA[{fake_response}]]></response>
        <hresult></hresult>
        <message></message>
    """)

    # Step 4: close session
    soap("closeConnection", f"<ticket>{ticket}</ticket>")
    print("\n=== Simulation complete ===")


if __name__ == "__main__":
    company = sys.argv[1] if len(sys.argv) > 1 else "acoustical"
    user_map = {
        "acoustical":    ("wc_acoustical",    "wc-acoustical-2026"),
        "architectural": ("wc_architectural", "wc-architectural-2026"),
    }
    user, password = user_map.get(company, user_map["acoustical"])
    run_simulation(company=company, user=user, password=password)
