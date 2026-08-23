"""
Generates a second, fictional seed document - an IT & Data Security policy - deliberately
in a different domain from the HR policy. Used to demo that retrieval picks the right
document for a question even when the vector store holds chunks from more than one source.

    python tools/generate_it_security_policy_pdf.py
"""
import os

from reportlab.lib.pagesizes import LETTER
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer

OUTPUT_PATH = os.path.join(
    os.path.dirname(__file__), "..", "src", "main", "resources", "documents", "IT-Security-Policy.pdf"
)

SECTIONS = [
    ("1. Password and Account Security", """
    All employees must use a unique password of at least 12 characters for their Northwind
    corporate account, combining upper and lower case letters, numbers, and a symbol. Passwords
    must be changed immediately if an employee suspects their account has been compromised, and
    must never be shared with another employee, contractor, or third party under any circumstance.
    Multi-factor authentication (MFA) is mandatory for all accounts with access to customer data,
    financial systems, or source code repositories. Accounts left inactive for more than 90 days
    are automatically disabled and must be reactivated through the IT service desk.
    """),
    ("2. Device and Endpoint Management", """
    Company-issued laptops must have full-disk encryption, an approved endpoint security agent,
    and automatic screen lock after 5 minutes of inactivity enabled at all times. Employees may
    not install unapproved software or disable security agents on company devices without written
    authorization from the IT Security team. Lost or stolen devices must be reported to IT Security
    within 24 hours so the device can be remotely locked and wiped. Personal devices used to access
    company email or chat must be enrolled in the mobile device management (MDM) system.
    """),
    ("3. Data Classification and Handling", """
    Company data is classified into three tiers: Public, Internal, and Restricted. Restricted data
    includes customer personal information, payment details, and unreleased financial results, and
    may only be stored in approved company systems - never on personal cloud storage, personal
    email, or removable USB drives. Internal data may be shared with employees and approved
    contractors but not posted publicly. Any suspected exposure of Restricted data must be reported
    immediately to the Data Protection Officer, regardless of how minor it may seem.
    """),
    ("4. Acceptable Use of Company Systems", """
    Company systems, network access, and email are provided for business purposes. Limited
    personal use is permitted provided it does not interfere with work, consume excessive
    bandwidth, or involve illegal, discriminatory, or offensive content. Employees must not use
    company systems to access or distribute pirated software, unlicensed media, or any material
    that violates Northwind's Code of Conduct. IT Security reserves the right to monitor network
    traffic and system logs for security purposes in accordance with applicable law.
    """),
    ("5. Incident Reporting", """
    Any suspected phishing email, malware infection, unauthorized access attempt, or other security
    incident must be reported to security@northwindretail.example within 1 hour of discovery, and
    the affected device disconnected from the network if malware is suspected. Do not attempt to
    investigate or remediate a suspected breach yourself - IT Security will coordinate the
    response. Employees who report a genuine mistake in good faith (such as clicking a phishing
    link) will not face disciplinary action for the report itself.
    """),
    ("6. Remote Access and VPN", """
    Remote access to internal systems requires the company VPN client with MFA enabled; direct
    remote desktop connections to internal servers from outside the corporate network are
    prohibited. Split-tunneling must be disabled while connected to the VPN so that all traffic is
    routed through Northwind's security monitoring. Public Wi-Fi (airports, cafes, hotels) may only
    be used to access company systems through the VPN, never as a direct connection.
    """),
]

styles = getSampleStyleSheet()
title_style = ParagraphStyle("PolicyTitle", parent=styles["Title"], fontSize=20, spaceAfter=6)
subtitle_style = ParagraphStyle("PolicySubtitle", parent=styles["Normal"], fontSize=10, textColor="#555555", spaceAfter=24)
heading_style = ParagraphStyle("PolicyHeading", parent=styles["Heading2"], fontSize=13, spaceBefore=16, spaceAfter=8)
body_style = ParagraphStyle("PolicyBody", parent=styles["BodyText"], fontSize=10.5, leading=15, spaceAfter=6)


def build_pdf():
    os.makedirs(os.path.dirname(OUTPUT_PATH), exist_ok=True)
    doc = SimpleDocTemplate(
        OUTPUT_PATH,
        pagesize=LETTER,
        leftMargin=0.9 * inch,
        rightMargin=0.9 * inch,
        topMargin=0.9 * inch,
        bottomMargin=0.9 * inch,
        title="Northwind Retail Group - IT & Data Security Policy",
    )

    flow = [
        Paragraph("Northwind Retail Group", title_style),
        Paragraph("IT &amp; Data Security Policy &mdash; Effective January 1, 2026 (fictional, for demo purposes only)", subtitle_style),
    ]
    for heading, body in SECTIONS:
        flow.append(Paragraph(heading, heading_style))
        flow.append(Paragraph(" ".join(body.split()), body_style))
        flow.append(Spacer(1, 4))

    doc.build(flow)
    print(f"Wrote {OUTPUT_PATH}")


if __name__ == "__main__":
    build_pdf()
