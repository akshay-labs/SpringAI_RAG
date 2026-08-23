"""
Generates the fictional HR Company Policy PDF used as the seed document for the
RAG demo. Re-run this any time you want to change the policy text:

    python tools/generate_hr_policy_pdf.py

Requires: pip install reportlab
"""
import os

from reportlab.lib.pagesizes import LETTER
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer

OUTPUT_PATH = os.path.join(
    os.path.dirname(__file__), "..", "src", "main", "resources", "documents", "HR-Company-Policy.pdf"
)

SECTIONS = [
    ("1. Working Hours and Attendance", """
    Standard working hours at Northwind Retail Group are 9:00 AM to 5:30 PM, Monday through Friday,
    with a 30-minute unpaid lunch break. Employees are expected to log their attendance through the
    company time-tracking portal by 9:15 AM each working day. Flexible start times between 8:00 AM and
    10:00 AM are permitted with prior manager approval, provided the employee completes a full 8-hour
    workday. Persistent unexplained lateness of more than 15 minutes on three or more occasions within
    a rolling 30-day period will trigger an informal conversation with the employee's manager, followed
    by a formal attendance improvement plan if the pattern continues.
    """),
    ("2. Annual Leave and Time Off", """
    Full-time employees accrue 18 days of paid annual leave per calendar year, credited at a rate of
    1.5 days per completed month of service. Unused annual leave may be carried over into the next
    calendar year up to a maximum of 5 days; any balance beyond that is forfeited on December 31st.
    Leave requests must be submitted through the HR portal at least 5 working days in advance for
    requests of 3 days or fewer, and at least 15 working days in advance for requests longer than 3
    days. In addition to annual leave, employees are entitled to 10 paid sick days per year and 5 days
    of paid bereavement leave for the death of an immediate family member. New parents are eligible for
    16 weeks of paid parental leave, which may be taken continuously or split into two blocks within the
    first 12 months following the birth or adoption of a child.
    """),
    ("3. Code of Conduct", """
    All employees are expected to treat colleagues, customers, and vendors with respect and
    professionalism at all times. Harassment, discrimination, or retaliation of any kind is strictly
    prohibited and will result in disciplinary action up to and including immediate termination.
    Employees must avoid conflicts of interest and disclose to their manager any outside business
    activity, financial interest, or personal relationship that could reasonably be seen to influence
    their judgment at work. Confidential company information, including customer data, unreleased
    product plans, and internal financial figures, must never be shared outside the company without
    written authorization from a director-level employee or above.
    """),
    ("4. Remote and Hybrid Work", """
    Employees in eligible roles may work remotely up to 3 days per week, with a minimum of 2 days per
    week spent at their assigned office location to support in-person collaboration. Remote work
    arrangements must be agreed upon with the employee's direct manager and documented in the HR
    system. Employees working remotely are expected to be reachable during core collaboration hours of
    10:00 AM to 4:00 PM local time and must join scheduled meetings with video enabled unless a
    connectivity issue prevents it. The company provides a one-time home-office stipend of $500 for
    new hybrid employees to cover the cost of a monitor, chair, or similar equipment.
    """),
    ("5. Expense Reimbursement", """
    Business-related expenses such as client meals, travel, and approved training courses are
    reimbursable provided a valid receipt is submitted through the expense portal within 30 days of the
    purchase. Meal expenses while traveling on company business are capped at $60 per day. Airfare must
    be booked in economy class for flights under 6 hours; business class is permitted for flights over
    6 hours with prior manager approval. Reimbursement requests submitted more than 60 days after the
    expense date will not be honored except in cases of documented extenuating circumstances approved
    by HR.
    """),
    ("6. Termination and Resignation", """
    Employees wishing to resign are asked to provide a minimum of 2 weeks' written notice for
    individual-contributor roles and 4 weeks' written notice for management roles, submitted to their
    manager and HR. During the notice period, the company may, at its discretion, ask the employee to
    continue working, transition to garden leave, or a combination of both, while continuing to pay the
    employee's regular salary through the end of the notice period. In cases of termination for cause,
    including gross misconduct, employment may end immediately without notice or pay in lieu of notice.
    For all other involuntary terminations, the company provides severance pay of one week's salary per
    completed year of service, with a minimum of 2 weeks and a maximum of 12 weeks.
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
        title="Northwind Retail Group - HR Company Policy",
    )

    flow = [
        Paragraph("Northwind Retail Group", title_style),
        Paragraph("HR Company Policy Handbook &mdash; Effective January 1, 2026 (fictional, for demo purposes only)", subtitle_style),
    ]
    for heading, body in SECTIONS:
        flow.append(Paragraph(heading, heading_style))
        flow.append(Paragraph(" ".join(body.split()), body_style))
        flow.append(Spacer(1, 4))

    doc.build(flow)
    print(f"Wrote {OUTPUT_PATH}")


if __name__ == "__main__":
    build_pdf()
