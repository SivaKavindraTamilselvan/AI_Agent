"""
============================================================
  Requirements Analyzer  —  Groq + SMTP Email Sender
  ─────────────────────────────────────────────────────────
  Reads a plain-text requirements file, sends it to Groq
  (LLaMA 3 model) for structured analysis, then emails the
  formatted report to the client via Gmail SMTP.
============================================================
"""

import os
import re
import smtplib
import datetime
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from groq import Groq
from docx import Document


# ─────────────────────────────────────────────────────────
#  CONFIGURATION  — fill these before running
# ─────────────────────────────────────────────────────────

# Path to your plain-text requirements file
REQUIREMENTS_FILE = "requirment.docx"

# Groq API key — get yours free at https://console.groq.com
GROQ_API_KEY = "API_KEY"

# Groq model — llama-3.3-70b-versatile is fast and accurate
GROQ_MODEL = "llama-3.3-70b-versatile"

# Gmail credentials for the SENDER account
SENDER_EMAIL    = "MYEMAIL@gmail.com"
SENDER_PASSWORD = "EMAILPASSWORD"   # Gmail App Password, NOT your real password

# ── Replace with the actual client email ──
CLIENT_EMAIL = "CLIENTEMAIL@gmail.com"

# Optional: project / company name shown in the email subject
PROJECT_NAME = "Fashion E-Commerce Platform"


# ─────────────────────────────────────────────────────────
#  STEP 1 — READ THE REQUIREMENTS FILE
# ─────────────────────────────────────────────────────────

def read_requirements(filepath: str) -> str:
    """
    Opens the text file and returns its content as a string.
    Raises a clear error if the file is missing.
    """
    if not os.path.exists(filepath):
        raise FileNotFoundError(
            f"Requirements file not found: '{filepath}'\n"
            "Make sure the file is in the same folder as this script."
        )
    doc = Document(filepath)

    content = "\n".join(
        paragraph.text for paragraph in doc.paragraphs
        if paragraph.text.strip()
    )

    if not content.strip():
        raise ValueError(
            f"The file '{filepath}' is empty."
        )

    print(f"[✓] Requirements file loaded  ({len(content)} characters)")
    return content


# ─────────────────────────────────────────────────────────
#  STEP 2 — BUILD THE GROQ PROMPT
# ─────────────────────────────────────────────────────────

SYSTEM_PROMPT = """
You are a senior business analyst and software architect.
Your task is to analyse raw client requirements and produce a
structured, professional Software Requirements Analysis report.

Always respond in clean plain text using the EXACT section
headings listed below. Use numbered lists inside each section.
Do NOT use markdown symbols like **, ##, or *.

OUTPUT FORMAT (use these exact headings, nothing else):
================================================================
PROJECT OVERVIEW
================================================================
[2–3 sentence summary of what the system does and who uses it]

================================================================
FUNCTIONAL REQUIREMENTS
================================================================
[Numbered list — each item is a specific, testable feature]

================================================================
NON-FUNCTIONAL REQUIREMENTS
================================================================
[Numbered list — performance, security, scalability, usability,
 availability, compliance etc.]

================================================================
RISKS & MITIGATION
================================================================
[Numbered list — risk description | likelihood | impact | mitigation]

================================================================
ASSUMPTIONS
================================================================
[Numbered list — things assumed to be true for this project]

================================================================
QUESTIONS FOR CLIENT (Clarifications Needed)
================================================================
[Numbered list — specific questions that must be answered before
 development can begin, grouped by topic if possible]

================================================================
END OF REPORT
================================================================
"""

def build_user_message(raw_requirements: str) -> str:
    """Wraps the raw requirements in a clear instruction for the model."""
    return (
        "Below are the raw client requirements. Analyse them and produce "
        "the full structured report as instructed.\n\n"
        "--- RAW REQUIREMENTS START ---\n"
        f"{raw_requirements}\n"
        "--- RAW REQUIREMENTS END ---"
    )


# ─────────────────────────────────────────────────────────
#  STEP 3 — CALL GROQ API
# ─────────────────────────────────────────────────────────

def analyse_with_groq(raw_requirements: str) -> str:
    """
    Sends the requirements to Groq and returns the
    structured analysis as a plain-text string.
    """
    print("[→] Connecting to Groq API …")

    client = Groq(api_key=GROQ_API_KEY)

    response = client.chat.completions.create(
        model=GROQ_MODEL,
        messages=[
            {"role": "system",  "content": SYSTEM_PROMPT},
            {"role": "user",    "content": build_user_message(raw_requirements)},
        ],
        temperature=0.3,       # Lower = more consistent / professional output
        max_tokens=4096,        # Enough for a detailed report
    )

    analysis = response.choices[0].message.content.strip()
    print(f"[✓] Groq analysis received  ({len(analysis)} characters)")
    return analysis


# ─────────────────────────────────────────────────────────
#  STEP 4 — FORMAT THE EMAIL BODY
# ─────────────────────────────────────────────────────────

def format_email_body(analysis: str, project_name: str) -> str:
    today = datetime.date.today().strftime("%B %d, %Y")

    header = f"""
Dear Client,

Thank you for sharing your project requirements with us.
We have completed an initial analysis of your brief for the
"{project_name}" project.

Please find the structured requirements report below.
Kindly review it — especially the "QUESTIONS FOR CLIENT"
section — and provide your answers so we can proceed with
the full specification and development plan.

Date of Analysis : {today}
Project          : {project_name}

{'=' * 64}

"""

    footer = """
================================================================
Generated by Requirements Analyzer

If you have any questions or would like to discuss any section
in detail, please reply to this email or schedule a call.

Best regards,
Development Team
"""

    return header + analysis + footer
# ─────────────────────────────────────────────────────────
#  STEP 5 — SEND EMAIL VIA GMAIL SMTP
# ─────────────────────────────────────────────────────────

def send_email(
    body_text: str,
    project_name: str,
    sender_email: str,
    sender_password: str,
    recipient_email: str,
) -> None:
    """
    Sends a plain-text email using Gmail SMTP with TLS.

    Gmail requires an App Password, NOT your normal login password.
    Generate one at: Google Account → Security → 2-Step Verification
    → App Passwords → Select app: Mail → Generate.
    """
    print(f"[→] Preparing email to {recipient_email} …")

    subject = f"Requirements Analysis Report — {project_name}"

    # Build the MIME message
    message = MIMEMultipart("alternative")
    message["Subject"] = subject
    message["From"]    = sender_email
    message["To"]      = recipient_email

    # Plain-text part
    plain_part = MIMEText(body_text, "plain", "utf-8")
    message.attach(plain_part)

    # HTML part — makes it look polished in email clients
    html_body = plain_to_html(body_text, project_name)
    html_part = MIMEText(html_body, "html", "utf-8")
    message.attach(html_part)   # HTML is preferred by email clients when present

    # Send via Gmail SMTP (port 587, STARTTLS)
    print("[→] Connecting to Gmail SMTP …")
    with smtplib.SMTP("smtp.gmail.com", 587) as server:
        server.ehlo()
        server.starttls()                               # Encrypt the connection
        server.login(sender_email, sender_password)
        server.sendmail(sender_email, recipient_email, message.as_string())

    print(f"[✓] Email sent successfully to {recipient_email}")


def plain_to_html(plain_text: str, project_name: str) -> str:
    """
    Converts the plain-text report into a clean, styled HTML email.
    Section headings (lines of '=') are turned into visual dividers.
    """
    lines = plain_text.split("\n")
    html_lines = []

    html_lines.append(f"""
<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<style>
  body       {{ font-family: Arial, sans-serif; font-size: 14px;
                color: #2c2c2c; background: #f7f8fa; margin: 0; padding: 0; }}
  .wrapper   {{ max-width: 760px; margin: 30px auto; background: #ffffff;
                border-radius: 8px; overflow: hidden;
                box-shadow: 0 2px 12px rgba(0,0,0,0.08); }}
  .top-bar   {{ background: #1A3C5E; color: #ffffff; padding: 28px 36px; }}
  .top-bar h1{{ margin: 0; font-size: 20px; font-weight: 700; }}
  .top-bar p {{ margin: 4px 0 0; font-size: 13px; opacity: 0.75; }}
  .content   {{ padding: 32px 36px; }}
  .section-title {{
    background: #1A3C5E; color: #ffffff;
    padding: 8px 16px; border-radius: 4px;
    font-size: 13px; font-weight: 700;
    letter-spacing: 0.08em; margin: 28px 0 12px;
    text-transform: uppercase;
  }}
  .divider   {{ border: none; border-top: 2px solid #e0e6ef; margin: 20px 0; }}
  p          {{ margin: 6px 0; line-height: 1.7; }}
  .footer    {{ background: #f0f4f8; padding: 20px 36px;
                font-size: 12px; color: #7a8899; border-top: 1px solid #dde3eb; }}
</style>
</head>
<body>
<div class="wrapper">
  <div class="top-bar">
    <h1>Requirements Analysis Report</h1>
    <p>{project_name}</p>
  </div>
  <div class="content">
""")

    skip_next = False
    for i, line in enumerate(lines):
        stripped = line.strip()

        if skip_next:
            skip_next = False
            continue

        # Detect section headings surrounded by '===' lines
        if set(stripped) == {"="} and len(stripped) >= 10:
            # Check if next line is the actual heading text
            if i + 1 < len(lines) and lines[i + 1].strip() and set(lines[i + 1].strip()) != {"="}:
                heading_text = lines[i + 1].strip()
                if heading_text != "END OF REPORT":
                    html_lines.append(f'<div class="section-title">{heading_text}</div>')
                else:
                    html_lines.append('<hr class="divider">')
                skip_next = True   # Skip the heading text line (already consumed)
            continue

        if stripped == "":
            html_lines.append("<br>")
        elif re.match(r"^\d+\.", stripped):
            html_lines.append(f"<p>&nbsp;&nbsp;{stripped}</p>")
        else:
            html_lines.append(f"<p>{stripped}</p>")

    html_lines.append("""
  </div>
  <div class="footer">
    This report was generated automatically based on the requirements brief provided.
    Please reply with any clarifications or corrections.
  </div>
</div>
</body>
</html>
""")

    return "\n".join(html_lines)


# ─────────────────────────────────────────────────────────
#  STEP 6 — SAVE REPORT LOCALLY
# ─────────────────────────────────────────────────────────

def save_report(analysis: str, project_name: str) -> str:
    """
    Saves the analysis as a local .txt file for your records.
    Returns the filename.
    """
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    safe_name = project_name.replace(" ", "_").lower()
    filename  = f"report_{safe_name}_{timestamp}.txt"

    with open(filename, "w", encoding="utf-8") as f:
        f.write(f"Requirements Analysis Report\n")
        f.write(f"Project : {project_name}\n")
        f.write(f"Generated : {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
        f.write("=" * 64 + "\n\n")
        f.write(analysis)

    print(f"[✓] Report saved locally as  '{filename}'")
    return filename


# ─────────────────────────────────────────────────────────
#  MAIN — orchestrates all steps
# ─────────────────────────────────────────────────────────

def main():
    print("\n" + "=" * 56)
    print("  Requirements Analyzer  —  Groq + Gmail")
    print("=" * 56 + "\n")

    # ── Step 1: Read the requirements file ──
    raw_requirements = read_requirements(REQUIREMENTS_FILE)

    # ── Step 2 & 3: Analyse with Groq ──
    analysis = analyse_with_groq(raw_requirements)

    # ── Step 4: Save locally ──
    save_report(analysis, PROJECT_NAME)

    # ── Step 5: Format and send email ──
    email_body = format_email_body(analysis, PROJECT_NAME)
    send_email(
        body_text       = email_body,
        project_name    = PROJECT_NAME,
        sender_email    = SENDER_EMAIL,
        sender_password = SENDER_PASSWORD,
        recipient_email = CLIENT_EMAIL,
    )

    print("\n" + "=" * 56)
    print("  All done! Check your sent mail and the local report file.")
    print("=" * 56 + "\n")


if __name__ == "__main__":
    main()
