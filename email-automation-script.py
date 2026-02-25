import pandas as pd
from datetime import datetime, timedelta
import os
import ollama
import re
from email.message import EmailMessage

EXCEL_FILE = "job_tracker.xlsx"

RESUMES = {
    "1": "resume_flutter.pdf",
    "2": "resume_backend.pdf",
    "3": "resume_general.pdf"
}

# === INPUT ===
company = input("Company: ")
post = input("Role/Post: ")
to_email = input("HR Email: ")
jd = input("\nPaste Job Description:\n")

# === AI ANALYSIS ===
analysis_prompt = f"""
Extract structured info from this JD.

Return EXACTLY in this format:

Post:
Experience:
Location:
Remote:
Package:
Fit Score:

JD:
{jd}
"""

analysis = ollama.chat(
    model="mistral",
    messages=[{"role": "user", "content": analysis_prompt}]
)

analysis_text = analysis['message']['content']

print("\n=== ANALYSIS ===\n", analysis_text)

# === PARSE FIELDS SAFELY ===
def extract(field):
    pattern = rf"{field}:\s*(.*)"
    match = re.search(pattern, analysis_text, re.IGNORECASE)
    return match.group(1).strip() if match else ""

exp_required = extract("Experience")
location = extract("Location")
remote_text = extract("Remote").lower()
remote = remote_text in ["yes", "true", "remote"]
package = extract("Package")
fit_score = extract("Fit Score")

# === DECISION ===
apply_decision = input("\nApply? (y/n): ")
if apply_decision.lower() != "y":
    print("Skipped.")
    exit()

# === RESUME CHOICE ===
print("\nSelect Resume:")
for k, v in RESUMES.items():
    print(f"{k}. {v}")

resume_path = RESUMES.get(input("Choice: "), None)

if not resume_path:
    print("Invalid choice. Exiting.")
    exit()

# === EMAIL GENERATION ===
email_prompt = f"""
Write a concise job application email.

Role: {post}
Company: {company}
JD:
{jd}
"""

email_response = ollama.chat(
    model="mistral",
    messages=[{"role": "user", "content": email_prompt}]
)

email_body = email_response['message']['content']

print("\n=== EMAIL ===\n", email_body)

action = input("\nSend / edit / cancel? (s/e/c): ")

if action == "e":
    email_body = input("Paste edited email:\n")
elif action == "c":
    print("Cancelled.")
    exit()

# === PREVIEW EMAIL (sending disabled) ===
msg = EmailMessage()
msg["Subject"] = f"Application for {post} — {company}"
msg["To"] = to_email
msg.set_content(email_body)

print("\n📨 Email ready (sending disabled in test mode)")

# === LOG TO EXCEL ===
follow_up_date = (datetime.now() + timedelta(days=7)).strftime("%Y-%m-%d")

new_entry = pd.DataFrame([{
    "Post": post,
    "Company": company,
    "Date of Application": datetime.now().strftime("%Y-%m-%d"),
    "Status": "Applied",
    "Applied with": "Mail",
    "Remote": remote,
    "Location": location,
    "Package": package,
    "HR / Contact Name": "",
    "Exp required": exp_required,
    "Fit Score": fit_score,
    "Follow-up Date": follow_up_date
}])

if os.path.exists(EXCEL_FILE):
    df = pd.read_excel(EXCEL_FILE)
    df = pd.concat([df, new_entry], ignore_index=True)
else:
    df = new_entry

df.to_excel(EXCEL_FILE, index=False)

print("📊 Logged to Excel")