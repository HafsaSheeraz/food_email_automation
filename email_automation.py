import pandas as pd
import smtplib
import imaplib
import email
from email.message import EmailMessage
from datetime import datetime
import os
import time
from threading import Thread
from dotenv import load_dotenv

# Load variables from .env
load_dotenv()

# ✉️ Email credentials
EMAIL_ADDRESS = os.getenv("EMAIL_ADDRESS")
EMAIL_PASSWORD = os.getenv("EMAIL_PASSWORD")
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587
IMAP_SERVER = "imap.gmail.com"
EXCEL_FILE = "sample.xlsx"

def load_patients():
    try:
        df = pd.read_excel(EXCEL_FILE)
        df.dropna(how="all", inplace=True)
        print("[INFO] Excel file loaded successfully.")
        return df
    except Exception as e:
        print(f"[ERROR] Failed to load Excel file: {e}")
        return pd.DataFrame()

def save_patients(df, prefix="updated"):
    try:
        base, ext = os.path.splitext(EXCEL_FILE)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        new_file = f"{prefix}_{base}{ext}"
        df.to_excel(new_file, index=False)
        print(f"[INFO] Data saved to {os.path.abspath(new_file)}")
    except Exception as e:
        print(f"[ERROR] Failed to save Excel file: {e}")

def clean_excel(filename):
    df = pd.read_excel(filename)
    df.dropna(how="all", inplace=True)
    df = df.loc[:, ~df.columns.duplicated()]
    df.to_excel(filename, index=False)

def send_emails():
    df = load_patients()
    if df.empty:
        print("[ERROR] No data found in Excel.")
        return

    sent_count = 0
    fail_count = 0
    if "First_Email_Sent" not in df.columns:
        df["First_Email_Sent"] = ""
    df["First_Email_Sent"] = df["First_Email_Sent"].astype(str)

    try:
        print("[INFO] Connecting to SMTP server...")
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT, timeout=10) as smtp:
            smtp.starttls()
            smtp.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
            print("[✅] Logged in to SMTP")

            for index, row in df.iterrows():
                name = row.get("Name", "there")
                to_email = str(row.get("Email", "")).strip()

                if not to_email or "@" not in to_email:
                    print(f"[⚠️] Skipping invalid or missing email for {name}: '{to_email}'")
                    fail_count += 1
                    continue
                if row.get("First_Email_Sent", "") == "Yes":
                    print(f"[SKIP] Already sent to {to_email}")
                    continue
              
                msg = EmailMessage()
                msg["Subject"] = "Important: Food Support Inquiry"
                msg["From"] = EMAIL_ADDRESS
                msg["To"] = to_email
                msg.set_content(f"""Hi {name},

We're reaching out to support your food needs. Please reply "yes" to this message if any of the following statements apply to you:

• In the last 12 months, have you worried whether your food would run out before you had money to buy more?
• Did the food you bought just not last, and you didn't have money to get more?

Your answers will help us connect you with local food support resources.

Best regards,
Baylor Scott Nutrition Team
""")

                try:
                    smtp.send_message(msg)
                    df.at[index, "First_Email_Sent"] = "Yes"
                    print(f"[📤] Sent to {name} <{to_email}>")
                    sent_count += 1
                    time.sleep(1)
                except Exception as e:
                    print(f"[❌] Failed to send to {to_email}: {e}")
                    fail_count += 1

        df.to_excel("sample.xlsx", index=False)
        clean_excel("sample.xlsx")  # 🧹 Clean after update
        print(f"\n[📊] Summary: {sent_count} sent, {fail_count} failed.")
        print("[✅] Email sending completed.")
    except Exception as e:
        print(f"[❌] SMTP connection error: {e}")

def send_followup(to_email, name="there"):
    try:
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT, timeout=10) as smtp:
            smtp.starttls()
            smtp.login(EMAIL_ADDRESS, EMAIL_PASSWORD)

            msg = EmailMessage()
            msg["Subject"] = "Next Steps: Food Resource Support"
            msg["From"] = EMAIL_ADDRESS
            msg["To"] = to_email

            form_link = "https://forms.cloud.microsoft/r/BXwc88PcMw"  # MS Forms link here

            msg.set_content(f"""Hi {name},

Thank you for confirming your food support needs.

To connect you with the right local resources, please fill out this quick form:

👉 {form_link}

It will take less than 2 minutes to complete.

Best regards,  
Baylor Scott Nutrition Team
""")
            
            smtp.send_message(msg)
            print(f"[📨] Follow-up form sent to {to_email}")
    except Exception as e:
        print(f"[❌] Failed to send follow-up form to {to_email}: {e}")

def check_responses():
    df = load_patients()
    if df.empty:
        print("[ERROR] No patient data loaded.")
        return

    df["Email"] = df["Email"].astype(str).str.strip().str.lower()
    if "Food Insecurity" not in df.columns:
        df["Food Insecurity"] = ""
    else:
        df["Food Insecurity"] = df["Food Insecurity"].astype(str)

    if "Form_Link_Sent" not in df.columns:
        df["Form_Link_Sent"] = ""
        df["Form_Link_Sent"] = df["Form_Link_Sent"].astype(str)

    try:
        print("[INFO] Connecting to IMAP...")
        mail = imaplib.IMAP4_SSL(IMAP_SERVER)
        mail.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
        mail.select("inbox")

        print("[INFO] Searching for unseen replies...")
        status, messages = mail.search(None, '(UNSEEN SUBJECT "Important: Food Support Inquiry")')
        if status != "OK" or not messages[0]:
            print("[INFO] No new replies found.")
            mail.logout()
            return

        print(f"[INFO] Found {len(messages[0].split())} new replies.")
        for num in messages[0].split():
            status, data = mail.fetch(num, "(RFC822)")
            if status != "OK":
                print(f"[ERROR] Failed to fetch message {num}")
                continue

            raw_email = data[0][1]
            msg = email.message_from_bytes(raw_email)
            sender = email.utils.parseaddr(msg["From"])[1].strip().lower()
            body = ""

            if msg.is_multipart():
                for part in msg.walk():
                    if part.get_content_type() == "text/plain" and "attachment" not in str(part.get("Content-Disposition")):
                        body = part.get_payload(decode=True).decode(errors="ignore")
                        break
            else:
                body = msg.get_payload(decode=True).decode(errors="ignore")

            clean_response = body.split("\n")[0].strip().lower()
            print(f"[REPLY] {sender} ➜ \"{clean_response}\"")
          
            match = df[df["Email"] == sender]

            if match.empty:
                print(f"[⚠️] Email not found in patient file: {sender}")
                continue

            index = match.index[0]
            if any(word in clean_response for word in ["yes", "ok", "yeah"]):
                # Add this once before assigning string values to the column
                df["Form_Link_Sent"] = df["Form_Link_Sent"].astype(object)

                df.at[index, "Food Insecurity"] = "Yes"
                name = match["Name"].values[0] if "Name" in match.columns else "there"
                print(f"[✅] Marked YES for {sender}")

                if df.at[index, "Form_Link_Sent"] != "Yes":
                    send_followup(sender, name=name)
                    df.at[index, "Form_Link_Sent"] = "Yes"
            else:
                df.at[index, "Food Insecurity"] = "No"
                print(f"[INFO] Marked NO for {sender}")

        df.to_excel("sample.xlsx", index=False)
        mail.logout()
        clean_excel("sample.xlsx")  # 🧹 Clean after response updates
        print("[✅] Response check complete and saved to sample.xlsx.")
    except Exception as e:
        print(f"[❌ IMAP ERROR] {e}")

def merge_form_responses(form_file="form_responses.csv"):

    try:
        print("[INFO] Loading original patient file...")
        patients_df = pd.read_excel("sample.xlsx")
        patients_df["Email"] = patients_df["Email"].astype(str).str.lower().str.strip()

        print(f"[INFO] Loading form responses from {form_file}...")
        if form_file.endswith(".csv"):
            form_df = pd.read_csv(form_file)
        else:
            form_df = pd.read_excel(form_file)

        form_df.rename(columns=lambda x: x.strip().replace('\n', ' ').replace('"', ''), inplace=True)

        # Try to find 'Email2' first, then fallback to 'Email'
        email_column = None
        for col in form_df.columns:
            if col.strip().lower() == "email2":
                email_column = col
                break
        if not email_column:
            for col in form_df.columns:
                if "email" in col.lower():
                    email_column = col
                    break

        form_df[email_column] = form_df[email_column].astype(str).str.lower().str.strip()

        columns_to_copy = {
            "ZIP Code": "ZIP",
            "Preferred Contact Method (Text / Call)": "Preferred Contact Method",
            "Household Size": "Household Size",
            "Dietary Needs": "Dietary Needs",
            "Do you have access to transportation?": "Access to Transportation"
        }

        for form_col in columns_to_copy.keys():
            if form_col not in form_df.columns:
                print(f"[⚠️] Missing column in form: '{form_col}', adding empty column.")
                form_df[form_col] = ""

        for final_col in columns_to_copy.values():
            if final_col not in patients_df.columns:
                print(f"[ℹ️] Adding missing column in patient file: '{final_col}'")
                patients_df[final_col] = ""

        print("[INFO] Merging data...")
        for _, row in form_df.iterrows():
            email = row[email_column]
            if pd.isna(email) or "@" not in email:
                print(f"[⚠️] Skipping invalid email in form: {email}")
                continue

            match_idx = patients_df[patients_df["Email"] == email].index
            if not match_idx.empty:
                for form_col, final_col in columns_to_copy.items():
                    val = str(row.get(form_col, "")).strip()
                    if val:
                        # Ensure column is string type to avoid dtype conflict
                        patients_df[final_col] = patients_df[final_col].astype(str)
                        patients_df.loc[match_idx, final_col] = val
                print(f"[✅] Updated form data for: {email}")
            else:
                print(f"[⚠️] Email not found in main sheet: {email}")

        patients_df.to_excel("sample.xlsx", index=False)
        # ↓ Add cleanup after saving
        clean_excel("sample.xlsx")

        print(f"[🧹] Cleaned up empty rows and duplicate columns in sample.xlsx")
        print(f"[✅] Data merged and saved to sample.xlsx")

    except Exception as e:
        print(f"[❌] Merge failed: {e}")
# 🔁 Optional Scheduler (Auto-check every 10 minutes)
def auto_check_scheduler(interval_minutes=10):
    print(f"[SCHEDULER] Starting auto-check every {interval_minutes} minutes...")
    while True:
        try:
            check_responses()
        except Exception as e:
            print(f"[SCHEDULER ERROR] {e}")
        time.sleep(interval_minutes * 60)
if __name__ == "__main__":
    # send_emails()
    # check_responses()
    merge_form_responses()
