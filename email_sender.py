import asyncio
import smtplib
import os
import time
import json
import random
import pandas as pd
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from anthropic import Anthropic
from datetime import datetime


# ─────────────────────────────────────────────────────────────
# EMAIL GENERATOR — uses Claude to write personalized emails
# ─────────────────────────────────────────────────────────────
class EmailGenerator:
    def __init__(self, anthropic_api_key: str):
        self.client = Anthropic(api_key=anthropic_api_key)

    def generate(
        self,
        lead: dict,
        your_name: str,
        your_company: str,
        your_service: str,
        your_website: str = "",
    ) -> dict:
        """Generate a personalized cold email for one lead."""

        prompt = f"""
You are an expert cold email copywriter. Write a short, highly personalized 
cold email for the following lead. 

RULES:
- Subject line: max 8 words, curiosity-driven, no clickbait
- Email body: max 120 words
- Sound human, not robotic or salesy
- Reference something specific about their business
- ONE clear call to action at the end (reply, book call, or visit website)
- No "I hope this email finds you well"
- No generic openers
- Sign off naturally

LEAD INFO:
Business Name : {lead.get('business_name', 'N/A')}
Category      : {lead.get('category', 'N/A')}
Location      : {lead.get('address', 'N/A')}
Rating        : {lead.get('rating', 'N/A')}
Website       : {lead.get('website', 'N/A')}
ML Score      : {lead.get('ml_score', 'N/A')} / 100
ML Grade      : {lead.get('ml_grade', 'N/A')}

SENDER INFO:
Your Name     : {your_name}
Your Company  : {your_company}
Your Service  : {your_service}
Your Website  : {your_website}

Respond ONLY as a JSON object with exactly these keys:
{{
  "subject": "...",
  "body": "..."
}}
"""
        try:
            response = self.client.messages.create(
                model="claude-haiku-4-5-20251001",
                max_tokens=512,
                messages=[{"role": "user", "content": prompt}]
            )
            text = response.content[0].text.strip()

            # Strip markdown code blocks if present
            text = text.replace("```json", "").replace("```", "").strip()
            parsed = json.loads(text)

            return {
                "subject": parsed.get("subject", f"Quick question about {lead.get('business_name','your business')}"),
                "body":    parsed.get("body", ""),
            }

        except Exception as e:
            print(f"  → Claude generation failed: {e}. Using template.")
            return self._fallback_template(lead, your_name, your_company, your_service)

    def _fallback_template(self, lead, your_name, your_company, your_service):
        """Backup template if Claude API fails."""
        biz  = lead.get("business_name", "your business")
        city = lead.get("address", "").split(",")[-1].strip() or "your area"
        return {
            "subject": f"Quick idea for {biz}",
            "body": (
                f"Hi,\n\n"
                f"I came across {biz} and was genuinely impressed — "
                f"businesses like yours in {city} are exactly who we love working with.\n\n"
                f"At {your_company}, we help {your_service}. "
                f"I think there's a real opportunity here worth a 10-minute conversation.\n\n"
                f"Would you be open to a quick call this week?\n\n"
                f"Best,\n{your_name}"
            ),
        }


# ─────────────────────────────────────────────────────────────
# GMAIL SMTP SENDER
# Uses your Gmail to send emails (free, no API key needed)
# ─────────────────────────────────────────────────────────────
class GmailSender:
    def __init__(self, gmail_address: str, gmail_app_password: str):
        """
        gmail_app_password: NOT your Gmail password.
        Generate at: myaccount.google.com → Security → App Passwords
        """
        self.gmail   = gmail_address
        self.password = gmail_app_password

    def send(self, to_email: str, subject: str, body: str) -> bool:
        """Send one email. Returns True if successful."""
        try:
            msg = MIMEMultipart("alternative")
            msg["Subject"] = subject
            msg["From"]    = self.gmail
            msg["To"]      = to_email

            # Plain text version
            msg.attach(MIMEText(body, "plain"))

            # HTML version (simple formatting)
            html_body = body.replace("\n", "<br>")
            html = f"""
            <html><body style="font-family: Arial, sans-serif; font-size: 14px; 
                               color: #333; max-width: 600px;">
                {html_body}
            </body></html>
            """
            msg.attach(MIMEText(html, "html"))

            with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
                server.login(self.gmail, self.password)
                server.sendmail(self.gmail, to_email, msg.as_string())

            return True

        except smtplib.SMTPRecipientsRefused:
            print(f"    → Invalid email address: {to_email}")
            return False
        except smtplib.SMTPAuthenticationError:
            print("    → Gmail auth failed. Check your App Password.")
            return False
        except Exception as e:
            print(f"    → Send failed: {e}")
            return False


# ─────────────────────────────────────────────────────────────
# COLD EMAIL CAMPAIGN RUNNER
# ─────────────────────────────────────────────────────────────
class ColdEmailCampaign:

    def __init__(
        self,
        anthropic_api_key: str,
        gmail_address: str,
        gmail_app_password: str,
        your_name: str,
        your_company: str,
        your_service: str,
        your_website: str = "",
    ):
        self.generator = EmailGenerator(anthropic_api_key)
        self.sender    = GmailSender(gmail_address, gmail_app_password)
        self.your_name    = your_name
        self.your_company = your_company
        self.your_service = your_service
        self.your_website = your_website

        # Log file to track sent emails
        self.log_path = os.path.join(
            os.path.dirname(os.path.abspath(__file__)),
            "emails_sent_log.csv"
        )

    def _already_emailed(self, email: str) -> bool:
        """Check if we already sent to this address (prevent duplicates)."""
        if not os.path.exists(self.log_path):
            return False
        log = pd.read_csv(self.log_path)
        return email in log["to_email"].values

    def _log_sent(self, lead: dict, subject: str, status: str):
        """Save sent email to log."""
        entry = {
            "timestamp":     datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "to_email":      lead.get("email", ""),
            "business_name": lead.get("business_name", ""),
            "subject":       subject,
            "ml_score":      lead.get("ml_score", ""),
            "ml_grade":      lead.get("ml_grade", ""),
            "status":        status,
        }
        df_new = pd.DataFrame([entry])

        if os.path.exists(self.log_path):
            df_existing = pd.read_csv(self.log_path)
            df_combined = pd.concat([df_existing, df_new], ignore_index=True)
        else:
            df_combined = df_new

        df_combined.to_csv(self.log_path, index=False)

    def run(
        self,
        scored_df: pd.DataFrame,
        min_grade: str = "B",
        dry_run: bool = True,
        delay_seconds: tuple = (45, 90),
        daily_limit: int = 30,
    ):
        """
        Run the cold email campaign.

        scored_df     : DataFrame output from LeadScoringModel.score()
        min_grade     : Only email leads with this grade or better
                        "A" = top leads only
                        "B" = good leads (recommended)
                        "C" = all reasonable leads
        dry_run       : True  = preview emails, DON'T actually send
                        False = actually send emails
        delay_seconds : Random delay between emails (min, max) in seconds
                        Keeps you out of spam folders
        daily_limit   : Max emails to send per run (Gmail allows 500/day)
        """

        # Filter by grade
        GRADE_ORDER = {"A": 0, "B": 1, "C": 2, "D": 3}
        min_val     = GRADE_ORDER.get(min_grade, 1)
        eligible    = scored_df[
            scored_df["ml_grade"].apply(lambda g: GRADE_ORDER.get(g, 99)) <= min_val
        ].copy()

        # Only leads with verified emails
        has_email = eligible[eligible["email"] != "N/A"]

        print(f"\n{'='*55}")
        print(f"  COLD EMAIL CAMPAIGN")
        print(f"{'='*55}")
        print(f"  Total scored leads  : {len(scored_df)}")
        print(f"  Grade {min_grade}+ leads      : {len(eligible)}")
        print(f"  Have email address  : {len(has_email)}")
        print(f"  Mode                : {'🔍 DRY RUN (preview only)' if dry_run else '🚀 LIVE SEND'}")
        print(f"  Daily limit         : {daily_limit}")
        print(f"{'='*55}\n")

        if len(has_email) == 0:
            print("  ✗ No leads with email addresses found.")
            print("  → Add your Hunter.io API key to enrich leads with emails first.")
            return

        sent_count   = 0
        skip_count   = 0
        failed_count = 0

        for _, lead in has_email.iterrows():
            if sent_count >= daily_limit:
                print(f"\n  → Daily limit of {daily_limit} reached. Run again tomorrow.")
                break

            email = lead["email"]
            biz   = lead["business_name"]

            # Skip if already emailed
            if self._already_emailed(email):
                print(f"  ⏭  Already emailed: {biz} ({email})")
                skip_count += 1
                continue

            print(f"\n  [{sent_count+1}] {biz}")
            print(f"      Email : {email}")
            print(f"      Score : {lead['ml_score']} | Grade: {lead['ml_grade']}")

            # Generate personalized email via Claude
            print(f"      Generating email via Claude...")
            email_content = self.generator.generate(
                lead          = lead.to_dict(),
                your_name     = self.your_name,
                your_company  = self.your_company,
                your_service  = self.your_service,
                your_website  = self.your_website,
            )

            print(f"      Subject : {email_content['subject']}")
            print(f"      Preview : {email_content['body'][:80]}...")

            if dry_run:
                print(f"      Status  : 🔍 DRY RUN — not sent")
                self._log_sent(lead.to_dict(), email_content["subject"], "dry_run")
                sent_count += 1

            else:
                success = self.sender.send(
                    to_email = email,
                    subject  = email_content["subject"],
                    body     = email_content["body"],
                )

                if success:
                    print(f"      Status  : ✅ Sent!")
                    self._log_sent(lead.to_dict(), email_content["subject"], "sent")
                    sent_count += 1

                    # Random delay between emails — critical to avoid spam filters
                    wait = random.uniform(*delay_seconds)
                    print(f"      Waiting : {wait:.0f}s before next email...")
                    time.sleep(wait)

                else:
                    print(f"      Status  : ❌ Failed")
                    self._log_sent(lead.to_dict(), email_content["subject"], "failed")
                    failed_count += 1

        # Summary
        print(f"\n{'='*55}")
        print(f"  Campaign Complete")
        print(f"  Sent     : {sent_count}")
        print(f"  Skipped  : {skip_count} (already emailed)")
        print(f"  Failed   : {failed_count}")
        print(f"  Log      : {self.log_path}")
        print(f"{'='*55}")
