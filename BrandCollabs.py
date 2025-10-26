import os
import re
import pandas as pd
import smtplib
import time
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

# ----------------------------
# SETTINGS
# ----------------------------
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587
SENDER_EMAIL = "hello@sophiefamilytravel.com"
# Tip: put your app password into an environment variable for security
SENDER_PASSWORD = os.getenv("MAIL_APP_PASSWORD", "egiy xrde wrax mfzw")
# Signature (Modify with your profile image & links)
signature = """
<table style="border: none; font-family: Arial, sans-serif;">
    <tr>
        <td style="padding-right: 10px;">
            <img src="https://static.wixstatic.com/media/f7a29a_386c54bd31c4400783b75053217586f0~mv2.jpg/v1/crop/x_0,y_79,w_960,h_1121/fill/w_600,h_700,al_c,q_85,usm_0.66_1.00_0.01,enc_avif,quality_auto/thumbnail_image.jpg" 
                 width="80" height="80" 
                 style="border-radius: 50%;">
        </td>
        <td>
            <p style="margin: 0; font-size: 14px;">
                <a href="https://www.instagram.com/SophieFamilyTravel" 
                   style="color: #0073e6; text-decoration: none; font-weight: bold;">
                   @SophieFamilyTravel
                </a>
            </p>
            <p style="margin: 0; font-size: 12px; color: #333;">Family Travel Content Creator</p>
            <p style="margin: 0; font-size: 12px;">
                <a href="https://www.SophieFamilyTravel.com" 
                   style="color: #b3005e; text-decoration: none; font-weight: bold;">
                   www.SophieFamilyTravel.com
                </a>
            </p>
            <p style="margin: 0;">
                <img src="https://upload.wikimedia.org/wikipedia/commons/a/a5/Instagram_icon.png" 
                     width="18" height="18" 
                     style="vertical-align: middle;">
            </p>
        </td>
    </tr>
</table>
"""
CSV_FILE = "Life.csv"
ATTACHMENT_PATH = "Media kit SophieFamilyTravel - LifestyleProducts.pdf"
SLEEP_BETWEEN_EMAILS_SEC = 2
DRY_RUN = False   # True = preview only, False = actually send
DEFAULT_PROSERV = "products and services"
DEFAULT_ACCOUNTTYPE = "travel"


# ----------------------------
# HELPERS
# ----------------------------
def normalize_space(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "").strip())

def collapse_leading_brand_duplicates(brand: str, product: str) -> str:
    """Collapse repeated brand at start of product string."""
    b = brand or ""
    p = product or ""
    if not b or not p:
        return p

    # Collapse multiple leading brand occurrences (case-insensitive)
    pattern = r"^(?:\s*" + re.escape(b) + r"\s*){2,}"
    p2 = re.sub(pattern, b + " ", p, flags=re.IGNORECASE)

    # If begins with Brand then separators, trim punctuation
    p2 = re.sub(r"^(" + re.escape(b) + r")\s*[-–—:|]\s*", r"\1 ", p2, flags=re.IGNORECASE)

    return normalize_space(p2)

def product_reference(brand_name: str, brand_product: str) -> str:
    """
    Returns a clean product reference without brand duplication.
    """
    bn = normalize_space(brand_name)
    bp_raw = normalize_space(brand_product)

    if not bp_raw and bn:
        return bn
    if not bn:
        return bp_raw

    bp = collapse_leading_brand_duplicates(bn, bp_raw)

    if bn.lower() in bp.lower():
        return bp  # already contains brand
    return normalize_space(f"{bn} {bp}")

def build_email_html(brand_name: str, brand_product: str, brand_paragraph: str, proserv: str, accounttype: str,) -> str:
    product_ref = product_reference(brand_name, brand_product)

    custom_para_html = ""
    if brand_paragraph and brand_paragraph.strip().lower() not in {"nan", "none"}:
        custom_para_html = f"<p>{brand_paragraph.strip()}</p>"

    html = f"""\
    <html>
      <body style="font-family: Arial, sans-serif; line-height:1.5;">
        <p>Hi,</p>
        <p>I hope you are having a great day so far!</p>

        <p>
          My name is Sophie, and I am the content creator behind
          <strong>@SophieFamilyTravel</strong> where I engage with over 230k followers between Instagram and YouTube.
          You can explore my account
          <a href="https://www.instagram.com/sophiefamilytravel" target="_blank">here</a>.
          {custom_para_html}
        </p>

       
        <p>
          My account focuses on family {accounttype}, and my audience consists mainly of mums looking for recommendations on
          the best products and services for their family. I reach millions of accounts each month and I believe these are also
          your ideal customers.
        </p>

        <p>
          I would love to explore a collaboration with <strong>{brand_name}</strong>. I can create content showcasing
          <strong>{product_ref}</strong> while demonstrating the benefits of using your
          {proserv or DEFAULT_PROSERV}. Additionally, the high-resolution content I create can be delivered with usage and ad rights for
          your future use on social media, your website, or ads.
        </p>

        <p>
          I am confident that my content creation skills and my engaged audience can help promote your brand
          effectively. I have attached my Media Kit, which provides more information about myself and the value I can
          bring to {brand_name}.
        </p>

        <p>
          Please let me know if this is something you'd like to pursue, and we can discuss the details further via
          email, phone call, or Zoom.
        </p>

        <p>Thank you so much for your time and consideration!</p>

        <p>
        {signature}  <!-- Signature already contains black font -->
    </body>
    </html>
    """
    return html

# ----------------------------
# LOAD CSV (+ fix common typos)
# ----------------------------
print(f"\n📥 Reading CSV file: {CSV_FILE}")
try:
    df = pd.read_csv(CSV_FILE)
    print("✅ CSV loaded successfully.")
    print("🧾 Columns found:", df.columns.tolist())

    # Auto-correct common header typos
    rename_map = {}
    if "Brand Paragrah" in df.columns:
        rename_map["Brand Paragrah"] = "Brand Paragraph"
        print("✏️ Correcting column name: 'Brand Paragrah' → 'Brand Paragraph'")

    if rename_map:
        df.rename(columns=rename_map, inplace=True)

    print(f"📊 Total rows: {len(df)}\n")
except Exception as e:
    print(f"❌ Failed to read CSV: {e}")
    raise SystemExit

# Required columns
required_columns = {"Brand", "Email", "Brand Paragraph", "Brand Product"}
if not required_columns.issubset(df.columns):
    print(f"❌ CSV is missing required columns: {required_columns}")
    raise SystemExit

# ----------------------------
# CONNECT (unless DRY_RUN)
# ----------------------------
if not DRY_RUN:
    print("🔐 Connecting to email server...")
    try:
        server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
        server.starttls()
        server.login(SENDER_EMAIL, SENDER_PASSWORD)
        print("✅ Logged in to email server.\n")
    except Exception as e:
        print(f"❌ Failed to connect or login: {e}")
        raise SystemExit
else:
    server = None
    print("🧪 DRY_RUN enabled — emails will be printed, not sent.\n")

# ----------------------------
# SEND LOOP
# ----------------------------
sent_count = 0
for _, row in df.iterrows():
    recipient_email = normalize_space(str(row["Email"]))
    brand_name = normalize_space(str(row["Brand"]))
    brand_paragraph = str(row["Brand Paragraph"])
    brand_product = str(row["Brand Product"])
    proserv = normalize_space(str(row.get("Proserv", ""))) or DEFAULT_PROSERV 
    accounttype = str(row["accounttype"]) or DEFAULT_ACCOUNTTYPE


    if not recipient_email or recipient_email.lower() in {"nan", ""}:
        print(f"⚠️ Skipping {brand_name or '[No Brand]'} — no valid email.")
        continue

    subject = f"Collaboration between SophieFamilyTravel & {brand_name}"
    body_html = build_email_html(brand_name, brand_product, brand_paragraph, proserv, accounttype)



    # Build message
    msg = MIMEMultipart()
    msg["From"] = SENDER_EMAIL
    msg["To"] = recipient_email
    msg["Subject"] = subject
    # set reply-to properly now
    msg.add_header("Reply-To", SENDER_EMAIL)
    msg.attach(MIMEText(body_html, "html"))

    # Attach media kit if present
    if ATTACHMENT_PATH:
        if os.path.exists(ATTACHMENT_PATH):
            with open(ATTACHMENT_PATH, "rb") as f:
                part = MIMEBase("application", "octet-stream")
                part.set_payload(f.read())
            encoders.encode_base64(part)
            part.add_header("Content-Disposition",
                            f"attachment; filename={os.path.basename(ATTACHMENT_PATH)}")
            msg.attach(part)
        else:
            print(f"⚠️ Attachment not found → {ATTACHMENT_PATH}")

    print(f"📨 {'Previewing' if DRY_RUN else 'Sending'} → {brand_name} <{recipient_email}>")
    if DRY_RUN:
        print(f"Subject: {subject}")
        preview = re.sub(r"\s+", " ", body_html)
        print(f"Body preview: {preview[:400]}...")
        print("-" * 60)
    else:
        try:
            server.sendmail(SENDER_EMAIL, recipient_email, msg.as_string())
            sent_count += 1
            print(f"✅ Sent to {recipient_email}")
            time.sleep(SLEEP_BETWEEN_EMAILS_SEC)
        except Exception as e:
            print(f"❌ Failed to send to {recipient_email}: {e}")

# ----------------------------
# CLEANUP
# ----------------------------
if not DRY_RUN and server is not None:
    server.quit()
print(f"\n📬 Done. {sent_count} email(s) {'would have been sent' if DRY_RUN else 'sent'}.")
