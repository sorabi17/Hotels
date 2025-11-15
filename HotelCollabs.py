import smtplib
import pandas as pd
import os
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

# Load CSV
df = pd.read_csv("C:\\Temp\\Email\\sheet19.csv")
df.columns = df.columns.str.strip().str.lower()  # Normalize column names

# Email account settings (update with your credentials)
SMTP_SERVER = "smtp.gmail.com"  # For Gmail; change for other providers
SMTP_PORT = 587
SENDER_EMAIL = "hello@sophiefamilytravel.com"
SENDER_PASSWORD = "egiy xrde wrax mfzw"

# Specify the file to attach (update with your file path)
ATTACHMENT_PATH = r"C:\Temp\Email\Media kit SophieFamilyTravel - Hotel.pdf"

# Check if the file exists
if not os.path.exists(ATTACHMENT_PATH):
    print(f"Error: Attachment {ATTACHMENT_PATH} not found.")
    exit()

# Start the email server
server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
server.starttls()
server.login(SENDER_EMAIL, SENDER_PASSWORD)

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

# Loop through each row in the CSV and send personalized emails
for _, row in df.iterrows():
    recipient_email = str(row["email"]).strip() if pd.notna(row["email"]) else "missing@example.com"
    hotel_name = row["hotel name"].strip() if pd.notna(row["hotel name"]) else "Unknown Hotel"
    subject = f"Collaboration between SophieFamilyTravel & {hotel_name}"  # Dynamic subject line
    city = row["city"].strip() if pd.notna(row["city"]) else "your city"
    hotel_type = row["type"].strip().lower() if pd.notna(row["type"]) else ""  # Hotel group type

    # Email body with a hyperlink
    profile_link = "https://www.instagram.com/SophieFamilyTravel"

    # --- Conditional hotel group paragraph ---
    if "marriott" in hotel_type:
        hotel_paragraph = (
            "I have had very successful collaborations with Marriott hotels in the past in Abu Dhabi, "
            "Cologne and Frankfurt and I would love to work with the Marriott group again."
        )
    elif "hilton" in hotel_type:
        hotel_paragraph = (
            "I have had 5 very successful collaborations with Hilton hotels in the past "
            "(Bangkok, Dubai, Portland, Sacramento and Frankfurt)."
        )
    elif "ihg" in hotel_type:
        hotel_paragraph = (
            "I have recently had a very successful collaboration with an IHG hotel in Dubai – "
            "the Intercontinental Residences and Suites."
        )
    elif "accor" in hotel_type:
        hotel_paragraph = (
            "I have worked with Accor Hotels 3 times in the past in Dubai, Germany and Singapore "
            "and would love to work with you again!"
        )
    else:
        hotel_paragraph = ""  # No group mention if not part of the above

    body = f"""\
<html>
    <body style="color: black; font-family: Arial, sans-serif;">
        <p style="color: black;">Hello,</p>
        <p style="color: black;">I hope you are having a great day so far!</p>
        <p style="color: black;">
            My name is Sophie, and I am the content creator behind <strong>@SophieFamilyTravel</strong>, 
            where I engage with over 230,000 enthusiastic followers between my Instagram and YouTube. 
            <a href="{profile_link}" style="color: #0073e6; text-decoration: none; font-weight: bold;">
                You can explore my profile here
            </a>.
        </p>
        <p style="color: black;">
            I will be coming to {city} mid January to create valuable content for my audience.
            {" " + hotel_paragraph if hotel_paragraph else ""} 
        </p>
        <p style="color: black;">
            <strong>{hotel_name}</strong> immediately caught my eye, and I would love to create some content for you, 
            showcasing all you have to offer! My content would be perfect for you to advertise your hotel 
            as an amazing destination for families!
        </p>
        <p style="color: black;">
            More specifically, I would like to offer you a variety of options that include 
            <strong>High Resolution photos and videos</strong> that you can use for your social media platforms 
            and website, as well as posts on my own profile showcasing your hotel to my audience. 
            I would love to provide this for free, in exchange for a stay at your stunning place 
            whilst I am shooting content.
        </p>
        <p style="color: black;">
            I hope this is something you are interested in. My previous collaborations have 
            driven significant engagement and inspired many of my followers to explore and book 
            holidays and hotels. Please find my <strong>Media Kit</strong> attached with more 
            details about what I can offer.
        </p>
        <p style="color: black;">I would love to answer any questions you have.</p>
        <p style="color: black;">Thank you so much for your time!</p>

        {signature}
    </body>
</html>
"""

    # Construct the email
    msg = MIMEMultipart()
    msg["From"] = SENDER_EMAIL
    msg["To"] = recipient_email
    msg["Subject"] = subject
    msg.attach(MIMEText(body, "html"))  # Sending as HTML for hyperlink support

    # Attach the Media Kit
    with open(ATTACHMENT_PATH, "rb") as attachment:
        part = MIMEBase("application", "octet-stream")
        part.set_payload(attachment.read())

    encoders.encode_base64(part)
    part.add_header("Content-Disposition", f"attachment; filename={os.path.basename(ATTACHMENT_PATH)}")
    msg.attach(part)

    # Send the email with error handling
    try:
        server.sendmail(SENDER_EMAIL, recipient_email, msg.as_string())
        print(f"✅ Email sent to {recipient_email} for {hotel_name}")
    except Exception as e:
        print(f"❌ Failed to send email to {recipient_email}: {e}")


# Close the server connection
server.quit()
print("✅ All emails sent successfully!")
