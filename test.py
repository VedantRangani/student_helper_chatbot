import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import os 
from dotenv import load_dotenv

load_dotenv()
# 🔑 CONFIG
sender_email = "vedantranganir@gmail.com"
receiver_email = "vedantranganir@gmail.com"
password = os.getenv("MAIL_PASSWORD")

# 📧 MESSAGE
msg = MIMEMultipart()
msg["From"] = sender_email
msg["To"] = receiver_email
msg["Subject"] = "Test Email"

body = "This is a test email from your chatbot project."
msg.attach(MIMEText(body, "plain"))

try:
    # 🔥 CONNECT TO SMTP SERVER
    server = smtplib.SMTP("smtp.gmail.com", 587)
    server.starttls()
    server.login(sender_email, password)

    server.send_message(msg)
    server.quit()

    print("✅ Email sent successfully!")

except Exception as e:
    print("❌ Error:", e)