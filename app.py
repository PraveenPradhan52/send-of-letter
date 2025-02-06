from flask import Flask, render_template, request
import fitz  # PyMuPDF for PDF editing
import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

app = Flask(__name__)

# Email Configuration
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587
EMAIL_SENDER = "your_email@gmail.com"
EMAIL_PASSWORD = "your_app_password"  # Use App Password for security

# Function to modify the PDF template
def generate_offer_letter(name, start_date, duration):
    template_path = "offer_template.pdf"
    output_path = f"offer_letter_{name}.pdf"

    doc = fitz.open(template_path)

    for page in doc:
        text = page.get_text("text")
        text = text.replace("{{NAME}}", name)
        text = text.replace("{{START_DATE}}", start_date)
        text = text.replace("{{DURATION}}", duration)
        
        # Clear and rewrite the modified text
        page.clean_contents()
        page.insert_text((50, 100), text, fontsize=12)

    doc.save(output_path)
    doc.close()
    return output_path

# Function to send an email with the PDF
def send_email(recipient_email, name, pdf_path):
    msg = MIMEMultipart()
    msg["From"] = EMAIL_SENDER
    msg["To"] = recipient_email
    msg["Subject"] = "Internship Offer Letter"

    body = f"Dear {name},\n\nCongratulations! Please find your internship offer letter attached.\n\nBest Regards,\nCompany HR"
    msg.attach(MIMEText(body, "plain"))

    with open(pdf_path, "rb") as attachment:
        pdf_attachment = MIMEApplication(attachment.read(), _subtype="pdf")
        pdf_attachment.add_header("Content-Disposition", f"attachment; filename=offer_letter.pdf")
        msg.attach(pdf_attachment)

    try:
        server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
        server.starttls()
        server.login(EMAIL_SENDER, EMAIL_PASSWORD)
        server.sendmail(EMAIL_SENDER, recipient_email, msg.as_string())
        server.quit()
        print("Email sent successfully!")
    except Exception as e:
        print(f"Error sending email: {e}")

@app.route("/", methods=["GET", "POST"])
def internship_form():
    if request.method == "POST":
        name = request.form["name"]
        email = request.form["email"]
        start_date = request.form["start_date"]
        duration = request.form["duration"]

        pdf_path = generate_offer_letter(name, start_date, duration)
        send_email(email, name, pdf_path)

        return "Offer Letter Sent Successfully!"
    return render_template("form.html")

if __name__ == "__main__":
    app.run(debug=True)
