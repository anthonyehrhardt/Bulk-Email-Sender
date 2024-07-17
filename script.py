import pandas as pd
import smtplib

# Define your email and password (or app-specific password if 2FA is enabled)
SenderAddress = "aehrhardt@gmail.com"
password = "your_app_specific_password"

# Read the emails from the Excel file
e = pd.read_excel("Email.xlsx")
emails = e['Emails'].values

# Create an SMTP session
server = smtplib.SMTP("smtp.gmail.com", 587)
server.starttls()

try:
    # Log in to your email account
    server.login(SenderAddress, password)
    print("Logged in successfully")

    # Define the email content
    msg = "Hello, this is an email from ARTI"
    subject = "Hello world"
    body = "Subject: {}\n\n{}".format(subject, msg)

    # Send the email to each address
    for email in emails:
        server.sendmail(SenderAddress, email, body)
        print(f"Email sent to {email}")

except smtplib.SMTPAuthenticationError as e:
    print("Failed to log in")
    print(e)
finally:
    # Terminate the SMTP session
    server.quit()
