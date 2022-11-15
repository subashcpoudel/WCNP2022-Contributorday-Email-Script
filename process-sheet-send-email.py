import os, pandas, smtplib, email, ssl

#Read Contributors List
contributors = pandas.read_excel('list.xlsx', sheet_name='Sheet1', header=None)

#Setup Email
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText




#Read Gmail Username, Pass from .env
from dotenv import load_dotenv
load_dotenv()
sender_email_id = os.getenv('sender_email_id')
sender_email_id_password = os.getenv('sender_email_id_password')


# iterate through each row of contributors
for index, row in contributors.iterrows():

    subject = "WordCamp Nepal 2022 - Contributor Day"
    body = "This is an email with attachment sent from Python"
    sender_email = "nepal@wordcamp.org"
    receiver_email = row[4]

    # Create a multipart message and set headers
    message = MIMEMultipart()
    message["From"] = sender_email
    message["To"] = receiver_email
    message["Subject"] = subject

    # Add body to email
    message.attach(MIMEText(body, "plain"))
    filename = "Contributor Day Certificate-"+row[0]+".jpg"  # In same directory as script

    # Open PDF file in binary mode
    with open(filename, "rb") as attachment:
        # Add file as application/octet-stream
        # Email client can usually download this automatically as attachment
        part = MIMEBase("application", "octet-stream")
        part.set_payload(attachment.read())

    # Encode file in ASCII characters to send by email    
    encoders.encode_base64(part)

    # Add header as key/value pair to attachment part
    part.add_header(
        "Content-Disposition",
        f"attachment; filename= {filename}",
    )

    # Add attachment to message and convert message to string
    message.attach(part)
    text = message.as_string()

    # Log in to server using secure context and send email
    context = ssl.create_default_context()
    with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
        server.login(sender_email_id, sender_email_id_password)
        server.sendmail(sender_email, receiver_email, text)