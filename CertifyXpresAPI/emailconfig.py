from typing import List

from fastapi_mail import FastMail, MessageSchema, ConnectionConfig, MessageType
from pydantic import BaseModel, EmailStr
from dotenv import load_dotenv
import os

conf = ConnectionConfig(
    MAIL_USERNAME="singhbishts08@gmail.com",
    MAIL_PASSWORD="ifcs sycj tkwq xalo",
    MAIL_FROM="singhbishts08@gmail.com",
    MAIL_PORT=587,
    MAIL_SERVER="smtp.gmail.com",
    MAIL_STARTTLS=True,
    MAIL_SSL_TLS=False,
    USE_CREDENTIALS=True,
    VALIDATE_CERTS=True
)

class EmailSchema(BaseModel):
    email: List[EmailStr]

async def send_email(emails: List[EmailStr], certificates_paths: List[str]):

    load_dotenv()
    fm = FastMail(conf)

    template = """
        <html>
            <body>
                <h1>Certificate PDF</h1>
                <p>Dear user,</p>
                <p>Please find attached your certificate PDF.</p>
                <p>Best regards,</p>
                <p>The CertifyXpress Team</p>
            </body>
        </html>
    """

    for email, certificate_path in zip(emails, certificates_paths):
        message = MessageSchema(
            subject="Certificate PDF",
            recipients=[email],
            body=template,
            subtype=MessageType.html,
            attachments=[certificate_path]
        )
        await fm.send_message(message=message)
