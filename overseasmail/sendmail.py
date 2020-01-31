import smtplib
import re
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from typing import List


class OverseasMail:
    """Helper class to send mails over the SMTP protocol using overseas own SMTP server.

    Attributes:
        subject(str): Subject string in mail format
        body(str): Body of the email.
    """

    _HOST = 'overseas-be.mail.protection.outlook.com'
    _INVALID_MAIL_ERROR = 'Not a valid E-mail address.'

    _sender: str
    _to: List[str]
    _cc: List[str]
    _subject: str
    message: str

    def __init__(self, **kwargs):
        """
        Keyword Args:
            sender (str): Mail address of the sender.
            to (str): Mail address of the receiver.
            cc (str): Mail address in cc.
            subject (str): Subject of the mail.
        """
        self._sender = kwargs.get('sender', '')

        first_to = kwargs.get('to', False)
        self._to = [first_to] if first_to else []
        first_to = kwargs.get('cc', '')
        self._cc = [first_to] if first_to else []
        self.subject = kwargs.get('subject', '')
        self.message = kwargs.get('body', '')
        self._attachments = []

    @property
    def sender(self):
        return self._sender

    @sender.setter
    def sender(self, email):
        if not self.check_valid_mail(email):
            raise ValueError(self._INVALID_MAIL_ERROR)
        self._sender = email

    def add_receiver(self, mail_address: str):
        """Adds a mail address To send to.

        Args:
            mail_address (str): A mail address we want to send to
        """
        if not self.check_valid_mail(mail_address):
            raise ValueError(self._INVALID_MAIL_ERROR)
        self._to.append(mail_address)

    def add_cc(self, mail_address: str):
        """

        Args:
            mail_address (str): A mail address in CC.
        """
        if not self.check_valid_mail(mail_address):
            raise ValueError(self._INVALID_MAIL_ERROR)
        self._cc.append(mail_address)

    def add_attachment(self, attachment_path: str):
        """ Adds an attachment to the mail
        Args:
            attachment_path(path):
        """

        self._attachments.append(attachment_path)

    def send_mail(self):
        """ Sends the defined mail.
        """
        mailer = smtplib.SMTP(self._HOST)
        # path = str(pathlib.Path(__file__).parent.absolute())
        # context = ssl.create_default_context()
        # context.load_cert_chain( keyfile=path + "/overseas_wildcard.pem", password='Azerty123')
        # mailer.starttls(context=context)

        message = MIMEMultipart()
        message["from"] = self._sender
        message["to"] = ','.join(self._to)
        message['subject'] = self.subject
        message["Cc"] = ','.join(self._cc)
        message.attach(MIMEText(self.message, 'plain'))

        for file in self._attachments:
            with open(file, "rb") as attachment:
                part = MIMEBase('application', 'octet-stream')
                part.set_payload(attachment.read())
                encoders.encode_base64(part)
                part.add_header('Content-Disposition', f'attachment; filename="{file}"')
                message.attach(part)
        mailer.sendmail(self._sender, self._to, message.as_string())
        mailer.quit()

    @staticmethod
    def check_valid_mail(mail_address: str) -> bool:
        """ Check if a mail address if valid formed. Returns True if the form is valid, Else False

        Args:
            mail_address (str):

        Returns (bool): True for well formed mail address, False for not well formed mail address
        """
        regex = '^\w+([\.-]?\w+)*@\w+([\.-]?\w+)*(\.\w{2,3})+$'
        return True if re.search(regex, mail_address) is not None else False
