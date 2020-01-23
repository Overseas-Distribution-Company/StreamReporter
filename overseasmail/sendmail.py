import smtplib
import ssl
import re
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

    def send_mail(self):
        """ Sends the defined mail.
        """
        mailer = smtplib.SMTP(self._HOST)
        context = ssl.create_default_context()
        mailer.starttls()

        message = MIMEMultipart()
        message["from"] = self._sender
        message["to"] = ','.join(self._to)
        message['subject'] = self.subject
        message["Cc"] = ','.join(self._cc)
        message.attach(MIMEText(self.message, 'plain'))
        print(self._to)
        mailer.sendmail(self._sender, self._to, message.as_string())

    @staticmethod
    def check_valid_mail(mail_address: str) -> bool:
        """ Check if a mail address if valid formed. Returns True if the form is valid, Else False

        Args:
            mail_address (str):

        Returns (bool): True for well formed mail address, False for not well formed mail address
        """
        regex = '^\w+([\.-]?\w+)*@\w+([\.-]?\w+)*(\.\w{2,3})+$'
        return True if re.search(regex, mail_address) is not None else False


test_obj = OverseasMail()

test_obj.sender = 'testing@overseas.be'
test_obj.add_receiver('rheirman@overseas.be')
test_obj.add_cc('robbeheirman@msn.com')
test_obj.subject = 'plain testing purposes'
test_obj.message = 'Dit is een test mail'
test_obj.send_mail()
