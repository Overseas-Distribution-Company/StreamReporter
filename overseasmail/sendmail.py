import smtplib
import ssl
from typing import List


class OverseasMail:
    """Helper class to send mails over the SMTP protocol using overseas own SMTP server."""

    HOST = 'overseas-be.mail.protection.outlook.com'

    _mailer: smtplib.SMTP
    _sender: str
    _to: List[str]
    _cc: List[str]
    _subject: str

    def __init__(self, **kwargs):
        """

        Keyword Args:
            sender (str): Mail address of the sender.
            to (str): Mail address of the receiver.
            cc (str): Mail address in cc.
            subject (str): Subject of the mail.
        """
        self._mailer = smtplib.SMTP(self.__class__.HOST)
        context = ssl.create_default_context()
        self._mailer.starttls(context=context)
        self.sender = kwargs.get('sender', '')
        self._to = [kwargs.get('to', '')]
        self._cc = [kwargs.get('cc', '')]
        self._subject = kwargs.get('subject', '')

    def add_receiver(self, mail_address):
        """Adds a mail address To send to.

        Args:
            mail_address (str): A mail address we want to send to
        """
        self._to.append(mail_address)

    
test_obj = OverseasMail()
