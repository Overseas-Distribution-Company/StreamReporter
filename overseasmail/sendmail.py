import smtplib
import ssl
import re
from typing import List


class OverseasMail:
    """Helper class to send mails over the SMTP protocol using overseas own SMTP server.

    Attributes:
        sender(str): Sender string in mail format.
        subject(str): Subject string in mail format
        -

    """

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
        self.subject = kwargs.get('subject', '')

    def add_receiver(self, mail_address: str):
        """Adds a mail address To send to.

        Args:
            mail_address (str): A mail address we want to send to
        """
        self._to.append(mail_address)

    def add_cc(self, mail_address: str):
        """

        Args:
            mail_address (str): A mail addres in CC.
        """
        self._cc.append(mail_address)

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
