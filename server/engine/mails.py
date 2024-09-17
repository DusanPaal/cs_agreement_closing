"""
The module provides interface for sending and managing emails.
It also uses the exchangelib library to connect to the Exchange
server via Exchange Web Services (EWS) in order to retrieve
messages and save message attachment under a specified account.
"""

import os
import re
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from os.path import exists, isfile, join, split
from smtplib import SMTP
from typing import Union

import exchangelib as xlib
from exchangelib import (
    Account, Build, Configuration, Identity,
    Message, OAuth2Credentials, Version
)


# custom message classes
class SmtpMessage(MIMEMultipart):
    """Wraps MIMEMultipart objects
    that represent messages to sent
    via an SMTP server.
    """

# custom exceptions
class AttachmentSavingError(Exception):
    """When any exception is caught
    during writing of attachment
    data to a file.
    """

class FolderNotFoundError(Exception):
    """A local file directory doesn't exist."""

class MessageNotFoundError(Exception):
    """Message with a given ID doesn't exist."""

class UndeliveredError(Exception):
    """Message delivery failes."""


def _get_credentials(acc_name: str) -> OAuth2Credentials:
    """Returns credentails for a given account."""

    cred_dir = join(os.environ["APPDATA"], "bia")
    cred_path = join(cred_dir, f"{acc_name.lower()}.token.email.dat")

    if not isfile(cred_path):
        raise FileNotFoundError(f"Credentials file not found: {cred_path}")

    with open(cred_path, encoding = "utf-8") as stream:
        lines = stream.readlines()

    params = {
        "client_id": None,
        "client_secret": None,
        "tenant_id": None,
        "identity": Identity(primary_smtp_address = acc_name)
    }

    for line in lines:

        if ":" not in line:
            continue

        tokens = line.split(":")
        param_name = tokens[0].strip()
        param_value = tokens[1].strip()

        if param_name == "Client ID":
            key = "client_id"
        elif param_name == "Client Secret":
            key = "client_secret"
        elif param_name == "Tenant ID":
            key = "tenant_id"

        params[key] = param_value

    # verify loaded parameters
    if params["client_id"] is None:
        raise ValueError("Parameter 'client_id' not found!")

    if params["client_secret"] is None:
        raise ValueError("Parameter 'client_secret' not found!")

    if params["tenant_id"] is None:
        raise ValueError("Parameter 'tenant_id' not found!")

    # params OK, create credentials
    creds = OAuth2Credentials(
        params["client_id"],
        params["client_secret"],
        params["tenant_id"],
        params["identity"]
    )

    return creds

def _sanitize_emails(addr: Union[str,list]) -> list:
    """Trims email address(es) and checks whether their 
    name format complies to the company's naming standards.
    """

    mails = []
    validated = []

    if isinstance(addr, str):
        mails = [addr]
    elif isinstance(addr, list):
        mails = addr
    else:
        raise TypeError(f"Argument 'addr' has invalid type: '{type(addr)}'!")

    for mail in mails:

        stripped = mail.strip()
        validated.append(stripped)

        # check if email is Ledvance-specific
        if re.search(r"\w+\.\w+@ledvance.com", stripped) is None:
            raise ValueError(f"Invalid email address format: '{stripped}!")

    return validated

def create_message(
        from_addr: str, to_addr: Union[str,list], subj: str,
        body: str, att: Union[str,list] = None) -> SmtpMessage:
    """Creates a SMTP message.

    Params:
    -------
    from_addr:
        Email address of the sender.

    to_addr:
        Recipient address(es).
        If a string email address is used, the message 
        will be sent to that specific address. If multiple 
        addresses are used, then the message will be sent 
        to all of the recipients.

    subj:
        Message subject.

    body:
        Message body in HTML format.

    att:
        Any valid path(s) to message atachment file(s).
        If None is used (default value), then a message without
        any file attachments will be created. If a file path is used,
        then this file will be attached to the message. If multiple
        paths are used, these will be attached as multiple attachments
        to the message.

    Raises:
    -------
    FileNotFoundError:
        If any of the attachment paths used is not found.

    Returns:
    --------
    A SmtpMessage object representing the message.
    """

    if not isinstance(to_addr, str) and len(to_addr) == 0:
        raise ValueError("No message recipients provided!")

    if att is None:
        att_paths = []
    elif isinstance(att, list):
        att_paths = att
    elif isinstance(att, str):
        att_paths = [att]
    else:
        raise TypeError(f"Argument 'att' has invalid type: '{type(att)}'!")

    for att_path in att_paths or []:
        if not isfile(att_path):
            raise FileNotFoundError(f"Attachment not found: '{att_path}'!")

    # sanitize input
    recips = _sanitize_emails(to_addr)

    # process
    email = SmtpMessage()
    email["Subject"] = subj
    email["From"] = from_addr
    email["To"] = ";".join(recips)
    email.attach(MIMEText(body, "html"))

    for att_path in att_paths:

        with open(att_path, "rb") as file:
            payload = file.read()

        # The content type "application/octet-stream" means
        # that a MIME attachment is a binary file
        part = MIMEBase("application", "octet-stream")
        part.set_payload(payload)
        encoders.encode_base64(part)

        # get file name
        file_name = split(att_path)[1]

        # Add header
        part.add_header(
            "Content-Disposition",
            f"attachment; filename = {file_name}"
        )

        # Add attachment to the message
        # and convert it to a string
        email.attach(part)

    return email

def send_message(msg: SmtpMessage, host: str, port: int) -> None:
    """Sends a message using SMTP server.

    Params:
    -------
    msg:
        A SmtpMessage object representing the message to be sent.

    host:
        Name of the SMTP host server used for message sending.

    port:
        Number o the SMTP server port.

    Raises:
    -------
    UndeliveredError:
        When message fails to reach all the required recipients.

    TimeoutError:
        When attempt to connect to the SMTP server times out.
    """

    with SMTP(host, port, timeout = 30) as smtp_conn:
        smtp_conn.set_debuglevel(0) # off = 0; verbose = 1; timestamped = 2
        send_errs = smtp_conn.sendmail(msg["From"], msg["To"], msg.as_string())

    if len(send_errs) != 0:
        failed_recips = ";".join(send_errs.keys())
        raise UndeliveredError(f"Message undelivered to: {failed_recips}")

def save_attachments(msg: Message, folder_path: str, ext: str = None) -> list:
    """Saves message attachments of a specific type to a local file.

    Params:
    -------
    msg:
        An exchangelib.Message object
        that represents the email.

    folder_path:
        Path to the folder where attachments will be stored.

    ext:
        File extension, that determines which attachments to consider (default None).
        By default, all attachments will be downloaded, regardless of the file type.
        If a file extension (e.g. '.pdf') is used, then only attachments having that
        particular file type will be downloaded.

    Returns:
    --------
    A list of file paths to the stored attachments.

    Rasises:
    --------
    FolderNotFoundError:
        When 'folder_path' argument refers to an non-existitg folder.

    AttachmentSavingError:
        When writing attachemnt data to a file fails.
    """

    if not exists(folder_path):
        raise FolderNotFoundError(f"Folder does not exist: {folder_path}")

    file_paths = []

    for att in msg.attachments:

        file_path = join(folder_path, att.name)

        if not (ext is None or file_path.lower().endswith(ext)):
            continue

        try:
            with open(file_path, "wb") as a_file:
                a_file.write(att.content)
        except Exception as exc:
            raise AttachmentSavingError(str(exc)) from exc

        if not isfile(file_path):
            raise AttachmentSavingError(f"Error writing attachment data to file: {file_path}")

        file_paths.append(file_path)

    return file_paths

def get_account(mailbox: str, name: str, x_server: str) -> Account:
    """Returns an account for a shared mailbox.

    Params:
    -------
    mailbox:
        Name of the shared mailbox.

    name:
        Name of the account for which
        the credentails will be obtained.

    x_server:
        Name of the MS Exchange server.

    Returns:
    --------
    An exchangelib.Account object.
    """

    build = Build(major_version = 15, minor_version = 20)

    cfg = Configuration(
        _get_credentials(name),
        server = x_server,
        auth_type = xlib.OAUTH2,
        version = Version(build)
    )

    acc = Account(
        mailbox,
        config = cfg,
        autodiscover = False,
        access_type = xlib.IMPERSONATION
    )

    return acc

def get_message(acc: Account, email_id: str) -> Message:
    """Returns a message from an account.

    Params:
    -------
    acc:
        Account object containing the message.

    email_id:
        A unique string ID, the 'message_id' property of the message to fetch.

    Returns:
    --------
    An exchangelib:Message object representing the message.
    If no message is found, then None is returned.

    Raises:
    -------
    MessageNotFoundError:
        When message with a given ID doesn't exist.
    """

    # sanitize input
    if not email_id.startswith("<"):
        email_id = f"<{email_id}"
    if not email_id.endswith(">"):
        email_id = f"{email_id}>"

    # process
    emails = acc.inbox.walk().filter(message_id = email_id).only(
        "subject", "text_body", "headers", "sender",
        "attachments", "datetime_received", "message_id"
    )

    if emails.count() == 0:
        raise MessageNotFoundError(f"Could not find a message with ID: '{email_id}'!")

    return emails[0]
