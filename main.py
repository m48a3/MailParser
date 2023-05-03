import imaplib
import email
from email.header import decode_header
import os
import webbrowser

imap = imaplib.IMAP4_SSL("imap.mail.ru")  # establish connection

imap.login("ivankov2001@bk.ru", "Nz2cNLSLUb3UhKf6eQvb")  # login

# print(imap.list())  # print various inboxes
status, messages = imap.select("INBOX")  # select inbox

numOfMessages = int(messages[0])  # get number of messages


def clean(text):
    # clean text for creating a folder
    return "".join(c if c.isalnum() else "_" for c in text)


def obtain_header(msg):
    # decode the email subject
    subject, encoding = decode_header(msg["Subject"])[0]
    if isinstance(subject, bytes):
        subject = subject.decode(encoding)

    # decode email sender
    From, encoding = decode_header(msg.get("From"))[0]
    if isinstance(From, bytes):
        From = From.decode(encoding)

    print("Subject:", subject)
    print("From:", From)
    return subject, From


def download_attachment(part):
    # download attachment
    filename = part.get_filename()

    if filename:
        folder_name = clean(subject)
        if not os.path.isdir(folder_name):
            # make a folder for this email (named after the subject)
            os.mkdir(folder_name)
            filepath = os.path.join(folder_name, filename)
            # download attachment and save it
            open(filepath, "wb").write(part.get_payload(decode=True))


for i in range(numOfMessages, numOfMessages - 10, -1):
    res, msg = imap.fetch(str(i), "(RFC822)")  # fetches the email using it's ID


    for response in msg:
            if isinstance(response, tuple):
                msg = email.message_from_bytes(response[1])
                name, address = email.utils.parseaddr(msg['From'])
                if(address.find("ivankov2001@gmail.com") != -1):

                    # print(email.utils.parseaddr(msg['From']))
                    subject, From = obtain_header(msg)

                    if msg.is_multipart():
                        # iterate over email parts
                        for part in msg.walk():
                            # extract content type of email
                            content_type = part.get_content_type()
                            content_disposition = str(part.get("Content-Disposition"))

                            try:
                                body = part.get_payload(decode=True).decode()
                            except:
                                pass

                            if content_type == "text/plain" and "attachment" not in content_disposition:
                                start_idx = body.find("приход") + len("приход")
                                end_idx = body.rfind("Наличные",start_idx)
                                print(body[start_idx:end_idx].strip())
                            elif "attachment" in content_disposition:
                                download_attachment(part)
                    else:
                        # extract content type of email
                        content_type = msg.get_content_type()
                        # get the email body
                        body = msg.get_payload(decode=True).decode()
                        if content_type == "text/plain":
                            start_idx = body.find("приход") + len("приход")
                            end_idx = body.rfind("Наличные", start_idx)
                            print(body[start_idx:end_idx].strip())


                    print("=" * 100)


imap.close()