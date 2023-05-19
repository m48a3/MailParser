import imaplib
import email
from email.header import decode_header
import os
import webbrowser
from receipt_parser import RuleBased
import pandas as pd
import re

imap = imaplib.IMAP4_SSL("imap.mail.ru")  # Устанавливаем соединение

imap.login("ivankov2001@bk.ru", "Nz2cNLSLUb3UhKf6eQvb")  # Входим в наш почтовый ящик

# print(imap.list())  # print various inboxes
status, messages = imap.select("INBOX")  # Выбираем входящие

numOfMessages = int(messages[0])


def clean(text):

    return "".join(c if c.isalnum() else "_" for c in text)


def obtain_header(msg):

    subject, encoding = decode_header(msg["Subject"])[0]
    if isinstance(subject, bytes):
        subject = subject.decode(encoding)


    From, encoding = decode_header(msg.get("From"))[0]
    if isinstance(From, bytes):
        From = From.decode(encoding)


    return subject, From


def download_attachment(part):

    filename = part.get_filename()

    if filename:
        folder_name = clean(subject)
        if not os.path.isdir(folder_name):

            os.mkdir(folder_name)
            filepath = os.path.join(folder_name, filename)

            open(filepath, "wb").write(part.get_payload(decode=True))


class Product:
    __slots__ = ['id','name','price']

    def __init__(self, id, name, price):
        self.id = id
        self.name = name
        self.price = price

for i in range(numOfMessages, numOfMessages - 10, -1):
    res, msg = imap.fetch(str(i), "(RFC822)")


    for response in msg:
            if isinstance(response, tuple):
                msg = email.message_from_bytes(response[1])
                name, address = email.utils.parseaddr(msg['From'])
                if(address.find("ivankov2001@gmail.com") != -1):


                    subject, From = obtain_header(msg)

                    if msg.is_multipart():

                        for part in msg.walk():

                            content_type = part.get_content_type()
                            content_disposition = str(part.get("Content-Disposition"))

                            try:
                                body = part.get_payload(decode=True).decode()
                            except:
                                pass

                            if content_type == "text/plain" and "attachment" not in content_disposition:
                                start_idx = body.find("приход") + len("приход")
                                end_idx = body.rfind("Итог",start_idx)
                                # print(body[start_idx:end_idx].strip())
                                pattern = re.compile(
                                    r'(\d+)\s+(.+?)\s+Цена\*Кол\s+(\d+\.\d+)\s+(\d+)\s+Сумма\s+(\d+\.\d+)\s+Способ расчёта\s+(.+?)\s+Предмет расчёта\s+(.+)')

                                products = []
                                data = {'name':[]}
                                df = pd.DataFrame(data)
                                for match in pattern.finditer(body[start_idx:end_idx].strip()):
                                    id = match.group(1)
                                    name = match.group(2)
                                    price = float(match.group(3))
                                    product = Product(id, name, price)
                                    products.append(product)
                                for product in products:
                                    # print(f"ID: {product.id} Название: {product.name} Цена: {product.price}")
                                    new_row={'name': [product.name]}
                                    new_rows_df = pd.DataFrame(new_row)
                                    df = pd.concat([df, new_rows_df], ignore_index=True)

                                # data = {'name':['ДОБР.Нап.КОЛА б/алк.ПЭТ 1.5л ','COLG.Паста ТР.ДЕЙСТ.100мл ']}
                                # df = pd.DataFrame(data)
                                # rb = RuleBased()
                                # print(body[start_idx:end_idx].strip())
                                # print(rb.parse(df))
                                rb = RuleBased()
                                print(rb.parse(df))
                                rb.parse(df).to_csv('my_data.csv', index=False, encoding='utf-8')
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
                            product = 'ДОБР.Нап.КОЛА б/алк.ПЭТ 1.5л'
                            rb = RuleBased()
                            # print(rb.parse(product),"...........")

                    print("=" * 100)


imap.close()