import csv
import imaplib
import email
from email.header import decode_header
import os
from fpdf import FPDF
import plotly.express as px
import pandas as pd


def test_emailvalidation():
    try:
        username = 'nagma.peerzade.17@gmail.com'
        password = 'ammu@0910'
        # create an IMAP4 class with SSL
        server = imaplib.IMAP4_SSL("imap.gmail.com")
        # authenticate
        server.login(username, password)

    except Exception as e:
        print("invalid email or password please enter again")
        test_emailvalidation()
        print("----------------------------------------------")
    firstMailExtraction(server, username)
    matplot(server)


def firstMailExtraction(imap, username):
    z = open("omkar.txt", "w")
    body = ""
    csvfile = open('Names.csv', 'w')
    field_names = ['Subject', 'From', 'Date', 'Body']
    writer = csv.DictWriter(csvfile, fieldnames=field_names)
    writer.writeheader()
    status, messages = imap.select("INBOX")
    N = 8
    # total number of emails
    messages = int(messages[0])

    for i in range(messages, messages - N, -1):
        # fetch the email message by ID
        res, msg = imap.fetch(str(i), "(RFC822)")
        for response in msg:
            if isinstance(response, tuple):
                # parse a bytes email into a message object
                msg = email.message_from_bytes(response[1])
                # decode the email subject
                subject, encoding = decode_header(msg["Subject"])[0]
                if isinstance(subject, bytes):
                    # if it's a bytes, decode to str
                    subject = subject.decode(encoding)
                # decode email sender
                From, encoding = decode_header(msg.get("From"))[0]
                if isinstance(From, bytes):
                    From = From.decode(encoding)
                Date, encoding = decode_header(msg.get("Date"))[0]
                if isinstance(Date, bytes):
                    Date = Date.decode(encoding)
                abc = Date[0:17]
                # print("Subject:", subject)
                z.write("Subject = ")
                z.write(subject)
                # print("From:", From)
                z.write("From = ")
                z.write(From)
                z.write("Date = ")

                z.write(abc)

                # if the email message is multipart
                if msg.is_multipart():
                    # iterate over email parts
                    for part in msg.walk():
                        # extract content type of email
                        content_type = part.get_content_type()
                        content_disposition = str(part.get("Content-Disposition"))
                        try:
                            # get the email body
                            body = part.get_payload(decode=True).decode()
                        except:
                            pass
                        if content_type == "text/plain" and "attachment" not in content_disposition:
                            # print text/plain emails and skip attachments
                            z.write("body = ")
                            z.write(body)
                            # print(body)
                            field_names = ['Subject', 'From', 'Date', 'Body']
                            dict = {"Subject": subject, "From": From, "Date": abc, "Body": body}
                            l = [dict]
                            with open('Names.csv', 'a') as csvfile:
                                writer = csv.DictWriter(csvfile, fieldnames=field_names)
                                writer.writeheader()
                                writer.writerows(l)

                        elif "attachment" in content_disposition:
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
                else:

                    # extract content type of email
                    content_type = msg.get_content_type()
                    # get the email body
                    body = msg.get_payload(decode=True).decode()
                    if content_type == "text/plain":
                        # print only text email parts
                        print(body)


    # close the connection and logout
    z.close()
    pdf = FPDF()
    # Add a page
    pdf.add_page()
    # set style and size of font
    # that you want in the pdf
    pdf.set_font("Arial", size=15)
    # open the text file in read mode
    f = open("omkar.txt", "r")
    # insert the texts in pdf
    for x in f:
        pdf.cell(200, 10, txt=x, ln=1, align='C')
    # save the pdf with name
    pdf.output("omkar.pdf")
    print(".....conversion done successfully......")
    imap.close()


def matplot(imap):
    a = []
    body = ''

    csvfile = open('Names.csv', 'w')
    field_names = ['Subject', 'From', 'Date']
    writer = csv.DictWriter(csvfile, fieldnames=field_names)
    writer.writeheader()
    status, messages = imap.select("INBOX")
    # total number of emails
    messages = int(messages[0])
    tea = messages
    for i in range(messages, messages - tea, -1):
        # fetch the email message by ID
        res, msg = imap.fetch(str(i), "(RFC822)")
        for response in msg:
            if isinstance(response, tuple):
                # parse a bytes email into a message object
                msg = email.message_from_bytes(response[1])
                subject, encoding = decode_header(msg["Subject"])[0]
                if isinstance(subject, bytes):
                    subject = subject.decode(encoding)
                From, encoding = decode_header(msg.get("From"))[0]
                if isinstance(From, bytes):
                    From = From.decode(encoding)

                Date, encoding = decode_header(msg["Date"])[0]
                if isinstance(Date, bytes):
                    Date = Date.decode(encoding)
                abc = Date[0:17]

                dicti = {"Subject": subject, "From": From, "Date": abc}
                g = [dicti]
                writer.writerows(g)


def clean(text):
    # clean text for creating a folder
    return "".join(c if c.isalnum() else "_" for c in text)


test_emailvalidation()
df = pd.read_csv('Names.csv')
fig_pie = px.bar(data_frame=df, x='Subject', y='Date')
fig_pie.show()
