from flask import Flask, request
from docx import Document
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.mime.application import MIMEApplication
from email import encoders
import threading
import os
import time
import uuid
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
import subprocess
import configparser


def send_email_with_attachment(name,sender_email, sender_password, receiver_email, subject, body, file_path):
    # Setup the MIME
    message = MIMEMultipart()
    message['From'] = sender_email
    message['To'] = receiver_email
    message['Subject'] = subject

    # Attach body
    message.attach(MIMEText(body, 'plain'))

    with open(file_path, 'rb') as file:
        attach = MIMEApplication(file.read(),_subtype="pdf")
        attach.add_header('Content-Disposition','attachment',filename=str(name)+'.pdf')
    message.attach(attach)

    # Connect to the SMTP server and send the email
    with smtplib.SMTP('smtp.gmail.com', 587) as server:
        server.starttls()
        server.login(sender_email, sender_password)
        server.sendmail(sender_email, receiver_email, message.as_string())

app = Flask(__name__)

@app.route('/webhook', methods=['POST'])
def handle_webhook():
    config = configparser.ConfigParser()
    config.read('setting.ini')
    data = request.json 
    print("Received form data:", data)
    form_data = data['data']
    doc = Document('template.docx')
    for problem in config['PROBLEMS']:
        if problem in form_data:
            find_replace_word(doc, config['PROBLEMS'][problem], form_data[problem])
            break

    receiver_email= data['email']
    email_subject = config['EMAIL']['subject']
    email_body= config['EMAIL']['body']
    uuid4 = str(uuid.uuid4())
    doc.save(uuid4 + '.docx')
    fiiename= uuid4 + '.docx'
    pdfname = "Type your pdf name here"
    task_thread = threading.Thread(target=send_request_and_process, args=(pdfname, receiver_email, email_subject, email_body, fiiename))
    task_thread.start()
    return "Received the data successfully!"

def send_request_and_process(name, receiver_email, email_subject, email_body, fiiename):
    config = configparser.ConfigParser()
    config.read('setting.ini')
    while not os.path.exists(fiiename):
        print('wait for ' + fiiename + '...')
        time.sleep(1)
    time.sleep(1)
    pdfname= fiiename[:-4] + 'pdf'
    subprocess.run(['doc2pdf', fiiename])
    print('send email to ' + receiver_email + '...')
    send_email_with_attachment(name, config["EMAIL_SETTING"]["email_address"], config['EMAIL_SETTING']["password"] , receiver_email, email_subject, email_body, pdfname) 
    print('send email to ' + receiver_email + ' successfully!')

def find_replace_word(doc, old_word, new_word):
    for paragraph in doc.paragraphs:
        if old_word in paragraph.text:
            paragraph.text = paragraph.text.replace(old_word, new_word)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                if old_word in cell.text:
                    cell.text = cell.text.replace(old_word, new_word)

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)