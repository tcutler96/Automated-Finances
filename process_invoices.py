import imaplib
import PyPDF2
import email
import json
import os


class ProcessInvoices:
    def __init__(self, test=True):
        if test:
            self.base_folder = os.path.dirname(os.path.abspath(__file__))
        else:
            self.base_folder = 'C:\\Users\Tom\Important\Self Employed\Yodel'
        self.folder_path = os.path.join(self.base_folder, 'Invoices 2021-22')
        self.host = 'imap-mail.outlook.com'
        self.subjects = {'YODEL Pre-Advice Document': ['Pre-Advice Report', '.csv'],
                         'YODEL Self-Billed Invoice': ['Self Billing Invoice', '.pdf'],
                         'Yodel Third Party Billing': ['Insurance Invoice', '.pdf']}
        self.senders = ['equitas.billing@yodel.co.uk', '<Equitas_Billing@Yodel.co.uk>']
        with open('credentials.json', 'r') as f:
            file = json.load(f)
            self.username = file['Email Username']
            self.password = file['Email Password']
        self.connection = self.open_connection()
        self.process_emails()

    def open_connection(self):
        imap = imaplib.IMAP4_SSL(self.host)
        imap.login(self.username, self.password)
        imap.select('INBOX')
        return imap

    def process_emails(self):
        for subject in self.subjects:
            subject_data = self.subjects[subject]
            data = self.find_emails(subject)
            counter = 0
            for email_id in data[0].split():
                counter += 1
                content = self.fetch_email(email_id)
                file = self.get_attachment(content)
                self.save_attachment(file, subject_data, counter)
        self.close_connection()

    def find_emails(self, subject):
        result, data = self.connection.search(None, '(SUBJECT "' + subject + '")')
        if result == 'OK':
            return data

    def fetch_email(self, email_id):
        result, data = self.connection.fetch(email_id, '(RFC822)')
        if result == 'OK':
            content = email.message_from_bytes(data[0][1])
            if content['from'] in self.senders:
                return content

    def get_attachment(self, content):
        for part in content.walk():
            if part.get_content_maintype() == 'multipart':
                continue
            if part.get('Content-Disposition') is None:
                continue
            return part.get_payload(decode=True)

    def save_attachment(self, file, subject_data, counter):
        if not os.path.isdir(self.folder_path):
            os.mkdir(self.folder_path)
        if subject_data[0] == 'Pre-Advice Report':
            file_date = file[172: 198].decode('utf-8')
            file_name = subject_data[0] + '. ' + file_date + subject_data[1]
            file_path = os.path.join(self.folder_path, str(counter) + '. ' + file_date)
            if not os.path.isdir(file_path):
                os.mkdir(file_path)
            file_path = os.path.join(file_path, file_name)
        else:
            temp_folder = os.path.join(self.folder_path, 'temp')
            if not os.path.isdir(temp_folder):
                os.mkdir(temp_folder)
            temp_file = os.path.join(temp_folder, 'temp.pdf')
            with open(temp_file, 'wb') as f:
                f.write(file)
            with open(temp_file, 'rb') as f:
                pdf = PyPDF2.PdfFileReader(f)
                if pdf.isEncrypted:
                    pdf.decrypt('')
                content = pdf.getPage(0).extractText()
                if subject_data[0] == 'Self Billing Invoice':
                    file_date = content[330:334] + content[334:337].lower() + '20' + content[337:339] + ' to ' + \
                        content[353:357] + content[357:360].lower() + '20' + content[360:362]
                else:
                    file_date = content[354:379].replace(' ', '-').replace('---', ' to ')
                    file_dates = file_date.split(' to ')
            if os.path.isfile(temp_file):
                os.remove(temp_file)
            if os.path.isdir(temp_folder):
                os.rmdir(temp_folder)
            for folder in os.listdir(self.folder_path):
                stop = False
                if subject_data[0] == 'Self Billing Invoice':
                    if file_date.split(' to ')[1] in folder:
                        stop = True
                else:
                    for date in file_dates:
                        if date in folder:
                            stop = True
                file_name = subject_data[0] + '. ' + file_date + subject_data[1]
                file_path = os.path.join(self.folder_path, folder, file_name)
                if stop:
                    break
        with open(file_path, 'wb') as f:
            f.write(file)

    def close_connection(self):
        self.connection.close()
        self.connection.logout()


if __name__ == '__main__':
    ProcessInvoices()
