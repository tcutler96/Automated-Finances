import os
from io import BytesIO
from PyPDF2 import PdfReader
from imaplib import IMAP4, IMAP4_SSL
from email import message_from_bytes
from utilities import get_base_path, get_email_credentials, get_financial_year, get_fortnights


class ProcessInvoices:
    def __init__(self, test=False):
        self.base_path = get_base_path(test=test)
        self.invoice_path = os.path.join(self.base_path, 'CutlerT Invoices')
        self.host = 'imap-mail.outlook.com'
        self.email_username, self.email_password = get_email_credentials()
        self.subjects = {'YODEL Pre-Advice Document': ['Pre-Advice Report', '.csv'],
                         'YODEL Self-Billed Invoice': ['Self Billing Invoice', '.pdf'],
                         'Yodel Third Party Billing': ['Insurance Invoice', '.pdf']}
        self.senders = ['equitas.billing@yodel.co.uk', '<Equitas_Billing@Yodel.co.uk>']
        self.invoices_saved = 0

        self.connection = self.open_connection()
        if self.connection:
            self.process_invoices()
            self.close_connection()

    def open_connection(self):
        imap = IMAP4_SSL(self.host)
        try:
            imap.login(self.email_username, self.email_password)
        except IMAP4.error:
            print('Failed to login to email...')
            return None
        imap.select('INBOX')
        return imap

    def close_connection(self):
        self.connection.close()
        self.connection.logout()

    def process_invoices(self):
        for subject in self.subjects:
            subject_data = self.subjects[subject]
            email_ids = self.find_emails(subject)
            for email_id in email_ids:
                content = self.fetch_email(email_id)
                file = self.get_attachment(content)
                if self.save_attachment(file, subject_data):
                    break
        if self.invoices_saved:
            print(f'Successfully saved {self.invoices_saved} new invoice{"s" if self.invoices_saved > 1 else ""}...')
        else:
            print('No new invoices found...')

    def find_emails(self, subject):
        result, email_ids = self.connection.search(None, '(SUBJECT "' + subject + '")')
        if result == 'OK':
            email_ids = email_ids[0].split()
            email_ids.reverse()
            return email_ids

    def fetch_email(self, email_id):
        result, data = self.connection.fetch(message_set=email_id, message_parts='(RFC822)')
        if result == 'OK':
            content = message_from_bytes(data[0][1])
            if content['from'] in self.senders:
                return content

    def get_attachment(self, content):
        for part in content.walk():
            if part.get_content_maintype() == 'multipart':
                continue
            if part.get('Content-Disposition') is None:
                continue
            return part.get_payload(decode=True)

    def save_attachment(self, file, subject_data):
        if subject_data[0] == 'Pre-Advice Report':
            file_date = file[172: 198].decode('utf-8')
            file_date_end = file_date.split(' to ')[1]
            financial_year, financial_year_start, _ = get_financial_year(file_date_end)
            file_path = os.path.join(self.invoice_path, financial_year)
            if not os.path.isdir(file_path):
                os.mkdir(file_path)
            fortnights = get_fortnights(financial_year_start, file_date_end)
            file_path = os.path.join(file_path, str(fortnights) + '. ' + file_date)
            if not os.path.isdir(file_path):
                os.mkdir(file_path)
            file_name = subject_data[0] + ' ' + file_date + subject_data[1]
            file_path = os.path.join(file_path, file_name)
        else:
            pdf = PdfReader(BytesIO(file))
            content = pdf.pages[0].extract_text()
            if subject_data[0] == 'Self Billing Invoice':
                file_date = content[content.find('From:') + 6:content.find('From:') + 15] + ' ' + content[content.find('To:') + 4:content.find('To:') + 13]
                file_date = file_date[0:4] + file_date[4:6].lower() + '-20' + file_date[7:9] + ' to ' + file_date[10:14] + file_date[14:16].lower() + '-20' + file_date[17:]
            else:
                file_date = content[content.find('period') + 7:content.find('period') + 32].replace(' ', '-').replace('---', ' to ')
            file_date_end = file_date.split(' to ')[1]
            financial_year, financial_year_start, _ = get_financial_year(file_date_end)
            stop = False
            for folder in os.listdir(os.path.join(self.invoice_path, financial_year)):
                for date in file_date.split(' to '):
                    if date in folder:
                        stop = True
                if stop:
                    file_name = subject_data[0] + ' ' + file_date + subject_data[1]
                    file_path = os.path.join(self.invoice_path, financial_year, folder, file_name)
                    break
        if os.path.isfile(file_path):
            return True
        with open(file_path, 'wb') as f:
            f.write(file)
            self.invoices_saved += 1


if __name__ == '__main__':
    ProcessInvoices(test=True)
