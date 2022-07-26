import xlsxwriter
import PyPDF2
import json
import os


class ProcessIncome:
    def __init__(self, test=True):
        if test:
            self.base_folder = os.path.dirname(os.path.abspath(__file__))
        else:
            self.base_folder = 'C:\\Users\Tom\Important\Self Employed\Yodel'
        self.payment_path = os.path.join(self.base_folder, 'Natwest Transactions 2021-22.pdf')
        self.payment_name = 'YDN PAYMENTS'
        self.parcel_folder = os.path.join(self.base_folder, 'Invoices 2021-22')
        self.parcel_file_name = 'Self Billing Invoice'
        self.output_path = os.path.join(self.base_folder, 'CutlerT Income 2021-22.xlsx')
        self.output_header = [['Date Paid', 12], ['Amount', 10], ['Sales', 10], ['Insurance', 10], ['Adjustments', 12], ['Dates Worked', 24], ['1st Parcels', 12], ['2nd Parcels', 12], ['Total Parcels', 12]]
        self.insurance_amount = 6
        self.income_data = {}
        with open('credentials.json', 'r') as f:
            file = json.load(f)
            self.customer_number = file['Bank Customer Number']
            self.pin = file['Bank Pin']
            self.password = file['Bank Password']
        self.process_income()

    def process_income(self):
        self.get_payment_data()
        self.get_parcel_data()
        self.write_output_file()

    def get_payment_data(self):
        with open(self.payment_path, 'rb') as f:
            pdf, pdf_pages = self.read_pdf(f)
            counter = 0
            for pdf_page in range(pdf_pages - 1, -1, -1):
                content = pdf.getPage(pdf_page).extractText()
                payments = [i for i in range(len(content)) if content.startswith(self.payment_name, i)]
                payments.reverse()
                for payment in payments:
                    if content[payment - 2] == 'C':
                        date = content[payment - 16:payment - 5].replace('\n', '0').replace(' ', '-')
                        amount = float(content[payment + 14:payment + 22].replace(',', '').replace('-', '').replace('\n', ''))
                        self.income_data[counter] = [date, amount]
                        counter += 1

    def get_parcel_data(self):
        folders = os.listdir(self.parcel_folder)
        for folder in folders:
            for file in os.listdir(os.path.join(self.parcel_folder, folder)):
                if file.startswith(self.parcel_file_name):
                    with open(os.path.join(self.parcel_folder, folder, file), 'rb') as f:
                        pdf, pdf_pages = self.read_pdf(f)
                        parcels = []
                        adjustments = 0
                        for pdf_page in range(pdf_pages):
                            content = pdf.getPage(pdf_page).extractText()
                            parcel_index = content.find('Parcel Stop Total')
                            if parcel_index > 0:
                                parcels.append(int(content[parcel_index + 23:parcel_index + 26]))
                            if pdf_page == pdf_pages - 1:
                                total_index = content.find('Parcel Stops')
                                total = float(content[total_index + 30:total_index + 38].replace(' ', '').replace(',', ''))
                                adjustment_index = content.find('Manual Adjustments')
                                adjustments += float(content[adjustment_index + 32:adjustment_index + 38].replace(' ', ''))
                        for key, value in self.income_data.items():
                            if value[1] == total + adjustments - self.insurance_amount:
                                self.income_data[key] = self.income_data[key] + [total, -self.insurance_amount, adjustments, file[22:46], parcels[0], parcels[1], sum(parcels)]

    def read_pdf(self, f):
        pdf = PyPDF2.PdfFileReader(f)
        if pdf.isEncrypted:
            pdf.decrypt('')
        pdf_pages = pdf.getNumPages()
        return pdf, pdf_pages

    def write_output_file(self):
        workbook = xlsxwriter.Workbook(self.output_path)
        worksheet = workbook.add_worksheet()
        bold = workbook.add_format({'bold': True})
        right = workbook.add_format({'align': 'right'})
        for col, header in enumerate(self.output_header):
            worksheet.write(0, col, header[0], bold)
            worksheet.set_column(col, col + 1, header[1])
        for row, data in enumerate(self.income_data):
            for col, cell in enumerate(self.income_data[data]):
                worksheet.write(row + 1, col, cell, bold if col == 0 else right)
        workbook.close()


if __name__ == '__main__':
    ProcessIncome()
