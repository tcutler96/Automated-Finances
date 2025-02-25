import os
from csv import reader
from utilities import get_base_path, Workbook


class ProcessIncome:
    def __init__(self, test=False):
        self.base_path = get_base_path(test=test)
        self.input_path = os.path.join(self.base_path, 'CutlerT Income Transactions')
        self.income_path = os.path.join(self.base_path, 'CutlerT Income.xlsx')
        self.header_data = [['Date', 15], ['Amount', 15]]
        self.incomes_recorded = 0

        self.process_income()

    def process_income(self):
        workbook = Workbook(self.income_path)
        folders = [[f.path, f.name] for f in os.scandir(self.input_path) if f.is_dir()]
        for folder_path, financial_year in folders:
            file_name = [f for f in os.listdir(folder_path) if f.startswith('Income Transactions') & f.endswith('.csv')]
            if file_name:
                file_name = file_name[0]
                if financial_year in workbook.workbook.sheetnames:
                    workbook.set_worksheet(financial_year)
                else:
                    workbook.add_worksheet(financial_year, self.header_data)
                    workbook.write_cell(row=2, col=1, value='Total', form='bold')
                income_data = []
                with open(os.path.join(folder_path, file_name), 'rt', encoding='utf8') as f:
                    for row in reader(f):
                        if row[0] != 'Date':
                            income_data.append([row[0].replace(' ', '-'), float(row[3])])
                update_total = False
                for row, data in enumerate(income_data):
                    row += 2
                    if data[0] != workbook.worksheet.cell(row=row, column=1).value:
                        workbook.write_row(row=row, values=[[data[0], 'right'], [data[1], 'money']], insert=True)
                        self.incomes_recorded += 1
                        update_total = True
                if update_total:
                    workbook.write_cell(row=row + 1, col=2, value=f'=SUM(B2:B{row})', form='money')
        workbook.save_workbook(self.income_path)
        if self.incomes_recorded:
            print(f'Successfully recorded {self.incomes_recorded} new incomes...')
        else:
            print('No new incomes found...')


if __name__ == '__main__':
    ProcessIncome(test=True)
