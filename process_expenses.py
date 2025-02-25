import os
from datetime import datetime
from utilities import get_base_path, copy_file, get_financial_year, Workbook


class ProcessExpenses:
    def __init__(self, new_expenses=None, reset=None, test=False):
        self.base_path = get_base_path(test=test)
        self.expense_name = 'CutlerT Expenses.xlsx'
        self.expense_path = os.path.join(self.base_path, self.expense_name)
        self.header_data = [['Date', 15], ['Amount', 15], ['Reason', 15], ['Note', 15]]
        self.copy_path = os.path.join(self.base_path, 'CutlerT Expenses Copies')
        self.new_expenses = new_expenses
        self.reset = reset

        self.process_expenses()

    def process_expenses(self):
        if self.reset:
            copy_file(os.path.join(self.copy_path, self.expense_name.split('.')[0] + f' ({self.reset}).' + self.expense_name.split('.')[1]), self.expense_path)
            print(f'Successfully restored a previous expense file...')
        else:
            copy_file(self.expense_path, os.path.join(self.copy_path, self.expense_name.split('.')[0] + f' ({len(os.listdir(self.copy_path)) + 1}).' + self.expense_name.split('.')[1]))
            print(f'Successfully made a copy of current expense file...')
            if self.new_expenses:
                workbook = Workbook(self.expense_path)
                for new_expense in self.new_expenses:
                    if len(new_expense) < 4:
                        new_expense.append(None)
                    expense_date = datetime.strptime(new_expense[0], '%d/%m/%y').strftime('%d-%b-%Y')
                    financial_year = get_financial_year(expense_date)[0]
                    if financial_year in workbook.workbook.sheetnames:
                        workbook.set_worksheet(financial_year)
                    else:
                        workbook.add_worksheet(financial_year, self.header_data)
                        workbook.write_cell(row=2, col=1, value='Total', form='bold')
                    for row, values in enumerate(workbook.worksheet.iter_rows(min_row=2)):
                        row += 2
                        value = values[0].value
                        if value != 'Total':
                            if datetime.strptime(expense_date, '%d-%b-%Y') < datetime.strptime(value, '%d-%b-%Y'):
                                break
                        else:
                            break
                    workbook.write_row(row=row, values=[[expense_date, 'right'], [new_expense[1], 'money'], [new_expense[2], 'right'], [new_expense[3], None]], insert=True)
                for sheetname in workbook.workbook.sheetnames:
                    workbook.set_worksheet(sheetname)
                    for row, values in enumerate(workbook.worksheet.iter_rows(min_row=2)):
                        row += 2
                        value = values[0].value
                        if value == 'Total':
                            workbook.write_cell(row=row, col=2, value=f'=SUM(B2:B{row - 1})', form='money')
                workbook.save_workbook(self.expense_path)
                print(f'Successfully recorded {len(self.new_expenses)} new expense{"s" if len(self.new_expenses) > 1 else ""}...')


if __name__ == '__main__':
    # ProcessExpenses(new_expenses=[[date (dd/mm/yy) {str}, amount {int/float}, reason {str}, note (optional) {str}]])
    # ProcessExpenses(new_expenses=[['15/01/25', 37.31, 'Fuel']])
    # ProcessExpenses(reset=1)
    ProcessExpenses(test=True)
