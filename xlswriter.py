from logging import exception

import xlsxwriter

data1 = [
    ["Value", 10000, 5000, 8000, 6000],
    ["Pears", 2000, 3000, 4000, 5000],
    ["Bananas", 6000, 6000, 6500, 6000],
    ["Oranges", 500, 300, 200, 700],
]

class XlsWriter:
    def __init__(self, path):
        if path.endswith('xlsx'):
            self.path = path
    def create_tab(self, data):
        try:
            with xlsxwriter.Workbook(str(self.path)) as workbook:
                worksheet1 = workbook.add_worksheet()
                caption = "Default table with data."

                # Set the columns widths.
                worksheet1.set_column("B:G", 12)

                # Write the caption.
                worksheet1.write("B1", caption)

                # Add a table to the worksheet.
                a = len(data)
                b = len(data[0])
                # initial position
                start_l = "B"
                start_d = 3

                first_row = start_d - 1
                first_col = ord(start_l) - ord('A')
                last_row = first_row + b - 1
                last_col = first_col + a - 1

                worksheet1.add_table(first_row, first_col, last_row, last_col, {"data": data})
        except Exception as e:
            print(f"Nie udalo sie utworzyc tabeli prosze zamknac plik: {self.path}")

bot = XlsWriter("tabelka.xlsx")
bot.create_tab(data1)

