from logging import exception
from pathlib import Path
import xlsxwriter

class XlsWriter:
    def __init__(self, path):
        p = Path(path)
        if p.suffix.lower() != ".xlsx":
            p = p.with_suffix(".xlsx")
        p.parent.mkdir(parents=True, exist_ok=True)
        self.path = p

    def create_tab(self, data):
        if not data or not data[0]:
            raise ValueError("Pusta tabela: brak danych lub kolumn.")

        n_rows = len(data)        # liczba WIERSZY
        n_cols = len(data[0])     # liczba KOLUMN

        start_l = "B"
        start_d = 3
        first_col = ord(start_l) - ord('A')
        first_row = start_d - 1
        last_row = first_row + n_rows
        last_col = first_col + n_cols - 1

        try:
            with xlsxwriter.Workbook(str(self.path)) as wb:
                ws = wb.add_worksheet()
                ws.set_column(first_col, last_col, 12)
                ws.write("B1", "Result.")

                cols = [{"header": "Products"}, {"header": "Amount"}]

                ws.add_table(first_row, first_col, last_row, last_col, {
                    "data": data,
                    "columns": cols,
                })
        except PermissionError:
            print(f"Plik jest otwarty w Excelu, zamknij go: {self.path}")
        except Exception as e:
            print(f"Błąd tworzenia tabeli: {e}")