import openpyxl
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.worksheet.table import Table, TableStyleInfo
import pandas as pd


START_COLUMN, INDENT, START_ROW = 2, 4, 2
DATE_COLUMN = ['Дата', 'Дата документа', 'Получен']
HEADERS_FONT = Font(
        name='Calibri',
        size=16,
        bold=True,
        italic=False,
        vertAlign=None,
        underline='none',
        strike=False,
        color='FF000000'
            )

def decorating_excel_list(report_dict: dict, savename: str) -> None:
    wb = openpyxl.Workbook()
    sheet = wb.get_sheet_by_name(wb.sheetnames[0])
    sheet.title = 'Сверка'
    for title, res in report_dict.items():
        add_smart_table(sheet, res, title)
    set_max_width(report_dict, sheet)
    wb.save(savename)

def add_smart_table(ws, result: pd.DataFrame, title: str) -> None:
    global START_ROW
    result = edit_date_columns(result)
    ws.cell(row=START_ROW, column=START_COLUMN).value = title
    ws.cell(row=START_ROW, column=START_COLUMN).font = HEADERS_FONT
    length = len(result)
    if not result.empty:
        rows = dataframe_to_rows(result, index=False)
        for r_idx, row in enumerate(rows, START_ROW + 1):
            for c_idx, value in enumerate(row, START_COLUMN):
                ws.cell(row=r_idx, column=c_idx, value=value)

        length = len(result)
        col_length = len(result.columns) - 1
        table_range = f'{get_column_letter(START_COLUMN)}{START_ROW + 1}:{get_column_letter(START_COLUMN + col_length)}{START_ROW + length + 1}'
        table = Table(displayName=title.replace(' ', '_'), ref=table_range)
        style = TableStyleInfo(name="TableStyleMedium2", showFirstColumn=False,
                               showLastColumn=False, showRowStripes=True, showColumnStripes=True)
        table.tableStyleInfo = style
        ws.add_table(table)
    START_ROW += length + INDENT

def edit_date_columns(table: pd.DataFrame) -> pd.DataFrame:
    for col in DATE_COLUMN:
        if col in table.columns:
            table[col] = table[col].apply(str)
    return table


def set_max_width(report_dict: dict, ws) -> None:
    width_dict = dict()
    for title, report in report_dict.items():
        for i, column in enumerate(report.columns, START_COLUMN):
            data_list = [len(str(value)) for value in report[column]]
            data_list.append(len(column))
            max_width = max(data_list) + INDENT
            current_dict = width_dict.get(i, 1)
            if max_width > current_dict:
                width_dict[i] = max_width
    for column, width in width_dict.items():
        ws.column_dimensions[get_column_letter(column)].width = width


# jr = read_file(r"C:\Users\cyril\Desktop\Сверка НДС\Журнал авансовых счетов-фактур 4кв23.xlsx")
# dc = read_file(r"C:\Users\cyril\Desktop\Сверка НДС\Diadoc 14.12.23 19.08.csv", True)
# decorating_excel_list(jr, dc)