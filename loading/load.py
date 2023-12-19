from datetime import datetime
import chardet as chardet
import pandas as pd

filename = r"C:\Users\cyril\Desktop\Сверка НДС\Журнал авансовых счетов-фактур 4кв23.xlsx"
diadoc = r"C:\Users\cyril\Desktop\Сверка НДС\1Diadoc 14.12.23 19.08.csv"
report = r"C:\Users\cyril\Desktop\Сверка НДС\Отчет по незакрытым авансам.xlsx"
DEFAULT_FORMAT = '.xlsx'
def parse_data(date_str):
    return datetime.strptime(date_str, '%d.%m.%Y')

def get_encoding(filename: str) -> str:
    enc = chardet.detect(open(filename, 'rb').read())
    return enc['encoding']

def read_file(filename, is_csv=False):
    sep = ';'
    if not is_csv:
        df = pd.read_excel(filename)
        df[['Получен', 'Дата']] = df[['Получен', 'Дата']].apply(pd.to_datetime, format='%d.%m.%Y')
    else:
        encode = get_encoding(filename)
        df = pd.read_csv(filename, sep=sep, encoding=encode)
        df['Дата документа'] = df['Дата документа'].apply(pd.to_datetime, format='%d.%m.%Y')
        for column in ['ИНН', 'Номер документа']:
            df[column] = df[column].apply(lambda x: str(x).replace('=', '').replace('"', ''))
        df['Всего'] = df['Всего'].apply(lambda x: float(str(x).replace(' ', '').replace(',', '.')))
    return df

def create_save_filename(path: str, name: str, format=DEFAULT_FORMAT) -> str:
    filename = f'{path}\\{name}{DEFAULT_FORMAT}'
    return filename


# dc = read_file(r"C:\Users\cyril\Desktop\Сверка НДС\Diadoc 14.12.23 19.08.csv", True)
# print(dc)