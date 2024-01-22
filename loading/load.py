from datetime import datetime
import chardet as chardet
import pandas as pd
import numpy as np

from logs import log

report = r"C:\Users\cyril\Desktop\Сверка НДС\Отчет по незакрытым авансам Жилстрой-МО.xlsx"
DEFAULT_FORMAT = '.xlsx'
COLUMNS = ['Документ', 'Вх. номер', 'Вх. дата', 'Сумма', 'СФ',	'Проведен',	'Пометка удаления', 'Договор субподрядчика', 'Проект',
           'Контрагент', 'Договор', 'Технический заказчик', 'Сумма незакрытого аванса']
CORRECT_ORDER_COLUMNS = ['Субподрядчик', 'Документ', 'Вх. номер', 'Вх. дата', 'Сумма','СФ',	'Проведен',	'Пометка удаления', 'Договор субподрядчика', 'Проект',
           'Контрагент', 'Договор', 'Технический заказчик', 'Сумма незакрытого аванса', 'Сумма НДС']
ADVACNE_COL = 'Сумма незакрытого аванса'
VAT_COL = 'Сумма НДС'
PERCENT = 0.8
def parse_data(date_str):
    return datetime.strptime(str(date_str), '%d.%m.%Y')
def date_to_str(dt):
    return dt.strftime('%d.%m.%Y')

def get_encoding(filename: str) -> str:
    enc = chardet.detect(open(filename, 'rb').read())
    return enc['encoding']
def read_data(filename, is_csv=False):
    sep = ';'
    if not is_csv:
        df = pd.read_excel(filename)
    else:
        encode = get_encoding(filename)
        df = pd.read_csv(filename, sep=sep, encoding=encode)
    return df
def read_file(filename, is_csv=False):
    df = read_data(filename, is_csv)
    if not is_csv:
        log.info('Считывание данных журнала с/ф')
        df[['Получен', 'Дата']] = df[['Получен', 'Дата']].apply(pd.to_datetime, format='%d.%m.%Y')
    else:
        log.info('Считывание данных файла диадок')
        df['Дата документа'] = df['Дата документа'].apply(pd.to_datetime, format='%d.%m.%Y')
        for column in ['ИНН', 'Номер документа']:
            df[column] = df[column].apply(lambda x: str(x).replace('=', '').replace('"', ''))
        df['Всего'] = df['Всего'].apply(lambda x: float(str(x).replace(' ', '').replace(',', '.')))
    return df

def create_save_filename(path: str, name: str, format=DEFAULT_FORMAT) -> str:
    filename = f'{path}\\{name}{format}'
    return filename

def read_report(filename:str) -> pd.DataFrame:
    log.info('Считывание данных по авансам')
    raw_data = pd.read_excel(filename)
    raw_data = drop_na_columns(raw_data)
    raw_data.columns = COLUMNS
    drop_rows = []
    raw_data['Субподрядчик'], subagent = '', ''
    for i in range(len(raw_data)):
        if not str(raw_data.iloc[i, 0]).__contains__('Списание'):
            subagent = str(raw_data.iloc[i, 0])
        count_na_values = sum(raw_data.iloc[i].isna())
        percent = round(count_na_values / len(raw_data.iloc[i]))
        if percent >= PERCENT:
            drop_rows.append(i)
        raw_data['Субподрядчик'].iloc[i] = subagent
    raw_data.drop(drop_rows, axis=0, inplace=True)
    edit_df = raw_data[raw_data['Документ'].str.contains('Списание')]
    edit_df[VAT_COL] = edit_df[ADVACNE_COL]*20/120
    edit_df[VAT_COL] = edit_df[VAT_COL].apply(float)
    edit_df[VAT_COL] = np.around(edit_df[VAT_COL],2)
    edit_df = edit_df[CORRECT_ORDER_COLUMNS]
    return edit_df


def drop_na_columns(df: pd.DataFrame) -> pd.DataFrame:
    drop_columns = []
    for col in df.columns:
        na_count = sum(df.loc[:, col].isna())
        percent = round(na_count/len(df))
        if percent>=PERCENT:
            drop_columns.append(col)
    df.drop(drop_columns, axis=1, inplace=True)
    return df

def load_historic_data(filename: str) -> pd.DataFrame:
    log.info('Считывание исторических данных')
    df = pd.read_excel(filename)
    df['Дата'] = df['Дата'].apply(pd.to_datetime, format='%d.%m.%Y')
    df['Дата'] = df['Дата'].apply(date_to_str)
    return df


def add_row_to_frame(hist_data: pd.DataFrame, add_data: pd.DataFrame, filename:str) -> pd.DataFrame:
    log.info('Добавление новой строки в исторические данные')
    new_data = pd.concat([hist_data, add_data])
    new_data.to_excel(filename, index=False)
    return new_data

# dc = read_report(report)
# print(dc)