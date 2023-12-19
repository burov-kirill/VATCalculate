import re

import pandas as pd


COLUMNS = ['Дата', 'Номер', 'Сумма', 'Валюта', 'Контрагент', 'Вид счета-фактуры', 'Организация']
DOC_COLUMN = 'Документ-основание'
RECEIVD_COL = 'Получен'
DATE_COL = 'Дата'
NUMBER_COL = 'Номер'
JOURNAL_COLUMNS = ['ИНН', 'Номер', 'Дата', 'Сумма']
DIADOC_COLUMNS = ['ИНН', 'Номер документа', 'Дата документа', 'Всего']
TITLE_TABLES = ['Проверка на наличие дублей', 'Ошибки в датах', 'Некорректные номера документов', 'Проверка дубликатов по столбцу Документ Основание']
COMPARING_TITLE = 'Сверка данных из журнала и Диадок'
def duplicate_check(frame:pd.DataFrame) -> pd.DataFrame:
    return frame[frame.duplicated(COLUMNS)]

def dates_check(frame: pd.DataFrame) -> pd.DataFrame:
    return frame[frame[RECEIVD_COL]<frame[DATE_COL]]

def number_check(frame: pd.DataFrame) ->pd.DataFrame:
    return frame[(frame[NUMBER_COL].str.contains('F'))| (~frame[NUMBER_COL].str.contains('0|1|2|3|4|5|6|7|8|9'))]

def duplicate_document_check(frame: pd.DataFrame) -> pd.DataFrame:
    return frame[frame.duplicated(DOC_COLUMN)]

def extract_VAT(string: str) -> str:
    pattern = '\((\d+)\/\d+\)'
    matches = re.search(pattern, string)
    if matches != None:
        return matches.group(1)
    else:
        return string
def add_VAT_column(frame: pd.DataFrame) -> pd.DataFrame:
    frame['ИНН'] = frame['Контрагент'].apply(extract_VAT)
    return frame

def compare_tables(journal: pd.DataFrame, diadoc: pd.DataFrame) -> pd.DataFrame:
    journal = add_VAT_column(journal)
    journal = journal[JOURNAL_COLUMNS]
    diadoc = diadoc[DIADOC_COLUMNS]
    compare_result = pd.merge(journal, diadoc, left_on=JOURNAL_COLUMNS, right_on=DIADOC_COLUMNS)
    return compare_result

def create_all_report(journal: pd.DataFrame, diadoc: pd.DataFrame) -> dict:
    FUNCTION_LIST = [duplicate_check, dates_check, number_check, duplicate_document_check]
    frame_dict = dict()
    for key, func in zip(TITLE_TABLES, FUNCTION_LIST):
        frame_dict[key] = func(journal)
    frame_dict[COMPARING_TITLE] = compare_tables(journal, diadoc)
    return frame_dict
