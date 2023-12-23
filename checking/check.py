import datetime
import re

import numpy as np
import pandas as pd

from loading.load import read_file, read_report, parse_data
from logs import log

COLUMNS = ['Дата', 'Номер', 'Сумма', 'Валюта', 'Контрагент', 'Вид счета-фактуры', 'Организация']
DOC_COLUMN = 'Документ-основание'
RECEIVD_COL = 'Получен'
DATE_COL = 'Дата'
NUMBER_COL = 'Номер'
JOURNAL_COLUMNS = ['ИНН', 'Номер', 'Дата', 'Сумма']
DIADOC_COLUMNS = ['ИНН', 'Номер документа', 'Дата документа', 'Всего']
TITLE_TABLES = ['Проверка на наличие дублей', 'Ошибки в датах', 'Некорректные номера документов', 'Проверка дубликатов по столбцу Документ Основание']
COMPARING_TITLE = 'Сверка данных из журнала и Диадок'
COLUMNS_FOR_GROUPS = ['Субподрядчик_ext','Сумма НДС', 'Сумма незакрытого аванса']
COLUMNS_FOR_REGROUPS = ['Проект', 'Контрагент', 'Документ', 'Вх. номер', 'Вх. дата', 'Сумма']
TITLE_TABLES_FOR_REPORT = ['REPORT_Сумма НДС', 'REPORT_ТОП10', 'REPORT_Анализ', 'REPORT_Средняя разница', 'REPORT_Перегрупировка']
FACTURE_COLUMN = 'СФ'
VAT_COLUMN = 'Сумма НДС'
def duplicate_check(frame:pd.DataFrame) -> pd.DataFrame:
    log.info('Вызов функкции обнаружения дубликатов')
    return frame[frame.duplicated(COLUMNS)]

def dates_check(frame: pd.DataFrame) -> pd.DataFrame:
    log.info('Вызов функкции сверки дат')
    return frame[frame[RECEIVD_COL]<frame[DATE_COL]]

def number_check(frame: pd.DataFrame) ->pd.DataFrame:
    log.info('Вызов функкции сверки корректного наименования номера договора')
    return frame[(frame[NUMBER_COL].str.contains('F'))| (~frame[NUMBER_COL].str.contains('0|1|2|3|4|5|6|7|8|9'))]

def duplicate_document_check(frame: pd.DataFrame) -> pd.DataFrame:
    log.info('Вызов функкции обнаружения дубликатов в документах')
    return frame[frame.duplicated(DOC_COLUMN)]

def structure_analysis_table(frame: pd.DataFrame) -> pd.DataFrame:
    log.info('Создание таблицы разбиения НДС')
    amount = sum(frame[VAT_COLUMN])
    VAT_n = sum(frame[pd.isna(frame[FACTURE_COLUMN])][VAT_COLUMN])
    VAT_y = sum(frame[~pd.isna(frame[FACTURE_COLUMN])][VAT_COLUMN])
    f_row = {'Тип': 'Сумма НДС по непредоставленным документам', 'Сумма': round(VAT_n/amount, 2)}
    s_row = {'Тип': 'Сумма НДС по предоставленным документам', 'Сумма': round(VAT_y/amount, 2)}
    res = pd.DataFrame([f_row, s_row])
    return res

def get_grouped_data(frame: pd.DataFrame) -> pd.DataFrame:
    log.info('Группировка данных')
    sep_frame = frame[pd.isna(frame[FACTURE_COLUMN])][COLUMNS_FOR_GROUPS]
    group_data = sep_frame.groupby(COLUMNS_FOR_GROUPS[0], as_index=False).agg(sum)
    group_data.sort_values(by=COLUMNS_FOR_GROUPS[1], ascending=False, inplace=True)
    return group_data

def describe_table(frame: pd.DataFrame) -> pd.DataFrame:
    log.info('Создание таблицы описания данных по авансам')
    VAT = sum(frame[pd.isna(frame[FACTURE_COLUMN])][VAT_COLUMN])
    ADVANCE = sum(frame[pd.isna(frame[FACTURE_COLUMN])]['Сумма незакрытого аванса'])
    COUNT = len(frame[pd.isna(frame[FACTURE_COLUMN])])
    ALL_ADVANCE = sum(frame['Сумма незакрытого аванса'])
    f_row = {'Тип': 'Сумма незакрытого аванса', 'Сумма': round(ADVANCE/1_000_000_000, 1)}
    s_row = {'Тип': 'Сумма НДС с незакрытого аванса', 'Сумма': round(VAT / 1_000_000, 1)}
    t_row = {'Тип': 'Количество непредоставленных документов', 'Сумма': COUNT}
    fr_row = {'Тип': 'Сумма авансирования', 'Сумма': round(ALL_ADVANCE / 1_000_000_000, 1)}
    result = pd.DataFrame([f_row, s_row, t_row, fr_row])
    return result
def get_only_n_values(frame: pd.DataFrame,  n: int=10) -> pd.DataFrame:
    log.info('Создание таблицы ТОП 10')
    frame = add_column(frame, 'Субподрядчик_ext', 'Субподрядчик', '(.*)\(\d*\/\d*\)')
    frame = get_grouped_data(frame).iloc[:n, :]
    for col in COLUMNS_FOR_GROUPS[1:]:
        frame[col] = frame[col].astype(float)
        frame[col] = np.around(frame[col]/1_000_000,1)
    frame.sort_values(by=COLUMNS_FOR_GROUPS[1], ascending=True, inplace=True)
    return frame

def regroup_data(frame: pd.DataFrame) -> pd.DataFrame:
    return frame[COLUMNS_FOR_REGROUPS].groupby(by=COLUMNS_FOR_REGROUPS[:-1], as_index=False).agg(sum)


def extract_value(string: str, pattern) -> str:
    # pattern = '\((\d+)\/\d+\)'
    matches = re.search(pattern, string)
    if matches != None:
        return str(matches.group(1)).strip()
    else:
        return string

def add_column(frame: pd.DataFrame, new_col, old_col, pattern) -> pd.DataFrame:
    frame[new_col] = frame[old_col].apply(extract_value, args=[pattern])
    return frame

def compare_tables(journal: pd.DataFrame, diadoc: pd.DataFrame) -> pd.DataFrame:
    log.info('Сверки данных журнала и диадок')
    journal = add_column(journal, 'ИНН', 'Контрагент', '\((\d+)\/\d+\)')
    journal = journal[JOURNAL_COLUMNS]
    diadoc = diadoc[DIADOC_COLUMNS]
    compare_result = pd.merge(journal, diadoc, left_on=JOURNAL_COLUMNS, right_on=DIADOC_COLUMNS)
    return compare_result

def calc_date_diff(dt: datetime.datetime, current_date: datetime.datetime) -> int:
    diff = current_date - dt
    return diff.days

def get_avg_date(frame: pd.DataFrame, curr_dt: datetime.datetime) -> pd.DataFrame:
    log.info('Вызов функции получения средней даты непредоставления документов')
    frame = add_column(frame, 'Agent_dt', 'Документ', '(\d{2}\.\d{2}\.\d{4})')
    frame[frame['Agent_dt'].str.contains('Списание')].loc[:, 'Agent_dt'] = frame[frame['Agent_dt'].str.contains('Списание')].loc[:, 'Вх. дата']
    frame['Agent_dt'] = frame['Agent_dt'].apply(parse_data)
    frame['date_diff'] = frame['Agent_dt'].apply(calc_date_diff, args=[curr_dt])
    mean_dt = frame[pd.isna(frame[FACTURE_COLUMN])]['date_diff'].mean()
    sum_VAT = round(frame[pd.isna(frame[FACTURE_COLUMN])][VAT_COLUMN].sum()/1_000_000,2)
    mean_dt = int(round(mean_dt))
    str_dict = {'Дата': curr_dt.strftime('%d.%m.%Y'), 'Разница': mean_dt, VAT_COLUMN: sum_VAT}
    res = pd.DataFrame([str_dict])
    return res
def create_all_report(journal_path: str=None, diadoc_path: str=None, report_path: str=None, curr_dt:datetime.datetime=None,  opt: str=None) -> dict:
    log.info('Создание словаря документов')
    frame_dict = dict()
    if opt in ('only_errors', 'all'):
        diadoc, journal = read_file(diadoc_path, True), read_file(journal_path)
        FUNCTION_LIST = [duplicate_check, dates_check, number_check, duplicate_document_check]
        for key, func in zip(TITLE_TABLES, FUNCTION_LIST):
            frame_dict[key] = func(journal)
        frame_dict[COMPARING_TITLE] = compare_tables(journal, diadoc)
    if opt in ('only_advances', 'all'):
        report = read_report(report_path)
        FUNCTION_LIST = [describe_table, get_only_n_values, structure_analysis_table, get_avg_date, regroup_data]
        for key, func in zip(TITLE_TABLES_FOR_REPORT, FUNCTION_LIST):
            if key != 'REPORT_Средняя разница':
                frame_dict[key] = func(report)
            else:
                frame_dict[key] = func(report, curr_dt)
    return frame_dict

