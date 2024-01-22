import datetime
import itertools
import re
from collections import Counter

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
RES_COLUMNS = ['ИНН', 'Номер', 'Дата', 'Сумма', 'Номер документа', 'Дата документа', 'Всего', 'Ссылка']
TITLE_TABLES = ['Проверка на наличие дублей', 'Ошибки в датах', 'Некорректные номера документов', 'Проверка дубликатов по столбцу Документ Основание']
COMPARING_TITLE = 'Запросы на аннулирование'
COLUMNS_FOR_GROUPS = ['Субподрядчик_ext','Сумма НДС', 'Сумма незакрытого аванса']
COLUMNS_FOR_REGROUPS = ['Проект', 'Контрагент', 'Документ', 'Вх. номер', 'Вх. дата', 'Субподрядчик', 'Договор субподрядчика', 'Договор',  'Сумма']
TITLE_TABLES_FOR_REPORT = ['REPORT_Сумма НДС', 'REPORT_ТОП10', 'REPORT_Анализ', 'REPORT_Средняя разница', 'REPORT_Перегрупировка', 'REPORT_Все контрагенты']
FACTURE_COLUMN = 'СФ'
VAT_COLUMN = 'Сумма НДС'
INSPECTIONS_DICT = {
        'Камеральная, Встречная': 'Камеральная, по поручению',
        'Камеральная': 'Камеральная, прямая',
        'Вне рамок проверок, Встречная': 'Вне рамок проверок, прямое',
        'Выездная, Встречная': 'Выездная, по поручению',
        'Вне рамок проверок': 'Вне рамок проверок, прямое',
        'Встречная': 'Прочее',
        'Выездная': 'Выездная, по поручению',
        'Вне рамок проверок, Камеральная': 'Вне рамок проверок, прямое',
        'Камеральная, Вне рамок проверок': 'Камеральная, прямая',
        'Выездная, прямая': 'Прочее'

}

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
    ALL_ADVANCE = sum(frame['Сумма незакрытого аванса'])*20/120
    f_row = {'Тип': 'Сумма незакрытого аванса', 'Сумма': round(ADVANCE/1_000_000_000, 1)}
    s_row = {'Тип': 'Сумма НДС с незакрытого аванса', 'Сумма': round(VAT / 1_000_000, 1)}
    t_row = {'Тип': 'Количество непредоставленных документов', 'Сумма': COUNT}
    fr_row = {'Тип': 'Сумма авансирования', 'Сумма': round(ALL_ADVANCE / 1_000_000_000, 1)}
    result = pd.DataFrame([f_row, s_row, t_row, fr_row])
    return result
def get_only_n_values(frame: pd.DataFrame,  n: int=10) -> pd.DataFrame:
    log.info('Создание таблицы ТОП 10')
    frame = add_column(frame, 'Субподрядчик_ext', 'Субподрядчик', '(.*)\(\d*\/\d*\)')
    if n == 10:
        frame = get_grouped_data(frame).iloc[:n, :]
    else:
        frame = get_grouped_data(frame)
    for col in COLUMNS_FOR_GROUPS[1:]:
        frame[col] = frame[col].astype(float)
        frame[col] = np.around(frame[col]/1_000_000,1)
    if n == 10:
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
    compare_result = pd.merge(journal, diadoc, left_on=JOURNAL_COLUMNS, right_on=DIADOC_COLUMNS)
    compare_result = compare_result[RES_COLUMNS]
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
        FUNCTION_LIST = [describe_table, get_only_n_values, structure_analysis_table, get_avg_date, regroup_data, get_only_n_values]
        for key, func in zip(TITLE_TABLES_FOR_REPORT, FUNCTION_LIST):
            if key != 'REPORT_Средняя разница' and key!='REPORT_Все контрагенты':
                frame_dict[key] = func(report)
            elif key == 'REPORT_Все контрагенты':
                frame_dict[key] = func(report, 1)
            else:
                frame_dict[key] = func(report, curr_dt)
    return frame_dict

def get_quarter(month: int) -> int:
    if month in (1, 2, 3):
        return 1
    elif month in (4, 5, 6):
        return 2
    elif month in (7, 8, 9):
        return 3
    else:
        return 4
def get_quarter_from_string(string: str) -> str:
    dt = datetime.datetime.strptime(string, '%d.%m.%y')
    if dt.year>=2021:
        year = str(dt.year)[2:]
        quarter = get_quarter(dt.month)
        return f'{quarter}кв{year}г.'
    else:
        return f'{dt.year}г.'

def get_year_from_string(string: str) -> int:
    dt = datetime.datetime.strptime(string, '%d.%m.%y')
    return dt.year
def get_quarter_report(frame: pd.DataFrame, current_year: int) -> pd.DataFrame:
    skip_years = ['2014г.', '2015г.', '2016г.']
    frame['Квартал'] = frame['Дата получения'].apply(get_quarter_from_string)
    frame = frame[~frame['Квартал'].isin(skip_years)]
    quarter_frame = frame[['Квартал']].value_counts().to_frame()
    quarter_frame.reset_index(inplace=True)
    quarter_frame.columns = ['Квартал', 'Количество документов от ИФНС']
    quarter_frame.sort_values(by='Квартал', inplace=True, key=lambda col: col.map(lambda x: (x[-4:-2], x[0])))
    return quarter_frame

def rename_docs(doc: str) -> str:
    new_docs = {'Требование о представлении документов (информации)': 'Требование о представлении документов',
                'Требование о представлении пояснений': 'Требование о представлении пояснений',
                'Уведомление о вызове налогоплательщика (плательщика сбора, налогового агента)': 'Уведомление о вызове налогоплательщика',
                'Решение об отмене приостановления операций по счетам налогоплательщика': 'Решение об отмене приостановления операций по счетам',
                'Решение о привлечении лица к ответственности за налоговое правонарушение,'
                ' предусмотренное Налоговым кодексом Российской Федерации (за исключением'
                ' налогового правонарушения, дело о выявлении которого рассматривается'
                ' в порядке, установленном статьей 101 Налогового кодекса Российской Федерации': 'Решение о привлечении лица к ответственности за налоговое правонарушение'}
    return new_docs[doc]
def get_type_doc_report(frame: pd.DataFrame, current_year: int) -> pd.DataFrame:
    docs = ['Требование о представлении документов (информации)', 'Требование о представлении пояснений',
            'Уведомление о вызове налогоплательщика (плательщика сбора, налогового агента)', 'Решение об отмене приостановления операций по счетам налогоплательщика',
            'Решение о привлечении лица к ответственности за налоговое правонарушение, предусмотренное Налоговым кодексом Российской Федерации (за исключением налогового правонарушения, дело о выявлении которого рассматривается в порядке, установленном статьей 101 Налогового кодекса Российской Федерации']
    frame['Год'] = frame['Дата получения'].apply(get_year_from_string)
    current_year_frame = frame[frame['Год'] == current_year]
    current_year_frame = current_year_frame[['Вид документа']]
    current_year_frame = current_year_frame[current_year_frame['Вид документа'].isin(docs)]
    res = current_year_frame.value_counts().to_frame()
    res.reset_index(inplace=True)
    res.columns = ['Вид документа', 'Количество документов от ИФНС']
    res['Вид документа'] = res['Вид документа'].apply(rename_docs)
    return res
def get_percent(cnt_doc:int, common_sum: int) -> float:
    return round(cnt_doc/common_sum, 2)
def get_type_org_report(frame: pd.DataFrame, current_year: int) -> pd.DataFrame:
    frame['Год'] = frame['Дата получения'].apply(get_year_from_string)
    current_year_frame = frame[(frame['Год'] == current_year) & (frame['Организация'] != 'Не удалось определить название организации')]
    current_year_frame['Организация'] = current_year_frame['Организация'].apply(str.strip)
    current_year_frame = current_year_frame[['Организация']]
    res = current_year_frame.value_counts().to_frame()
    res.reset_index(inplace=True)
    res.columns = ['Наименование подрядчика', 'Количество, шт.']
    common_sum = res['Количество, шт.'].sum()
    res['Количество, шт.'] = res['Количество, шт.'].apply(get_percent, args=[common_sum])
    return res

def rename_taxes(tax: str) -> str:
    if tax == 'НДС':
        return tax
    else:
        return 'ННП'
def get_check_declaration_report(frame: pd.DataFrame, current_year: int) -> pd.DataFrame:
    taxes = ['НДС', 'Прибыль']
    frame['Год'] = frame['Дата получения'].apply(get_year_from_string)
    current_year_frame = frame[(frame['Вид проверки'].str.contains('Камеральная')) & (frame['Вид налога'].isin(taxes)) & (frame['Год'] == current_year)]
    current_year_frame = current_year_frame[['Вид налога']]
    current_year_frame['Вид налога'] = current_year_frame['Вид налога'].apply(rename_taxes)
    res = current_year_frame.value_counts().to_frame()
    res.reset_index(inplace=True)
    res.columns = ['Вид налога', 'Количество']
    common_sum = res['Количество'].sum()
    res['Количество'] = res['Количество'].apply(get_percent, args=[common_sum])
    return res

def get_type_inspection(insp: str) -> str:
    type_insp = INSPECTIONS_DICT.get(insp, 'Пусто')
    return type_insp
def get_inspection_report(frame: pd.DataFrame, current_year: int) -> pd.DataFrame:
    frame['Год'] = frame['Дата получения'].apply(get_year_from_string)
    frame['Тип'] = frame['Вид проверки'].apply(get_type_inspection)
    current_year_frame = frame[(frame['Год'] == current_year) & (frame['Тип'] != 'Пусто')]
    current_year_frame = current_year_frame[['Тип']]
    res = current_year_frame.value_counts().to_frame()
    res.reset_index(inplace=True)
    res.columns = ['Вид налоговой проверки', 'Количество, шт.']
    return res

def get_analysis_report(frame: pd.DataFrame, current_year: int) -> pd.DataFrame:
    frame['Год'] = frame['Дата получения'].apply(get_year_from_string)
    years = [current_year - 2, current_year - 1, current_year]
    frame = frame[(frame['Год'].isin(years)) & (~frame['ИНН контрагентов'].isna())]
    lst = []
    for year in years:
        cnt_doc = frame[frame['Год'] == year]['Вид документа'].count()
        cnt_org = get_count_identifier(frame[frame['Год'] == year]['ИНН контрагентов'].values)
        d = {'Год': year, 'Количество требований': cnt_doc, 'Количество подрядчиков': cnt_org}
        lst.append(d)
    res = pd.DataFrame(lst)
    return res

def get_count_identifier(ident_arr) -> int:
    ident_arr = list(map(lambda x: x.split(', ') if len(x.split(', ')) == 1 or len(x.split(', ')) >= 5 else [x.split(', ')[0]], ident_arr))
    flat_list = list(itertools.chain(*ident_arr))
    return len(set(flat_list))

def date_row_filter(row, column) -> bool:
    dt = row[column]
    try:
        dt = datetime.datetime.strptime(dt, '%d.%m.%y')
    except:
        return False
    else:
        return True
def plan_fact_report(frame: pd.DataFrame, current_year: int) -> pd.DataFrame:
    docs = ['Требование о представлении документов (информации)', 'Требование о представлении пояснений']
    frame['Год'] = frame['Дата получения'].apply(get_year_from_string)
    current_year_frame = frame[frame['Год'] == current_year]
    current_year_frame.columns = list(map(lambda x: x.replace('\n', ''), current_year_frame.columns))
    fact = current_year_frame.apply(date_row_filter, axis=1, args=['Квитанция:факт'])
    plan = current_year_frame.apply(date_row_filter, axis=1, args=['Квитанция:план'])
    res = fact & plan
    depart_frame = current_year_frame[res]
    depart_frame['Квитанция:факт'] = depart_frame['Квитанция:факт'].apply(lambda x: datetime.datetime.strptime(x, '%d.%m.%y'))
    depart_frame['Квитанция:план'] = depart_frame['Квитанция:план'].apply(lambda x: datetime.datetime.strptime(x, '%d.%m.%y'))
    depart_frame['Depart'] = depart_frame['Квитанция:факт'] <= depart_frame['Квитанция:план']
    depart = depart_frame[depart_frame['Depart'] == True]['Depart'].count()/len(depart_frame)

    fact = current_year_frame.apply(date_row_filter, axis=1, args=['Ответ:факт'])
    plan = current_year_frame.apply(date_row_filter, axis=1, args=['Ответ:план'])
    res = fact & plan
    response_frame = current_year_frame[res]
    response_frame['Ответ:факт'] = response_frame['Ответ:факт'].apply(lambda x: datetime.datetime.strptime(x, '%d.%m.%y'))
    response_frame['Ответ:план'] = response_frame['Ответ:план'].apply(lambda x: datetime.datetime.strptime(x, '%d.%m.%y'))
    response_frame['Response'] = response_frame['Ответ:факт'] <= response_frame['Ответ:план']
    response = response_frame[(response_frame['Response'] == True) & (response_frame['Вид документа'].isin(docs))]['Response'].count() / len(response_frame)

    delay = len(response_frame[(response_frame['Response'] == False) & (response_frame['Вид документа'].isin(docs))])
    fine = delay*5000

    row1 = {'Тип': 'Доля квитанций отправленных вовремя', 'Сумма': int(round(depart, 2)*100)}
    row2 = {'Тип': 'Доля ответов на Требования, отправленных вовремя', 'Сумма': int(round(response, 2)*100)}
    row3 = {'Тип': 'Количество требований, отправленных позже срока', 'Сумма': delay}
    row4 = {'Тип': 'Прогноз начисления штрафа', 'Сумма': round(fine, 1)}

    res = pd.DataFrame([row1, row2, row3, row4])
    return res

def agents_report(frame: pd.DataFrame, year: int) -> pd.DataFrame:
    frame['Год'] = frame['Дата получения'].apply(get_year_from_string)
    frame = frame[(frame['Год'] == year) & (~frame['ИНН контрагентов'].isna())]
    ident_arr = list(map(lambda x: x.split(', ') if len(x.split(', '))>=5 else [x.split(', ')[0]], frame[frame['Год'] == year]['ИНН контрагентов'].values))
    cnt_dict = Counter(list(itertools.chain(*ident_arr)))
    res_lst = list()
    for i, (ident, cnt) in enumerate(sorted(cnt_dict.items(), key=lambda x: x[1], reverse=True)):
        if i == 10:
            break
        else:
            res_lst.append({'Наименование подрядчика': ident, 'Количество, шт.': cnt})
    res = pd.DataFrame(res_lst)
    res.sort_values(by=['Количество, шт.'], inplace=True)
    return res

def create_report_dict(frame: pd.DataFrame, year: int) -> dict:
    TITLE_DICT = {
        'Информативный блок': plan_fact_report,
        'Поквартальный анализ полученных документов': get_quarter_report,
        'Анализ в разрезе видов документов': get_type_doc_report,
        'Анализ по виду организаций': get_type_org_report,
        'Анализ по виду налога': get_check_declaration_report,
        'Анализ по виду проверки': get_inspection_report,
        'Основные подрядчики': agents_report,
        'Анализ требований': get_analysis_report,

    }
    RESULT_DICT = dict()
    for title, func in TITLE_DICT.items():
        res = func(frame, year)
        RESULT_DICT[title] = res
    return RESULT_DICT