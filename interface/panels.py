import datetime
import sys
import time

import win32com.client
import PySimpleGUI as sg

from loading.load import parse_data
from updating.update import call_updater, check_version

CHECK_DICT = {
    'only_errors': ['journal', 'diadoc', 'save', 'err_name'],
    'only_advances': ['report', 'save','adv_name'],
    'all': ['journal', 'diadoc', 'report', 'save', 'err_name', 'adv_name']
}
sg.LOOK_AND_FEEL_TABLE['SamoletTheme'] = {
                                        'BACKGROUND': '#007bfb',
                                        'TEXT': '#FFFFFF',
                                        'INPUT': '#FFFFFF',
                                        'TEXT_INPUT': '#000000',
                                        'SCROLL': '#FFFFFF',
                                        'BUTTON': ('#FFFFFF', '#007bfb'),
                                        'PROGRESS': ('#354d73', '#FFFFFF'),
                                        'BORDER': 1, 'SLIDER_DEPTH': 0,
                                        'PROGRESS_DEPTH': 0, }
KEYS_DICT = {
    'report': 'Отчет по незакрытым авансам',
    'journal': 'Журнал с/ф',
    'diadoc': 'Выгрузка из Диадок',
    'save': 'Папка для сохранения',
    'errors_name': 'Имя файла отчета по ошибкам',
    'report_name': 'Имя файла отчета НДС'
}
DEFAULT_FILENAME = 'Результат'
sg.theme('SamoletTheme')
def start():
    y, m, d = datetime.datetime.now().year, datetime.datetime.now().month, datetime.datetime.now().day-1
    str_now_dt = datetime.datetime.now().strftime('%d.%m.%Y')
    UPD_FRAME = [sg.Column([[sg.Button('Проверка', key='check_upd'), sg.Text('Нет обновлений', key='not_upd_txt'),
                  sg.Push(),
                  sg.pin(sg.Text('Доступно обновление', justification='center', visible=False, key='upd_txt')),
                  sg.Push(),
                  sg.pin(sg.Button('Обновить', key='upd_btn', visible=False))],
                 ],
                           size=(420, 50))]
    MAIN_PANEL = [
        sg.Column([
            [sg.Radio(text='Только ошибки', default=True, group_id='how_do', key='only_errors', enable_events=True),
             sg.Radio(text='Только НДС', default=False, group_id='how_do', key='only_advances', enable_events=True),
             sg.Radio(text='Все', default=False, group_id='how_do', key='all', enable_events=True)],
            [sg.pin(sg.Column(
                [
                    [sg.Text('Дата отчета', font='bold')],
                    [sg.Input(key='curr_dt', default_text=str_now_dt), sg.CalendarButton('Выбрать', target='curr_dt', close_when_date_chosen=True, no_titlebar=False,
                                                                default_date_m_d_y = (m, d, y), format="%d.%m.%Y", locale='ru')]
                ], key='date_col', visible=False))],
            [sg.pin(sg.Column(
            [[sg.Text('Журнал с/ф', font='bold')],
            [sg.Input(key='journal'), sg.FileBrowse(button_text='Выбрать')],
            [sg.Text('Выгрузка из Диадок', font='bold')],
            [sg.Input(key='diadoc'), sg.FileBrowse(button_text='Выбрать')]], key='errors_col'))],
            [sg.pin(sg.Column(
            [
                [sg.Text('Отчет по незакрытым авансам', font='bold')],
                [sg.Input(key='report'), sg.FileBrowse(button_text='Выбрать')],
                [sg.Text('Файл с историческими данными', font='bold')],
                [sg.Input(key='ist_file'), sg.FileBrowse(button_text='Выбрать')]
            ], visible=False, key='report_col'))],
            [sg.Column([[sg.Text('Папка для сохранения', font='bold')],
                        [sg.Input(key='save'), sg.FolderBrowse(button_text='Выбрать')]])],
            [sg.pin(sg.Column(
            [[sg.Text('Имя файла отчета по ошибкам', font='bold')],
            [sg.Input(key='err_name', default_text=DEFAULT_FILENAME)]], key='errors_name'))],
            [sg.pin(sg.Column(
            [[sg.Text('Имя файла отчета НДС', font='bold')],
                 [sg.Input(key='adv_name', default_text='Отчет')]], key='report_name', visible=False))],
            [sg.OK(button_text='Далее'), sg.Cancel(button_text='Выход')]
        ], key='-FILE_PANEL-', visible=True, size=(420, 650))
    ]
    layout = [
            [sg.Frame(layout=[UPD_FRAME], title='Обновление', key='--UPD_FRAME--')],
            [sg.Frame(layout=[MAIN_PANEL], title='Выбор файлов')]]
    yeet = sg.Window('Сверка данных файлов', layout=layout)
    check, upd_check = False, True
    while True:
        event, values = yeet.read(100)
        if check:
            upd_check = check_version()
            check = False
        if event in ('Выход', sg.WIN_CLOSED):
            sys.exit()
        if event == 'check_upd':
            check = True
        if not upd_check:
            yeet['not_upd_txt'].Update(visible=False)
            yeet['upd_txt'].Update(visible=True)
            yeet['upd_btn'].Update(visible=True)
        if event == 'upd_btn':
            yeet.close()
            call_updater('pocket')
        if event == 'Далее':
            break
        if event == 'only_advances':
            yeet['date_col'].Update(visible=True)
            yeet['report_col'].Update(visible=True)
            yeet['report_name'].Update(visible=True)
            yeet['errors_col'].Update(visible=False)
            yeet['errors_name'].Update(visible=False)
            yeet.refresh()
            yeet['-FILE_PANEL-'].contents_changed()
        if event == 'only_errors':
            yeet['date_col'].Update(visible=False)
            yeet['report_name'].Update(visible=False)
            yeet['errors_name'].Update(visible=True)
            yeet['report_col'].Update(visible=False)
            yeet['errors_col'].Update(visible=True)
            yeet.refresh()
            yeet['-FILE_PANEL-'].contents_changed()
        if event == 'all':
            yeet['date_col'].Update(visible=True)
            yeet['report_name'].Update(visible=True)
            yeet['errors_name'].Update(visible=True)
            yeet['report_col'].Update(visible=True)
            yeet['errors_col'].Update(visible=True)
            yeet.refresh()
            yeet['-FILE_PANEL-'].contents_changed()
    yeet.close()
    check_report, values = check_user_values(data=values)
    if check_report:
        # edit_values = edit_values_dict(values)
        return values
    else:
        check_input_error = input_error_panel(values)
        if check_input_error:
            return start()

def check_user_values(data: dict) -> tuple:
    for ipt_type in CHECK_DICT.keys():
        if data[ipt_type] == True:
            data['type'] = ipt_type
            for k, v in data.items():
                if k in CHECK_DICT[ipt_type]:
                    if v == '':
                        return False, k
            break
    data['curr_dt'] = parse_data(data['curr_dt'])
    return True, data

# def edit_values_dict(values_dict: dict) -> dict:
#     result_dict = dict()
#     for k, v in values_dict.items():
#         if k in KEYS and v != '':
#             result_dict[k] = v
#     return result_dict

def input_error_panel(key):
    event = sg.popup(f'''При вводе данных возникла ошибка
Не выбран следующий ключ <!{KEYS_DICT[key]}!>
Вы хотите повторить ввод данных?''',
                     title='Ошибка', custom_text=('Да', 'Нет'))
    if event == 'Да':
        return True
    else:
        sys.exit()

def end(path):
    event = sg.popup('Сверка завершена\nОткрыть обработанный файл?',
                     title='Завершение работы', custom_text=('Да', 'Нет'))
    if event == 'Да':
        Excel = win32com.client.Dispatch("Excel.Application")
        Excel.Visible = True
        Excel.Workbooks.Open(Filename=path)
        time.sleep(5)
        del Excel
    else:
        sys.exit()

def error():
    sg.popup_auto_close('При выполнении сверки возникла непредвиденная ошибка\nПодробности можно посмотреть в лог файле',
                                title='Выход с исключением', auto_close_duration = 15)
    sys.exit()