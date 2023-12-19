import sys
import time

import win32com.client
import PySimpleGUI as sg

from updating.update import call_updater, check_version

KEYS = ['journal', 'diadoc', 'report', 'save', 'name']
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
DEFAULT_FILENAME = 'Результат'
sg.theme('SamoletTheme')
def start():
    UPD_FRAME = [sg.Column([[sg.Button('Проверка', key='check_upd'), sg.Text('Нет обновлений', key='not_upd_txt'),
                  sg.Push(),
                  sg.pin(sg.Text('Доступно обновление', justification='center', visible=False, key='upd_txt')),
                  sg.Push(),
                  sg.pin(sg.Button('Обновить', key='upd_btn', visible=False))],
                 ],
                           size=(420, 50))]
    MAIN_PANEL = [
        sg.Column([
            [sg.Text('Журнал с/ф', font='bold')],
            [sg.Input(key='journal'), sg.FileBrowse(button_text='Выбрать')],
            [sg.Text('Выгрузка из Диадок', font='bold')],
            [sg.Input(key='diadoc'), sg.FileBrowse(button_text='Выбрать')],
            # [sg.Text('Отчет по незакрытым авансам', font='bold')],
            # [sg.Input(key='report'), sg.FileBrowse(button_text='Выбрать')],
            [sg.Text('Папка для сохранения', font='bold')],
            [sg.Input(key='save'), sg.FolderBrowse(button_text='Выбрать')],
            [sg.Text('Имя файла', font='bold')],
            [sg.Input(key='name', default_text=DEFAULT_FILENAME, pad=((5, 10), (5, 20)))],
            [sg.OK(button_text='Далее'), sg.Cancel(button_text='Выход')]
        ], key='-FILE_PANEL-', visible=True, size=(420, 300))
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
    yeet.close()
    check_values = check_user_values(data=values)
    if check_values:
        edit_values = edit_values_dict(values)
        return edit_values
    else:
        check_input_error = input_error_panel()
        if check_input_error:
            return start()

def check_user_values(data: dict) -> bool:
    for k, v in data.items():
        if k in KEYS and v == '':
            return False
    return True

def edit_values_dict(values_dict: dict) -> dict:
    result_dict = dict()
    for k, v in values_dict.items():
        if k in KEYS and v != '':
            result_dict[k] = v
    return result_dict

def input_error_panel():
    event = sg.popup('Ошибка ввода', 'При вводе данных возникла ошибка.\nВы хотите повторить ввод данных?',
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