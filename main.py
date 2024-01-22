import datetime

from checking.check import create_all_report, create_report_dict, plan_fact_report, agents_report
from decorating.decore import create_result_files
from interface.panels import start, end, error
from loading.load import create_save_filename, load_historic_data, add_row_to_frame, read_data
from logs import log


log.info('Начало обработки данных')
values = start()
if values['SLCT_OPT'] != 'Отчет по документам ФНС':
    type_report = values['type']
    diadoc_path, journal_path, report_path = values['diadoc'], values['journal'], values['report']
    report_date = values['curr_dt']
    historic_path = values['ist_file']
    try:
        save_err_name = create_save_filename(values['save'], values['err_name'])
        save_adv_name = create_save_filename(values['save'], values['adv_name'])
        report_dict = create_all_report(journal_path, diadoc_path, report_path, report_date, type_report)
        if type_report in ('only_advances', 'all'):
            historic_data = load_historic_data(historic_path)
            historic_data = add_row_to_frame(historic_data, report_dict['REPORT_Средняя разница'], historic_path)
            report_dict['REPORT_Средняя разница'] = historic_data
        create_result_files(report_dict, save_err_name, save_adv_name, type_report)
    except Exception as exp:
        log.info('Возникла ошибка')
        log.exception(exp)
        error()
    else:
        log.info('Обработка данных закончена')
        if type_report in ('only_errors', 'all'):
            end(save_err_name)
        else:
            end(save_adv_name)
else:
    try:
        path = values['doc_report']
        df = read_data(path)
        report_year = values['year']
        d = create_report_dict(df, report_year)
        pres_savename = create_save_filename(values['save_pres'], values['pres_name'])
        create_result_files(report_dict=d, adv_name=pres_savename)
    except Exception as exp:
        log.info('Возникла ошибка')
        log.exception(exp)
        error()
    else:
        log.info('Обработка данных закончена')
        end(pres_savename)

