from checking.check import create_all_report
from decorating.decore import decorating_excel_list
from interface.panels import start, end, error
from loading.load import create_save_filename, read_file

values = start()
diadoc_path = values['diadoc']
journal_path = values['journal']
try:
    save_filename = create_save_filename(values['save'], values['name'])
    diadoc, journal = read_file(diadoc_path, True), read_file(journal_path)
    report_dict = create_all_report(journal, diadoc)
    decorating_excel_list(report_dict, save_filename)
except Exception as exp:
    error()
else:

    end(save_filename)


