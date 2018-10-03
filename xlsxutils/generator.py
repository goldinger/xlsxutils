import os
import stat
import xlsxwriter
from xlsxutils.writer import write


def generate_excel(data, file_path, sheet_name):
    folder = '/'.join(file_path.split('/')[:-1]) + '/'
    os.makedirs(folder, exist_ok=True)
    file_name = file_path
    try:
        os.chmod(file_name, stat.S_IWRITE)
        os.remove(file_name)
    except Exception:
        pass
    wb = xlsxwriter.Workbook(file_name, {'remove_timezone': True})
    ws = wb.add_worksheet(sheet_name)
    ws.set_column('A:ZZ', 25)

    date_format = wb.add_format({'num_format': 'yyyy-mm-dd'})
    hour_format = wb.add_format({'num_format': 'hh:mm:ss'})

    title_blue_format = wb.add_format()
    title_blue_format.set_bold(True)
    title_blue_format.set_align('center')
    title_blue_format.set_font_size(16)
    title_blue_format.set_bg_color('blue')
    title_blue_format.set_color('white')

    date_blue_format = wb.add_format({'num_format': 'yyyy-mm-dd'})
    date_blue_format.set_bold(True)
    date_blue_format.set_align('center')
    date_blue_format.set_font_size(16)
    date_blue_format.set_bg_color('blue')
    date_blue_format.set_color('white')

    row = 0
    col = 0

    for title in data.get('columns'):
        write(ws, row, col, title, format=title_blue_format, date_format=date_blue_format, hour_format=hour_format)
        col += 1
    col = 0
    row += 1

    for values in data.get('rows', []):
        if values is not None:
            for value in values:
                write(ws, row, col, value, date_format=date_format)
                col += 1
            col = 0
            row += 1

    wb.close()
    # os.chmod(file_name, stat.S_IREAD | stat.S_IRGRP | stat.S_IROTH)
