from datetime import datetime, timedelta


def write(ws, row, col, val, format=None, date_format=None, hour_format=None):
    if val is None:
        pass
    elif type(val) is str:
        ws.write_string(row, col, val, format)
    elif type(val) is bool:
        ws.write_boolean(row, col, val, format)
    elif type(val) is int or type(val) is float:
        ws.write_number(row, col, val, format)
    elif type(val) is datetime:
        ws.write_datetime(row, col, val, date_format)
    elif type(val) is timedelta:
        ws.write_datetime(row, col, val, hour_format)
    elif type(val) is list:
        ws.write(row, col, '; '.join([str(v) for v in val]), format)
    else:
        ws.write(row, col, val, format)
