from generator import generate_excel


class ExcelGenerator:
    sheet_name = "Sheet1"

    @classmethod
    def get_displayable_data(cls):
        return {
            "columns": cls.get_column_titles(),
            "rows": [r for r in cls.get_displayable_rows() if cls.is_row_valid(r)]
        }

    @classmethod
    def generate_file(cls):
        generate_excel(cls.get_displayable_data(), cls.get_file_name(), cls.get_sheet_name())

    @classmethod
    def get_value(cls, data, title):
        raise NotImplementedError

    @classmethod
    def get_displayable_row(cls, *args):
        response = []
        for title in cls.get_column_titles():
            response.append(cls.get_value(*args, title))
        return response

    @classmethod
    def get_column_titles(cls):
        raise NotImplementedError

    @classmethod
    def get_displayable_rows(cls):
        raise NotImplementedError

    @classmethod
    def get_file_name(cls):
        raise NotImplementedError

    @classmethod
    def get_sheet_name(cls):
        return cls.sheet_name

    @classmethod
    def is_row_valid(cls, row):
        return True
