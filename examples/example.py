from xlsxutils.ExcelGenerator import ExcelGenerator


class MyGen(ExcelGenerator):
    @classmethod
    def get_value(cls, data, title):
        if title == 'number':
            return data
        elif title == 'number + 1':
            return data + 1
        else:
            return None

    @classmethod
    def get_column_titles(cls):
        return [
            'number',
            'number + 1'
        ]

    @classmethod
    def get_displayable_rows(cls):
        rows = []
        for i in range(20):
            rows.append(cls.get_displayable_row(i))
        return rows

    @classmethod
    def get_file_name(cls):
        return 'insert/your/path/here'


if __name__ == '__main__':
    MyGen.generate_file()
