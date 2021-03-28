from pprint import pprint
import pyexcel

zsd_path = 'C:\\Users\\by059491\\Downloads\\zsd.xlsx'
lx02_path = 'C:\\Users\\by059491\\Downloads\\a64ca054-85a8-441e-a0f2-9698644439c4.xlsx'


class Subtotal:
    """
    Читает эксель файл и подсчитывает количество продукции по каждому материалу.
    Первой строкой в файле должны быть заголовки столбцов.
    excel_file - путь к файлу
    material - имя колонки, который содержит коды материалов
    quantity - имя колонки, которая содержит количество, необходимое для подсчета
    """
    def __init__(self, excel_file, material, quantity, typ=None, value_to_ignore=None):
        self.excel_array = pyexcel.get_records(file_name=excel_file)
        self.output = {}
        for row_index, row in enumerate(self.excel_array):
            if typ is not None and row[typ] == value_to_ignore:
                continue
            else:
                if row_index == 0 or typ is not None and self.excel_array[row_index - 1][typ] != row[typ]:
                    previous_row = None
                else:
                    previous_row = self.excel_array[row_index - 1]
                if row[material] not in self.output:
                    self.output.setdefault(row[material], row[quantity])
                if previous_row is not None and previous_row[material] == row[material]:
                    self.output[row[material]] += row[quantity]

    def __getitem__(self, item):
        return self.output[item]

    def __sub__(self, other):
        """
        Метод вычисления. Отнимает количество по каждому материалу, если тот есть в other.
        Если полученное количество отрицательное - игнорирует его.
        Возвращает словарь.
        """
        sub_output = {}
        for key, val in self.output.items():
            try:
                new_val = self.output[key]
                new_val -= other[key]
                if new_val > 0:
                    sub_output.setdefault(key, new_val)
            except KeyError as exc:
                print(exc)
        return sub_output

    def __str__(self):
        return f'{self.output}'


if __name__ == '__main__':
    subtotal_zsd = Subtotal(zsd_path, material='Material', quantity='Delivery Quantity')
    subtotal_lx02 = Subtotal(lx02_path, material='Material', quantity='Available stock', typ='Storage Type', value_to_ignore='110')
    print(subtotal_zsd['5608'])
    print(subtotal_zsd)
    print(subtotal_lx02)
    res = subtotal_zsd - subtotal_lx02
    pprint(res)