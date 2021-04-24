import os
from datetime import datetime, timedelta
from dataclasses import dataclass, astuple
from operator import attrgetter
from math import ceil
import pyexcel
from chrome_driver import OutOfStock, Lx02, BASE_PATH


MATERIAL = 'Material'
DESCRIPTION = 'Material description'
BIN = 'Storage Bin'
TYP = 'Storage Type'
QUANTITY = 'Available stock'
DATE = 'SLED/BBD'
QUARANTINE = 'Storage location'

pallet_size = {
    '819310': 48,
    '819611': 48,
    '372112': 48,
    '819410': 48,
    '1533902': 48,
    '735207': 60,
    '750009': 60,
    '752008': 60,
    '822603': 65,
    '865503': 65,
    '371503': 65,
    '840707': 65,
    '428921': 120,
    '14692': 120,
}


@dataclass
class BinData:
    bin_name: str
    quantity: int
    bin_date: datetime
    bin_desc: str

    def __str__(self):
        return f'{self.bin_desc} - {self.quantity} - {self.bin_name} - {self.bin_date.strftime("%d.%m.%Y ")}'

    def __iter__(self):
        return iter(astuple(self))


class BinDeterminate:
    """
    Читает эксель файл и определяет идущее в отгрузку место для каждого материала.
    Формирует словарь, ключ: материал, значение: список объектов BinData, содержащих место, количество продукции,
    дату и описание материала.
    """

    # словарь исключений для эксель файла 436907
    ignore_dict = {
        MATERIAL: ['10018454',
                   '10018455',
                   '10018456',
                   '10048062',
                   '10060654',
                   '10060655',
                   '10060656',
                   '10064536',
                   '10077352', ],
        BIN: ['W2', ],
        TYP: ['200', ],
        QUARANTINE: ['1500', ],
        DATE: ['', ],
    }

    def __init__(self, excel_file):
        """
        :param excel_file: принимает pyexcel.get_records(путь к файлу)
        """
        self.excel_array = excel_file
        self.output = {}
        for row in self.excel_array:
            bin_data = BinData(bin_name=row[BIN], quantity=row[QUANTITY], bin_date=row[DATE], bin_desc=row[DESCRIPTION])
            if not any(row[k] in v for k, v in self.ignore_dict.items()):
                if row[MATERIAL] not in self.output:
                    self.output.setdefault(row[MATERIAL], [bin_data, ])
                else:
                    material_list = self.output[row[MATERIAL]]
                    try:
                        exist_index = material_list.index(
                            next(x for x in material_list if x.bin_name == bin_data.bin_name))
                        material_list[exist_index].quantity += bin_data.quantity
                        if material_list[exist_index].bin_date > bin_data.bin_date:
                            material_list[exist_index].bin_date = bin_data.bin_date
                    except StopIteration:
                        material_list.append(bin_data)

    def get_sorted_array(self):
        """
        Сортирует значения полученного словаря по дате и затем по количеству.
        Возращает новый словарь.

        """
        sorted_output = {k: v for k, v in sorted(self.output.items(), key=lambda x: x[1][0].bin_desc)}
        for key, value in sorted_output.items():
            value.sort(key=attrgetter('bin_date', 'quantity'))
        return sorted_output

    def get_current_bin_txt(self):
        """
        Записывает файл output.txt, в котором перечислены все материалы и первое отгружаемое место.
        """
        with open('output.txt', 'w') as txt:
            for key, value in self.get_sorted_array().items():
                output = f'{key:<10} {value[0].bin_desc:<40} :: {value[0].bin_name:<6} :: ' \
                         f'{value[0].bin_date.strftime("%d.%m.%Y")}\n'
                txt.write(output)

    def __str__(self):
        return f'{self.output}'


class Subtotal:
    """
    Читает эксель файл и подсчитывает общее количество продукции по каждому материалу.
    Первой строкой в файле должны быть заголовки столбцов.
    """

    def __init__(self, excel_file, material, quantity, value_to_ignore=None):
        """
        :param excel_file: принимает pyexcel.get_records(путь к файлу)
        :param material: имя колонки, который содержит коды материалов
        :param quantity: имя колонки, которая содержит количество, необходимое для подсчета
        :param value_to_ignore: словарь, ключ - имя колонки, значение - что исключать
        """
        self.excel_array = excel_file
        self.output = {}
        for row in self.excel_array:
            if value_to_ignore is not None and any(row[x] == y for x, y in value_to_ignore.items()):
                continue
            else:
                if row[material] not in self.output:
                    self.output.setdefault(row[material], row[quantity])
                    continue
                else:
                    self.output[row[material]] += row[quantity]

    def __getitem__(self, item):
        return self.output[item]

    def __sub__(self, other):
        """
        Метод вычисления. Отнимает количество по каждому материалу, если тот есть в other.
        Если полученное количество отрицательное - игнорирует его.
        Возвращает словарь.
        zsd_oos - lx_02
        :param other: объект Subtotal
        :return: словарь. ключ: материал, значение: разница по количеству
        """
        sub_output = {}
        for key, val in self.output.items():
            try:
                new_val = self.output[key]
                new_val -= other[key]
                if new_val > 0:
                    sub_output.setdefault(key, new_val)
            except KeyError:
                pass
        return sub_output

    def __str__(self):
        return f'{self.output}'


def main(date_from=None, date_to=None):
    zsd_oos = OutOfStock(date_from=date_from, date_to=date_to)
    zsd_path = zsd_oos.get_file()
    lx_02 = Lx02()
    lx02_path = lx_02.get_file()
    zsd_array = pyexcel.get_records(file_name=zsd_path)
    lx02_array = pyexcel.get_records(file_name=lx02_path)
    subtotal_zsd = Subtotal(zsd_array, material='Material', quantity='Cumltv Confd Qty(SU)')
    subtotal_lx02 = Subtotal(lx02_array,
                             material='Material',
                             quantity='Available stock',
                             value_to_ignore={'Storage Type': '110'})

    bin_determinate = BinDeterminate(lx02_array)
    difference = subtotal_zsd - subtotal_lx02
    with open('Результат.txt', 'w') as file:
        for mat, mat_value in bin_determinate.get_sorted_array().items():
            if mat in pallet_size:
                try:
                    amount = difference[mat] / pallet_size[mat]
                    line = f'{mat:<8} {mat_value[0].bin_desc:<40} | {mat_value[0].bin_name:<8} | ' \
                           f'{ceil(amount):<2} | {difference[mat]}\n'
                    file.write(line)
                    print(line, end='')
                except KeyError:
                    pass
    os.startfile(BASE_PATH / 'Результат.txt', 'print')


if __name__ == '__main__':
    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument('-f', '--friday', action='store_true')
    args = parser.parse_args()
    if args.friday:
        date_1 = datetime.today() + timedelta(days=1)
        date_2 = date_1 + timedelta(days=2)
        main(date_from=date_1.strftime('%d.%m.%Y'), date_to=date_2.strftime('%d.%m.%Y'))
    else:
        main()
