from pprint import pprint
from datetime import datetime
from dataclasses import dataclass, astuple
from operator import attrgetter

import pyexcel
import openpyxl

MATERIAL = 'Material'
DESCRIPTION = 'Material description'
BIN = 'Storage Bin'
TYP = 'Storage Type'
QUANTITY = 'Available stock'
DATE = 'GR Date'
QUARANTINE = 'Storage location'

juice_list = {'819310': ('1.0 BRK X12 DOBRIY APPLE JC GEMINA BY', 48),
              '819611': ('1.0 BRK X12 DOBRIY MULTIFRUIT NECT G BY', 48),
              '372112': ('1.0 BRK X12 DOBRIY ORANGE NECT GEMIN BY', 48),
              '819410': ('1.0 BRK X12 DOBRIY TOMATO JC GEMINA BY', 48),
              '735207': ('1.0 BRK X12 RICH ORANGE JUICE BY', 60),
              '750009': ('1.0 BRK X12 RICH APPLE JUICE BY', 60),
              '752005': ('1.0 BRK X12 RICH CHERRY NECT BYRU', 60),
              '822603': ('2.0 BRK X6 DOBRIY APPLE JC BYRU', 65),
              '865503': ('2.0 BRK X6 DOBRIY MULTIFRUIT NECT BYRU', 65),
              '371503': ('2.0 BRK X6 DOBRIY ORANGE NECT BYRU', 65),
              '840707': ('2.0 BRK X6 DOBRIY TOMATO SALT JC BYRU', 65), }


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
    }

    def __init__(self, excel_file):
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
                        if material_list[exist_index].bin_date < bin_data.bin_date:
                            material_list[exist_index].bin_date = bin_data.bin_date
                    except StopIteration:
                        material_list.append(bin_data)

    def get_sorted_array(self):
        sorted_output = {k: v for k, v in sorted(self.output.items(), key=lambda x: x[1][0].bin_desc)}
        for key, value in sorted_output.items():
            value.sort(key=attrgetter('bin_date', 'quantity'))
        return sorted_output

    def get_current_bin_excel(self):
        book_name = 'example.xlsx'
        try:
            sheet = pyexcel.get_sheet(file_name=book_name)
        except FileNotFoundError:
            book = openpyxl.Workbook()
            book.save(filename=book_name)
            sheet = pyexcel.get_sheet(file_name=book_name)
        sheet.row += ["Material", "Description", "Bin", "Date", "Quantity"]
        for key, value in self.get_sorted_array().items():
            sheet.row += [key, value[0].bin_desc, value[0].bin_name, value[0].bin_date, value[0].quantity]
        sheet.save_as(filename=book_name)

    def get_current_bin_txt(self):
        with open('output.txt', 'w') as file:
            for key, value in self.get_sorted_array().items():
                output = f'{key:<10} {value[0].bin_desc:<40} :: {value[0].bin_name:<6} :: ' \
                         f'{value[0].bin_date.strftime("%d.%m.%Y")}\n'
                file.write(output)

    def __str__(self):
        return f'{self.output}'


class Subtotal:
    """
    Читает эксель файл и подсчитывает количество продукции по каждому материалу.
    Первой строкой в файле должны быть заголовки столбцов.
    excel_file - принимает pyexcel.get_records(путь к файлу)
    material - имя колонки, который содержит коды материалов
    quantity - имя колонки, которая содержит количество, необходимое для подсчета
    value_to_ignore - словарь, ключ - имя колонки, значение - что исключать
    """

    def __init__(self, excel_file, material, quantity, value_to_ignore=None):
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
    # zsd_path = 'C:\\Users\\by059491\\Downloads\\55408ca5-0618-4227-85b2-17ad9265b7dd.xlsx'
    lx02_path = 'C:\\Users\\by059491\\Downloads\\60b9c892-9784-4c26-b3fe-0892eded0b88.xlsx'
    # zsd_array = pyexcel.get_records(file_name=zsd_path)
    lx02_array = pyexcel.get_records(file_name=lx02_path)

    # subtotal_zsd = Subtotal(zsd_array, material='Material', quantity='Cumltv Confd Qty(SU)')
    # subtotal_lx02 = Subtotal(lx02_array,
    #                          material='Material',
    #                          quantity='Available stock',
    #                          value_to_ignore={'Storage Type': '110'})
    # print(subtotal_zsd['5608'])
    # print(subtotal_zsd)
    # print(subtotal_lx02)
    # res = subtotal_zsd - subtotal_lx02
    # pprint(res)
    # for mat, desc in juice_list.items():
    #     pallet_amount = round(res[mat] / desc[1], 1)
    #     print(f'{desc[0]} - {pallet_amount}')

    bin_determinate = BinDeterminate(lx02_array)
    pprint(bin_determinate.get_current_bin_txt())
