from pprint import pprint

import pyexcel

from time import time

start = time()

zsd_path = 'C:\\Users\\by059491\\Downloads\\zsd.xlsx'
lx02_path = 'C:\\Users\\by059491\\Downloads\\lx02.xlsx'


class Subtotal:
    def __init__(self, excel_file):
        self.excel_array = pyexcel.get_records(file_name=excel_file)
        self.output = {}

    def __call__(self, material, quantity):
        for row_index, row in enumerate(self.excel_array):
            previous_row = self.excel_array[row_index - 1] if row_index != 0 else None
            if row[material] not in self.output:
                self.output.setdefault(row[material], row[quantity])
            if previous_row is not None and previous_row[material] == row[material]:
                self.output[row[material]] += row[quantity]
        return self.output


if __name__ == '__main__':
    subtotal = Subtotal(zsd_path)
    print(subtotal(material='Material', quantity='Delivery Quantity'))
