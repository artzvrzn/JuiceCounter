import pyexcel
from time import time

start = time()

file_path = 'C:\\Users\\by059491\\Downloads\\ea9391ce-8b48-4772-9147-4f4ca6a2acf2.xlsx'


class JuiceCounter:
    def __init__(self, zsd_output):
        self.output = {}
        self.zsd_output = zsd_output

    def get_subtotal(self):
        file_array = pyexcel.get_array(file_name=self.zsd_output)
        for row_index, row in enumerate(file_array):
            material = row[2]
            amount = row[4]
            if material not in self.output:
                self.output.setdefault(material, amount)
            if file_array[row_index - 1][2] == material:
                self.output[material] += amount
        if 'Material' in self.output: del self.output['Material']

