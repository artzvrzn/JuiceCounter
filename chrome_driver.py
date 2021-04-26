import glob
import os
from selenium import webdriver
from datetime import date, timedelta
from selenium.webdriver.common.keys import Keys
from time import sleep
from pathlib import Path
from getpass import getuser
from logger import logger

BASE_PATH = Path(__file__).resolve().parent


def get_last_file_path():
    """
    Возвращает путь к последнему файлу формата xlsx.
    """
    try:
        last_file = max(glob.glob(f'C:\\Users\\{getuser()}\\Downloads\\*.xlsx'), key=os.path.getctime)
    except ValueError:
        last_file = 'No file'
        logger.debug(last_file)
    return last_file


class _GetPage:
    """
    Базовый класс для работы с транзакциями.
    """
    BASE_URL = 'https://cuvl0301.eur.cchbc.com:8204/'

    def __init__(self, page_url):
        """
        Инициализирует вебдрайвер chrome и переходит на url транзакции.
        :param page_url: url страницы
        """
        self.driver = webdriver.Chrome(BASE_PATH / 'chromedriver.exe')
        self.driver.minimize_window()
        self.driver.implicitly_wait(10)
        self.driver.get(page_url)
        self.driver.switch_to.frame('application-Shell-startGUI')
        self.fill_parameters = {}
        self.last_file_name = get_last_file_path()
        self.output_file_name = None

    def _fill_start_page(self):
        """
        Метод, используемый для заполнения стартовой страницы транзакции.
        """
        for id_key, id_val in self.fill_parameters.items():
            element = self.driver.find_element_by_id(id_key)
            element.clear()
            element.send_keys(id_val)

    def _export_file(self):
        """
        Метод, используемый для экспорта файла формата xlsx.
        """
        while get_last_file_path() == self.last_file_name:
            sleep(0.5)
        self.output_file_name = get_last_file_path()
        logger.info(f'{self.__class__.__name__} = {self.output_file_name}')

    def get_file(self):
        self._fill_start_page()
        self._export_file()
        self.driver.quit()
        return self.output_file_name


class OutOfStock(_GetPage):
    """
    Открывает страницу транзакции ZSD_OOS в браузере.
    Константы - id элементов на странице, с которыми необходимо взаимодействие.
    """
    SALES_ORG = 'M0:46:::1:34'
    DEL_DATE_FROM = 'M0:46:::3:34'
    DEL_DATE_TO = 'M0:46:::3:59'
    DOC_TYPE = 'M0:46:::6:34'
    CHECKBOX_ONE = 'M0:46:::16:3'
    CHECKBOX_TWO = 'M0:46:::17:3'
    SUBMIT_BTN = 'M0:50::btn[8]'
    EXPORT_BTN = '_MB_EXPORT102-r'
    SPREADSHEET = 'menu_MB_EXPORT102_1_1-r'
    CONTINUE_BTN = 'M1:50::btn[0]'
    PAGE_URL = _GetPage.BASE_URL + 'sap/bc/ui2/flp#Shell-startGUI?sap-ui2-tcode=ZSD_OOS&sap-system=LOCAL'

    def __init__(self, date_from=None, date_to=None):
        """
        :param date_from: опционально - дата, с которой будет выполнена выборка в транзакции.
        По умолчанию - завтрашний день. Формат %d.%m.%Y
        :param date_to: опционально - дата, до которой будет выполнена выборка.
        """
        super().__init__(self.PAGE_URL)
        self.date_from = date.today() + timedelta(days=1)
        self.date_to = date_to
        self.fill_parameters = {
            self.SALES_ORG: '2300',
            self.DEL_DATE_FROM: self.date_from.strftime('%d.%m.%Y'),
            self.DOC_TYPE: 'ZOR',
        }
        self.checkboxes = [
            self.driver.find_element_by_id(self.CHECKBOX_ONE),
            self.driver.find_element_by_id(self.CHECKBOX_TWO),
        ]
        if date_from is not None:
            self.fill_parameters[self.DEL_DATE_FROM] = date_from
        if date_to is not None:
            self.fill_parameters.setdefault(self.DEL_DATE_TO, self.date_to)

    def _fill_start_page(self):
        super()._fill_start_page()
        for checkbox in self.checkboxes:
            checkbox.click()
        self.driver.find_element_by_id(self.SUBMIT_BTN).click()

    def _export_file(self):
        self.driver.find_element_by_id(self.EXPORT_BTN).click()
        self.driver.find_element_by_id(self.SPREADSHEET).find_element_by_tag_name('tr').click()
        self.driver.find_element_by_id(self.CONTINUE_BTN).click()
        super()._export_file()


class Lx02(_GetPage):
    """
    Открывает страницу транзакции LX_02 в браузере.
    Константы - id элементов на странице, с которыми необходимо взаимодействие.
    """
    WAREHOUSE_NUMBER = 'M0:46:::0:34'
    STORAGE_TYPE_FROM = 'M0:46:::1:34'
    STORAGE_TYPE_TO = 'M0:46:::1:59'
    PLANT = 'M0:46:::6:34'
    LAYOUT = 'M0:46:::18:34'
    SUBMIT_BTN = 'M0:50::btn[8]'
    CONTINUE_BTN = 'M1:50::btn[0]'
    PAGE_URL = _GetPage.BASE_URL + 'sap/bc/ui2/flp#Shell-startGUI?sap-ui2-tcode=LX02&sap-system=LOCAL'

    def __init__(self):
        super().__init__(self.PAGE_URL)
        self.fill_parameters = {
            self.WAREHOUSE_NUMBER: '270',
            self.STORAGE_TYPE_FROM: '110',
            self.STORAGE_TYPE_TO: '200',
            self.PLANT: '2310',
            self.LAYOUT: 'SLED/BBD',
        }

    def _fill_start_page(self):
        super()._fill_start_page()
        self.driver.find_element_by_id(self.SUBMIT_BTN).click()

    def _export_file(self):
        body = self.driver.find_element_by_tag_name('body')
        sleep(0.7)
        body.send_keys(Keys.SHIFT + Keys.F4)
        self.driver.find_element_by_id(self.CONTINUE_BTN).click()
        super()._export_file()


if __name__ == '__main__':
    zsd_oos = OutOfStock()
    zsd_oos.get_file()
    lx02 = Lx02()
    lx02.get_file()
