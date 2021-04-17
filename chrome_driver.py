import pyautogui as gui
from selenium import webdriver
from datetime import date, timedelta
from selenium.webdriver.common.keys import Keys
from time import sleep
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec
from selenium.common.exceptions import NoSuchFrameException


class GetPage:
    def __init__(self, page_url):
        self.driver = webdriver.Chrome('C:\\Users\\by059491\\PycharmProjects\\JuiceCounter\\chromedriver.exe')
        self.driver.implicitly_wait(10)
        self.driver.get(page_url)
        self.driver.switch_to.frame('application-Shell-startGUI')
        self.fill_parameters = {}

    def fill_start_page(self):
        for id_key, id_val in self.fill_parameters.items():
            element = self.driver.find_element_by_id(id_key)
            element.clear()
            element.send_keys(id_val)


class OutOfStock(GetPage):
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

    def __init__(self, page_url, date_from=None, date_to=None):
        super().__init__(page_url)

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

    def fill_start_page(self):
        super().fill_start_page()
        for checkbox in self.checkboxes:
            checkbox.click()
        self.driver.find_element_by_id(self.SUBMIT_BTN).click()

    def export_file(self):
        self.driver.find_element_by_id(self.EXPORT_BTN).click()
        self.driver.find_element_by_id(self.SPREADSHEET).find_element_by_tag_name('tr').click()
        self.driver.find_element_by_id(self.CONTINUE_BTN).click()


class Lx02(GetPage):
    WAREHOUSE_NUMBER = 'M0:46:::0:34'
    STORAGE_TYPE_FROM = 'M0:46:::1:34'
    STORAGE_TYPE_TO = 'M0:46:::1:59'
    PLANT = 'M0:46:::6:34'
    LAYOUT = 'M0:46:::18:34'
    SUBMIT_BTN = 'M0:50::btn[8]'
    CONTINUE_BTN = 'M1:50::btn[0]'

    def __init__(self, page_url):
        super().__init__(page_url)
        self.fill_parameters = {
            self.WAREHOUSE_NUMBER: '270',
            self.STORAGE_TYPE_FROM: '110',
            self.STORAGE_TYPE_TO: '200',
            self.PLANT: '2310',
            self.LAYOUT: 'SLED/BBD',
        }

    def fill_start_page(self):
        super().fill_start_page()
        self.driver.find_element_by_id(self.SUBMIT_BTN).click()

    def export_file(self):
        body = self.driver.find_element_by_tag_name('body')
        sleep(0.7)
        body.send_keys(Keys.SHIFT + Keys.F4)
        self.driver.find_element_by_id(self.CONTINUE_BTN).click()


if __name__ == '__main__':
    zsd_oos = OutOfStock('https://cuvl0301.eur.cchbc.com:8204/sap/bc/ui2/flp#Shell-startGUI?sap-ui2-tcode=ZSD_OOS&sap-system=LOCAL')
    zsd_oos.fill_start_page()
    zsd_oos.export_file()
    lx02 = Lx02('https://cuvl0301.eur.cchbc.com:8204/sap/bc/ui2/flp#Shell-startGUI?sap-ui2-tcode=LX02&sap-system=LOCAL')
    lx02.fill_start_page()
    lx02.export_file()