import pyautogui as gui
from selenium import webdriver
from datetime import date, timedelta
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec
from selenium.common.exceptions import NoSuchFrameException

from time import sleep

driver = webdriver.Chrome('C:\\Users\\by059491\\PycharmProjects\\JuiceCounter\\chromedriver.exe')
driver.implicitly_wait(10)


class GetPage:
    def __init__(self, page_url):
        driver.get(page_url)
        self.fill_parameters = {}
        self.submit = None

    def fill_start_page(self):
        for id_key, id_val in self.fill_parameters.items():
            element = driver.find_element_by_id(id_key)
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
        driver.switch_to.frame('application-Shell-startGUI')
        self.date_from = date.today() + timedelta(days=1)
        self.date_to = date_to
        self.fill_parameters = {
            self.SALES_ORG: '2300',
            self.DEL_DATE_FROM: self.date_from.strftime('%d.%m.%Y'),
            self.DOC_TYPE: 'ZOR',
        }
        self.checkboxes = [
            driver.find_element_by_id(self.CHECKBOX_ONE),
            driver.find_element_by_id(self.CHECKBOX_TWO),
        ]
        if date_from is not None:
            self.fill_parameters[self.DEL_DATE_FROM] = date_from
        if date_to is not None:
            self.fill_parameters.setdefault(self.DEL_DATE_TO, self.date_to)

    def fill_start_page(self):
        super().fill_start_page()
        for checkbox in self.checkboxes:
            checkbox.click()
        submit = driver.find_element_by_id(self.SUBMIT_BTN)
        submit.click()

    def export_file(self):
        driver.find_element_by_id(self.EXPORT_BTN).click()
        driver.find_element_by_id(self.SPREADSHEET).find_element_by_tag_name('tr').click()
        driver.find_element_by_id(self.CONTINUE_BTN).click()


# SALES_ORG_ID = 'M0:46:::1:34'
# DEL_DATE_FROM_ID = 'M0:46:::3:34'
# DEL_DATE_TO_ID = 'M0:46:::3:59'
# DOC_TYPE_ID = 'M0:46:::6:34'
# CHECKBOX_ONE_ID = 'M0:46:::16:3'
# CHECKBOX_TWO_ID = 'M0:46:::17:3'
# 'M0:50::btn[8]-cnt'
# 'thead'
# 'menu_MB_EXPORT102_1_1-r' # cписок
# 'u1BC56D'
# EXPORT_BTN_ID = '_MB_EXPORT102-r'
#
#
#
#
# driver.get('https://cuvl0301.eur.cchbc.com:8204/sap/bc/ui2/flp#Shell-startGUI?sap-ui2-tcode=ZSD_OOS&sap-system=LOCAL')
# driver.switch_to.frame(FRAME)
#
#
#
# checkbox_1 = driver.find_element_by_id(CHECKBOX_ONE_ID)
# checkbox_1.click()
# checkbox_2 = driver.find_element_by_id(CHECKBOX_TWO_ID)
# checkbox_2.click()
# submit = driver.find_element_by_id('M0:50::btn[8]')
# submit.click()
# export = driver.find_element_by_id(EXPORT_BTN_ID)
# export.click()
# spreadsheet = driver.find_element_by_id('menu_MB_EXPORT102_1_1-r').find_element_by_tag_name('tr')
# spreadsheet.click()
# continue_btn = driver.find_element_by_id('M1:50::btn[0]')
# continue_btn.click()


zsd_oos = OutOfStock('https://cuvl0301.eur.cchbc.com:8204/sap/bc/ui2/flp#Shell-startGUI?sap-ui2-tcode=ZSD_OOS&sap-system=LOCAL')
zsd_oos.fill_start_page()
zsd_oos.export_file()
