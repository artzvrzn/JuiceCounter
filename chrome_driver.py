import pyautogui as gui
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec
from selenium.common.exceptions import NoSuchFrameException
from time import sleep


FRAME = 'application-Shell-startGUI'
TABLE_IDS = {
    'M0:46:::1:34': '2300',
    'M0:46:::3:34': '17.04.2021',
    'M0:46:::6:34': 'ZOR',
}
SALES_ORG_ID = 'M0:46:::1:34'
DEL_DATE_FROM_ID = 'M0:46:::3:34'
DEL_DATE_TO_ID = 'M0:46:::3:59'
DOC_TYPE_ID = 'M0:46:::6:34'
CHECKBOX_ONE_ID = 'M0:46:::16:3'
CHECKBOX_TWO_ID = 'M0:46:::17:3'
'M0:50::btn[8]-cnt'
'thead'
'menu_MB_EXPORT102_1_1-r' # cписок
'u1BC56D'
EXPORT_BTN_ID = '_MB_EXPORT102-r'



driver = webdriver.Chrome('C:\\Users\\by059491\\PycharmProjects\\JuiceCounter\\chromedriver.exe')
driver.implicitly_wait(10)
driver.get('https://cuvl0301.eur.cchbc.com:8204/sap/bc/ui2/flp#Shell-startGUI?sap-ui2-tcode=ZSD_OOS&sap-system=LOCAL')
driver.switch_to.frame(FRAME)

for id_key, id_val in TABLE_IDS.items():
    element = driver.find_element_by_id(id_key)
    element.clear()
    element.send_keys(id_val)

checkbox_1 = driver.find_element_by_id(CHECKBOX_ONE_ID)
checkbox_1.click()
checkbox_2 = driver.find_element_by_id(CHECKBOX_TWO_ID)
checkbox_2.click()
submit = driver.find_element_by_id('M0:50::btn[8]')
submit.click()
export = driver.find_element_by_id(EXPORT_BTN_ID)
export.click()
spreadsheet = driver.find_element_by_id('menu_MB_EXPORT102_1_1-r').find_element_by_tag_name('tr')
spreadsheet.click()
continue_btn = driver.find_element_by_id('M1:50::btn[0]')
continue_btn.click()

