import selenium
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from openpyxl import Workbook, load_workbook 
import time

service = Service()
option = webdriver.ChromeOptions()
driver = webdriver.Chrome(service=service, options=option)
wb = load_workbook('apartments-list.xlsx')
ws = wb.active
max_row = ws.max_row

url = "https://ss.lv"
delay = 1
max_price = 80000
number_of_rooms = "2"

driver.get(url)
time.sleep(delay)

find = driver.find_element(By.ID, "mtd_59").click()
time.sleep(delay)
find = driver.find_element(By.ID, "ahc_14195").click()
time.sleep(delay)
find = driver.find_element(By.ID, "ahc_1088").click()
time.sleep(delay)
find = driver.find_element(By.CLASS_NAME, "l100")
select = Select(find)
select.select_by_visible_text("Pārdod")
time.sleep(delay)
find = driver.find_element(By.NAME, "topt[1][min]")
select = Select(find)
select.select_by_visible_text(number_of_rooms)
find = driver.find_element(By.NAME, "topt[1][max]")
select = Select(find)
select.select_by_visible_text(number_of_rooms)
time.sleep(delay)
find = driver.find_element(By.ID, "f_o_8_max").send_keys(max_price)
find = driver.find_element(By.CLASS_NAME, "s12").click()
time.sleep(delay)


table = driver.find_element(By.XPATH, "//form/table[@align='center']")
rows = table.find_elements(By.TAG_NAME, "tr")

table_data = []

for row in rows[1:-1]:
    cells = row.find_elements(By.TAG_NAME, "td")[2:]
    row_data = [cell.text for cell in cells]
    table_data.append(row_data)
    ws.append(row_data)

wb.save('apartments-list.xlsx')
wb.close()
print(table_data)