from selenium import webdriver
import openpyxl
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from time import sleep

blacklistalert_url = "https://blacklistalert.org/"

driver = webdriver.Chrome("./chromedriver")
wait = WebDriverWait(driver=driver, timeout=60)
driver.get(blacklistalert_url)

#読み込みexcelファイル
wb = openpyxl.load_workbook("ip.xlsx")
ws = wb.worksheets[0]
#書き込みexcelファイル
wb2 = openpyxl.load_workbook("res.xlsx")
ws2 = wb2.worksheets[0]

ip_list = []
index = 1

for col in ws["A"]:
  ip_list.append(col.value.splitlines()[0])

for ip in ip_list:
  sleep(2)
  form_xpath = '/html/body/center/font/form/input[1]'
  click_button_xpath = '/html/body/center/font/form/input[3]'

  form = driver.find_element(By.XPATH, form_xpath)
  form.clear()
  form.send_keys(ip)
  click_button = driver.find_element(By.XPATH,click_button_xpath)
  click_button.click()
  print("clicked")
  wait.until(EC.text_to_be_present_in_element((By.TAG_NAME, 'body'),'sponsored'))
  ip_tables = driver.find_elements(By.TAG_NAME,'table')
  trs = ip_tables[0].find_elements(By.TAG_NAME, "tr")
  
  NG_count = 0
  NG_sites = []
  trs_count = 0
  for tr in trs:
    trs_count += 1
    site = tr.find_element(By.CLASS_NAME,'left').text
    status = tr.find_elements(By.TAG_NAME,"strong")
    #Listedかどうか
    if len(status) == 0:
      status = "NG"
    else:
      status = "OK"
    if status == "NG":
      NG_count += 1
      NG_sites.append(site)
  #table3つ目が存在する場合これも追加
  if len(ip_tables) == 3:
    trs = ip_tables[2].find_elements(By.TAG_NAME, "tr")
    for tr in trs:
      trs_count += 1
      site = tr.find_element(By.CLASS_NAME,'left').text
      status = tr.find_elements(By.TAG_NAME,"strong")
      #Listedかどうか
      if len(status) == 0:
        status = "NG"
      else:
        status = "OK"
      if status == "NG":
        NG_count += 1
        NG_sites.append(site)

  ws2.cell(row=index,column=1).value = ip
  if NG_count != 0:
    ws2.cell(row=index,column=3).value = ", ".join(NG_sites)
    ws2.cell(row=index,column=2).value = "NG"
    ws2[f'B{index}'].font = openpyxl.styles.fonts.Font(color='FF0000')
  else:
    ws2.cell(row=index,column=2).value = "OK"
    ws2[f'B{index}'].font = openpyxl.styles.fonts.Font(color='00793D')
  wb2.save("./res.xlsx")
  index += 1
  print(trs_count)
