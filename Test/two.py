from selenium import webdriver
from time import sleep
from datetime import datetime
import xlrd

book = xlrd.open_workbook('test.xlsx')
sheet = book.sheet_by_name('login_info')
values = sheet.row_values(0)
driver = webdriver.Chrome()
driver.get(values[0])
driver.find_element_by_xpath('//*[@id="login_img"]').click()
driver.find_element_by_xpath('//*[@id="myForm"]/table/tbody/tr/td/table/tbody/tr[3]/td[3]/input').send_keys(values[1])
driver.find_element_by_xpath('//*[@id="myForm"]/table/tbody/tr/td/table/tbody/tr[5]/td[3]/input').send_keys(
    int(values[2]))
driver.find_element_by_xpath('//*[@id="myForm"]/table/tbody/tr/td/table/tbody/tr[7]/td[3]/a[1]/img').click()

sleep(2)
# 新建策略
left_frame = driver.find_element_by_xpath('//*[@id="oa_left_middle"]')
driver.switch_to.frame(left_frame)
driver.find_element_by_xpath('//*[@id="outlooktitle3"]/table/tbody/tr/td').click()
driver.find_element_by_xpath('//*[@id="outlookdivin3"]/span[1]/table/tbody/tr/td/a').click()
driver.switch_to.default_content()
sleep(2)

date_format = '%H%M%S'
date = datetime.strftime(datetime.now(), date_format)  # 调用datetime里面的strftime 方法 把时间格式化
oa_main = driver.find_element_by_xpath('//*[@id="oa_main"]')
driver.switch_to.frame(oa_main)
driver.find_element_by_xpath('//*[@id="addPolicyInfoBtn"]/span/span').click()
driver.find_element_by_xpath('//*[@id="addPolicyInfoDialog"]/div[1]/div/div/form/div/fieldset/div/input[3]'). \
    send_keys(date)
driver.find_element_by_xpath('//*[@id="addPolicyInfoDialog"]/div[2]/a[1]/span/span').click()
driver.find_element_by_xpath('/html/body/div[13]/div[2]/div[4]/a/span/span').click()

# 配置策略
driver.find_element_by_xpath('//*[@id="queryPolicyInfoBtn"]/span/span').click()
driver.find_element_by_xpath('//*[@id="policyInfoQueryDialog"]/div[1]/div/div/form/div/fieldset/div/input[1]'). \
    send_keys(date)
sleep(2)
driver.find_element_by_xpath('//*[@id="policyInfoQueryDialog"]/div[2]/a[1]/span/span').click()
sleep(2)

driver.find_element_by_xpath('//*[@id="datagrid-row-r1-2-0"]/td[6]/div/a').click()
panel_noscroll = driver.find_element_by_xpath('/html/body/div[15]')
iframes = panel_noscroll.find_elements_by_tag_name('iframe')
driver.switch_to.frame(iframes[0])
sleep(2)
driver.find_element_by_xpath('//*[@id="tabsok"]/ul/li[2]/a/b').click()
