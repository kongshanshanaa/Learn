from selenium import webdriver
driver = webdriver.Chrome
driver.get()
driver.find_element_by_xpath('//*[@id="login_img"]').click()
driver.find_element_by_xpath('').send_keys('')
