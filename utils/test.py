import time
from selenium import webdriver
from selenium.webdriver import ActionChains
from selenium.webdriver.common.keys import Keys

driver = webdriver.Chrome()
driver.maximize_window()
driver.get('http://web-dev2.chiefclouds.com/!/audience/tag/first?groupId=all')
time.sleep(10)

element = driver.find_element_by_xpath("//*[@id='all_anchor']")
actionChains = ActionChains(driver)
actionChains.context_click(element).send_keys(Keys.ARROW_DOWN).send_keys(
    Keys.ENTER).perform()
