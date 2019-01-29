#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
File: WaitUtil.py
Author: Gene Jiang
Email: zhengrong.jiang@chiefclouds.com
Description: Wait to display the elements smartly
"""

from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


class WaitUtil(object):
    """
    Wait to display the elements smartly
    """
    def __init__(self, driver):
        self.location_type_dict = {
            "xpath": By.XPATH,
            "id": By.ID,
            "name": By.NAME,
            "css_selector": By.CSS_SELECTOR,
            "class_name": By.CLASS_NAME,
            "tag_name": By.TAG_NAME,
            "link_text": By.LINK_TEXT,
            "partial_link_text": By.PARTIAL_LINK_TEXT
        }
        # initialize the driver object
        self.driver = driver
        #  Wait the elements explicitly
        self.wait = WebDriverWait(self.driver, 10)

    def presenceOfElementLocated(self, locatorMethod, locatorExpression, *arg):
        """
        显式等待页面元素出现在DOM中，但并不一定可见，存在则返回该页面元素对象
        """
        try:
            if self.location_type_dict.has_key(locatorMethod.lower()):
                element = self.wait.until(EC.presence_of_element_located((
                    self.location_type_dict[locatorMethod.lower()],
                    locatorExpression)))
                return element
            else:
                raise TypeError(u"未找到定位方式，请确认定位方法是否写正确")
        except Exception as e:
            raise e

    def frameToBeAvailableAndSwitchToIt(self, locationType, locatorExpression,
                                        *args):
        """
        检查frame是否存在， 存在则切换进frame控件 中
        :param locationType:
        :param locatorExpression:
        :param args:
        :return:
        """
        try:
            self.wait.until(EC.frame_to_be_available_and_switch_to_it((
                self.location_type_dict[locationType.lower()],
                locatorExpression)))
        except Exception as e:
            raise e

    def visibilityOfElementLocated(self, locationType, locatorExpression,
                                   *args):
        """
        显式等待页面元素出现在DOM 中， 并且可见， 存在返回该页面元素对象
        :param locationType:
        :param locatorExpression:
        :param args:
        :return:
        """
        ''''''

        try:
            element = self.wait.until(EC.visibility_of_element_located((
                self.location_type_dict[locationType.lower()],
                locatorExpression)))
            return element
        except Exception as e:
            raise e

    def locate_presence_element(self, location_type,
                                 locator_expression, *arg):
        """
        Description: Wait to display the elements explicitly, but the element
        may be not existed. if the elements is existed, return the elements

        :locator_method: xpath, id, name, css_selector, class_name etc

        :locator_expression:xpath value, css_selector value etc
        
        :args:
        """
        try:
            if location_type.lower() in self.location_type_dict:
                element = self.wait.until(
                      EC.presence_of_element_located((
                        self.location_type_dict[location_type.lower()],
                        locator_expression
                    ))
                )

                return element
            else:
                raise TypeError("The locator method is not found!")
        except Exception as e:
            raise e

    def locate_visibility_element(self, location_type,
                                  locator_expression, *args):
        """
        Description: Wait to display the elements explicitly, but the element
        may be not existed. if the elements is existed, return the elements

        :locator_method: xpath, id, name, css_selector, class_name etc

        :locator_expression:xpath value, css_selector value etc

        :args:
        """
        try:
            element = self.wait.until(
                EC.visibility_of_element_located((
                    self.location_type_dict[location_type.lower()],
                    locator_expression
                ))
            )

            return element

        except Exception as e:
            raise e


if __name__ == "__main__":
    from selenium import webdriver
    import time

    # 进行单元测试

    driver = webdriver.Chrome(
        executable_path=r"d:\Selenium_driver\Chromedriver.exe")

    driver.get("http://mail.126.com")
    waitUtil = WaitUtil(driver)
    time.sleep(5)

    driver.find_element_by_xpath("//*[contains(@id,'x-URS-iframe')]")

    waitUtil.frameToBeAvailableAndSwitchToIt("xpath", "//*[starts-with(@id,'x-URS-iframe']")
    waitUtil.visibilityOfElementLocated("xpath", "//input@name='email'")
    driver.quit()

    waitUtil = WaitUtil(driver)
    element_visible_one = waitUtil.locate_presence_element("id", "kw")
    element_visible_two = waitUtil.locate_visibility_element("id", "kw")

    print(element_visible_one)
    print(element_visible_two)

    driver.quit()



