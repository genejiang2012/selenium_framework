#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
File: object_map.py
Author: Gene Jiang
Email: genejiang2012@outlook.com
Github: https://github.com/genejiang2012
Description: 
"""

from selenium.webdriver.support.ui import WebDriverWait


# 获取单个页面元素对象 
def getElement(driver, locationType, locatorExpression): 
    try: 
        element = WebDriverWait(driver, 30).until(
            lambda x: x.find_element(by=locationType, value=locatorExpression)) 
        return element 
    except Exception as e: 
        raise e


# 获取多个相同页面元素对象，以 list 返回 
def getElements(driver, locationType, locatorExpression): 
    try:
        elements = WebDriverWait(driver, 30). until(
            lambda x: x.find_elements(by=locationType, value=locatorExpression)) 
        
        return elements
    except Exception as e: 
        raise e


if __name__ == '__main__': 
    from selenium import webdriver 
    
    # 进行单元测试 
    driver = webdriver.Firefox(
        executable_path=r"d:\Selenium_driver\geckodriver.exe")
    driver.get("http://www.baidu.com") 
    searchBox = getElement(driver, "id", "kw")
    # 打印页面对象的标签
    print(searchBox.tag_name) 
    aList = getElements(driver, "tag name", "a")
    print(aList)
    print(len(aList))
    driver.quit()


