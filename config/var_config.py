#!/usr/bin/env python3
# -*- coding: utf-8 -*-

""" 
File: var_config.py
Author: Gene Jiang
Email: zhengrong.jiang@chiefclouds.com
Github: https://github.com/yourname
Description: 
"""

import os
from utils.parse_excel import *


ieDriverFilePath = "D:\Selenium_driver\IEDriverServer.exe"
chromeDriverFilePath = "D:\Selenium_driver\chromedriver.exe"
firefoxDriverFilePath = "D:\Selenium_driver\geckodriver.exe"

# 获取当前文件所在目录的父目录的绝对 路径

parentDirPath = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# 异常图片存放目录
screenPicturesDir = parentDirPath + "\\exceptionpictures\\"

# 测试数据文件存放绝对路径
dataFilePath = parentDirPath + u"\\testData\\test_126.xlsx"

# 测试数据文件中，测试用例表中部分列对应的数字序号
testCase_testCaseName = 2
testCase_frameWorkName = 4
testCase_testStepSheetName = 5
testCase_dataSourceSheetName = 6
testCase_isExecute = 7
testCase_runTime = 8
testCase_testResult = 9

# 用例步骤表中，部分列对应的数字序号
testStep_testStepDescribe = 1
testStep_keyWords = 2
testStep_locationType = 3
testStep_locatorExpression = 4
testStep_operateValue = 5
testStep_runTime = 6
testStep_testResult = 7
testStep_errorInfo = 8
testStep_errorPic = 9

# 数据源表中， 是否执行列对应的数字编号
dataSource_isExecute = 6
dataSource_email = 2
dataSource_runTime = 7
dataSource_result = 8

# # Get the directory of project
# project_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
#
# # Get the directory of data
# data_dir = "{}\\{}\\".format(project_dir, 'data')
#
# # Test Data Path
# test_data_dir = "{}\\{}\\".format(project_dir, 'testdata')
#
# # Test Case Path
# test_case_file = "{}\\{}\\{}".format(project_dir,
#                                      'testdata', 'interface_automation.xlsx')
#
# # Get the path of the db config file
# db_config_path = "{}\\{}\\".format(project_dir, 'config')
#
# # Get the config file path
# db_config_file_path = "{}\\{}\\{}".format(project_dir,
#                                           'config', "db_config.ini")
#
# # get the base url and data from excel
# pe = ParseExcel()
# pe.load_work_book(test_case_file)
#
# # Get the config data for qinyuan
# sheet_name = pe.get_sheet_by_used_name("Config")
#
# DEFINE_TEST_SITE = 5            # Dev2 web site
#
# if DEFINE_TEST_SITE == 1:
#     # Test for Qinyuan
#     qinyuan_config_data = pe.get_row(sheet_name, row_no=2)
#
#     BASE_URL = qinyuan_config_data[1].value
#     login_url = qinyuan_config_data[2].value
#     data = eval(qinyuan_config_data[3].value)
# elif DEFINE_TEST_SITE == 2:
#     # Test for Miteno
#     miteno_config_data = pe.get_row(sheet_name, row_no=3)
#
#     BASE_URL = miteno_config_data[1].value
#     login_url = miteno_config_data[2].value
#     data = eval(miteno_config_data[3].value)
# elif DEFINE_TEST_SITE == 3:
#     # web dev platform
#     web_dev_config_data = pe.get_row(sheet_name, row_no=4)
#
#     BASE_URL = web_dev_config_data[1].value
#     login_url = web_dev_config_data[2].value
#     data = eval(web_dev_config_data[3].value)
# elif DEFINE_TEST_SITE == 4:
#     # PICC
#     picc_config_data = pe.get_row(sheet_name, row_no=5)
#
#     BASE_URL = picc_config_data[1].value
#     login_url = picc_config_data[2].value
#     data = eval(picc_config_data[3].value)
# elif DEFINE_TEST_SITE == 5:
#     # Web Dev2
#     web_dev2_config_data = pe.get_row(sheet_name, row_no=6)
#
#     BASE_URL = web_dev2_config_data[1].value
#     login_url = web_dev2_config_data[2].value
#     data = eval(web_dev2_config_data[3].value)
#     config_token = web_dev2_config_data[4].value
#     config_usid = web_dev2_config_data[5].value
#     print(type(data))


if __name__ == "__main__":
    # print('The project directory is %s' % project_dir)
    # print('The data directory is %s' % data_dir)
    # print("The directory of test data is {}".format(test_data_dir))
    # print("The directory of test case is {}".format(test_case_file))
    # print("The directory of config file is {}".format(db_config_file_path))
    pass
