#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
File: WriteTestResult.py
Author: Gene Jiang
Email: genejiang2012@outlook.com
Github: https://github.com/genejiang2012
Description: 
"""

import time
import traceback

from pageaction.PageAction import *
from utils.parse_excel import ParseExcel
from config.var_config import *


# 创建解析 Excel对象
excelObj = ParseExcel()
# 将Excel数据文件加载到内存
excelObj.loadWorkBook(dataFilePath)


def writeTestResult(sheetName, rowNo, colsNo, testResult, errorInfo = None,
                    picPath = None):
    """
    用例或用例步骤执行结束后,向Excel中写执行结果信息
    :param sheetObj:
    :param rowNo:
    :param colsNo:
    :param testResult:
    :param errorInfo:
    :param picPath:
    :return:
    """

    # 测试通过结果信息为绿色， 失败为红色 
    colorDict = {"pass": 'green', "faild": "red", "": None}
    
    # 因为"测试用例"工作表和"用例步骤sheet表"中都有测试执行时间和
    # 测试结果列，定义此字典对象是为了区分具体应该写哪个工作表 
    
    colsDict = {"testCase": [testCase_runTime, testCase_testResult],
                "caseStep": [testStep_runTime, testStep_testResult],
                "dataSheet": [dataSource_runTime, dataSource_result]}

    try:
        # 在测试步骤sheet中，写入测试结果
        sheet_obj = excelObj.getSheetByName(sheetName)

        excelObj.writeCell(sheet_obj, content=testResult, rowNo=rowNo,
                           colsNo=colsDict[colsNo][1],
                           style=colorDict[testResult])
        if testResult == "": 
            # 清空时间单元格内容 
            excelObj.writeCell(sheet_obj, content="", rowNo=rowNo,
                               colsNo=colsDict[colsNo][0])
        else: 
            # 在测试步骤sheet中，写入测试时间
            excelObj.writeCellCurrentTime(sheet_obj, rowNo=rowNo,
                                          colsNo=colsDict[colsNo][0]) 
        
        if errorInfo and picPath: 
            # 在测试步骤sheet中，写入异常信息
            excelObj.writeCell(sheet_obj, content=errorInfo, rowNo=rowNo,
                               colsNo=testStep_errorInfo) 
            
            # 在测试步骤sheet中， 写入异常截图路径
            excelObj.writeCell(sheet_obj, content=picPath, rowNo=rowNo,
                               colsNo=testStep_errorPic)
        else: 
            if colsNo == "caseStep": 
                # 在测试步骤sheet中， 清空异常信息单元格 
                excelObj.writeCell(sheet_obj, content="", rowNo=rowNo,
                                   colsNo=testStep_errorInfo)
                # 在 试步骤sheet中， 清空异常信息单元格 
                excelObj.writeCell(sheet_obj, content="", rowNo=rowNo,
                                   colsNo=testStep_errorPic) 
    except Exception as e:
        print(u"写excel时,发生异常")
        traceback.print_exc()


if __name__ == "__main__":
    writeTestResult("TestCase", rowNo=2, colsNo="testCase",
                    testResult="pass")
