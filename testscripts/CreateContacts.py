#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
File: CreateContacts.py
Author: Gene Jiang
Email: genejiang2012@outlook.com
Github: https://github.com/genejiang2012
Description: 
"""

from . import *
from testscripts.WriteTestResult import writeTestResult

def dataDriverFun(dataSourceSheetObj, stepSheetObj): 
    try: 
        # 获取数据源表中是否执行列对象
        dataIsExecuteColumn = excelObj.getColumn(dataSourceSheetObj, 
                                                 dataSource_isExecute) 
        # 获取数据源表中"电子邮箱"列对象 
        emailColumn = excelObj.getColumn(dataSourceSheetObj, dataSource_email) 
        
        # 获取测试步骤表中存在数据区域的行数
        stepRowNums = excelObj.getRowsNumber(stepSheetObj) 
        
        # 记录成功执行的数据条数
        successDatas = 0 
        
        # 记录被设置为执行的数据条数
        requiredDatas = 0 
        
        for idx, data in enumerate(dataIsExecuteColumn[1:]): 
            
            # 遍历数据源表， 准备进行数据驱动测试 
            # 因为第一行是标题行，所以从第二行开始遍历 
            if data.value == "y": 
                print ("开始添加联系人%s"%emailColumn[idx + 1].value)
                
                requiredDatas += 1 
                
                # 定义记录执行成功步骤数变量
                successStep = 0 
                
                for index in range(2, stepRowNums + 1):
                       # 获取数据驱动测试步骤表中 
                       # 第index行对象 
                       rowObj = excelObj.getRow(stepSheetObj, index)
                       excelObj.writeCell(sheetObj, content = "", 
                                          rowNo = rowNo, 
                                          colsNo = colsDict[colsNo][0]) 
            else: 
                # 在测试步骤sheet中， 写入测试时间
                excelObj.writeCellCurrentTime(sheetObj, rowNo = rowNo, 
                                              colsNo = colsDict[colsNo][0]) 
            
            if errorInfo and picPath: 
                # 在测试步骤sheet中， 写入异常信息
                excelObj.writeCell(sheetObj, content = errorInfo, 
                                   rowNo = rowNo, colsNo = testStep_errorInfo) 
                # 在测试步骤sheet中，写入异常截图路径
                excelObj.writeCell(sheetObj, content = picPath, 
                                   rowNo = rowNo, colsNo = testStep_errorPic) 
            else: 
                if colsNo == "caseStep": 
                       # 在测试步骤sheet中，清空异常信息单元格
                       excelObj.writeCell(sheetObj, content = "", 
                                          rowNo = rowNo, 
                                          colsNo = testStep_errorInfo) 
                       # 在测试步骤sheet中，清空异常信息单元格
                       excelObj.writeCell(sheetObj, content = "", 
                                          rowNo = rowNo, 
                                          colsNo = testStep_errorPic)
    except Exception as e: 
        print("写excel时发生异常")
        print(traceback.print_exc())



