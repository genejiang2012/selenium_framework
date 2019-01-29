# encoding= utf-8

import time
import traceback

from pageaction.PageAction import *
from utils.parse_excel import ParseExcel
from config.var_config import *


# 创建解析 Excel对象
excelObj = ParseExcel()
# 将Excel数据文件加载到内存
excelObj.loadWorkBook(dataFilePath)
