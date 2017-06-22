# -*- coding:utf-8 -*-
class CommonPath(object):
	def __init__(self):
		self.DATA_PATH = 'data'
		self.INPUT_EXCEL_PATH = 'inputexcel'
commonPath = CommonPath()
import sys
sys.path.append(commonPath.DATA_PATH)

import exceltools
from defs import tableDef_Sheet

exceltools.convert(u'测试.xlsx', [tableDef_Sheet], commonPath)
