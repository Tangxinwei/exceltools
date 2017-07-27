# -*- coding:utf-8 -*-
import os
import xlrd
import excelconst
import codecs
import csv

def convertTuple(d, ret):
	ret.append('(')
	for value in d:
		convertObject(value, ret)
		ret.append(',')
	ret.append(')')

def convertList(d, ret):
	ret.append('[')
	for value in d:
		convertObject(value, ret)
		ret.append(',')
	ret.append(']')

def convertDict(d, ret):
	ret.append('{')
	for key, value in d.iteritems():
		convertObject(key, ret)
		ret.append(':')
		convertObject(value, ret)
		ret.append(',')
	ret.append('}')

def convertObject(d, ret):
	if type(d) is dict:
		convertDict(d, ret)
	elif type(d) is list:
		convertList(d, ret)
	elif type(d) is tuple:
		convertTuple(d, ret)
	elif type(d) in (unicode, str):
		ret.append(repr(d))
	elif type(d) in (float, int):
		ret.append(str(d))
	else:
		raise Exception('conver object error %s %s' % (type(d), d))

def dealWithRow(table, row, sheetModule):
	if row:
		if sheetModule.INPUT_COLLAYOUT == excelconst.TableLayOut.AS_LIST:
			table.append(row)

def convert(excelName, sheetModules, commonPaths):
	if excelName.endswith('csv'):
		_convertCsv(excelName, sheetModules, commonPaths)
	else:
		_convertExcel(excelName, sheetModules, commonPaths)

class CsvSheet(object):
	def __init__(self, allRows):
		self.allRows = allRows
		self.nrows = len(allRows)

	def row_values(self, idx):
		return self.allRows[idx]

	def cell_type(self, i, j):
		return xlrd.XL_CELL_TEXT

class CsvBook(object):
	pass

def _convertCsv(csvName, sheetModules, commonPaths):
	#csv has only one sheet
	csvName = os.path.join(commonPaths.INPUT_EXCEL_PATH, csvName)
	sheetModule = sheetModules[0]
	f = open(csvName)
	sheet = csv.reader(f)
	allRow = []
	for row in sheet:
		for idx, m in enumerate(row):
			row[idx] = row[idx].decode(getattr(sheetModule, 'ENCODING', 'gbk'))
		allRow.append(row)
	f.close()
	_convertSheet(csvName, sheetModule, commonPaths, CsvSheet(allRow), CsvBook())


def _convertExcel(excelName, sheetModules, commonPaths):
	excelPath = os.path.join(commonPaths.INPUT_EXCEL_PATH, excelName)
	book = xlrd.open_workbook(excelPath)
	print 'convert excel', excelPath
	for sheetModule in sheetModules:
		if getattr(sheetModule, 'SHEET_NAME', None):
			sheet = book.sheet_by_name(sheetModule.SHEET_NAME)
		else:
			sheet = book.sheet_by_index(sheetModule.SHEET_IDX)
		_convertSheet(excelName, sheetModule, commonPaths, sheet, book)

def _convertSheet(excelName, sheetModule, commonPaths, sheet, book):
	if sheetModule.INPUT_COLLAYOUT == excelconst.TableLayOut.AS_LIST:
		table = []

	if sheetModule.INPUT_ROWLAYOUT == excelconst.TableLayOut.AS_DICT:
		currentRowIdx = 0
		idxMap = {}
		if getattr(sheetModule, 'CUSTOM_KEY_MAP', {}):
			keyValueMap = sheetModule.CUSTOM_KEY_MAP
			excelRowValues = sheet.row_values(0)
			for j in xrange(len(excelRowValues)):
				value = unicode(excelRowValues[j])
				if value in keyValueMap:
					idxMap[j] = keyValueMap[value]
			currentRowIdx += 1
		elif getattr(sheetModule, 'CUSTOM_KEY_TUPLE', {}):
			idxMap = sheetModule.CUSTOM_KEY_TUPLE
		currentRowIdx += getattr(sheetModule, 'SKIP_ROW_NUMBER', 0)

		for i in xrange(currentRowIdx, sheet.nrows):
			row = {}
			excelRowValues = sheet.row_values(i)
			for j, (keyName, dataConvert) in idxMap.iteritems():
				v = excelRowValues[j]
				if sheet.cell_type(i, j) == xlrd.XL_CELL_DATE:
					dateTuple = xlrd.xldate_as_tuple(v, book.datemode)
					v = '-'.join([str(data) for data in dateTuple])
				v = unicode(v)
				if v:
					dataConvert.setValue(v)
					row[keyName] =  dataConvert.convert()
			if getattr(sheetModule, 'convert_row', None):
				row = sheetModule.convert_row(row)
			dealWithRow(table, row, sheetModule)
	elif sheetModule.INPUT_ROWLAYOUT == excelconst.TableLayOut.AS_LIST:
		currentRowIdx = getattr(sheetModule, 'SKIP_ROW_NUMBER', 0)
		idxMap = sheetModule.CUSTOM_KEY_TUPLE
		for i in xrange(currentRowIdx, sheet.nrows):
			row = [None] * len(sheetModule.CUSTOM_VALUE_TUPLE)
			excelRowValues = sheet.row_values(i)
			for j in xrange(len(sheetModule.CUSTOM_VALUE_TUPLE)):
				v = excelRowValues[j]
				if sheet.cell_type(i, j) == xlrd.XL_CELL_DATE:
					dateTuple = xlrd.xldate_as_tuple(v, book.datemode)
					v = '-'.join([str(data) for data in dateTuple])
				v = unicode(v)
				dataConvert = sheetModule.CUSTOM_VALUE_TUPLE[j]
				if v:
					dataConvert.setValue(v)
				row[j] = dataConvert.convert()
			if getattr(sheetModule, 'convert_row', None):
				row = sheetModule.convert_row(row)
			dealWithRow(table, row, sheetModule)
		
	if getattr(sheetModule, 'convert_table', None):
		table = sheetModule.convert_table(table)

	if table:
		saveFile = os.path.join(commonPaths.DATA_PATH, sheetModule.SAVE_FILE)
		f = codecs.open(saveFile, 'w', 'utf-8')
		f.write('# -*- coding:utf-8 -*-\n')
		if sheetModule.INPUT_COLLAYOUT == excelconst.TableLayOut.AS_LIST:
			f.write('data = [\n')
			for row in table:
				f.write('%s,\n' % row)
			f.write(']')
		f.close()
