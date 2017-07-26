# -*- coding:utf-8 -*-
from exceltools import excelconst
from exceltools.DataTypeDef import *

#sheet的下标
#SHEET_IDX和SHEET_NAME二选一,必填
SHEET_IDX = 0
#csv的编码,仅对csv生效
ENCODING = 'gbk'

#输入每一行的格式,当选择为Dict的时候  这种情况只有定义了key而且excel中真的有值的列,才会被导出
#									如果定义了CUSTOM_KEY_TUPLE,那么excel的每一列都会使用CUST_KEY里面的内容
#									如果定义了CUSTOM_KEY_MAP,那么会把excel的第一行映射到一个字符串上,并跳过第一行
#				当选择为List的时候
#									需要定义CUSTOM_VALUE_TUPLE,用于定义每一行的字符值(这种情况会默认每一行都是满的,没填则是默认值)
INPUT_ROWLAYOUT = excelconst.TableLayOut.AS_DICT
CUSTOM_KEY_MAP = {
				u'测试' : (u'test', DFloat()),
				u'测试日期': (u't', DDate()),
}

'''CUSTOM_KEY_TUPLE = {
	0 : (u'test', DFloat()),
	1 : (u't', DStr()),
}'''

#跳过前面几行
SKIP_ROW_NUMBER = 0

#输入每一列的格式,当选择为List的时候,每一行都会被存为list中的一项
INPUT_COLLAYOUT = excelconst.TableLayOut.AS_LIST

#保存的路径名
SAVE_FILE = 'test.py'

#自定义的每一行的转换
def convert_row(rowObj):
	return rowObj

#自定义的导表转换结束之后的操作
def convert_table(table):
	return table


