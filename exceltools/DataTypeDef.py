# -*- coding:utf-8 -*-
class DataTypeBase(object):
	def __init__(self, default):
		self.default = default

	def setValue(self, v):
		self.value = v

class DInt(DataTypeBase):
	def __init__(self, default = 0):
		super(DInt, self).__init__(default)

	def convert(self):
		return int(getattr(self, 'value', self.default))

class DFloat(DataTypeBase):
	def __init__(self, default = 0.0):
		super(DFloat, self).__init__(default)

	def convert(self):
		return float(getattr(self, 'value', self.default))

class DStr(DataTypeBase):
	def __init__(self, default = ''):
		super(DStr, self).__init__(default)

	def convert(self):
		return unicode(getattr(self, 'value', self.default))

class DDate(DataTypeBase):
	def __init__(self, default = '2011-10-1'):
		self.default = default

	def convert(self):
		v = getattr(self, 'value', self.default)
		data = v.split('-')
		while len(data) < 3:
			data.append(u'0')
		if len(data) > 3:
			data = data[0:3]
		return unicode(u'-'.join(data))
