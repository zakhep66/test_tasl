import datetime

from openpyxl import load_workbook


wb_val = load_workbook(filename="task_support.xlsx", data_only=True)  # записываю в эту переменную файл excel

sheetVal = wb_val['Tasks']  # указываю какая страница мне нужна


def num1():
	"""
	:return: возчращает количество чётных элементов столбца
	"""
	count = 0
	for i in range(3, 1003):
		if sheetVal['B' + str(i)].value % 2 == 0:
			count += 1
	return count


def num2():
	"""
	:return: возвращает количество простых чисел
	"""
	def simple_dimple(arg):
		"""
		Проверка на простое число
		:param arg: получает на вход значение ячейки
		:return: возвращает единицу если число простое и ноль если не простое
		"""
		k = 0
		for j in range(2, arg // 2 + 1):
			if arg % j == 0:
				k += 1
				break
		if k == 0:
			return 1
		return 0

	count = 0
	for i in range(3, 1003):
		if simple_dimple(sheetVal['C' + str(i)].value) == 1:
			count += 1
	return count


def num3():
	"""
	:return: возвращает количество простых чисел в столбце
	"""
	def convert(arg):
		"""
		:param arg: получает на вход значение ячейки
		:return: возвращает понятное для питона число
		"""
		arg = arg.replace(' ', '')
		arg = arg.replace(',', '.')
		return float(arg)

	count = 0
	for i in range(3, 1003):
		if convert(sheetVal['D' + str(i)].value) < 0.5:
			count += 1
	return count


def date1():
	"""
	:return: возвращает количество вторников
	"""
	count = 0
	for i in range(3, 1003):
		if sheetVal['E' + str(i)].value[:3] == 'Thu':
			count += 1
	return count


def date2():
	def convert_date(arg):
		"""
		:param arg: получает значение ячейки
		:return: возвращает индекс дня недели где понедельник - 0 воскресенье - 6
		"""
		return datetime.datetime.strptime(arg, '%Y-%m-%d %H:%M:%S.%f').weekday()

	count = 0
	for i in range(3, 1003):
		if convert_date(sheetVal['F' + str(i)].value) == 1:
			count += 1
	return count


def date3():
	"""
	если это вторник прибавляем - неделю (7 дней), если это окажется следующий месяц значит этот вторник был последним
	:return: возвращает количество последних вторников месяца
	"""
	def convert_date(arg):
		"""
		:param arg: значение ячейки
		:return: дату в формате datetime
		"""
		return datetime.datetime.strptime(arg, '%m-%d-%Y')

	def concatenation_date(arg):
		"""
		:param arg: дата
		:return: следующую неделю с этого дня
		"""
		return arg + datetime.timedelta(days=7)

	count = 0
	for i in range(3, 1003):
		if convert_date(sheetVal['G' + str(i)].value).weekday() == 1:
			if concatenation_date(convert_date(sheetVal['G' + str(i)].value)).strftime('%m') != convert_date(
					sheetVal['G' + str(i)].value).strftime('%m'):
				count += 1
	return count
