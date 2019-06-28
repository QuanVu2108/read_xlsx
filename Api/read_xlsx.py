import openpyxl
import os

JSON_MIME_TYPE = 'application/json'

dir_path = os.path.dirname(os.path.realpath(__file__))
xlsx_path = dir_path + '/temp/Test.xlsx'
book = openpyxl.load_workbook(xlsx_path)
sheet = book.active

def name_int_chr(col_id):
	if col_id < 27:
		return chr(64 + col_id )
	else:
		return chr(64 + col_id/26) + chr(64 + col_id%26)
	
def name_chr_int(col_id):
	if len(col_id) == 1:
		return ord(col_id) - 64 -1

def check_col(col_id, lim):
	if col_id == 'all' or (col_id.isnumeric() and int(col_id) < lim) or (col_id.isalpha() and len(col_id) == 1 and ord(col_id) - 64 <= lim):
		return True
	else:
		return False
			
row_min = sheet.min_row
row_max = sheet.max_row
col_min = sheet.min_column
col_max = sheet.max_column
cells = sheet[name_int_chr(col_min) + str(row_min): name_int_chr(col_max) + str(row_max)]


def get_data(col_id):
	data = dict()
	if check_col(col_id, col_max):
		if col_id == 'all':
			for i in range(col_max):
				for j in range(row_max):
					data.update({name_int_chr(i + 1) + str(j + 1) : str(cells[j][i].value)})
					print(type(data))
		elif col_id.isnumeric():
				col_id_int = int(col_id)
				for j in range(row_max):
					data.update({name_int_chr(col_id_int + 1) + str(j + 1) : str(cells[j][int(col_id_int)].value)})
		else :
			for j in range(row_max):
				data.update({col_id + str(j + 1) : str(cells[j][name_chr_int(col_id)].value)})
	return data
		
	