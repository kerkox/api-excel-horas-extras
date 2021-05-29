import uuid
import os
import pandas as pd
import numpy as np
import zipfile
from pandas.core.frame import DataFrame

PATH_OUTPUT = ".\\data\\output\\"
HORAS_EXTRAS_HEDO = ["H.E.D.O", "H.E.O.D"]
HORAS_EXTRAS_HENO = ["H.E.N.O", "H.E.O.N"]
HORAS_EXTRAS_HEDF = ["H.E.D.F", "H.E.F.D"]
HORAS_EXTRAS_HENF = ["H.E.N.F", "H.E.F.N"]
HORAS_EXTRAS_RN = ["R.N"]
HORAS_EXTRAS_CONSTANTES_ARRAY = [
	HORAS_EXTRAS_HEDO[0],
	HORAS_EXTRAS_HENO[0],
	HORAS_EXTRAS_HEDF[0],
	HORAS_EXTRAS_HENF[0],
	HORAS_EXTRAS_RN[0]
]
HORA_EXTRAS_CODE = {
	HORAS_EXTRAS_HEDO[0]: HORAS_EXTRAS_HEDO[0]+'_700',
	HORAS_EXTRAS_HENO[0]: HORAS_EXTRAS_HENO[0]+'_701',
	HORAS_EXTRAS_HEDF[0]: HORAS_EXTRAS_HEDF[0]+'_702',
	HORAS_EXTRAS_HENF[0]: HORAS_EXTRAS_HENF[0]+'_703',
	HORAS_EXTRAS_RN[0]: HORAS_EXTRAS_RN[0]+'_704'
}
KEY_CEDULA = 'cedula'
KEY_NOMBRE = 'nombre'
KEY_APELLIDO = 'apellido'
data_validar = [
	HORAS_EXTRAS_HEDO,
	HORAS_EXTRAS_HENO
]

TEXTOS_FILA_INVALIDA = [
	'TOTAL HORAS EXTRAS',
		'COMP',
		'COMPENSATORIOS',
		'DOM',
		'DOMINICALES',
		'VIAT',
		'VIATICOS'
]


def readAllSheets(filename):
	if not os.path.isfile(filename):
		return None

	xls = pd.ExcelFile(filename)
	sheets = xls.sheet_names
	xls.close()
	return sheets


def print_excel(excel_data: DataFrame):
	print("*"*20)
	print(obtener_info_horas_persona(excel_data))
	print()


def obtener_info_persona(columna_data, data, index_data):
	# print("+"*30)
	# print("obtener_info_persona, columna: ", columna_data)
	NOMBRE = "NOMBRE"
	point = np.where(columna_data == NOMBRE)
	persona = {}
	if len(point[0]) >= 1:
		index = point[0][0]
		persona[KEY_NOMBRE] = columna_data[index+2]
		columna_data = data[:, index_data+1]
		persona[KEY_APELLIDO] = columna_data[index+2]
		columna_data = data[:, index_data+2]
		persona[KEY_CEDULA] = str(
			columna_data[index+2]).replace(",", "").replace(".", "")
	# print("persona: ", persona)
	return persona


def is_columna_valida_horas_extras(lista_columna):
	constantes_encotradas = 0

	for data in data_validar:
		if (data[0] in lista_columna or data[1] in lista_columna):
			constantes_encotradas += 1
	# print("Lista: ", lista_columna)
	return constantes_encotradas == 1


def obtener_horas_extras_final(lista_columna, index_start, data, index_data):
	value_max = 0
	index_max = -1
	for index in range(index_start, len(lista_columna)):
		try:
			if float(lista_columna[index]) > value_max:
				value_max = float(lista_columna[index])
				index_max = index
		except ValueError:
			continue
	# print(f"Columna value_max:{value_max} index_max:{index_max}")
	# print(lista_columna)
	row = data[index_max, :]
	for texto_invalido in TEXTOS_FILA_INVALIDA:
		if texto_invalido in row:
			index_max -= 1
			value_max = float(lista_columna[index_max])
			break
	# print(f"Ajustes: value_max:{value_max} index_max:{index_max} ")
	# Se valida que la fila npo contenta el texto total horas extras

	return value_max, index_max


# Recibe una columna de datos de numpy array
def obtener_info_horas_extras(columna_data, data, index_data):
	# print("-"*30)
	# print("obtener_info_horas_extras, columna: ", columna_data)
	horas_extras = {}
	lista_columna = columna_data.tolist()

	if not is_columna_valida_horas_extras(lista_columna):
		# print("Se encontraron ambos por lo tanto no entra")
		return horas_extras
	# print("-"*20)
	index = -1
	point_0 = lista_columna.count(HORAS_EXTRAS_HEDO[0])
	point_1 = lista_columna.count(HORAS_EXTRAS_HEDO[1])
	if point_0 != 0:
		index = lista_columna.index(HORAS_EXTRAS_HEDO[0])
	elif point_1 != 0:
		index = lista_columna.index(HORAS_EXTRAS_HEDO[1])

	# point = np.where(
	# 	columna_data == HORAS_EXTRAS_HEDO[0] or columna_data == HORAS_EXTRAS_HEDO[1])
	if index >= 0:
		# print("columna original donde se encontro: ")
		# print(columna_data)
		value_max, index_max = obtener_horas_extras_final(
			lista_columna, index+1, data, index_data)
		# print(f"value_max: {value_max}, index_max: {index_max}")
		horas_extras[HORAS_EXTRAS_CONSTANTES_ARRAY[0]] = value_max

		for column_add in range(1, 5):
			columna = data[:, index_data+column_add]
			# print("--"*10)
			# print(f"COLUMNA {column_add+1}: ")
			# print(columna)
			if index_max == -1:
				value_max, index_max = obtener_horas_extras_final(
					columna, index+1, data, index_data)
			try:
				horas_extras[HORAS_EXTRAS_CONSTANTES_ARRAY[column_add]] = float(
					columna[index_max])
			except ValueError:
				horas_extras[HORAS_EXTRAS_CONSTANTES_ARRAY[column_add]] = 0
				index_max = -1

	return horas_extras


def obtener_info_horas_persona(excel_data: DataFrame):
	data = excel_data.to_numpy()
	info = {}
	horas_extras = {}
	data_horas_persona = {}
	for i in range(len(data[1, :])):
		columna = data[:, i]
		# divider = "*"*20
		# print(f"{divider} index: {i} ")
		# print(columna)
		if horas_extras.get(HORAS_EXTRAS_CONSTANTES_ARRAY[0]) == None:
			horas_extras = obtener_info_horas_extras(columna, data, i)
		if info.get('nombre') == None:
			info = obtener_info_persona(columna, data, i)

		if horas_extras.get(HORAS_EXTRAS_CONSTANTES_ARRAY[0]) != None and info.get('nombre') != None:
			break

	data_horas_persona.update(horas_extras)
	data_horas_persona.update(info)
	return data_horas_persona


# last_sheet = sheets[0]
# excel_data = pd.read_excel(file_name, last_sheet)
# print_excel(excel_data)

def show_all_sheets(file_name):
	for data in create_dictionary_extra_hours(file_name):
		print(data)


def create_dictionary_extra_hours(file_name):
	sheets = readAllSheets(file_name)
	list_extra_hours = []
	for sheet in sheets:
		excel_data = pd.read_excel(file_name, sheet)
		extract_hours = obtener_info_horas_persona(excel_data)
		if extract_hours:
			list_extra_hours.append(extract_hours)

	return list_extra_hours


def extraer_data_excel(list_files):
	horas_extras = []
	for file_excel in list_files:
		horas_extras += create_dictionary_extra_hours(file_excel)

	return horas_extras


def generar_listas_por_tipo_hora_extra_con_cedula(list_files):
	data = extraer_data_excel(list_files)
	info_horas_extras = {}
	for hora_extra_key in HORAS_EXTRAS_CONSTANTES_ARRAY:
		info_horas_extras[HORA_EXTRAS_CODE[hora_extra_key]] = []
		for hora_extra in data:
			if float(hora_extra[hora_extra_key]) > 0:
				peer = f"{hora_extra[KEY_CEDULA]};{int(hora_extra[hora_extra_key])}"
				info_horas_extras[HORA_EXTRAS_CODE[hora_extra_key]].append(
					peer)

	return info_horas_extras


def show_by_name_data(list_files):
	data = extraer_data_excel(list_files)
	for hora_extra_key in data:
		print(hora_extra_key)


files_name = ['.\\data\\test.xlsx', '.\\data\\test_2.xlsx']
# files_name = ['.\\data\\test_2.xlsx']


def write_txt_file_by_type_extra_hour(list_data, name_file):
	path = f"{PATH_OUTPUT}{name_file}.txt"
	file = open(path, "w+")
	file.write("\n".join(list_data))
	# for line in list_data:
	# 	file.write(line+"\n")
	file.close()


def generate_txt_file_data_extra_hours(list_files):
	result = generar_listas_por_tipo_hora_extra_con_cedula(list_files)
	for info in result:
		write_txt_file_by_type_extra_hour(result[info], info)
	code = zip_files_extra_hours_txt()
	return code

def clean_files():
	path_base = os.path.join(PATH_OUTPUT, f'PLANOS.zip')
	if os.path.exists(path_base):
		os.remove(path_base)
	for folder, subfolders, files in os.walk(PATH_OUTPUT):
		for file in files:
			if file.endswith('.txt'):
				path_file = os.path.join(folder, file)
				os.remove(path_file)

def zip_files_extra_hours_txt():
	# clean_files()
	code = uuid.uuid4()
	zip_path = f"{PATH_OUTPUT}\\PLANOS-{code}.zip"
	zip_txt_files = zipfile.ZipFile(zip_path, "w")
	for folder, subfolders, files in os.walk(PATH_OUTPUT):
		for file in files:
			if file.endswith('.txt'):
				path_file = os.path.join(folder, file)
				zip_txt_files.write(path_file, os.path.relpath(
				os.path.join(folder, file), PATH_OUTPUT), compress_type=zipfile.ZIP_DEFLATED)
	zip_txt_files.close()

	return code


def show_console_data_result(list_files):
	result = generar_listas_por_tipo_hora_extra_con_cedula(list_files)
	for info in result:
		print("*"*30)
		print(info)
		print(result[info])


# show_console_data_result(files_name)
# show_by_name_data(files_name)
# generate_txt_file_data_extra_hours(files_name)
