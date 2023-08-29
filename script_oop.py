'''The main purpose of the code here presented and thereby executed and used by me is to extract data from an excel file, 
perform some transformations, control the output, then parsing the data into a JSON structure to finally send a POST request 
to an HTML service'''

import pandas as pd 
import datetime
from datetime import timedelta 
import requests 
from requests.auth import HTTPBasicAuth 
import json 

# Simple popup message window.
import tkinter as tk
from tkinter import ttk

# Current Date
now = pd.Timestamp(datetime.date.today())


class MsgWindow():


	def __init__(self, my_font):
		self.font = my_font 


	def popupMsg(self, msg): 
		popup = tk.Tk() 

		popup.wm_title("This window's porpouse is to make me look more masculine") # Handy title 
		label = ttk.Label(popup, text=msg, font=self.font) # This object calls the module tkinter and uses the metod Label() to define the msg structure.
		label.pack(side="top", fill="x", padx=10, pady=20) 
		button = ttk.Button(popup, text="Enought", command=popup.destroy) 
		button = button.pack()

		popup.mainloop() 


def Display(rows, cols, width): # leaving this here in case somebody wants to use the console---------------------------------------------------------------------------
	pd.set_option('display.max_rows', rows)
	pd.set_option('display.max_columns', cols)
	pd.set_option('display.width', width)


#Display(20000, 20, 2160) # Uncomment this when working with the console.


class GetDataframe(): # This class reads an Excel file, gets data from the sheets, then parses data into pandas DataFrame.--------------------


	def __init__(self, path, sheetName):
		self.file_path = path
		self.sheetName = sheetName


	def ExceltoDataframe(self):
		try:
			df = pd.read_excel(self.file_path, sheet_name=self.sheetName, dtype=str)
		except:
			return print('Something wrong with your file address or the sheet name youve provided')
		return df


def CleanThisShit(): # Now we are going to clean and prepare the data-----------------------------------------------------------------------------------------------------
	StartingData = GetDataframe('G:\My Drive\Formulario Abastecimiento\Abastecimiento.xlsm', 'BD_Apr') 
	StartingData = StartingData.ExceltoDataframe() 

	LocationsOne = GetDataframe('G:\My Drive\Planos\Criterios.xlsx', 'Bodega') 
	LocationsOne = LocationsOne.ExceltoDataframe()

	LocationsTwo = GetDataframe('G:\My Drive\Planos\Criterios.xlsx', 'Bodega')
	LocationsTwo = LocationsTwo.ExceltoDataframe()

	LocationsOne = LocationsOne.drop(['descBodega', 'estado'], axis=1)
	LocationsTwo = LocationsTwo.drop(['descBodega', 'estado'], axis=1)
	StartingData = StartingData.drop(['Id_Apr', 'User_Apr', 'D_Item_Apr', 'Qty_kg_Apr', 'Factor_Um_Item_Apr', 'Peso_Um_Item_Apr', 'Usr_LE', 'Fecha_VbOk', 'If_ok', 'TB_LE', 'Orig_Bodega_Apr', 'Dest_Bodega_Apr', 'UEN', 'Sublinea'], axis=1)
	StartingData = StartingData.drop(StartingData.columns[StartingData.columns.str.contains('Unnamed', case=False)], axis=1)

	cols = {'Cod_Bodega_O_Apr': 'id_Bodega', 'Cod_Bodega_D_Apr': 'id_Bodega2'}
	cols2 = {'id_Bodega':'id_Bodega2'}
	StartingData.rename(columns=cols, inplace=True)
	LocationsTwo.rename(columns=cols2, inplace=True)

	StartingData = StartingData.merge(LocationsOne, on='id_Bodega')
	StartingData = StartingData.merge(LocationsTwo, on='id_Bodega2')

	StartingData['Date_Apr'] = StartingData['Date_Apr'].astype('datetime64[ns]') 
	StartingData = StartingData.loc[StartingData['Date_Apr'] == now] 
	StartingData = StartingData.loc[StartingData['Estado'] == 'OK']
	StartingData['Qty_Um_Apr'] = round(pd.to_numeric(StartingData['Qty_Um_Apr']),0) 
	StartingData = StartingData.loc[StartingData['Qty_Um_Apr'] > 0]
	StartingData = StartingData.loc[StartingData['id_Bodega'] != StartingData['id_Bodega2']] 
	StartingData = StartingData.loc[(StartingData['id_Bodega2'] != '00092') | (StartingData['id_Bodega2'] != '00095') | (StartingData['id_Bodega2'] != '00-91') | (StartingData['id_Bodega'] != '00-91')] #modify for IPIALES When needed
	StartingData =StartingData.drop(['Estado'], axis=1)

	CleanData = StartingData

	return CleanData


def ProcessThisShit(UnprocessedData): # Now that the data has been properly structured and cleaned, its time to build some calculated columns to accomplish our task.
	UnprocessedData['Date_entry'] = UnprocessedData['Date_Apr'] + timedelta(days=6) 
	UnprocessedData['ctrl1'] = UnprocessedData['id_Bodega'] + UnprocessedData['id_Bodega2'] 
	UnprocessedData = UnprocessedData.sort_values('ctrl1') 

	list1 = list(UnprocessedData['ctrl1']) # Pass this object to a function, find the unique combinations, then create a new column to store the result.


	def findUnique(Values):  # Function to find unique values 
		unique = []
		i = 0
		for everyItem in Values:
			if everyItem != Values[abs(i-1)]:
				unique.append('1')
			else:
				unique.append('0')
			i = i + 1
		unique[0] = '1'
		return unique 

	ctrl2 = findUnique(list1) 
	UnprocessedData.insert(loc=10, column='ctrl2', value=ctrl2) # Updating the dataframe


	def findConsec(Values): # Function to find consecutive values
		consec = []
		#consec[0] = '1'
		i = 0
		for everyItem in Values:
			if everyItem == '0':
				consec.append(str(i))
			else:
				consec.append('1')
				i = 1
			i = i + 1
		return consec


	def asignDocument(Values): # Function to find document values
		docs = []
		i = 0
		for everyItem in Values:
			if everyItem == '1':
				docs.append(str(i+1))
				i = i + 1
			else:
				docs.append(str(i))
		return docs

	def killHyphens(Values):
		no_hyphens = []
		for everyItem in Values:
			no_hyphens.append(everyItem.replace('-',''))
		return no_hyphens

	list2 = list(UnprocessedData['ctrl2'])
	ctrl3 = findConsec(list2) #Working function
	UnprocessedData.insert(loc=11, column='ctrl3', value=ctrl3) # Updating the dataframe
	ctrl4 = asignDocument(list2) #Working function
	UnprocessedData.insert(loc=12, column='ctrl4', value=ctrl4) # Updating the dataframe
	UnprocessedData['Date_Apr'] = UnprocessedData['Date_Apr'].astype(str)
	UnprocessedData['Date_entry'] = UnprocessedData['Date_entry'].astype(str)
	dates1 = list(UnprocessedData['Date_Apr'])
	dates2 = list(UnprocessedData['Date_entry'])
	UnprocessedData['Date_Apr'] = killHyphens(dates1) 
	UnprocessedData['Date_entry'] = killHyphens(dates2)

	ProcessedData = UnprocessedData

	return ProcessedData


def formatThisShitDOCS(UnformatedData): # Writing this function to independently handle the formating of the outputs needed.--------------------------------------------------------------------------------------------------------
	Documents = UnformatedData.loc[UnformatedData['ctrl2'] == '1']
	Documents = Documents.drop(['ctrl2', 'ctrl3', 'id_COpera_x', 'Um_Item_Apr', 'Item_Apr'], axis=1) 
	Documents = Documents[['id_COpera_y', 'ctrl4', 'Date_Apr', 'Date_entry', 'id_Bodega', 'id_Bodega2']] 
	new_names1 = {"id_COpera_y": "copera", "ctrl4":"numDoc", "Date_Apr":"doc_FechaDocumento", "Date_entry":"doc_FechaEntrega", "id_Bodega":"doc_BodegaSalida", "id_Bodega2":"doc_BodegaEntrada"} #The HTML service requires different naming
	Documents.rename(columns=new_names1, inplace=True)
	Documents = Documents.set_index('copera')
	Documents = Documents.reset_index()

	return Documents


def formatThisShitMOVS(UnformatedData): 
	Movements = UnformatedData
	Movements = Movements.drop(['ctrl1', 'Date_Apr', 'ctrl2', 'id_Bodega2'], axis=1)
	Movements = Movements[['id_COpera_y', 'ctrl4', 'ctrl3', 'Item_Apr', 'id_Bodega', 'Um_Item_Apr', 'Qty_Um_Apr', 'Date_entry', 'id_COpera_x']]
	new_names2 = {"id_COpera_y": "copera", "ctrl4":"numDoc", "ctrl3":"mov_Registro", "Item_Apr":"mov_Item", "id_Bodega":"mov_Bodega", "Um_Item_Apr":"mov_UMedida", "Qty_Um_Apr":"mov_Cantidad", "Date_entry":"mov_FechaEntrega", "id_COpera_x":"mov_Copera"}
	Movements.rename(columns=new_names2, inplace=True)
	Movements = Movements.set_index('copera')
	Movements = reset_index()

	return Movements


def exportThisShit(TakeThis, AlsoThis): # This Function can be utilized, if for whatever reason you decide to export the data for further evaluation or something like that.-------------------------------------------------------------------------------------------------
	#first we will discuss the naming of the file, im gonna call it "Rqi_(insert the current date here)"
	date_to_name = str(now) 
	date_to_name = date_to_name[0:10]
	file_name = 'Rqi ' + date_to_name + '.xlsx'
	write_data = pd.ExcelWriter(file_name, engine='xlsxwriter')
	TakeThis.to_excel(write_data, sheet_name='Documentos') # It had to be in spanish :D
	AlsoThis.to_excel(write_data, sheet_name='Movimientos')
	write_data.save()
	write_data.close()


def PostThisShit(docs, df): # Finally, define the JSON structure needed and send the request to POST data.
	emptyDIC1 = {}
	emptyDIC2 = {}
	fullDIC = {}
	listOUTTER =[]
	listINNER = []
	listOUTTER.append(emptyDIC1)
	listOUTTER.append(fullDIC)
	listOUTTER.append(emptyDIC2)	

	data = ''
	s=0
	j=0
	rows_docs = len(docs.index)
	rows_df = len(df.index)

	for k in range(rows_docs):
		id_doc = int(docs.iat[s,1])
		fullDIC.update( {"copera": docs.iat[s,0], "numDoc": docs.iat[s,1], "doc_FechaDocumento": docs.iat[s,2], "doc_FechaEntrega": docs.iat[s,3], "doc_BodegaSalida": docs.iat[s,4], "doc_BodegaEntrada": docs.iat[s,5]} )
		j=0
		for i in range(rows_df): # Built an identifier that corelates rows_df with rows_docs wich is ctrl 2 or numdoc as I've always done.
			id_mov = int(df.iat[j,1])
			if id_doc == id_mov:
				listINNER.append({"copera": (df.iat[j,0]), "numDoc": df.iat[j,1], "mov_Registro": (df.iat[j,2]), "mov_Item": (df.iat[j,3]), "mov_Bodega": (df.iat[j,4]), "mov_UMedida": (df.iat[j,5]), "mov_Cantidad": (df.iat[j,6]), "mov_FechaEntrega": (df.iat[j,7])})
			j=j+1
		# Adding a new key value pair, n values
		fullDIC.update( {'mov_' : listINNER} )
		data = json.dumps(str(listOUTTER))
		data = data.replace('\'', '\"')
		data = data[1:-1]
		listINNER = []
		s=s+1
		
		My_url = 'http://olympus.web.lan/Olympus/api/siesa/importar/rstOly_sob_aba_sie_pedInterno'
		My_auth = HTTPBasicAuth('dpto_user', 'pass2019')

		r = requests.post( url = My_url, data = data, auth = My_auth)
		response = r.text()
		return response # String with either confirmation or denial of the request

Data = CleanThisShit()
Data = ProcessThisShit(Data)
Documents = formatThisShitDOCS(Data)
Movements = formatThisShitMOVS(Data)
#exportThisShit(Documents, Movements) # Uncomment to export the data to an excel file.
response = PostThisShit(Documents, Movements)

windowOne = MsgWindow(("Helvetica Narrow", 14))
windowOne.popupMsg(" You're all set ;) " + "\n" + response)

