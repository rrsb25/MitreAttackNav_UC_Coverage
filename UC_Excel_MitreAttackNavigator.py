from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
import pathlib
import json

tacticas = []
mitre = {"name": "layer",	"versions": {		"attack": "13",		"navigator": "4.8.2",		"layer": "4.4"	},	"domain": "enterprise-attack",	"description": "",	"sorting": 0,	"layout": {		"layout": "side",		"aggregateFunction": "average",		"showID": "false",		"showName": "true",		"showAggregateScores": "false",		"countUnscored": "false"	},	"hideDisabled": "false",	"techniques": ""    }

file = input('Enter filename\n Example: C:/Users/richard/excel.xlsx\n')
xls = load_workbook(filename=file)

sheets_excel  = input('Your excel has more than one sheet? Yes or No\n')

if sheets_excel == 'yes' or sheets_excel == 'Yes':
    name_sheet = input('What is the name Excel sheet?\n')
    print(name_sheet)
    sheet = xls[name_sheet] #set sheet
    last_column = len(list(sheet.columns)) #get columns
    last_row = len(list(sheet.rows)) #get rows
    print("Filas: "+str(last_row)+" Columnas: "+str(last_column))
else:
    sheet = xls.active 
    last_column = len(list(sheet.columns))
    last_row = len(list(sheet.rows))

columtactica = input('What is the excel column that store tactics?\n')
columtecnica = input('What is the column that store techniques?\n')

for row in range(2, last_row + 1):
    my_dict = {} #set dictionary empty
    if sheet[columtactica + str(row)].value != None and sheet[columtecnica + str(row)].value != None:
        my_dict[sheet[columtactica + str(1)].value] = sheet[columtactica + str(row)].value 
        #key: value --> key=columtactica's first row, value=columtactica's for row
        my_dict[sheet[columtecnica + str(1)].value] = sheet[columtecnica + str(row)].value
        #key: value --> key=columtecnica's first row, value=columtecnica's for row
        my_dict["color"] = "#4348c8"
        tacticas.append(my_dict) #append each dictionary completed to a list
        print("RONDA ----")

mitre["techniques"] = tacticas #adding whole tacticas in key "technique"

pathcurrentdir = pathlib.Path(__file__).parent.resolve()
data = json.dumps(mitre, sort_keys=True, indent=4)
with open(str(pathcurrentdir)+'/MitreNavigator.json', 'w', encoding='utf-8') as f:
    f.write(data)
