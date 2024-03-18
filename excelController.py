import openpyxl
import xlwings as xw
import csv

positionColumn = 1 #const
teamColumn = 2 #const
playerColumn = 3 #const
fileName = "Barrow.xlsx" #read from config/arguments/something
search_value = "J. Mullings" #read from config/arguments/something

def find_row_by_value(sheet, column, value):
    for index, row in enumerate(sheet.iter_rows(min_row = 1, max_row=sheet.max_row, values_only=True)):
        if row[column-1] == value:
            return index
    return None            

def find_teamPart_row(positionValue, teamValue):
    sheet = create_openpyxl_sheet()
    positionRow = find_row_by_value(sheet, positionColumn, positionValue)
    if positionRow is not None:        
        for index, row in enumerate(sheet.iter_rows(min_row = positionRow, max_row=sheet.max_row, values_only=True)):
            if row[teamColumn-1] == teamValue:
                return positionRow + index
        return None     

def create_openpyxl_sheet():
    wb = openpyxl.load_workbook(fileName)
    return wb.active    

def insert_row_using_xlwings(fileName, positionValue, teamValue):
    app = xw.App(visible=False)
    wb = xw.Book(fileName)
    sheet = wb.sheets.active

    teamPartRow = find_teamPart_row(positionValue, teamValue)       

    sheet.range((teamPartRow + 1, 1)).api.EntireRow.Insert()

    wb.save()
    #wb.close can close opened excel file
    #wb.close()
    app.quit()

    return teamPartRow

def update_row_values(fileName, row_number, values):
    app = xw.App(visible=False)
    wb = xw.Book(fileName)
    sheet = wb.sheets.active

    #print(sheet.range((row_number, 1), (row_number, len(values))).value)# = values
    #print(sheet.range((row_number, 2), (row_number, 26)).value)# = values
    #updated columns range
    sheet.range((row_number, 1), (row_number, len(values))).value = values

    wb.save()
    #wb.close can close opened excel file
    #wb.close()
    app.quit()

def check_if_player_exist(value):
    wb = openpyxl.load_workbook(fileName)
    sheet = wb.active
    return find_row_by_value(sheet, playerColumn, value)

def create_csv(filename, data):
    with open(filename, 'w', newline='') as csvfile:
        writer = csv.writer(csvfile)
        for row in data:
            writer.writerow(row)

def read_csv(filename):
    data = []
    with open(filename, 'r', newline='') as csvfile:
        reader = csv.reader(csvfile)
        for row in reader:
            data.append(row)
    return data

def createTemplateFile():
    valuesTemplateFilename = 'template.csv'
    data = [
        ['Position', 'Team', 'Name', 'Age', 'Country', 'NONE', 'ToolRating_1', 'ToolRating_2', 'NONE', 'PositionMain_Name', 'PositionMain_Value', 'PositionSec_Name', 'PositionSec_Value', 'NONE', 
        'Progress', 'NONE', 'NONE', 'Determination', 'Potential', 'NoPermission', 'NONE', 'CA', 'CAChange', 'PA', 'PlayerStatus', 'RaportStatus', 'Info']    
    ]
    create_csv(valuesTemplateFilename, data)

#insert_row_using_xlwings(fileName, "BR", "REZERWA")
#print(check_if_player_exist(search_value))

readFileData = read_csv("playerInfo.csv")
search_value = readFileData[1][2]
playerIndex = check_if_player_exist(search_value)
if playerIndex is not None:
   update_row_values(fileName, playerIndex + 1, readFileData[1])
else:
    newRow = insert_row_using_xlwings(fileName, readFileData[1][0], readFileData[1][1])    
    update_row_values(fileName, newRow + 1, readFileData[1])            