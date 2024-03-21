import openpyxl
import xlwings as xw
import csv

positionColumn = 1 #const
teamColumn = 2 #const
playerColumn = 3 #const
fileName = "Barrow.xlsx" #read from config/arguments/something
allColumnsCount = 27

#TODO column by index, or finding index by column name? 

positionMainValueColumn = 10
positionSecondaryValueColumn = 12
positionProgressColumn = 14
CAValueColumn = 21
CAProgressColumn = 22

#TODO use only one: xlwings or openpyxl?

class Helper:
    def find_row_by_value(self, sheet, column, min_row_value, value):
        for index, row in enumerate(sheet.iter_rows(min_row = min_row_value, max_row=sheet.max_row, values_only=True)):
            if row[column-1] == value:
                return index
        return None            
        
class OpenpyxlController:
    def create_sheet(self):
        self.wb = openpyxl.load_workbook(fileName)
        return self.wb.active

    def close_controller(self):
        self.wb.close()    

class XlwingsController:
    def create_sheet(self, useApp):
        self.useApp = useApp
        if useApp:            
            self.app = xw.App(visible=False) #use it for "silent" open & close excel
        self.wb = xw.Book(fileName)
        return self.wb.sheets.active
    
    def close_controller(self, closeFile):
        self.wb.save()
        if(closeFile):
            self.wb.close()
        if(self.useApp):
            self.app.quit()

class CsvController:
    def create_csv(self, filename, data):
        with open(filename, 'w', newline='') as csvfile:
            writer = csv.writer(csvfile)
            for row in data:
                writer.writerow(row)

    def read_csv(self, filename):
        data = []
        with open(filename, 'r', newline='') as csvfile:
            reader = csv.reader(csvfile)
            for row in reader:
                data.append(row)
        return data
    
    def createTemplateFile(self):
        valuesTemplateFilename = 'template.csv'
        data = [
            ['Position', 'Team', 'Name', 'Age', 'Country', 'NONE', 'ToolRating_1', 'ToolRating_2', 'NONE', 'PositionMain_Name', 'PositionMain_Value', 'PositionSec_Name', 'PositionSec_Value', 'NONE', 
            'Progress', 'NONE', 'NONE', 'Determination', 'Potential', 'NoPermission', 'NONE', 'CA', 'CAChange', 'PA', 'PlayerStatus', 'RaportStatus', 'Info']    
        ]
        self.create_csv(valuesTemplateFilename, data)

class ExcelModificationsController:
    def __init__(self):
        self.helper = Helper()
        self.openpyxlController = OpenpyxlController()
        self.xlwingsController = XlwingsController()    
        self.csvController = CsvController()            

    def find_teamPart_row(self, positionValue, teamValue):
        sheet = self.openpyxlController.create_sheet()
        positionRow = self.helper.find_row_by_value(sheet, positionColumn, 1, positionValue)
        if positionRow is not None:        
            team_part_row = self.helper.find_row_by_value(sheet, teamColumn, positionRow, teamValue)
            self.openpyxlController.close_controller()
            if team_part_row is not None:
                return team_part_row + positionRow
            else:
                return None                    
        
    def insert_row(self, positionValue, teamValue):        
        sheet = self.xlwingsController.create_sheet(True)    
        teamPartRow = self.find_teamPart_row(positionValue, teamValue)       
        sheet.range((teamPartRow + 1, 1)).api.EntireRow.Insert()
        self.xlwingsController.close_controller(True)
        return teamPartRow
    
    def update_row_values(self, row_number, values):        
        sheet = self.xlwingsController.create_sheet(True)         
        sheet.range((row_number, 1), (row_number, len(values))).value = values
        self.xlwingsController.close_controller(True)

    def update_columns_values(self, rowNumber, columnsNumbers, values):
        sheet = self.xlwingsController.create_sheet(True)                
        for i in range(len(columnsNumbers)):            
            sheet.range((rowNumber, columnsNumbers[i]), (rowNumber, columnsNumbers[i])).value = values[i]            
        self.xlwingsController.close_controller(True)

    def get_player_row(self, value):        
        sheet = self.openpyxlController.create_sheet()                
        playerRow = self.helper.find_row_by_value(sheet, playerColumn, 1, value)
        self.openpyxlController.close_controller()
        return playerRow
    
    def get_player_data(self, value):
        sheet = self.xlwingsController.create_sheet(True)
        rowNumber = self.get_player_row(value) + 1              
        data = sheet.range((rowNumber, 1), (rowNumber, allColumnsCount)).value
        self.xlwingsController.close_controller(True)
        return data
    
    def insert_player_by_file(self):
        readFileData = self.csvController.read_csv("playerInfo.csv")
        searchValue = readFileData[1][2]
        playerRow = self.get_player_row(searchValue)
        if playerRow is not None:
            #check difference between position_values previous & current + difference between CA previous & current
            playerData = self.get_player_data(searchValue)

            readFileData[1][positionProgressColumn] = max(self.calculate_difference(playerData, readFileData, positionMainValueColumn), self.calculate_difference(playerData, readFileData, positionSecondaryValueColumn))         
            readFileData[1][CAProgressColumn] = self.calculate_difference(playerData, readFileData, CAValueColumn)

            self.update_row_values(playerRow + 1, readFileData[1])
        else:
            newRow = self.insert_row(readFileData[1][0], readFileData[1][1])
            self.update_row_values(newRow + 1, readFileData[1])        

    def update_player_by_file(self):
        #TODO check if player exist
        readTemplateData = self.csvController.read_csv("template.csv")
        readPlayerData = self.csvController.read_csv("playerInfoOnlyFew.csv")
        playerRow = self.get_player_row(readPlayerData[1][readPlayerData[0].index("Name")])
        columnsIndexes = []
        values = []
        for word in readPlayerData[0]:
            if word in readTemplateData[0]:                
                columnsIndexes.append(readTemplateData[0].index(word))
                values.append(readPlayerData[1][readPlayerData[0].index(word)])                        
        columnsIndexes = [index + 1 for index in columnsIndexes]                
        self.update_columns_values(playerRow + 1, columnsIndexes, values)

    def calculate_difference(self, previousPlayerData, currentPlayerData, columnValue):
        previousValue = float(previousPlayerData[columnValue])
        currentValue = float(currentPlayerData[1][columnValue]) 
        return (currentValue - previousValue)

excelModificationsController = ExcelModificationsController()
#excelModificationsController.insert_row("BR", "REZERWA")           
#excelModificationsController.insert_player_by_file()
excelModificationsController.update_player_by_file()