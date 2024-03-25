import unittest
from excelController import Helper
from excelController import OpenpyxlController

positionColumn = 1 #const
teamColumn = 2 #const
playerColumn = 3 #const
allColumnsCount = 27

class TestHelper(unittest.TestCase):
    def test_row_by_value(self):
        openpyxlController = OpenpyxlController()        
        sheet = openpyxlController.create_sheet()        
        value = "J. Mullings"
        helper = Helper()
        result = helper.find_row_by_value(sheet, playerColumn, 0, value)
        self.assertEqual(result, 3)
        openpyxlController.close_controller()

unittest.main()