import unittest
from excelController import Helper
from excelController import OpenpyxlController

positionColumn = 1 #const
teamColumn = 2 #const
playerColumn = 3 #const
allColumnsCount = 27

class TestHelper(unittest.TestCase):
    def setUp(self):
        self.helper = Helper()
        self.openpyxlController = OpenpyxlController()

    def test_row_by_value(self):           
        sheet = self.openpyxlController.create_sheet()        
        value = "J. Mullings"        
        result = self.helper.find_row_by_value(sheet, playerColumn, 0, value)
        self.assertEqual(result, 3)
        self.openpyxlController.close_controller()

unittest.main()