import unittest
from unittest.mock import Mock, patch
import sys
import pandas as pd
from excel2flapjack.main import X2F

# don't need any of this because we're using mock and patch
'''
xls = pd.ExcelFile("./test_files/flapjack_excel_converter_revamp.xlsx")
fj_url = "localhost:8000"
# fj_url = "flapjack.rudge-lab.org:8000"
fj_user = "dylan33smith"
fj_pass = "coco33"
'''

# define test case class which inherits from 'unittest.TestCase' that contains tests related to initialization of XDC class
class TestX2FInit(unittest.TestCase):
    # decorator replaces Flapjack class with a mock
    # any instantiation of Flapjack within the test will use this mock instead of the real class
    @patch('excel2flapjack.main.Flapjack')
    @patch('excel2flapjack.main.pd.ExcelFile')
    # patched Flapjack class is passed into the test method as the parameter MockFlapjack
    def test_init_method(self, MockExcelFile, MockFlapjack):
        # Create a mock ExcelFile object to simulate a pd.ExcelFile
        mock_xls = Mock()
        MockExcelFile.return_value = mock_xls
        print(mock_xls)
        
        # MockFlapjack.return_value is used to specify the mock object that should be returned whenever the Flapjack class is instantiated
        # mock_fj_instance represents this mock instance of Flapjack
        mock_fj_instance = MockFlapjack.return_value
        # configure log_in method to return 'None' simulating successful behavior
        mock_fj_instance.log_in.return_value = None  # Simulate successful login
        
        # define variables used as arguments when creating an instance of the 'XDC' class within a test
        test_url = "http://testurl.com" # URL without port 8000
        test_user = "testuser"
        test_pass = "testpass"
        
        # Instantiate XDC with mock excel file ann test credentials
        # this is the line that actually tests the __init__ method as it involves creating the XDC object
        xdc_instance = X2F(mock_xls, test_url, test_user, test_pass)
        

        # Assertions to verify the __init__ behavior
        
        # test Flapjack class was called exactly once with the url_base constructed by appending ':8000' to 'test_url'
        # verifies URL modification logic
        MockFlapjack.assert_called_once_with(url_base=test_url + ":8000")  # Check URL modification
        # test log_in method was called exactly once with the correct username and password
        mock_fj_instance.log_in.assert_called_once_with(username=test_user, password=test_pass)  # Check login call


if __name__ == '__main__':
    unittest.main()
