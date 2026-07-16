import unittest
from unittest.mock import MagicMock, patch
from excel2flapjack.main import X2F
import pandas as pd
import test_xcel_file

'''
xls = pd.ExcelFile("./test_files/xdc_test.xlsx")
fj_url = "localhost:8000"
# fj_url = "flapjack.rudge-lab.org:8000"
fj_user = "dylan33smith"
fj_pass = "coco33"
'''

class TestX2FStudyUpload(unittest.TestCase):


    def setUp(self):
        self.xls = MagicMock()

        # set up mock dataframe that can be further setup to test upload results
        self.data = {
            'name': [None],
            'description': [None],
            'pubchemid': [None],
            'object': [None],
            'flapjackid': [None],
            'chemical': [None],
            'concentration': [None],
            'dna': [None],
            'color': [None],
            'DOI': [None],
            'machine': [None],
            'temperature': [None],
            'study': [None],
            'row': [None],
            'col': [None],
            'assay': [None],
            'sampledesign': [None],
            'sample': [None],
            'signal': [None],
            'time': [None],
            'value': [None],
            'media': [None],
            'strain': [None],
            'vector': [None],
            'supplement': [None]
        }

        # patch the Flapjack class and start the patcher
        self.mock_fj_patcher = patch('excel2flapjack.main.Flapjack')
        self.mock_fj = self.mock_fj_patcher.start()
        self.mock_excel_patcher = patch('excel2flapjack.main.pd.ExcelFile')
        self.mock_excel_patcher.start().return_value = self.xls

        # Setup mock Flapjack instance
        self.mock_fj = self.mock_fj.return_value
        
        fj_return_df = {
            'id': [123],
            'is_owner': [True],
            'shared_with': [[]],  # Adjust according to the actual expected structure
            'name': ['test study name'],
            'description': ['test study description'],
            'doi': ['https://doi.org/thisisatest'],
            'sboluri': [''],
            'public': [False]
        }
        mock_fj_return_df = pd.DataFrame(fj_return_df)

        self.mock_fj.create.return_value = mock_fj_return_df


        self.x2f = X2F(self.xls, 'fakeurl.com', 'fake_user', 'fake_pass')
        # give instance empty df of right structure 
        self.x2f.df = pd.DataFrame(self.data)

    def tearDown(self):
        self.mock_fj_patcher.stop()
        self.mock_excel_patcher.stop()

        
    def test_studies_upload(self):
        self.x2f.df['name'] = 'test study name'
        self.x2f.df['description'] = 'test study description'
        self.x2f.df['object'] = 'Study'

        self.x2f.upload_studies()

        self.mock_fj.create.assert_called_with(
            'study',
            name='test study name',
            description='test study description',
            confirm=False,
            overwrite=False
        )

if __name__ == '__main__':
    unittest.main()
