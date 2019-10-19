import unittest
from ExcelLoader.Loader import XLSXLoader

class TestXLSXLoader(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        cls.loader = XLSXLoader('template.xlsx')

    def test_load_setting(self):
        setting = self.loader.setting(0)

        self.assertTrue(setting.cfg_name)
        self.assertTrue(setting.range)
        self.assertTrue(setting.field_row)
        self.assertTrue(setting.type_row)

    def test_get_raw_data(self):

        raw_data = self.loader.get_raw_data(0)

        self.assertTrue(raw_data.type_names)
        self.assertTrue(raw_data.field_names)
        self.assertTrue(raw_data.data)
        self.assertEqual(len(raw_data.type_names), len(raw_data.field_names))

    @classmethod
    def tearDownClass(cls):
        cls.loader.close()


if __name__ == '__main__':
    unittest.main()
