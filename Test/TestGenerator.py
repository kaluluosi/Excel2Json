import unittest
import os

from ExcelLoader.Loader import *
from ExcelLoader.Generator import Generator


class TestGenerator(unittest.TestCase):

    def test_export_json(self):
        gen = Generator('test.xlsx', 'config.json')
        gen.export_json()

        self.assertTrue(len(os.listdir('./data')) > 0)
        self.assertTrue(os.path.exists('./data/Test.json'))

