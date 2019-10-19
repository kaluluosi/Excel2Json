import json
import os

from ExcelLoader.Loader import ExcelLoader
from . import create_loader


class Generator(object):
    """
    导出器
    定义了导出过程，可以继承这个导出器，自己实现具体的加工rawdata的导出器。
    """

    def __init__(self, excel: str, config: str = None):

        if config:
            fp = open(config)

        self.config = json.load(fp)  # if config else{'export': '.'}
        self.excel = excel

        self.loader = create_loader(excel)  # type:ExcelLoader

    def export_json(self):
        dir_to_save = self.config['export']

        for raw_data in self.loader.all_raw_data:
            filename = os.path.join(dir_to_save, raw_data.cfg_file_name) + '.json'
            fp = open(filename, 'w', encoding='utf-8')
            json.dump(raw_data.data, fp, indent=True, ensure_ascii=False)

        self.loader.close()

