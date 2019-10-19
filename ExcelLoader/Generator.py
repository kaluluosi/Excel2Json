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

        self.excel = excel
        self.config = config
        self.loader = create_loader(excel)  # type:ExcelLoader

    @property
    def cfg(self):
        if self.config:
            fp = open(self.config)
            _cfg = json.load(fp)
        else:
            _cfg = {'export': '.'}
        return _cfg

    def export_json(self):
        dir_to_save = self.cfg['export']

        for raw_data in self.loader.all_raw_data:
            filename = os.path.join(dir_to_save, raw_data.cfg_file_name) + '.json'
            fp = open(filename, 'w', encoding='utf-8')
            json.dump(raw_data.data, fp, indent=True, ensure_ascii=False)

        self.loader.close()
