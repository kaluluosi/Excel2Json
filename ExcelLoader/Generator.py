import json
import os

from ExcelLoader.Loader import ExcelLoader
from . import create_loader


class Generator(object):
    """
    导出器
    定义了导出过程，可以继承这个导出器，自己实现具体的加工rawdata的导出器。
    """

    def __init__(self, excel: str, workspace: str):

        self.excel = excel
        self.workspace = os.path.abspath(workspace)
        self.loader = create_loader(excel)  # type:ExcelLoader

    @property
    def cfg(self):
        fp = open(os.path.join(self.workspace, 'config.json'))
        _cfg = json.load(fp)
        return _cfg

    def export_json(self):
        dir_to_save = os.path.join(self.workspace, self.cfg['export'])

        for raw_data in self.loader.all_raw_data:
            filename = os.path.join(self.workspace, dir_to_save, raw_data.cfg_file_name) + '.json'
            fp = open(filename, 'w', encoding='utf-8')
            json.dump(raw_data.data, fp, indent=4, ensure_ascii=False, separators=(',', ':'), skipkeys=True, check_circular=True, sort_keys=True)

        self.loader.close()
