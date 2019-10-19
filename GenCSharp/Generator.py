import os
import sys

from ExcelLoader import Generator


class CSharpGenerator(Generator):

    def __init__(self, excel, workspace):
        super(CSharpGenerator, self).__init__(excel, workspace)

    def export_csharp(self):
        provider_template = open(self.cfg['provider_template'])
        content = provider_template.read()

        dir_to_save = self.cfg['export_csharp']

        for raw_data in self.loader.all_raw_data:

            data_node_fields = []
            for i, field_name in enumerate(raw_data.field_names):
                if field_name == '':
                    continue
                data_node_fields.append(
                    "                public {type_name} {field_name};".format(field_name=field_name, type_name=raw_data.type_names[i]))

            code_str = content.format(ClassName=raw_data.cfg_file_name,
                                      CfgName=raw_data.cfg_file_name,
                                      DataNodeFields='\n'.join(data_node_fields))

            filename = os.path.join(dir_to_save, raw_data.cfg_file_name) + '.cs'
            with open(filename, 'w') as fp:
                fp.write(code_str)


if __name__ == '__main__':
    excel_file = r'..\Sample\template.xlsx'
    excel_dir = os.path.dirname(os.path.abspath(excel_file))

    gen = Generator(os.path.abspath(excel_file), os.path.abspath(os.path.join(excel_dir, 'config.json')))
    gen.export_json()
