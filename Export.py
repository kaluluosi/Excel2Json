import os
import json
import sys

from GenCSharp.Generator import Generator

excel_file = sys.argv[1]
excel_dir = os.path.dirname(os.path.abspath(excel_file))

default_config = {"export": ""}

config_file = os.path.join(excel_dir, 'config.json')

# 如果没有默认配置就自动创建一个
if not os.path.exists(config_file):
    fp = open(config_file, 'w')
    json.dump(default_config, fp, indent=True)
    fp.close()

cfg_file = open(config_file, 'r')
config = json.load(cfg_file)

# 如果没有data目录就自动创建一个
export_dir = os.path.join(excel_dir, config['export'])
if not os.path.exists(export_dir):
    os.mkdir(export_dir)

gen = Generator(os.path.abspath(excel_file), excel_dir)
gen.export_json()

print("导出>>", os.path.split(excel_file)[1])

