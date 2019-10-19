import json
import os
import sys
import argparse

from GenCSharp.Generator import CSharpGenerator

excel_file = sys.argv[1]
excel_dir = os.path.dirname(os.path.abspath(excel_file))

default_config = {"export": "", "provider_template": "CSharpDataProviderTemplate.cs", "export_csharp": "provider"}

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

provider_dir = os.path.join(excel_dir, config['export_csharp'])
if not os.path.exists(provider_dir):
    os.mkdir(provider_dir)

gen = CSharpGenerator(os.path.abspath(excel_file), excel_dir)
gen.export_json()
gen.export_csharp()

print("导出>>", os.path.split(excel_file)[1])