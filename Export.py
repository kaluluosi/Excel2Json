import os
import sys

from GenCSharp.Generator import Generator

excel_file = sys.argv[1]
excel_dir = os.path.dirname(os.path.abspath(excel_file))

gen = Generator(os.path.abspath(excel_file), excel_dir)
gen.export_json()