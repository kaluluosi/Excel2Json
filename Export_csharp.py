import os
import sys

from GenCSharp.Generator import CSharpGenerator

excel_file = sys.argv[1]
excel_dir = os.path.dirname(os.path.abspath(excel_file))

gen = CSharpGenerator(os.path.abspath(excel_file), excel_dir)
gen.export_json()
gen.export_csharp()