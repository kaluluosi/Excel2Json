
from .Loader import create_loader
from .Generator import Generator


import sys
import os

if __name__ == '__main__':

    excel_file = sys.argv[1]
    _dir = os.path.splitext(excel_file)[0]
    gen = Generator(sys.argv[1], os.path.join(_dir, 'config.json'))

