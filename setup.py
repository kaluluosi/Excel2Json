from setuptools import setup, find_packages

setup(name='Excel2Json',
      version='1.0',
      description='Excel转json导表工具',
      author='dengxuan',
      author_email='kaluluosi111@qq.com',
      url = 'https://github.com/kaluluosi/Excel2Json',
      packages=['ExcelLoader'],
      install_requires = ['openpyxl', 'xlrd']
      )
