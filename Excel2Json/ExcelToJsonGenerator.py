# coding: utf-8

import xlrd
import json
import sys
import string
import io
import  math
import  os
from xlrd.sheet import *
from xlrd.book import  *
import config

SETTING_END=3

# setting
class Setting():
    def __init__(self):
        self.description = 0
        self.typeFlag = 0
        self.fieldName = 0
        self.setting_end=3
        self.end = 0
        self.defFileName = 'none'
        self.cfgFileName = 'none'
        self.dataFileName = 'none'

class Range():
    def __init__(self):
        self.start = None
        self.end = None


class Field():
    def __init__(self):
        self.name = ''
        self.description = ''
        self.typeFlag = ''


'''
    打开excel文件
    找出每个可导的sheet
    导sheet
        找出每个sheet的设置
        找出配置表范围
        找出每个sheet的元数据
'''

'''Json type flag 2 C# type mapper'''

typeMapper={
    'int':'int',
    'float':'float',
    'string':'string',
    'dict':'Dictionary',
    'define':'int',
    'id':'int'
}

'''导表'''

#导出字典格式数据
def GenData(book:Book, sheet: Sheet)->dict:
    # 获取配置表的设置 setting
    setting = GetSetting(book, sheet)

    # 获取配置表的字段设置 metadata
    metadata = GetMetaData(book,sheet)

    # 数据范围
    fieldRange = GetRangeByColor(book, sheet, setting.fieldName)
    endRange = GetRangeByColor(book, sheet, setting.end)

    dataRange = Range()
    dataRange.start = fieldRange.start[:]
    dataRange.start[0] += 1
    dataRange.end = endRange.end[:];


    # 返回字典类型数据
    # 先将excel数据转成表
    data = dict()

    for r in range(dataRange.start[0], dataRange.end[0]):
        dataNode = dict()
        values  = sheet.row_values(r,dataRange.start[1], dataRange.end[1])

        for i in range(len(values)):
            dataNode[metadata[i].name] = Parse(metadata[i].type,values[i])

        if dataNode['id']==-1:
            continue
        else:
            data[int(dataNode['id'])] = dataNode


    return data


def GetMetaData(book:Book,sheet:Sheet)->dict:
    metadata = dict()
    setting = GetSetting(book,sheet)
    typeRange = GetRangeByColor(book, sheet, setting.typeFlag)

    colstart = typeRange.start[1]
    colend = typeRange.end[1]

    types = GetValuesByColor(book, sheet, setting.typeFlag)
    descs = GetValuesByColor(book,sheet,setting.description)
    names = GetValuesByColor(book,sheet,setting.fieldName)

    for i in range(len(types)):
        metadata[i] = Field()
        metadata[i].type = types[i]
        metadata[i].name = names[i]
        metadata[i].description = descs[i]

    return metadata


def ToJson(filepath,data):
    file = open(filepath,'w',encoding='utf-8')
    text = json.dump(data,file,indent=True,ensure_ascii=False)
    file.close()

#根据类型映射表将类型标志转换成对应的C#类型名
def TypeFlagToCCharpTypeName(typeFlag:str):
    if(typeFlag.startswith("dictionary")):
        startIndex = typeFlag.index('<')
        endIndex = typeFlag.index('>')
        keyValueTypes = typeFlag[startIndex+1:endIndex].split(',')
        return  "Dictionary<string,"+keyValueTypes[1]+">"
    elif typeFlag in typeMapper.keys():
        return  typeMapper[typeFlag]
    else:
        return  typeFlag


def GenCSharpDataProvider(book:Book,sheet:Sheet,filepath,cfgPath):
    metadata = GetMetaData(book,sheet)
    setting = GetSetting(book,sheet)
    file = open( os.path.dirname(sys.argv[0])+"\CSharpDataProviderTemplate.cs")
    fmt = file.read()
    className = "{ClassName}"
    cfgName ="{CfgName}"
    dataNodeFields ="{DataNodeFields}"

    fields = list()
    for id,meta in metadata.items():
        fieldFmt = "                    public {type} {field};"
        statement = fieldFmt.format(type=TypeFlagToCCharpTypeName(meta.type),field=meta.name)
        WriteLine(fields)
        fields.append(statement)
    #z字段代码
    fieldCode = ''.join(fields)

    code = fmt.replace(className, setting.dataFileName)
    code = code.replace(cfgName, setting.cfgFileName)
    code = code.replace(dataNodeFields,fieldCode)

    file = open(filepath,'w')
    file.write(code)
    file.close()

def WriteLine(code:list):
    code.append('\n')
'''获取设置信息'''


def GetSetting(book: Book, sheet: Sheet) -> Setting:
    setting = Setting()
    setting.description = GetCellColorIndex(book, sheet.cell(0, 1))
    setting.typeFlag = GetCellColorIndex(book, sheet.cell(1, 1))
    setting.fieldName = GetCellColorIndex(book, sheet.cell(2, 1))
    setting.end = GetCellColorIndex(book, sheet.cell(3, 1))
    setting.defFileName = sheet.cell(4, 1).value
    setting.cfgFileName = sheet.cell(5, 1).value
    setting.dataFileName = sheet.cell(6, 1).value
    return setting


def GetCellColorIndex(book: Book, cell: Cell):
    index = cell.xf_index
    xf = book.xf_list[index]
    return xf.background.pattern_colour_index

'''
根据字段类型标记做一次文本转python类型的操作
xlrd默认将数值转换成float，同时我们也需要区分数字字符和数值，因此需要根据标记再做一次转换。
'''
def Parse(type:str, value):

    if type == 'int':
        try:
            return int(value)
        except :
            return  -1

    if type=='string':
            return  str(value)

    if type.startswith('dict'):
        try:
            return  eval(value)
        except:
            return  None

    return  eval(value)




'''查找某颜色单元格所在的行列范围'''
def GetRangeByColor(book: Book, sheet: Sheet, color):
    rect = Range()

    for r in range(6, sheet.nrows):
        for c in range(sheet.ncols):
            if color == GetCellColorIndex(book, sheet.cell(r, c)):
                rect.start = [r, c]
                rect.end = [r, sheet.ncols]
                break

    if rect.start == None:
        raise Exception("配置表格式有问题，没有找到对应颜色单元格", color)

    for c in range(sheet.ncols):
        # print(GetCellColorIndex(book, sheet.cell(rect.start[0], c)))
        if GetCellColorIndex(book, sheet.cell(rect.start[0], c)) != color:
            rect.end = [rect.start[0], c]
            break
    return rect

def GetValuesByColor(book:Book,sheet:Sheet,color)->list:
    rect = GetRangeByColor(book,sheet,color)
    return  sheet.row_values(rect.start[0],rect.start[1],rect.end[1])


if __name__ == '__main__':
    print(sys.argv)

    filepath = sys.argv[1]
    sheetName = sys.argv[2]
    cfgDir = sys.argv[3]
    dataproviderDir = sys.argv[4]

    book = xlrd.open_workbook(filepath,formatting_info=True)
    if(sheetName == ""):
        print("导出所有"+sys.argv[1])
        folderpath = os.path.dirname(os.path.abspath(filepath))
        for sheet in book.sheets():
            data = GenData(book,sheet)
            setting = GetSetting(book,sheet)
            cfgPath = cfgDir+'\\'+setting.cfgFileName + '.json'
            dataprovideerPath =dataproviderDir+'\\'+setting.cfgFileName + '.cs'
            #导出json表
            ToJson(cfgPath,data)
            GenCSharpDataProvider(book,sheet,dataprovideerPath,cfgPath)
    else:
        sheet = book.sheet_by_name(sheetName)
        data = GenData(book,sheet)
        setting = GetSetting(book,sheet)
        cfgPath = cfgDir+'\\'+setting.cfgFileName + '.json'
        dataprovideerPath =dataproviderDir+'\\'+setting.cfgFileName + '.cs'
        #导出json表
        ToJson(cfgPath,data)
        GenCSharpDataProvider(book,sheet,dataprovideerPath,cfgPath)