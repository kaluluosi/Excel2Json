# README

## 更新日志

### ver 1.0

1. 增加config.json配置用来记录输出路径
2. 重写Excel Addin （因为源码已经丢失了）
3. 将原始数据获取和自定义导出过程（数据加工）剥离

## 如何使用

本工具依赖`openpyxl`和`xlrd`，先用pip安装这两个库

>pip install openpyxl

>pip install xlrd

Sample目录下有使用示例，直接把`template.xlsx`拖到`Export.bat`即可导出数据。

