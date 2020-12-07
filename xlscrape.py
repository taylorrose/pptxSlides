from pandas import *


def xlscrape(path):
    xls = ExcelFile(path)
    return xls.parse(xls.sheet_names[0])

