import os

import xlrd
import xlwt
import re
from xlutils.copy import copy


def getlist(workSpaceDir):
    fileNames = os.listdir(workSpaceDir)
    seachedFiles = []
    for fileName in fileNames:
        ele = re.search(r"(.*?).xls$", fileName)
        if ele:
            seachedFiles.append(ele.group())
    return seachedFiles


def get_row_data(fileList):
    data_list = []
    for file in fileList:
        name_true = file.split('.')[0]
        data = xlrd.open_workbook(file)
        table = data.sheets()[0]
        # print(table.nrows)
        for i in range(table.nrows):
            name = table.row(i)[4].value
            if name == name_true:
                # print(name)
                # print(table.row(i))
                data_list.append(table.row(i))
    return data_list


def write_rows(rowsData, outPut, oldRowNum):
    '''
    :param rowsData: the row data to be wrote
    :param outPut: the file you want
    :param oldRowNum: the row number you want to begin
    :return: status code
    '''
    workBook = copy(xlrd.open_workbook(outPut))
    workSheet = workBook.get_sheet(0)
    borders = xlwt.Borders()
    borders.left = xlwt.Borders.THIN
    borders.right = xlwt.Borders.THIN
    borders.top = xlwt.Borders.THIN
    borders.bottom = xlwt.Borders.THIN
    borders.bottom_colour = 0x3A

    style = xlwt.XFStyle()
    style.borders = borders
    i = oldRowNum
    for row in rowsData:
        print(row)
        for j in range(len(row)):
            print(row[j].value)
            workSheet.write(i, j, str(row[j].value), style)
        i += 1
    workBook.save(outPut)
    return 'successful !'


if __name__ == '__main__':
    path = os.getcwd()
    fileList = getlist(path)
    rows = get_row_data(fileList)
    output = 'output.xls'
    print(write_rows(rows, output, 2))
