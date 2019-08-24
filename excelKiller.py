import os

import xlrd
import xlwt
import re
from xlutils.copy import copy
import argparse


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
        for i in range(table.nrows):
            name = table.row(i)[4].value
            if name == name_true:
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
        for j in range(len(row)):
            workSheet.write(i, j, str(row[j].value), style)
        print('file {}.xls has been merged'.format(row[4].value))
        i += 1
    workBook.save(outPut)
    return 'successful !'


def addarg():
    parser = argparse.ArgumentParser()
    parser.add_argument('-workPath', type=str, help='the path of your excels', default=os.getcwd())
    parser.add_argument('-oldRowNum', type=str, help='the row number you want to begin', default=2)
    parser.add_argument('-outPutFiles', type=str, help='the file you want ro save the output, it must exist',
                        default='output.xls')
    args = parser.parse_args()
    return args


if __name__ == '__main__':
    args = addarg()
    output = args.outPutFiles
    path = args.workPath
    fileList = getlist(path)
    rows = get_row_data(fileList)
    print(write_rows(rows, output, args.oldRowNum))
