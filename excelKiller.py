#-*-coding:utf-8-*-
import os

import xlrd
import xlwt
from xlutils.copy import copy
import argparse


def getlist(workSpaceDir):# get the excel file list
    
    fileNames = os.listdir(workSpaceDir)
    seachedFiles = []
    for fileName in fileNames:
        if fileName.endswith('.xls'):
            seachedFiles.append(fileName)
    return seachedFiles


def get_row_data(fileList, workPath, nameColNum):
    data_list = []
    for file in fileList:
        name_true = file.split('.')[0] #get name from file name
        data = xlrd.open_workbook(os.path.join(workPath, file))
        table = data.sheets()[0]
        for i in range(table.nrows):
            name = table.row(i)[nameColNum].value # get name from row data
            if name == name_true:
                data_list.append(table.row(i))
    return data_list


def write_rows(rowsData, outPut, nameColNum, oldRowNum, nameList):
    '''
    :param rowsData: the row data to be wrote
    :param outPut: the file you want
    :param oldRowNum: the row number you want to begin
    :return: status code
    '''
    workBook = copy(xlrd.open_workbook( outPut))
    workSheet = workBook.get_sheet(0)

    #set borders styles
    borders = xlwt.Borders()
    borders.left = xlwt.Borders.THIN
    borders.right = xlwt.Borders.THIN
    borders.top = xlwt.Borders.THIN
    borders.bottom = xlwt.Borders.THIN
    borders.bottom_colour = 0x3A
    style = xlwt.XFStyle()
    style.borders = borders

    i = oldRowNum
    with open(nameList) as f:
        nameList = f.readline().split('，')

    for row in rowsData:
        for j in range(len(row)):
            workSheet.write(i, j, str(row[j].value), style)
        check(nameList, row[nameColNum].value)
        i += 1
    workBook.save(outPut)
    return nameList


def addarg():
    parser = argparse.ArgumentParser()
    parser.add_argument('-w', '--workPath', type=str, help='the path of your excels', default=os.getcwd())
    parser.add_argument('-o', '--oldRowNum', type=int, help='the row number you want to begin', default=2)
    parser.add_argument('-c', '--column', type=int, help='the column number of the name', default=3)
    parser.add_argument('-O', '--outPutFiles', type=str, help='the file you want ro save the output, it must exist',
                        default='output.xls')
    args = parser.parse_args()
    return args


def check(nameList, name):
    if name in nameList:
        nameList.remove(name)
        print('file {}.xls has been merged'.format(name))



if __name__ == '__main__':
    args = addarg()
    output = args.outPutFiles.strip('\'')
    path = args.workPath.strip('\'')
    nameList = 'nameList.txt'
    nameColNum = args.column
    fileList = getlist(path)
    rows = get_row_data(fileList, path, nameColNum)
    unDoneList = write_rows(rows, output, nameColNum, args.oldRowNum, nameList)
    print('*'*20 + 'unDone' + '*'*20)
    for ele in unDoneList:
        print(ele+' unDone')

