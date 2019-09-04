# -*-coding:utf-8-*-
import os

import xlrd
import xlwt
from xlutils.copy import copy
import argparse


class XlsKiller():
    def __init__(self, fileList, workPath, nameColNum, nameList, output, oldRowNum):
        self._fileList = fileList
        self._workPath = workPath
        self._nameColNum = nameColNum
        self._nameList = nameList
        self._output = output
        self._oldRowNum = oldRowNum



    def get_row_data(self):  # get the row data needed
        '''
        :param fileList: the file that the data exist
        :param workPath: the dictionary the excel files exist.
        :param nameColNum: the column number of the name column
        :return data_list: return a list that include all the data in rows.
        '''
        data_list = []
        with open(self._nameList) as f:
            nameList = f.readline().split('，')
        for file in self._fileList:
            # name_true = file.split('.')[0]  # get name from file name
            data = xlrd.open_workbook(os.path.join(self._workPath, file))
            table = data.sheets()[0]
            for i in range(table.nrows):
                name = table.row(i)[self._nameColNum].value  # get name from row data
                if name in nameList:
                    data_list.append(table.row(i))
        # print(data_list)
        return data_list


    def write_rows(self, rowsData):
        workBook = copy(xlrd.open_workbook(self._output))
        workSheet = workBook.get_sheet(0)

        # set borders styles
        borders = xlwt.Borders()
        borders.left = xlwt.Borders.THIN
        borders.right = xlwt.Borders.THIN
        borders.top = xlwt.Borders.THIN
        borders.bottom = xlwt.Borders.THIN
        borders.bottom_colour = 0x3A
        style = xlwt.XFStyle()
        style.borders = borders

        i = self._oldRowNum
        with open(self._nameList) as f:
            nameList = f.readline().split('，')

        for row in rowsData:
            for j in range(len(row)):
                workSheet.write(i, j, str(row[j].value), style)
            self.check(nameList, row[self._nameColNum].value)
            i += 1
        workBook.save(self._output)
        return nameList




    def check(self, nameList, name):  # used to check the name in the name list
        if name in nameList:
            nameList.remove(name)
            print('file {} has been merged'.format(name))

def getlist(workSpaceDir):  # get the excel file list
    fileNames = os.listdir(workSpaceDir)
    seachedFiles = []
    for fileName in fileNames:
        if fileName.endswith('.xls'):
            seachedFiles.append(fileName)
            fileType = '.xls'
        if fileName.endswith('.xlsx'):
            seachedFiles.append(fileName)
            fileType = '.xlsx'
    return seachedFiles, fileType

# if __name__ == "__main__":
#     output = 'D:\\OneDrive\\计算机1701\\19暑假返校\\output.xls'
#     path = 'D:\\OneDrive\\计算机1701\\19暑假返校'
#     nameList = 'nameList.txt'
#     nameColNum = 4
#     oldRowNum = 2
#     fileList, fileType = getlist(path)

#     killer = XlsKiller(fileList, path, nameColNum, nameList, output, oldRowNum)
#     dataList = killer.get_row_data()
#     unDoneList = killer.write_rows(dataList)
    # print('*'*20 + 'unDone' + '*'*20)
    # for ele in unDoneList:
    #     print(ele+' unDone')