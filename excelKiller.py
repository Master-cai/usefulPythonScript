# -*-coding:utf-8-*-
import os
import argparse
from xlsKiller import XlsKiller
from xlsxKiller import XlsxKiller


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



def addarg():
    parser = argparse.ArgumentParser()
    parser.add_argument('-w', '--workPath', type=str,
                        help='the path of your excels', default=os.getcwd())
    parser.add_argument('-o', '--oldRowNum', type=int,
                        help='the row number you want to begin', default=2)
    parser.add_argument('-c', '--column', type=int,
                        help='the column number of the name', default=3)
    parser.add_argument('-O', '--outPutFiles', 
                        type=str,
                        help='the file you want ro save the output, it must exist',
                        default='output.xls')
    args = parser.parse_args()
    return args


def main():
    args = addarg()
    output = args.outPutFiles.strip('\'')
    path = args.workPath.strip('\'')
    nameList = 'nameList.txt'
    nameColNum = args.column
    oldRowNum = args.oldRowNum
    fileList, fileType = getlist(path)
    if fileType == '.xls':
        killer = XlsKiller(fileList, path, nameColNum, nameList, output, oldRowNum)
        dataList = killer.get_row_data()
        unDoneList = killer.write_rows(dataList)
    if fileType == '.xlsx':
        killer = XlsxKiller(fileList, path, nameColNum, nameList, output, oldRowNum)
        dataList = killer.get_row_data()
        unDoneList = killer.write_rows(dataList)


    print('*'*20 + 'unDone' + '*'*20)
    for ele in unDoneList:
        print(ele+' unDone')


if __name__ == '__main__':
    main()
