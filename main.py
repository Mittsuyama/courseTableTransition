# main.py

import xlwt
import xlrd
import xlutils.copy
from bs4 import BeautifulSoup

class courseTableTransiton:
    workbook = None
    mySheet = None
    subName = []
    teacName = []
    durWeek = []
    posit = []
    timeTable = []

    def __init__(self):
        self.workbook = xlrd.open_workbook('course.xls')
        self.mySheet = self.workbook.sheet_by_index(0)

    def dealString(self, dataStr, weekNum, courseNum):
        judge = 0
        soup = BeautifulSoup(dataStr, features = 'html.parser')
        for content in soup.descendants:
            singleStr = str(content)
            if judge == 0:
                judge = 1
                self.subName.append(singleStr)
                self.timeTable.append([weekNum, courseNum])
            else:
                judge = 0
                for i in range(0, len(singleStr)):
                    if singleStr[i] == '[':
                        self.teacName.append(singleStr[:i])
                        break
                for i in range(0, len(singleStr)):
                    for j in range(i, len(singleStr)):
                        if singleStr[i] == '[' and singleStr[j] == ']':
                            self.durWeek.append(singleStr[i + 1:j])
                            break
                for i in range(0, len(singleStr)):
                    if singleStr[i] == ']':
                        self.posit.append(singleStr[i + 1:])

    def getData(self):
        for weekNum in range(1, 7):
            for courseNum in range(1, 6):
                self.dealString(self.mySheet.cell(weekNum + 1, courseNum + 1).value, weekNum, courseNum)
        # print debug
        print(self.subName)
        print(self.teacName)
        print(self.durWeek)
        print(self.posit)
        print(self.timeTable)

    def main(self):
        self.getData()

if __name__ == '__main__':
    courseTableTransiton().main()