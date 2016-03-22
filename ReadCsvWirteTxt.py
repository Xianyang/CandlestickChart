import os
import csv

def readDic():
    list_dic = os.walk("C:\\Users\\chester.luo\\Downloads")
    for root, dirs, files in list_dic:
        for f in files:
            if f.endswith(".csv"):
                print(f)
    return files

def createTxt(fileName, txtName):
    names = []
    isAddName = 0
    with open(fileName) as f:
        f_csv = csv.reader(f)
        for row in f_csv:
            if len(row) < 2:
                continue
            if row[1] == "Name":
                isAddName = 1
                continue
            if isAddName:
                names.append(row[0])

    txtFile = open("C:\\Users\\chester.luo\\txt\\ticker_"+txtName+".txt", "w")
    for name in names:
        txtFile.write(name + " US Equity" + "\n")
    txtFile.close()

def main():
    files = readDic()
    for f in files:
        if f.endswith(".csv"):
            createTxt(".\\csv\\" + f, f[0:len(f)-13])

main()