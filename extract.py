# import os
# import xlrd
import pandas as pd
import math

data = []

def extractData(filename):
    # directory = os.getcwd()
    # path = directory + filename

    file = pd.ExcelFile(filename)
    Summary = pd.read_excel(file, sheet_name=0)
    CoIp = pd.read_excel(file, sheet_name=1)
    PrismaUnmodified = pd.read_excel(file, sheet_name=2)
    PrismaSignunfilterd = pd.read_excel(file, sheet_name=3)
    PrismaPhospho = pd.read_excel(file, sheet_name=4)

    return CoIp

def makeFile(data):
    columns = data.columns
    genes = columns[3:-1]
    genesLine = data[columns[1]].values
    annotations = data[columns[0]].values

    indices = [i for i in range(len(data[genes[0]].values))]

    with open("interractions.sif", "w") as out:

        for gene in genes:
            print("GENE : ", gene)
            for value, i in zip(data[gene].values, indices) :
                if not math.isnan(value) :
                    # print(i, annotations[i][3:].replace(' ', '-'), genesLine[i], value  )
                    note = annotations[i][3:].replace(' ', '-')
                    if note == "NA":
                        line = gene + " " + genesLine[i]
                    else : 
                        line = gene + " " + note + " " + genesLine[i]
                    out.write(line + "\n")
    out.close()
                
                

def main():

    CoIp = extractData("Table_PPIcoIP.xlsx")
    makeFile(CoIp)


main()
# print(data[:5])