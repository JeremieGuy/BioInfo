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

    with open("interractions2.csv", "w") as out:
        stringGenes = ""
        for gene in genes: 
            stringGenes += gene + ","
        line = "GENE," + "Annotation," + stringGenes
        out.write(line + "\n")

        for i in indices:
            line = genesLine[i] + "," + annotations[i][3:]
            for gene in genes:
                if math.isnan(data[gene].values[i]):
                    line += "," + str(0)
                else : 
                    line += "," + str(data[gene].values[i])
            out.write(line + "\n")
           
    out.close()
                
                

def main():

    CoIp = extractData("Table_PPIcoIP.xlsx")
    makeFile(CoIp)


main()
# print(data[:5])