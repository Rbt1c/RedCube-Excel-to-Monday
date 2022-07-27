import xlrd

book = xlrd.open_workbook("SO_OpenOrders.xls")       #replace "test.xls" with desired xls file.
sh = book.sheet_by_index(0)

sourceFile = open('info.txt','w')

class listInit:
    def __init__(self, name, numberOfRowsRead, colum):
        self.name = name
        self.numberOfRowsRead = numberOfRowsRead
        self.colum = colum

        numberOfRowsRead  = int(numberOfRowsRead)
        colum   = int(colum)

        for x in range(numberOfRowsRead):
            cell =  sh.cell_value(rowx = x, colx = colum)
            if cell == "":
                pass
            elif cell == name:
                pass
            else:
                print("{0} : {1}".format(name, cell), file = sourceFile )

def EndProgram():
    sourceFile.close()
    exit()

def InitLists():
    nameOfListInit  = input("What is the name of the data you want to process?              :")
    numberOfRows    = input("How many rows do you wish to print?                            :")
    columNumber     = input("in what column is the data you want to process? (In decimal)   :")

    myList = listInit(nameOfListInit, numberOfRows, columNumber)

    createAnotherList()

def createAnotherList():
    AnotherList     = input("do you wish to create another list? [Yes/No] (default: No)     :")
    AnotherList     = str(AnotherList)
    if AnotherList == "no":
        EndProgram()
    elif AnotherList == "yes":
        InitLists()

InitLists()



