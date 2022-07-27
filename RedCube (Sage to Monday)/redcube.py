import xlrd

book = xlrd.open_workbook("SO_OpenOrders.xls")
sh = book.sheet_by_index(0)

sourceFile = open('info.txt','w')

print("the fox jumped over the hair!")

class listInit:
    def __init__(self, name, number, colum):
        self.name = name
        self.number = number
        self.colum = colum
        for x in range(number):
            cell =  sh.cell_value(rowx = x, colx = colum)
            if cell == "":
                pass
            else:
                print("{0} : {1}".format(name, cell), file = sourceFile )







sourceFile.close()



