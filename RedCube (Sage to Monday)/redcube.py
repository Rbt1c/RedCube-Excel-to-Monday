import xlrd

book = xlrd.open_workbook("test.xls")       #replace "test.xls" with desired xls file.
sh = book.sheet_by_index(0)

sourceFile = open('info.txt','w')

class listInit:
    def __init__(self, name, number, colum):
        self.name = name
        self.number = number
        self.colum = colum
        for x in range(number):
            cell =  sh.cell_value(rowx = x, colx = colum)
            if cell == "":
                pass
            if cell == name:
                pass
            else:
                print("{0} : {1}".format(name, cell), file = sourceFile )


c1 = listInit("People", 5, 0)
c2 = listInit("Salary", 5, 1)
c3 = listInit("Residence", 5, 2)

print("The Fox Jumped Over The Hair!")



sourceFile.close()



