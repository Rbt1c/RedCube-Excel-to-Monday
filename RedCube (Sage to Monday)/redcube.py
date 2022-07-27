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


c1 = listInit("Order Number :", 100, 3)
c2 = listInit("Order Date   :", 100, 1)
c3 = listInit("Customer     :", 100, 0)
c4 = listInit("Item Code    :", 100, 4)
c5 = listInit("Quantity     :", 100, 9)




sourceFile.close()



