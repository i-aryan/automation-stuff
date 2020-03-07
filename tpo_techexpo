import xlrd
from xlwt import Workbook
from googlesearch import search

loc = ("dbase.xlsx")

wb2 = xlrd.open_workbook(loc)
sheet = wb2.sheet_by_index(0)

wb = Workbook()
sheet1 = wb.add_sheet('Sheet 1')

for i in range(0,100):

    l=sheet.cell_value(i, 0)

    for p in search(l, tld="co.in", num=1, stop=1, pause=1):
        count=0
        p=p.replace("https://www.","")
        p=p.replace("http://www.","")
        p=p.replace("https://","")
        p=p.replace("http://","")
        for j in search(l + "training placement", tld="co.in", num=5,stop=5, pause=1):
                if p in j and count==0:
                    print(j)
                    sheet1.write(i, 0,j)
                    count=count+1

        if count==0:
            print("http://www."+p)
            sheet1.write(i, 0, "http://www."+p)


wb.save('links_output.xls')
















