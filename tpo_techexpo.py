import xlrd
from xlwt import Workbook
from googlesearch import search

loc = ("dbase.xlsx")

wb2 = xlrd.open_workbook(loc)
sheet = wb2.sheet_by_index(0)

wb = Workbook()
sheet1 = wb.add_sheet('Sheet 1')

for i in range(0,100): #100 = number of rows in sheet to be looped on

    l=sheet.cell_value(i, 0)

    try:
        for p in search(l, tld="co.in", num=1, stop=1, pause=1):
            count=0
            p=p.replace("https://www.","")
            p=p.replace("http://www.","")
            p=p.replace("https://","")
            p=p.replace("http://","")
            for j in search(l + "training placement", tld="co.in", num=5,stop=5, pause=1):
                    if p in j and count==0:
                        print(j,i)
                        sheet1.write(i, 0,j)
                        count=count+1

            if count==0:
                print("http://www."+p,i)
                sheet1.write(i, 0, "http://www."+p)
            if i%10==0: #saves sheet after every 10 enteries 
                try:
                    wb.save('links_output.xls')
                except:
                    print("error with saving")
    except:
        print("error occured")


wb.save('links_output.xls')
















