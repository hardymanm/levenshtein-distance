# Reading an excel file using Python 
import xlrd
#levenshtien algorithm
from Levenshtein import distance as Distance

# Give the location of the file 
loc= ("C:/Users/Admin/Desktop/unimas_kursus/test_malay.xlsx")

result=open("result_malay.txt","w")
  
# To open Workbook 
wb= xlrd.open_workbook(loc) 
sheet = wb.sheet_by_index(0) 
  
# For row 0 and column 0 
sheet.cell_value(0, 0)



#text1=['pretty','Sick','Thick','Your','Shark','Shoulder','Elephant','Name','Them','want']
text1=['Susu','Gula','Kerusi','Epal','Penyu','Isnin','Menyiram','Terjatuh','Berhati-hati','Buah-buahan']
cal_id=[]
cal_result=[]

def cal_distance(rows):
    
    for x in range(1,11):
        #print (sheet.cell_value(i,x))
        #check if the value is string if not ignore.
        if (isinstance(sheet.cell_value(rows,x),str)):
            #total=Distance(sheet.cell_value(i,x),text1[x-1])
            print(Distance(sheet.cell_value(rows,x),text1[x-1]))
            result.write(str(Distance(sheet.cell_value(rows,x),text1[x-1]))+" ")

    result.write("\n")
    

for i in range(sheet.nrows):
    textID= "[ID:" + str(i) +"]"
    print ("ID ",i)
    result.write(textID +" ")
    cal_distance(i)
    


result.close()
