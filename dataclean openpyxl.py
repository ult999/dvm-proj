from openpyxl import Workbook
wb=Workbook()
ws=wb.active

file1=open("G:\\messy_data.csv","r")


count1=0
str1=" "
while str1:
    count1=count1+1
    str1=file1.readline()
    print(str1)
    str1=str1.strip()  #to remove new line character
    l1=str1.split(",")

    if not str1: #to break loop if end of file has been reached
        break

    l2=[0,0,0,0]
    for i in l1:
        i=i.lstrip(" ") #some data entries have a space in 0 index which needs to be removed to properly scan the string
        
        if i[0].isalpha() and len(i)>1:
            l2[0]=i
        elif '2017' in i:   #assuming that all id no are from 2017 batch, this condition can be changed if some other quantity is fixed instead of year
            l2[1]=i
        elif i.isdigit():   
            l2[2]=i
        else:
            l2[3]=i
    print(l2)

    

    
    
    
    for j in (0,len(l2)-1):
        a=ws.cell(row=count1, column=j+1)
        a.value=str(l2[j])

    
        
        

    
file1.close()
wb.save('G://clean_data.xlsx')



    

