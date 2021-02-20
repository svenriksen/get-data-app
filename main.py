import pandas as pd
import xlrd
import shutil
import jinja2


path = 'a.xlsx'


#lay so dong va so cot

inpxlrd = xlrd.open_workbook(path)
mainsheet = inpxlrd.sheet_by_index(0)

cols = mainsheet.ncols
rows = mainsheet.nrows
rows=rows-1

#done
inp = pd.read_excel(path)

#khai bao bien

inp.index.names = ['delete']

while 'delete' in inp:
    inp.drop('delete',axis=1, inplace = True)

outtoan = inp.copy()
outly = inp.copy()
outhoa = inp.copy()
outsinh = inp.copy()
print(outtoan)
#xuly toan
rowtemp=0
i=0
while (i < outtoan.shape[0]):
    temp = outtoan.iloc[i,0]
    temp = str(temp)
    if (temp[0] != 'A'):
        outtoan.drop(index=rowtemp,inplace = True)
    else:
        i=i+1
        
    rowtemp=rowtemp+1

print(outtoan)
#xuly ly
rowtemp=0
i=0
while (i < outly.shape[0]):
    
    temp = outly.iloc[i,0]
    temp = str(temp)
    if (temp[0] != 'B'):
        outly.drop(index = rowtemp,inplace = True)
    else:
        i=i+1

    rowtemp=rowtemp+1


i = 0
rowtemp=0

#xuly hoa

while (i < outhoa.shape[0]):
   
    temp = outhoa.iloc[i,0]
    temp = str(temp)
    if (temp[0] != 'C'):
        outhoa.drop(index = rowtemp,inplace = True)            
    else:
        i=i+1

    rowtemp=rowtemp+1
#xuly sinh

rowtemp=0
i=0
while (i < outsinh.shape[0]):
    
    temp = outsinh.iloc[i,0]
    temp = str(temp)
    if (temp[0] != 'D'):
        outsinh.drop(index = rowtemp,inplace = True)
    else:
        i=i+1

    rowtemp=rowtemp+1
#done

#export back to excel file
writer = pd.ExcelWriter(path)
inp.to_excel(writer,'main')
outtoan.to_excel(writer,'toan')
outly.to_excel(writer,'ly')
outhoa.to_excel(writer,'hoa')
outsinh.to_excel(writer,'sinh')
writer.save()