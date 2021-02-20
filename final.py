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


hangnhap = inp.copy()
hangnoi = inp.copy()
cpnhaphang = inp.copy()
cpguihang = inp.copy()
cptiepkhach = inp.copy()
tscodinh = inp.copy()
luong = inp.copy()
chict = inp.copy()
print(inp)


#01


rowtemp=1
i=1
while (i < hangnhap.shape[0]):
    temp = hangnhap.iloc[i,1]
    temp = str(temp)
    if (temp != '01'):
        hangnhap.drop(index=rowtemp,inplace = True)
    else:
        i=i+1
        
    rowtemp=rowtemp+1

print(hangnhap)


#02

rowtemp=1
i=1
while (i < hangnoi.shape[0]):
    temp = hangnoi.iloc[i,1]
    temp = str(temp)
    if (temp != '02'):
        hangnoi.drop(index=rowtemp,inplace = True)
    else:
        i=i+1
        
    rowtemp=rowtemp+1

#03

rowtemp=1
i=1
while (i < cpnhaphang.shape[0]):
    temp = cpnhaphang.iloc[i,1]
    temp = str(temp)
    if (temp != '03'):
        cpnhaphang.drop(index=rowtemp,inplace = True)
    else:
        i=i+1
        
    rowtemp=rowtemp+1

#04

rowtemp=1
i=1
while (i < cpguihang.shape[0]):
    temp = cpguihang.iloc[i,1]
    temp = str(temp)
    if (temp != '04'):
        cpguihang.drop(index=rowtemp,inplace = True)
    else:
        i=i+1
        
    rowtemp=rowtemp+1

#05

rowtemp=1
i=1
while (i < cptiepkhach.shape[0]):
    temp = cptiepkhach.iloc[i,1]
    temp = str(temp)
    if (temp != '05'):
        cptiepkhach.drop(index=rowtemp,inplace = True)
    else:
        i=i+1
        
    rowtemp=rowtemp+1

#06

rowtemp=1
i=1
while (i < tscodinh.shape[0]):
    temp = tscodinh.iloc[i,1]
    temp = str(temp)
    if (temp != '06'):
        tscodinh.drop(index=rowtemp,inplace = True)
    else:
        i=i+1
        
    rowtemp=rowtemp+1

#07

rowtemp=1
i=1
while (i < luong.shape[0]):
    temp = luong.iloc[i,1]
    temp = str(temp)
    if (temp != '07'):
        luong.drop(index=rowtemp,inplace = True)
    else:
        i=i+1
        
    rowtemp=rowtemp+1

#08

rowtemp=1
i=1
while (i < chict.shape[0]):
    temp = chict.iloc[i,1]
    temp = str(temp)
    if (temp != '08'):
        chict.drop(index=rowtemp,inplace = True)
    else:
        i=i+1
        
    rowtemp=rowtemp+1



#export back to excel file
writer = pd.ExcelWriter(path)
inp.to_excel(writer,'NK CHUNG')
hangnhap.to_excel(writer,'HÀNG NHẬP')
hangnoi.to_excel(writer,'HÀNG NỘI')
cpnhaphang.to_excel(writer,'CHI PHÍ NHẬP HÀNG')
cpguihang.to_excel(writer,'CHI PHI GỬI HÀNG')
cptiepkhach.to_excel(writer,'CHI PHÍ TIẾP KHÁCH')
tscodinh.to_excel(writer,'TS CỐ ĐỊNH')
luong.to_excel(writer,'LƯƠNG')
chict.to_excel(writer,'CHI CT')


writer.save()