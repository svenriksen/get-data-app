import pandas as pd
import xlrd
import shutil
import jinja2

import numpy as np

path = 'b.xlsx'


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


inpout = inp

i = 1
print(inp.shape[0])
while (i<inp.shape[0]-1):
    temp = inp.iloc[i,5]
    #print(temp, ' ', type(temp))
    print(i)
    if (np.isnan(temp)):
        
        inp.iat[i,5] = inp.iloc[i+1,5]
        print(temp, ' ', inp.iloc[i+1,5])
        print(1)
    else:
        inp.iat[i,5] = np.nan
        print(0)
    i=i+1

print(inp)

tienmat = inp.copy()
tienguinh = inp.copy()
congcudungcu = inp.copy()
thuegtgt = inp.copy()
cpquanly = inp.copy()
ptcnb = inp.copy()
dthdtc = inp.copy()
hh = inp.copy()
dtbh = inp.copy()
tgtgtdm = inp.copy()
tgtgthnk = inp.copy()
txnk = inp.copy()


#01


rowtemp=1
i=1
while (i < tienmat.shape[0]):
    temp = tienmat.iloc[i,4]
    temp = str(temp)
    temp2 = str(tienmat.iloc[i,5])
    
    if (temp != '111.0' and temp2 != '111.0'):
        tienmat.drop(index=rowtemp,inplace = True)
        
    else:
        if (temp == '111.0'):
            tienmat.iloc[i,5] = np.nan
        else:
            
            tienmat.iloc[i,4] = np.nan
        i=i+1
        
    rowtemp=rowtemp+1



#2

rowtemp=1
i=1
while (i < tienguinh.shape[0]):
    temp = tienguinh.iloc[i,4]
    temp = str(temp)
    temp2 = str(tienguinh.iloc[i,5])
    
    if (temp != '112.0' and temp2 != '112.0'):
        tienguinh.drop(index=rowtemp,inplace = True)
    else:
        if (temp == '112.0'):
            tienguinh.iloc[i,5]=np.nan
        else:
            tienguinh.iloc[i,4]=np.nan
        i=i+1
        
    rowtemp=rowtemp+1


#3


rowtemp=1
i=1
while (i < congcudungcu.shape[0]):
    temp = congcudungcu.iloc[i,4]
    temp = str(temp)
    temp2 = str(congcudungcu.iloc[i,5])
    
    if (temp != '153.0' and temp2 != '153.0'):
        congcudungcu.drop(index=rowtemp,inplace = True)
    else:
        if (temp == '153.0'):
            congcudungcu.iloc[i,5]=np.nan
        else:
            congcudungcu.iloc[i,4]=np.nan
        i=i+1
        
    rowtemp=rowtemp+1
#4
rowtemp=1
i=1
while (i < thuegtgt.shape[0]):
    temp = thuegtgt.iloc[i,4]
    temp = str(temp)
    temp2 = str(thuegtgt.iloc[i,5])
    
    if (temp != '1331.0' and temp2 != '1331.0'):
        thuegtgt.drop(index=rowtemp,inplace = True)
    else:
        if (temp == '1331.0'):
            thuegtgt.iloc[i,5]=np.nan
        else:
            thuegtgt.iloc[i,4]=np.nan
        i=i+1
        
    rowtemp=rowtemp+1
#5

rowtemp=1
i=1
while (i < cpquanly.shape[0]):
    temp = cpquanly.iloc[i,4]
    temp = str(temp)
    
    temp2 = str(cpquanly.iloc[i,5])
    
    if (temp != '6422.0' and temp2 != '6422.0'):
        cpquanly.drop(index=rowtemp,inplace = True)
    else:
        if (temp == '6422.0'):
            cpquanly.iloc[i,5]=np.nan
        else:
            cpquanly.iloc[i,4]=np.nan
        i=i+1
        
    rowtemp=rowtemp+1

#6

rowtemp=1
i=1
while (i < ptcnb.shape[0]):
    temp = ptcnb.iloc[i,4]
    temp = str(temp)
    temp2 = str(ptcnb.iloc[i,5])
    
    if (temp != '331.0' and temp2 != '331.0'):
        ptcnb.drop(index=rowtemp,inplace = True)
    else:
        if (temp == '331.0'):
            ptcnb.iloc[i,5]=np.nan
        else:
            ptcnb.iloc[i,4]=np.nan
        i=i+1
        
    rowtemp=rowtemp+1
#7
rowtemp=1
i=1
while (i < dthdtc.shape[0]):
    temp = dthdtc.iloc[i,4]
    temp = str(temp)
    temp2 = str(dthdtc.iloc[i,5])
    
    if (temp != '515.0' and temp2 != '515.0'):
        dthdtc.drop(index=rowtemp,inplace = True)
    else:
        if (temp == '515.0'):
            dthdtc.iloc[i,5]=np.nan
        else:
            dthdtc.iloc[i,4]=np.nan
        i=i+1
        
    rowtemp=rowtemp+1
#8
rowtemp=1
i=1
while (i < hh.shape[0]):
    temp = hh.iloc[i,4]
    temp = str(temp)
    temp2 = str(hh.iloc[i,5])
    
    if (temp != '156.0' and temp2 != '156.0'):
        hh.drop(index=rowtemp,inplace = True)
    else:
        if (temp == '156.0'):
            hh.iloc[i,5]=np.nan
        else:
            hh.iloc[i,4]=np.nan
        i=i+1
        
    rowtemp=rowtemp+1
#9

rowtemp=1
i=1
while (i < dtbh.shape[0]):
    temp = dtbh.iloc[i,4]
    temp = str(temp)
    temp2 = str(dtbh.iloc[i,5])
    
    if (temp != '511.0' and temp2 != '511.0'):
        dtbh.drop(index=rowtemp,inplace = True)
    else:
        if (temp == '511.0'):
            dtbh.iloc[i,5]=np.nan
        else:
            dtbh.iloc[i,4]=np.nan
        i=i+1
        
    rowtemp=rowtemp+1
#10

rowtemp=1
i=1
while (i < tgtgtdm.shape[0]):
    temp = tgtgtdm.iloc[i,4]
    temp = str(temp)
    temp2 = str(tgtgtdm.iloc[i,5])
    
    if (temp != '33311.0' and temp2 != '33311.0'):
        tgtgtdm.drop(index=rowtemp,inplace = True)
    else:
        if (temp == '33311.0'):
            tgtgtdm.iloc[i,5]=np.nan
        else:
            tgtgtdm.iloc[i,4]=np.nan
        i=i+1
        
    rowtemp=rowtemp+1
#11

rowtemp=1
i=1
while (i < tgtgthnk.shape[0]):
    temp = tgtgthnk.iloc[i,4]
    temp = str(temp)
    temp2 = str(tgtgthnk.iloc[i,5])
    
    if (temp != '33312.0' and temp2 != '33312.0'):
        tgtgthnk.drop(index=rowtemp,inplace = True)
    else:
        if (temp == '33312.0'):
            tgtgthnk.iloc[i,5]=np.nan
        else:
            tgtgthnk.iloc[i,4]=np.nan
        i=i+1
        
    rowtemp=rowtemp+1
#12

rowtemp=1
i=1
while (i < txnk.shape[0]):
    temp = txnk.iloc[i,4]
    temp = str(temp)
    temp2 = str(txnk.iloc[i,5])
    
    if (temp != '3333.0' and temp2 != '3333.0'):
        txnk.drop(index=rowtemp,inplace = True)
    else:
        if (temp == '3333.0'):
            txnk.iloc[i,5]=np.nan
        else:
            txnk.iloc[i,4]=np.nan
        i=i+1
        
    rowtemp=rowtemp+1

#export back to excel filewriter = pd.ExcelWriter(path)
writer = pd.ExcelWriter(path)
inpout.to_excel(writer,'Sheet 1')
tienmat.to_excel(writer,'tien mat')
tienguinh.to_excel(writer,'tien gui ngan hang')
congcudungcu.to_excel(writer,'cong cu dung cu')
thuegtgt.to_excel(writer,'thue gtgt')
cpquanly.to_excel(writer,'CPQly')
ptcnb.to_excel(writer,'Phai tra cho nguoi ban')
dthdtc.to_excel(writer,'doanh thu hoat dong tai chinh')
hh.to_excel(writer,'hang hoa')
dtbh.to_excel(writer,'doanh thu ban hang')
tgtgtdm.to_excel(writer,'thue gtgt dau m')
tgtgthnk.to_excel(writer,'thue gtgt hang nhap khau')
txnk.to_excel(writer,'thue xuat nhap khau')
writer.save()