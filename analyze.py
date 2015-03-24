import os
import sys
import pandas as pd
import numpy

#inicialitzar data_frame de sortida
df_out=pd.DataFrame()
min=numpy.int64(sys.argv[1])
max=numpy.int64(sys.argv[2])


df_out['Machines']=pd.ExcelFile(os.listdir('.')[1]).sheet_names[2:]

for file in os.listdir('.')[1:]:
    print file
    sheets=pd.ExcelFile(file).sheet_names[2:]
        
    out=[]

    for sheet in sheets:
                
        df=pd.read_excel(file,sheetname=sheet)
        df.columns=['hour','data']
                
        i=0   
        for index in df.hour:
            if index==min:
                break
            i=i+1
        i_min=i
        
        
        i=0   
        for index in df.hour:
            if index==max:
                break
            i=i+1
        i_max=i
        
        out.append(df.data[i_min:i_max].mean())
    
    df_out[file]=out
    
df_out
        
#exportar a excel
writer = pd.ExcelWriter('out.xlsx')
df_out.to_excel(writer,'Sheet1',index=False)
writer.save()





