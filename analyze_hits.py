import os
import pandas as pd
import numpy

#inicialitzar data_frame de sortida
df_out=pd.DataFrame()
df_out['Machines']=ex=pd.ExcelFile(os.listdir('.')[0])

for file in os.listdir('.')[:-1]:
    print file
    ex=pd.ExcelFile(file)      
    out=[]

    for sheet in ex.sheet_names[2:]:
        df=ex.parse(sheet)        
        
        try:                                  
            i=df[df.columns[1]].argmax()
            i_min=i-10
            i_max=i+10
            
            out.append(df[df.columns[1]][i_min:i_max].mean())
        except:
            out.append(0)            
   
    df_out[file]=out

#exportar a excel
writer = pd.ExcelWriter('out.xlsx')
df_out.to_excel(writer,'Sheet1',index=False)
writer.save()