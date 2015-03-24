import os
import pandas as pd
import numpy as np
from datetime import *
from pandas import ExcelWriter

#llegir excel
df_day=pd.read_excel('C:/Users/u2d90072/Documents/feines/nmon/dia.xlsx',skiprows=1)

#inicialitzar data_frame de sortida
df_out=pd.DataFrame()


#aconseguir vector amb hores i minuts
hours=[]
for moment in df_day['Hora.1'][8:]:
    hour=int(np.array2string(moment)[0:2])
    minute=int(np.array2string(moment)[2:4])    
    hours.append(time(hour,minute))

#iterar maquines i fitxers nmon
for server in os.listdir('C:/Users/u2d90072/Documents/feines/nmon/nmon'):
  
  dia=0
  out=[]  
   
  for file in os.listdir('C:/Users/u2d90072/Documents/feines/nmon/nmon/'+server):
        
        print file
       
        df_cpu=pd.read_excel('C:/Users/u2d90072/Documents/feines/nmon/nmon/' + server+ '/' + file , sheetname='CPU_ALL')

        #buscar index hora - 10 minuts +10 minuts
        dt=datetime.combine(date.today(), hours[dia]) - timedelta(minutes=10)
        min=dt.time()
        dt=datetime.combine(date.today(), hours[dia]) + timedelta(minutes=10)
        max=dt.time()
        
        try:
            i=0
            for day in df_cpu[df_cpu.columns[0]]:
                if min<day.time():
                    break
                i=i+1 
            i_min=i
                
            i=0
            for day in df_cpu[df_cpu.columns[0]]:
                if max<day.time():
                    break
                i=i+1 
            i_max=i
                    #calcular mitjana 
            out.append(df_cpu['CPU%'][i_min:i_max].mean())
        except:
            out.append(out[-1])
        
        dia=dia+1
  #guardar al dataframe
  df_out[server]=out
    
    
df_out
#exportar a excel
writer = ExcelWriter('out.xlsx')
df_out.to_excel(writer,'Sheet1',index=False)
writer.save()





