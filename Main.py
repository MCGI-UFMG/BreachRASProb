# -*- coding: utf-8 -*-
"""
@author: Rodrigo Perdigão
"""

#Import Libraries

import win32com.client
import numpy as np
import h5py
import pandas as pd
from scipy.stats import laplace_asymmetric
import geopandas

# Creating empty lists and dataframes
poutflow=[] #Peak outflow list
depth_max=pd.DataFrame(columns=list(range(24219))) # Depth max in each cell and each simulation
depth_10=pd.DataFrame(columns=list(range(24219))) # 0 (dry) or 1 (wet) value in each cell for 10 minutes after dam breach
depth_20=pd.DataFrame(columns=list(range(24219))) # 0 (dry) or 1 (wet) value in each cell for 20 minutes after dam breach
depth_30=pd.DataFrame(columns=list(range(24219))) # 0 (dry) or 1 (wet) value in each cell for 30 minutes after dam breach
depth_40=pd.DataFrame(columns=list(range(24219))) # 0 (dry) or 1 (wet) value in each cell for 40 minutes after dam breach
depth_50=pd.DataFrame(columns=list(range(24219))) # 0 (dry) or 1 (wet) value in each cell for 50 minutes after dam breach
depth_60=pd.DataFrame(columns=list(range(24219))) # 0 (dry) or 1 (wet) value in each cell for 60 minutes after dam breach
depth_90=pd.DataFrame(columns=list(range(24219))) # 0 (dry) or 1 (wet) value in each cell for 90 minutes after dam breach
depth_120=pd.DataFrame(columns=list(range(24219))) # 0 (dry) or 1 (wet) value in each cell for 120 minutes after dam breach
depth_150=pd.DataFrame(columns=list(range(24219))) # 0 (dry) or 1 (wet) value in each cell for 150 minutes after dam breach
depth_180=pd.DataFrame(columns=list(range(24219))) # 0 (dry) or 1 (wet) value in each cell for 180 minutes after dam breach

#Breach Parameters used in each Monte Carlo Simulation
fwidth_samples=[] # Final Breach Width 
btelev_samples=[] # Botton Breach Elevation 
zleft_samples=[] # Breach Left Slope
zright_samples=[] # Breach Right Slope 
ftime_samples=[] # Breach Formation Time

bh = 61 # Breach Height
cl = 360 # Crest Length
celev = 272 # Crest Elevation
relev = 206.03 # Rock Bed Elevation

n=2000 #ESTABELECENDO N° DE REPETIÇÃO DO MONTE CARLO

for i in range(n):
    #Estimating Breach Parameters based in Da Silva and Eleutério (2023) https://doi.org/10.1111/jfr3.12900
    # Crest Elevation
    X = celev - laplace_asymmetric.rvs(loc=1, scale=0.06301609, kappa=0.69327604)*bh
    while X < relev:
        X = celev - laplace_asymmetric.rvs(loc=1, scale=0.06301609, kappa=0.69327604)*bh
    else: 
        btelev = X
    # Final Breaqch Width     
    Y = np.random.gamma(shape=0.2568,scale=1/1.6397)*cl 
    while Y > cl:
        Y = np.random.gamma(shape=0.2568,scale=1/1.6397)*cl
    else:
        fwidth = Y
        
    zleft=np.random.gamma(shape=0.4974,scale=1/0.2281)
    zright=np.random.gamma(shape=0.4974,scale=1/0.2281)
    topelev = fwidth+zleft*(celev-btelev)+zright*(celev-btelev)
    
    while topelev > 1.1*celev:
        zleft=np.random.gamma(shape=0.4974,scale=1/0.2281)
        zright=np.random.gamma(shape=0.4974,scale=1/0.2281)
        topelev = btelev+zleft*(celev-btelev)+zright*(celev-btelev)
    else:
        Z_esq = zleft
        Z_dir = zright
        
    TF=np.random.gamma(shape=1.5932,scale=1/1.5007)
    
    #NA PRÓXIMA LINHA, ALTERE O CAMINHO PARA O CAMINHO DO SEU ARQUIVO .P02
    #ATENÇÃO SÓ ALTERE O CAMINHO ENTRE ASPAS (""), MANTENHA A LETRA r ANTES DAS ASPAS.
    my_file = open(r"C:\Users\Rodrigo\Desktop\HECRAS\ESTUDO_ICOLD\RAS\bench_icold_artigo.p02")
    string_list = my_file.readlines()
    my_file.close()
    string_list[119]='Breach Geom=180,%.2f,%.2f,%.2f,%.2f,False,,,%.2f,1.7\n'%(LF,EL_FUN,Z_esq,Z_dir,TF)
    
    #NA PRÓXIMA LINHA, NOVAMENTE ALTERE O CAMINHO PARA O CAMINHO DO SEU ARQUIVO .P02
    #ATENÇÃO SÓ ALTERE O CAMINHO ENTRE ASPAS (""), MANTENHA A LETRA r ANTES DAS ASPAS.
    my_file = open(r"C:\Users\Rodrigo\Desktop\HECRAS\ESTUDO_ICOLD\RAS\bench_icold_artigo.p02", "w")
    new_file_contents = ''.join(string_list)
    my_file.write(new_file_contents)
    my_file.close()
    
    #ATRIBUINDO FUNCOES DO HECRASCONTROLLER AO OBJETO RC
    RC=win32com.client.Dispatch("RAS507.HECRASCONTROLLER")
    
    #ABRINDO E FECHANDO JANELA
    RC.ShowRAS()
    
    #ABRINDO PROJETO
    #NA PROXIMA LINHA ALTERE O CAMINHO PARA O CAMINHO DO SEU ARQUIVO .PRJ
    #ATENÇÃO SÓ ALTERE O CAMINHO ENTRE ASPAS (""), MANTENHA A LETRA r ANTES DAS ASPAS.
    RC.Project_Open(r"C:\Users\Rodrigo\Desktop\HECRAS\ESTUDO_ICOLD\RAS\bench_icold_artigo.prj")
    
    #EXECUTANDO SIMULACAO DO PLANO ATUAL
    Simulacao=RC.Compute_CurrentPlan(None,None,True)
    
    #SALVANDO PROJETO
    RC.Project_Save()
    
    #FECHANDO JANELA
    RC.QuitRAS()
    
    #NA PRÓXIMA LINHA ALTERE O CAMINHO PARA O CAMINHO DO SEU ARQUIVO .P02.HDF
    with h5py.File(r'C:\Users\Rodrigo\Desktop\HECRAS\ESTUDO_ICOLD\RAS\bench_icold_artigo.p02.hdf','r') as hdf:
       data=list(np.array(hdf.get('Results/Unsteady/Output/Output Blocks/Base Output/Unsteady Time Series/SA 2D Area Conn/barragem/Structure Variables')).T[0])
    
    #EXTRAINDO HIDROGRAMA DE RUPTURA
    qp=max(data)
    qp_list.append(qp)
    tp=data.index(qp)
    tp_list.append(tp)
    del data[:(tp+10)]
    data=list(np.diff(data)*-1)
    tb=(next((i for i,n in enumerate(data) if n < 10), len(data)))+tp
    tb_list.append(tb)
    
    #SALVANDO PARÂMETROS DE BRECHA AMOSTRADOS NA ITERACAO i
    LF_list.append(LF)
    EL_FUN_list.append(EL_FUN)
    Zesq_list.append(Z_esq)
    Zdir_list.append(Z_dir)
    TF_list.append(TF)
    end = time.time()
    tproces_list.append((end-start)/60)
    
    #SALVANDO DADOS COMPLETOS DE JUSANTE
    #AREA DA ENVOLTORIA MAXIMA
    area_max_list.append(float(geopandas.read_file(r"C:\Users\Rodrigo\Desktop\HECRAS\ESTUDO_ICOLD\RAS\dtm\Inundation Boundary (Max Value_0).shp")["Area"]))
    with h5py.File(r'C:\Users\Rodrigo\Desktop\HECRAS\ESTUDO_ICOLD\RAS\bench_icold_artigo.p02.hdf','r') as hdf:
        data=np.array(hdf.get('Results/Unsteady/Output/Output Blocks/Base Output/Unsteady Time Series/2D Flow Areas/jusante/Depth'))
        depth_max=(np.max(np.array(data),axis=0))
        depth_max=np.where(depth_max > 0.61, 1, 0)
        depth_10=np.where(data[9] > 0.61, 1, 0)
        depth_20=np.where(data[19] > 0.61, 1, 0)
        depth_30=np.where(data[29] > 0.61, 1, 0)
        depth_40=np.where(data[39] > 0.61, 1, 0)
        depth_50=np.where(data[49] > 0.61, 1, 0)
        depth_60=np.where(data[59] > 0.61, 1, 0)
        depth_90=np.where(data[89] > 0.61, 1, 0)
        depth_120=np.where(data[119] > 0.61, 1, 0)
        depth_150=np.where(data[149] > 0.61, 1, 0)
        depth_180=np.where(data[179] > 0.61, 1, 0)
        df_depth_max.loc[len(df_depth_max)]=depth_max
        df_depth_10.loc[len(df_depth_10)]=depth_10
        df_depth_20.loc[len(df_depth_20)]=depth_20
        df_depth_30.loc[len(df_depth_30)]=depth_30
        df_depth_40.loc[len(df_depth_40)]=depth_40
        df_depth_50.loc[len(df_depth_50)]=depth_50
        df_depth_60.loc[len(df_depth_60)]=depth_60
        df_depth_90.loc[len(df_depth_90)]=depth_90
        df_depth_120.loc[len(df_depth_120)]=depth_120
        df_depth_150.loc[len(df_depth_150)]=depth_150
        df_depth_180.loc[len(df_depth_180)]=depth_180

#POS_PROCESSAMENTO
    
resultados={"Nº Simulação":list(range(1,n+1,1)), "Qpico(m³/s)":qp_list,"Tpico(min)":tp_list,"Tbase(min)":tb_list,"Largura Final(m)":LF_list,"Elevação de Fundo(m)":EL_FUN_list,"Inclinação lateral esquerda":Zesq_list, "Inclinação lateral direita":Zdir_list,"Tempo de Formação(h)":TF_list}
result_df=pd.DataFrame(resultados)
result_df.set_index("Nº Simulação")

Q1=np.quantile(qp_list, 0.99)
Q10=np.quantile(qp_list, 0.9)   
Q50=np.quantile(qp_list, 0.5)
Q90=np.quantile(qp_list, 0.1)
Q95=np.quantile(qp_list, 0.05)
Q99=np.quantile(qp_list, 0.01)
Qp=[Q1,Q10,Q50,Q90,Q95,Q99]

vazoes={"Prob. excedência (%)":[1,10,50,90,95,99], "Vazões de pico (m³/s)":Qp}

QPs=pd.DataFrame(vazoes)
QPs.set_index("Prob. excedência (%)")

#NA PRÓXIMA LINHA, ALTERE O CAMINHO PARA O CAMINHO QUE VOCÊ DESEJA SALVAR A SUA PLANILHA DE SAÍDA
with pd.ExcelWriter(r"C:\Users\USUARIO\Desktop\output.xlsx") as writer:
    result_df.to_excel(writer,sheet_name="RESULTADOS")
    QPs.to_excel(writer,sheet_name="QUANTIS - QP")

