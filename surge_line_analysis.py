#!/usr/bin/env python
# -*- coding: utf-8 -*-
__author__ = "Al Shahriar"
__copyright__ = "Copyright Danfoss Turbocor 2023, The Meanline ML Project"
__credits__ = ["Al Shahriar"]
__license__ = "Private"
__version__ = "1.0.0"
__maintainer__ = "Al Shahriar"
__email__ = "al.shahriar@danfoss.com"
__status__ = "Pre-production"

"""Detail description

@Usage:
    Reading ansys csv file and reading the output
@Date: 
    July 14 2023
@Files
    output.csv
@Output
    print
"""
# %% Load libraries
import numpy as np
import pandas as pd
import tensorflow as tf
import pickle

# System or built-in
import os
import warnings
from datetime import datetime
import shutil
import time
import re

import xlsxwriter

import matplotlib.pyplot as plt

from pandas import read_csv
from pandas.plotting import scatter_matrix

from matplotlib import pyplot
# %%
if(1):
    fnames = ["mon_55k_110_07000","mon_55k_115_03170","mon_55k_120_00760","mon_55k_125_01710",\
              "mon_55k_130_01850","mon_55k_135_04500",\
              "mon_55k_140_03000","mon_55k_145_00890","mon_55k_150_02500","mon_55k_151_01750"]
    speedline_txt = "55_0k"    
else:
    fnames = ["mon_40k_060_20000","mon_40k_070_17500","mon_40k_080_17700","mon_40k_090_17900",\
              "mon_40k_100_18500","mon_40k_110_16630",\
              "mon_40k_120_19700","mon_40k_130_19800"]
    speedline_txt = "40_0k"

df_mean = pd.DataFrame()
df_std = pd.DataFrame()
df_max = pd.DataFrame()
df_min = pd.DataFrame()
df_abs_max = pd.DataFrame()
df_summary = pd.DataFrame()
mf_list = []

mf_out_txt = 'MF Outlet [kg s^-1]'
mf_in_txt = 'MF Inlet [kg s^-1]'
ps_out_txt = 'P Static Outlet [Pa]'
pt_out_txt = 'P Total Outlet [Pa]'
ps_in_txt = 'P Static Inlet [Pa]'

ts_out_txt = 'T Static Outlet [Pa]'
tt_out_txt = 'T Total Outlet [Pa]'
ts_in_txt = 'T Static Inlet [Pa]'

fx1p2 = "Fx1+Fx2";
fx1p2p3 = "Fx1+Fx2+Fx3"
fy1p2 = "Fy1+Fy2";
fy1p2p3 = "Fy1+Fy2+Fy3"
f_total_12 = "F_tot(1+2)"
f_total_123 = "F_tot(1+2+3)"

dirName = r"C:\Users\U423018\asr\surge_line_analysis\monitor_"+speedline_txt

writer = pd.ExcelWriter("processed_data_"+speedline_txt+".xlsx", engine='xlsxwriter')
nMF = len(fnames)
# %%
for i in range(nMF):
    fname = fnames[i]+".csv";
    case_name = fnames[i][4:11]
    print(case_name)
    fullname = os.path.join(dirName,fname);
    dfraw = pd.read_csv(fullname, skiprows=[0,1,2,3])
    timeStep = int(fnames[i][-5:])
    s_index = int(dfraw[dfraw["Accumulated Timestep"] == timeStep].index.values)
    e_index = len(dfraw.index)
    df = dfraw.truncate(before=s_index,after=e_index)
    df.reset_index(drop=True, inplace=True)
    col_names = df.columns.to_list()
    col_names = [lines.replace('Monitor Point: ','') for lines in col_names] # removing spaces
    df.columns = col_names;
    df = df.rename(columns={'Pres IMP1 IN [Pa]':ps_in_txt,'Pres VOL EX [Pa]':ps_out_txt,\
                            'TP VOL EX [Pa]':pt_out_txt,'MF IMP1 IN [kg s^-1]':mf_in_txt,\
                            'MF VOL EX [kg s^-1]':mf_out_txt, 'Temp IMP1 IN [K]': ts_in_txt,\
                            'Temp VOL EX [K]':ts_out_txt, 'TT VOL EX [K]': tt_out_txt })
    df[mf_out_txt] = df[mf_out_txt].abs()
    mf = abs(df[mf_out_txt][0])
    mf_txt = f'{mf:.3f}'
    mf_list.append(mf_txt)
    df['Torque1 [J]'] = df['Torque1 [J]'].astype(float)
    df['Torque2 [J]'] = df['Torque1 [J]'].astype(float)
    df['Torque3 [J]'] = df['Torque1 [J]'].astype(float)
    df[fx1p2] = df['Fx1 [N]']+df['Fx2 [N]']
    df[fx1p2p3] = df['Fx1 [N]']+df['Fx2 [N]']+df['Fx3 [N]']
    df[fy1p2] = df['Fy1 [N]']+df['Fy2 [N]']
    df[fy1p2p3] = df['Fy1 [N]']+df['Fy2 [N]']+df['Fy3 [N]']
    df[f_total_12] = np.linalg.norm(df[[fx1p2,fy1p2]].values,axis=1)
    df[f_total_123] = np.linalg.norm(df[[fx1p2p3,fy1p2p3]].values,axis=1)
    df_mean[case_name] = df.mean()
    df_std[case_name] = df.std()
    df_max[case_name] = df.max()
    df_min[case_name] = df.min()
    df_abs_max[case_name] = df.abs().max()
    # always put at the end
    df.loc["mean"] = df.mean()
    df.loc["std"] = df.std()
    df.loc["max"] = df.max()
    df.loc["min"] = df.min()
# df_mean.to_excel(writer, sheet_name="mean")
# df_std.to_excel(writer, sheet_name="std")
# df_max.to_excel(writer, sheet_name="max")
# df_min.to_excel(writer, sheet_name="min")
# df_abs_max.to_excel(writer, sheet_name="abs max")
# %% Summerize
data_to_add = [mf_in_txt,fx1p2,fy1p2,f_total_12,fy1p2,\
                          fy1p2p3,f_total_123,ps_out_txt,ts_out_txt]
data_to_add_to_summary_title = ["m","FX (Fx1 + Fx2)","FY (Fy1 + Fy2)","F_total","FX (Fx1 + Fx2 + Fx3)",\
                          "FY (Fy1 + Fy2 + Fy3)","F_total","Outlet Stastic Pressure","Outlet Stastic Temperature"]
n_data_summary = len(data_to_add)
mf_list.insert(0, " ")
mf_list.insert(1, "m")
count = 1
df_summary[str(count)] = mf_list
for i in range(n_data_summary):
    dft = pd.DataFrame()
    data_txt = data_to_add[i]
    # abs max
    col_temp = df_abs_max.loc[data_txt].to_list()
    col_temp.insert(0," ")
    col_temp.insert(1,"absolute max")    
    count = count + 1
    dft[str(count)] = col_temp
    # avg
    col_temp = df_mean.loc[data_txt].to_list()
    col_temp.insert(0,data_to_add_to_summary_title[i])
    col_temp.insert(1,"average")    
    count = count + 1
    dft[str(count)] = col_temp
    # std
    col_temp = df_std.loc[data_txt].to_list()
    col_temp.insert(0," ")
    col_temp.insert(1,"stdv")    
    count = count + 1
    dft[str(count)] = col_temp
    df_summary = pd.concat([df_summary,dft],axis=1)

df_summary.to_excel(writer, sheet_name="summary",index=False,header=False)
# %% raw data and processed data
for i in range(nMF):
    fname = fnames[i]+".csv";
    case_name = fnames[i][4:11]
    print(case_name)
    fullname = os.path.join(dirName,fname);
    dfraw = pd.read_csv(fullname, skiprows=[0,1,2,3])
    timeStep = int(fnames[i][-5:])
    s_index = int(dfraw[dfraw["Accumulated Timestep"] == timeStep].index.values)
    e_index = len(dfraw.index)
    df = dfraw.truncate(before=s_index,after=e_index)
    df.reset_index(drop=True, inplace=True)
    col_names = df.columns.to_list()
    col_names = [lines.replace('Monitor Point: ','') for lines in col_names] # removing spaces
    df.columns = col_names;
    df = df.rename(columns={'Pres IMP1 IN [Pa]':ps_in_txt,'Pres VOL EX [Pa]':ps_out_txt,\
                            'TP VOL EX [Pa]':pt_out_txt,'MF IMP1 IN [kg s^-1]':mf_in_txt,\
                            'MF VOL EX [kg s^-1]':mf_out_txt, 'Temp IMP1 IN [K]': ts_in_txt,\
                            'Temp VOL EX [K]':ts_out_txt, 'TT VOL EX [K]': tt_out_txt })
    df[mf_out_txt] = df[mf_out_txt].abs()
    mf = abs(df[mf_out_txt][0])
    mf_txt = f'{mf:.3f}'
    df['Torque1 [J]'] = df['Torque1 [J]'].astype(float)
    df['Torque2 [J]'] = df['Torque1 [J]'].astype(float)
    df['Torque3 [J]'] = df['Torque1 [J]'].astype(float)
    df.to_excel(writer, sheet_name=mf_txt+" raw",index=False)
    # df.loc["min"] = df.min()
    df[fx1p2] = df['Fx1 [N]'] + df['Fx2 [N]']
    df[fx1p2p3] = df['Fx1 [N]'] + df['Fx2 [N]'] + df['Fx3 [N]']
    df[fy1p2] = df['Fy1 [N]'] + df['Fy2 [N]']
    df[fy1p2p3] = df['Fy1 [N]'] + df['Fy2 [N]'] + df['Fy3 [N]']
    df["F_tot(1+2)"] = np.linalg.norm(df[[fx1p2,fy1p2]].values,axis=1)
    df["F_tot(1+2+3)"] = np.linalg.norm(df[[fx1p2p3,fy1p2p3]].values,axis=1)
    # must be at the end
    df.loc["mean"] = df.mean()
    df.loc["std"] = df.std()
    df.loc["abs max"] = df.abs().max()
    dft = df
    dft['Accumulated Timestep'].loc["mean"] = "mean"
    dft['Accumulated Timestep'].loc["std"] = "stdv"
    dft['Accumulated Timestep'].loc["abs max"] = "abs max"
    dft.to_excel(writer, sheet_name=mf_txt, index=False)
# %%    
writer.close()

# %%
plt.figure()
nMF = len(fnames)
for i in range(nMF):
    fname = fnames[i]+".csv";
    fullname = os.path.join(dirName,fname);
    dfraw = pd.read_csv(fullname, skiprows=[0,1,2,3])
    # put -1 for end value
    timeStep = int(fnames[i][-5:])
    s_index = int(dfraw[dfraw["Accumulated Timestep"] == timeStep].index.values)
    e_index = len(dfraw.index)
    df = dfraw.truncate(before=s_index,after=e_index)
    df.reset_index(drop=True, inplace=True)
    col_names = df.columns.to_list()
    col_names = [lines.replace('Monitor Point: ','') for lines in col_names] # removing spaces
    df.columns = col_names;
    df = df.rename(columns={'Pres IMP1 IN [Pa]':ps_in_txt,'Pres VOL EX [Pa]':ps_out_txt,\
                                  'TP VOL EX [Pa]':pt_out_txt,'MF IMP1 IN [kg s^-1]':mf_in_txt,\
                                  'MF VOL EX [kg s^-1]':mf_out_txt, 'Temp IMP1 IN [K]': ts_in_txt,\
                                  'Temp VOL EX [K]':ts_out_txt, 'TT VOL EX [K]': tt_out_txt })
    case_name = fnames[i][4:11]
    
    df[mf_out_txt] = df[mf_out_txt].abs()    
    df['MF SEAL EX [kg s^-1]'] = df['MF SEAL EX [kg s^-1]'].abs()        
    
    mf = abs(df[mf_out_txt][0])
    mf_txt = f'{mf:.2f}'
    title_txt =  "RPM: "+case_name[0:3]+" "+"MF: "+mf_txt
    
    plt.clf()
    df.plot(x='Accumulated Timestep',y=['Fx1 [N]', 'Fx2 [N]', 'Fx3 [N]'])
    plt.title(title_txt)
    plt.savefig(speedline_txt+"/"+"Fx"+case_name+r".png",dpi=300)
    plt.close()
    
    plt.clf()
    df.plot(x='Accumulated Timestep',y=['Fy1 [N]', 'Fy2 [N]', 'Fy3 [N]'])
    plt.title(title_txt)    
    plt.savefig(speedline_txt+"/"+"Fy"+case_name+r".png",dpi=300)
    plt.close()
    
    plt.clf()
    df.plot(x='Accumulated Timestep',y=[mf_in_txt, mf_out_txt])
    plt.title(title_txt)    
    plt.savefig(speedline_txt+"/"+"MF_main"+case_name+r".png",dpi=300)
    plt.close()
    
    plt.clf()
    df.plot(x='Accumulated Timestep',y=['MF SEAL EX [kg s^-1]', 'MF SEAL IN [kg s^-1]'])
    plt.title(title_txt)    
    plt.savefig(speedline_txt+"/"+"MF_seal"+case_name+r".png",dpi=300)
    plt.close()
    
    plt.clf()
    df.plot(x='Accumulated Timestep',y=[ps_out_txt,pt_out_txt])
    plt.title(title_txt)
    plt.savefig(speedline_txt+"/"+"Pr"+case_name+r".png",dpi=300)
    plt.close()
    
    plt.clf()
    df.plot(x='Accumulated Timestep',y=[tt_out_txt,ts_out_txt])
    plt.title(title_txt)
    plt.savefig(speedline_txt+"/"+"T_"+case_name+r".png",dpi=300)
    plt.close()
    
    
    df['Torque1 [J]'] = df['Torque1 [J]'].astype(float)
    df['Torque2 [J]'] = df['Torque1 [J]'].astype(float)
    df['Torque3 [J]'] = df['Torque1 [J]'].astype(float)
    
    plt.clf()
    df.plot(x='Accumulated Timestep',y=["Torque1 [J]"])
    plt.title(title_txt)
    plt.savefig(speedline_txt+"/"+"Torque1_"+case_name+r".png",dpi=300)    
    plt.close()
    
    plt.clf()
    df.plot(x='Accumulated Timestep',y=["Torque2 [J]","Torque3 [J]"])
    plt.title(title_txt)
    plt.savefig(speedline_txt+"/"+"Torque23_"+case_name+r".png",dpi=300)    
    plt.close()
plt.close()

# %%
df_mean = df_mean.T
df_std = df_std.T

# %%

nCol = len(df_mean.columns)
fig=plt.figure()
for iCol in range(nCol):
    plt.clf()
    x_txt = mf_out_txt;
    y_txt = df_mean.columns[iCol]
    y_txt_file =  y_txt.replace(' ','_')
    y_txt_file = re.sub('[!,*)@#%(&$?.^-]','', y_txt_file)
    y_txt_file =  y_txt_file.replace('[',''); y_txt_file =  y_txt_file.replace(']','')
    x = df_mean[x_txt]
    y = df_mean[y_txt]
    plt.plot(x,y,label=y_txt,marker='o')
    plt.ylabel(y_txt)
    plt.xlabel(x_txt)
    plt.grid()
    plt.title("Mean")
    plt.tight_layout()
    plt.savefig("mean_std/"+speedline_txt+"/"+"mean_"+y_txt_file+r".png",dpi=300)
# %% standard deviation
nCol = len(df_mean.columns)
fig=plt.figure()
for iCol in range(nCol):
    plt.clf()
    x_txt = mf_out_txt;
    y_txt = df_mean.columns[iCol]
    y_txt_file =  y_txt.replace(' ','_')
    y_txt_file = re.sub('[!,*)@#%(&$?.^-]','', y_txt_file)
    y_txt_file =  y_txt_file.replace('[',''); y_txt_file =  y_txt_file.replace(']','')
    x = df_mean[x_txt]
    y = df_std[y_txt]
    plt.plot(x,y,label=y_txt,marker='o')
    plt.ylabel(y_txt)
    plt.xlabel(x_txt)
    plt.grid()
    plt.title("STD")
    plt.tight_layout()
    plt.savefig("mean_std/"+speedline_txt+"/"+"std_"+y_txt_file+r".png",dpi=300)
    
# %% Three Fx and Fy one same plot both MEAN

df_mean.plot(x=mf_out_txt,y=['Fx1 [N]', 'Fx2 [N]', 'Fx3 [N]'],marker='o')
plt.ylabel("Forces")
plt.grid()
plt.title("Mean")
plt.tight_layout()
plt.savefig("mean_std/"+speedline_txt+"/"+"Fx"+"_avg"+r".png",dpi=300)

df_mean.plot(x=mf_out_txt,y=['Fy1 [N]', 'Fy2 [N]', 'Fy3 [N]'],marker='o')
plt.ylabel("Forces")
plt.grid()
plt.title("Mean")
plt.tight_layout()
plt.savefig("mean_std/"+speedline_txt+"/"+"Fy"+"_avg"+r".png",dpi=300)

# %% Three Fx and Fy one same plot both STD
df_std[mf_out_txt] = df_mean[mf_out_txt]
df_std.plot(x=mf_out_txt,y=['Fx1 [N]', 'Fx2 [N]', 'Fx3 [N]'],marker='o')
plt.ylabel("Forces [N]")
plt.grid()
plt.title("STD")
plt.tight_layout()
plt.savefig("mean_std/"+speedline_txt+"/"+"Fx"+"_std"+r".png",dpi=300)

df_std.plot(x=mf_out_txt,y=['Fy1 [N]', 'Fy2 [N]', 'Fy3 [N]'],marker='o')
plt.ylabel("Forces [N]")
plt.grid()
plt.title("STD")
plt.tight_layout()
plt.savefig("mean_std/"+speedline_txt+"/"+"Fy"+"_std"+r".png",dpi=300)

# %% Three Torque - one same plot MEAN
df_mean.plot(x=mf_out_txt,y=['Torque1 [J]', 'Torque2 [J]', 'Torque3 [J]'],marker='o')
plt.ylabel("Torque  [J]")
plt.grid()
plt.title("Mean")
plt.tight_layout()
plt.savefig("mean_std/"+speedline_txt+"/"+"Torque"+"_mean"+r".png",dpi=300)

# %% Three Torque - one same plot both STD
df_std[mf_out_txt] = df_mean[mf_out_txt]
df_std.plot(x=mf_out_txt,y=['Torque1 [J]', 'Torque2 [J]', 'Torque3 [J]'],marker='o')
plt.ylabel("Torque [J]")
plt.grid()
plt.title("STD")
plt.tight_layout()
plt.savefig("mean_std/"+speedline_txt+"/"+"Torque"+"_std"+r".png",dpi=300)


