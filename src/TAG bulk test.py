# -*- coding: utf-8 -*-
"""
Created on Wed Feb 14 09:14:17 2024

@author: BaolongSu
"""

import numpy as np
import pandas as pd
# from pyopenms import *
import os
# import tkinter as tk
# from tkinter import ttk
from tkinter import messagebox
# from tkinter.messagebox import showinfo
from tkinter import *
from tkinter import filedialog
import matplotlib.pyplot as plt
import matplotlib
import matplotlib.backends.backend_pdf
# from matplotlib.pyplot import cm
import seaborn as sns
# set directory
#os.chdir(file[0:file.rfind('/')])

#%%
file = "C:/Users/kevinwilliams/Documents/LipidyzerData/20240205_6500/Exp1/20240205_Exp1_DPDOX_10D_Cells.xlsx"
# Open Excel Document
lipid_all = pd.read_excel(file, sheet_name='Species Norm', index_col=0, na_values='.')
lipid_all = lipid_all.drop(['SampleNorm', 'NormType'], axis=1)
# Convert to floats
# lipid_all=lipid_all.replace(".",np.nan)
# lipid_all.astype(float)

# Transpose Document
species_all = lipid_all.T

# Create TAG DataFrame, get name TAG or TG
tag_all = species_all[species_all.index.str.startswith(('TAG', 'TG'))].copy()

# if no TAG, skip tag analysis
if len(tag_all) != 0:
    tag_all = tag_all.astype(float)
    #####################################
    # Create Index for TAG carbon length#
    #####################################
    # tag_all.columns = tag_all.loc['GroupName']
    tag_all['carbon'] = tag_all.index.str.slice(start=3, stop=5)
    tag_all.carbon = tag_all.carbon.astype(float)

    # Create a Series of all unique TAG carbon length
    # car_len = tag_all.carbon.values
    # ucar_len = np.unique(car_len)

    # Create DataFrame for each Carbon Length
    tag_36 = tag_all[(tag_all['carbon'] == 36)].copy()
    tag_37 = tag_all[(tag_all['carbon'] == 37)].copy()
    tag_38 = tag_all[(tag_all['carbon'] == 38)].copy()
    tag_39 = tag_all[(tag_all['carbon'] == 39)].copy()
    tag_40 = tag_all[(tag_all['carbon'] == 40)].copy()
    tag_41 = tag_all[(tag_all['carbon'] == 41)].copy()
    tag_42 = tag_all[(tag_all['carbon'] == 42)].copy()
    tag_43 = tag_all[(tag_all['carbon'] == 43)].copy()
    tag_44 = tag_all[(tag_all['carbon'] == 44)].copy()
    tag_45 = tag_all[(tag_all['carbon'] == 45)].copy()
    tag_46 = tag_all[(tag_all['carbon'] == 46)].copy()
    tag_47 = tag_all[(tag_all['carbon'] == 47)].copy()
    tag_48 = tag_all[(tag_all['carbon'] == 48)].copy()
    tag_49 = tag_all[(tag_all['carbon'] == 49)].copy()
    tag_50 = tag_all[(tag_all['carbon'] == 50)].copy()
    tag_51 = tag_all[(tag_all['carbon'] == 51)].copy()
    tag_52 = tag_all[(tag_all['carbon'] == 52)].copy()
    tag_53 = tag_all[(tag_all['carbon'] == 53)].copy()
    tag_54 = tag_all[(tag_all['carbon'] == 54)].copy()
    tag_55 = tag_all[(tag_all['carbon'] == 55)].copy()
    tag_56 = tag_all[(tag_all['carbon'] == 56)].copy()
    tag_57 = tag_all[(tag_all['carbon'] == 57)].copy()
    tag_58 = tag_all[(tag_all['carbon'] == 58)].copy()
    tag_59 = tag_all[(tag_all['carbon'] == 59)].copy()
    tag_60 = tag_all[(tag_all['carbon'] == 60)].copy()

    # Create Summary Dataframe
    tag_sum = pd.DataFrame(columns=tag_all.columns)
    tag_sum.loc['36'] = pd.Series(tag_36.sum(min_count=1))
    tag_sum.loc['37'] = pd.Series(tag_37.sum(min_count=1))
    tag_sum.loc['38'] = pd.Series(tag_38.sum(min_count=1))
    tag_sum.loc['39'] = pd.Series(tag_39.sum(min_count=1))
    tag_sum.loc['40'] = pd.Series(tag_40.sum(min_count=1))
    tag_sum.loc['41'] = pd.Series(tag_41.sum(min_count=1))
    tag_sum.loc['42'] = pd.Series(tag_42.sum(min_count=1))
    tag_sum.loc['43'] = pd.Series(tag_43.sum(min_count=1))
    tag_sum.loc['44'] = pd.Series(tag_44.sum(min_count=1))
    tag_sum.loc['45'] = pd.Series(tag_45.sum(min_count=1))
    tag_sum.loc['46'] = pd.Series(tag_46.sum(min_count=1))
    tag_sum.loc['47'] = pd.Series(tag_47.sum(min_count=1))
    tag_sum.loc['48'] = pd.Series(tag_48.sum(min_count=1))
    tag_sum.loc['49'] = pd.Series(tag_49.sum(min_count=1))
    tag_sum.loc['50'] = pd.Series(tag_50.sum(min_count=1))
    tag_sum.loc['51'] = pd.Series(tag_51.sum(min_count=1))
    tag_sum.loc['52'] = pd.Series(tag_52.sum(min_count=1))
    tag_sum.loc['53'] = pd.Series(tag_53.sum(min_count=1))
    tag_sum.loc['54'] = pd.Series(tag_54.sum(min_count=1))
    tag_sum.loc['55'] = pd.Series(tag_55.sum(min_count=1))
    tag_sum.loc['56'] = pd.Series(tag_56.sum(min_count=1))
    tag_sum.loc['57'] = pd.Series(tag_57.sum(min_count=1))
    tag_sum.loc['58'] = pd.Series(tag_58.sum(min_count=1))
    tag_sum.loc['59'] = pd.Series(tag_59.sum(min_count=1))
    tag_sum.loc['60'] = pd.Series(tag_60.sum(min_count=1))
    tag_sum = tag_sum.drop(['carbon'], axis=1)
    tag_summary = tag_sum.T
    tag_summary.insert(loc=0, column='GroupNum', value=np.array(species_all.loc['GroupNum']).astype(float))
    tag_summary.insert(loc=0, column='GroupName', value=np.array(species_all.loc['GroupName']))
    tag_summary.insert(loc=0, column='SampleID', value=np.array(species_all.loc['SampleID']))
    ##devide by 3
    tag_summary.iloc[:, 3:] = tag_summary.iloc[:, 3:] / 3

    # Sum Species
    tag_36.loc['36total'] = pd.Series(tag_36.sum(min_count=1))  # minimum 1 number, if all NA, return NA
    tag_37.loc['37total'] = pd.Series(tag_37.sum(min_count=1))
    tag_38.loc['38total'] = pd.Series(tag_38.sum(min_count=1))
    tag_39.loc['39total'] = pd.Series(tag_39.sum(min_count=1))
    tag_40.loc['40total'] = pd.Series(tag_40.sum(min_count=1))
    tag_41.loc['41total'] = pd.Series(tag_41.sum(min_count=1))
    tag_42.loc['42total'] = pd.Series(tag_42.sum(min_count=1))
    tag_43.loc['43total'] = pd.Series(tag_43.sum(min_count=1))
    tag_44.loc['44total'] = pd.Series(tag_44.sum(min_count=1))
    tag_45.loc['45total'] = pd.Series(tag_45.sum(min_count=1))
    tag_46.loc['46total'] = pd.Series(tag_46.sum(min_count=1))
    tag_47.loc['47total'] = pd.Series(tag_47.sum(min_count=1))
    tag_48.loc['48total'] = pd.Series(tag_48.sum(min_count=1))
    tag_49.loc['49total'] = pd.Series(tag_49.sum(min_count=1))
    tag_50.loc['50total'] = pd.Series(tag_50.sum(min_count=1))
    tag_51.loc['51total'] = pd.Series(tag_51.sum(min_count=1))
    tag_52.loc['52total'] = pd.Series(tag_52.sum(min_count=1))
    tag_53.loc['53total'] = pd.Series(tag_53.sum(min_count=1))
    tag_54.loc['54total'] = pd.Series(tag_54.sum(min_count=1))
    tag_55.loc['55total'] = pd.Series(tag_55.sum(min_count=1))
    tag_56.loc['56total'] = pd.Series(tag_56.sum(min_count=1))
    tag_57.loc['57total'] = pd.Series(tag_57.sum(min_count=1))
    tag_58.loc['58total'] = pd.Series(tag_58.sum(min_count=1))
    tag_59.loc['59total'] = pd.Series(tag_59.sum(min_count=1))
    tag_60.loc['60total'] = pd.Series(tag_60.sum(min_count=1))

    ##summary avg
    tag_all2 = lipid_all[lipid_all.columns[pd.Series(lipid_all.columns).str.startswith(('TAG', 'TG'))]].copy()
    tag_all2 = tag_all2.astype(float)
    tag_all2 = pd.concat([tag_all2, lipid_all.iloc[:, 1:4]], axis=1)
    tag_all2_drop = tag_all2.drop(['SampleID'], axis=1)
    tag_sum_avg = tag_all2_drop.groupby(['GroupName'], as_index=True).mean()
    # tag_summary_drop = tag_summary.drop(['SampleID'], axis=1)
    # tag_sum_avg = tag_summary_drop.groupby(['GroupName'], as_index=True).mean()
    tag_sum_avg = tag_sum_avg.sort_values(by=['GroupNum'])

    ##devide by 3
    tag_sum_avg.iloc[:, 0:(len(tag_sum_avg.columns) - 1)] =\
        tag_sum_avg.iloc[:, 0:(len(tag_sum_avg.columns) - 1)] / 3

    tag_sum_avgB = tag_sum_avg.copy()
    tag_sum_avg = tag_sum_avg.T
    tag_sum_avg = tag_sum_avg.drop(['GroupNum'], axis=0)
    tag_sum_avg['carbon'] = tag_sum_avg.index.str.slice(start=3, stop=5)
    # tag_sum_avg.carbon = tag_sum_avg.carbon.astype(float)
    tag_sum_avg = tag_sum_avg.groupby(['carbon'], as_index=True).sum().T
    tag_sum_avg.insert(loc=0, column='GroupNum', value=np.array(tag_sum_avgB['GroupNum']).astype(float))
    tag_sum_avg.columns = tag_sum_avg.columns
    tag_sum_avgCC = tag_sum_avg.copy()
    # tag_sum_avg['GroupNum'] =  tag_sum_avgB['GroupNum']

    tag_sum_avg = pd.DataFrame(index=tag_sum_avgCC.index, columns=np.arange(36, 61).astype(str))
    tag_sum_avg.insert(loc=0, column='GroupNum', value=tag_sum_avgCC['GroupNum'])
    for i in tag_sum_avgCC.columns:
        tag_sum_avg[i] = tag_sum_avgCC[i]

    tag_avg_di = tag_sum_avg['GroupNum'].to_dict()

    tag_summary_drop = tag_summary.drop(['SampleID'], axis=1)
    tag_sum_sd = tag_summary_drop.groupby(['GroupName'], as_index=True).std()
    tag_sum_sd['GroupNum'] = tag_sum_sd.index
    tag_sum_sd['GroupNum'] = tag_sum_sd['GroupNum'].replace(tag_avg_di)
    tag_sum_sd = tag_sum_sd.sort_values(by=['GroupNum'])
    # tag_sum_sd['ExpNum'] = tag_sum_avg['ExpNum']

#%% tag bulk
tag_all['bulk'] = tag_all.index.str.split('-').str[0]
tag_bulk = tag_all.drop('carbon', axis=1).groupby('bulk', as_index=True).sum().T



#%%
    # Write to Excel
    writer_tgbulk = pd.ExcelWriter(file[file.rfind('/') + 1:file.rfind('.')] + '_' + 'TG_bulk.xlsx')
    tag_bulk.to_excel(writer_tgbulk, 'TG bulk')
    writer_tgbulk.save()
    writer_tgbulk.close()
    
    writer = pd.ExcelWriter(file[file.rfind('/') + 1:file.rfind('.')] + '_' + 'TG_carbon.xlsx')
    tag_summary.to_excel(writer, 'Summary')
    tag_sum_avg.to_excel(writer, sheet_name='SumAvg')
    tag_sum_sd.to_excel(writer, sheet_name='SumAvg', startrow=tag_sum_sd.shape[0] + 5, startcol=0)
    tag_36.to_excel(writer, 'TAG36')
    tag_37.to_excel(writer, 'TAG37')
    tag_38.to_excel(writer, 'TAG38')
    tag_39.to_excel(writer, 'TAG39')
    tag_40.to_excel(writer, 'TAG40')
    tag_41.to_excel(writer, 'TAG41')
    tag_42.to_excel(writer, 'TAG42')
    tag_43.to_excel(writer, 'TAG43')
    tag_44.to_excel(writer, 'TAG44')
    tag_45.to_excel(writer, 'TAG45')
    tag_46.to_excel(writer, 'TAG46')
    tag_47.to_excel(writer, 'TAG47')
    tag_48.to_excel(writer, 'TAG48')
    tag_49.to_excel(writer, 'TAG49')
    tag_50.to_excel(writer, 'TAG50')
    tag_51.to_excel(writer, 'TAG51')
    tag_52.to_excel(writer, 'TAG52')
    tag_53.to_excel(writer, 'TAG53')
    tag_54.to_excel(writer, 'TAG54')
    tag_55.to_excel(writer, 'TAG55')
    tag_56.to_excel(writer, 'TAG56')
    tag_57.to_excel(writer, 'TAG57')
    tag_58.to_excel(writer, 'TAG58')
    tag_59.to_excel(writer, 'TAG59')
    tag_60.to_excel(writer, 'TAG60')
    
    writer.save()
    writer.close()
    print('carbon data saved')

    #################
    ##plot avg & sd##
    #################
    ### set plot size
    # plt.figure(num=None, figsize=(36, 12), dpi=80, facecolor='w', edgecolor='k')

    # melt data
    tag_avg_plot = tag_sum_avg.copy()
    tag_avg_plot['GroupName'] = tag_avg_plot.index
    tag_avg_plot = tag_avg_plot.drop(['GroupNum'], axis=1)
    # tag_avg_plot.dropna(axis=1, how='all', inplace=True)
    tag_avg_plot2 = pd.melt(tag_avg_plot, id_vars=["GroupName"], var_name="Carbon", value_name="Avg")

    tag_avg_sdplot = tag_sum_sd.copy()
    tag_avg_sdplot['GroupName'] = tag_avg_sdplot.index
    tag_avg_sdplot = tag_avg_sdplot.drop(['GroupNum'], axis=1)
    # tag_avg_sdplot.dropna(axis=1, how='all', inplace=True)
    tag_avg_sdplot2 = pd.melt(tag_avg_sdplot, id_vars=["GroupName"], var_name="Carbon", value_name="SD")

    tag_avg_plot3 = pd.merge(tag_avg_plot2, tag_avg_sdplot2, on=['GroupName', 'Carbon'])
    tag_avg_plot3['Error'] = tag_avg_plot3['SD'] * 1
    tag_avg_plot3 = tag_avg_plot3.fillna(0)
    tag_avg_plot3['Carbon'] = tag_avg_plot3['Carbon'].astype(int)
    tag_avg_plot3 = tag_avg_plot3.sort_values(by=['Carbon'])