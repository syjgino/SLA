# -*- coding: utf-8 -*-
"""
This file is part of the Shotgun Lipidomics Assistant (SLA) project.

Copyright 2020 Baolong Su (UCLA), Kevin Williams (UCLA), Lisa F. Bettcher (UW).

Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

    http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.
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
# import glob
# import re
# import statistics
import datetime


def imp_exp(exploc):
    exploc.configure(state="normal")
    exploc.delete(1.0, END)
    exp11 = filedialog.askopenfilename(filetypes=(("excel Files", "*.xlsx"), ("all files", "*.*")))
    exploc.insert(INSERT, exp11)
    exploc.configure(state="disabled")


def tagplot(exploc, CheckVar1, CheckVar2):
    matplotlib.use('Agg')
    start = datetime.datetime.now()
    file = exploc.get('1.0', 'end-1c')
    # root.destroy()

    # set directory
    os.chdir(file[0:file.rfind('/')])
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
        
        # tag bulk
        tag_bulk = tag_all.copy()
        tag_bulk['bulk'] = tag_bulk.index.str.split('-').str[0]
        tag_bulk = tag_bulk.drop('carbon', axis=1).groupby(['bulk'], as_index=True).sum().T

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

        ####################
        # TG plot function #
        ####################
        def grouped_barplot(df, cat, subcat, val, err, scat):
            # matplotlib.use('Agg') #dont show plot
            u = df[cat].unique()
            x = np.arange(len(u)) * 3
            y = scat
            subx = df[subcat].unique()
            offsets = 2.9 * (np.arange(len(subx)) - np.arange(len(subx)).mean()) / (len(subx) + 1.)
            xlist = [0.07, -0.07, 0.14, -0.14, 0.21, -0.21, 0.28, -0.28, 0.35, -0.35, 0.42, -0.42]
            while len(xlist) < 90:
                xlist = xlist + xlist

            if len(subx) == 1:
                width = 1.5
            else:
                width = np.diff(offsets).mean()

            # set plot size
            plt.figure(num=None, figsize=(40, 23), dpi=220, facecolor='w', edgecolor='k')
            plt.grid(True)
            # plt.ioff()  #turn off interactive mod, dont show plot
            if CheckVar1.get() == 0:
                if len(subx) <= 12:
                    for i, gr in enumerate(subx):
                        dfg = df[df[subcat] == gr]
                        plt.bar(x + offsets[i], dfg[val].values, width=width,
                                label="{} {}".format(subcat, gr), yerr=dfg[err].values,
                                # color = plt.cm.gist_ncar(i/len(subx)), alpha = .7,
                                color=plt.cm.Paired(i % 12), alpha=.7,
                                edgecolor='white', linewidth=1, capsize=10 * width
                                )
                    plt.ylim(bottom=0)
                    plt.xlabel(cat, fontsize=24)
                    plt.ylabel(val, fontsize=24)
                    plt.title('TG', fontsize=26)
                    plt.xticks(x, u, fontsize=24)
                    plt.yticks(fontsize=24)
                    plt.legend(subx, prop={'size': 24})
                    # plt.savefig(file[file.rfind('/')+1:file.rfind('.')]+'_TAG_'+cat+'.png')
                    # plt.show()
                elif len(subx) > 12:
                    for i, gr in enumerate(subx):
                        dfg = df[df[subcat] == gr]
                        plt.bar(x + offsets[i], dfg[val].values, width=width,
                                label="{} {}".format(subcat, gr), yerr=dfg[err].values,
                                color=plt.cm.gist_ncar(i / len(subx)), alpha=.7,
                                # color = plt.cm.Paired(i%12), alpha = .7,
                                edgecolor='white', linewidth=1, capsize=10 * width
                                )
                    plt.ylim(bottom=0)
                    plt.xlabel(cat, fontsize=24)
                    plt.ylabel(val, fontsize=24)
                    plt.title('TG', fontsize=26)
                    plt.xticks(x, u, fontsize=24)
                    plt.yticks(fontsize=24)
                    plt.legend(subx, prop={'size': 24})
                    # plt.savefig(file[file.rfind('/')+1:file.rfind('.')]+'_TAG_'+cat+'.png')
                    # plt.show()

            elif CheckVar1.get() == 1:
                if len(subx) <= 12:
                    for i, gr in enumerate(subx):
                        dfg = df[df[subcat] == gr]
                        plt.bar(x + offsets[i], dfg[val].values, width=width,
                                label="{} {}".format(subcat, gr),
                                # color = plt.cm.gist_ncar(i/len(subx)), alpha = .7,
                                color=plt.cm.Paired(i % 12), alpha=.7,
                                edgecolor='white', linewidth=1
                                )
                    plt.ylim(bottom=0)
                    plt.xlabel(cat, fontsize=24)
                    plt.ylabel(val, fontsize=24)
                    plt.title('TG', fontsize=26)
                    plt.xticks(x, u, fontsize=24)
                    plt.yticks(fontsize=24)
                    plt.legend(subx, prop={'size': 24})
                    # plt.savefig(file[file.rfind('/')+1:file.rfind('.')]+'_TAG_'+cat+'.png')
                elif len(subx) > 12:
                    for i, gr in enumerate(subx):
                        dfg = df[df[subcat] == gr]
                        plt.bar(x + offsets[i], dfg[val].values, width=width,
                                label="{} {}".format(subcat, gr),
                                color=plt.cm.gist_ncar(i / len(subx)), alpha=.7,
                                # color=plt.cm.Paired(i%12), alpha = .7,
                                edgecolor='white', linewidth=1
                                )
                    plt.ylim(bottom=0)
                    plt.xlabel(cat, fontsize=24)
                    plt.ylabel(val, fontsize=24)
                    plt.title('TAG', fontsize=26)
                    plt.xticks(x, u, fontsize=24)
                    plt.yticks(fontsize=24)
                    plt.legend(subx, prop={'size': 24})
                    # plt.savefig(file[file.rfind('/')+1:file.rfind('.')]+'_TAG_'+cat+'.png')

            # ADD dots
            if CheckVar2.get() == 1:
                for j in range(len(x)):
                    for i, gr in enumerate(subx):
                        yy = y[y.iloc[:, 1] == gr].iloc[:, j + 3]
                        yy = yy.sort_values()
                        gap = np.where(yy.diff() / (yy.max() - yy.min()) > 0.01)[0]
                        xx = xlist[0:len(yy)]
                        for ii in gap:
                            xx[ii:] = xlist[0:(len(yy) - ii)]
                        plt.scatter(x[j] + offsets[i] + pd.Series(xx) * width,
                                    yy, color='black', s=7)

            plt.savefig(file[file.rfind('/') + 1:file.rfind('.')] + '_TG_' + cat + '.pdf')

        # creat plot
        cat = "Carbon"
        subcat = "GroupName"
        val = "Avg"
        err = "Error"
        df = tag_avg_plot3
        scat = tag_summary
        grouped_barplot(df, cat, subcat, val, err, scat)

        ###################################
        # Create Index for TAG bond number#
        ###################################
        tag_all = tag_all.drop(['carbon'], axis=1)
        p = r':(.*)-'
        tag_all['bond'] = tag_all.index.str.extract(p, expand=False)
        tag_all.bond = tag_all.bond.astype(float)

        # Create a Series of all unique TAG bond number
        # bond_num = tag_all.bond.values
        # ubond_num = np.unique(bond_num)

        # Create DataFrame for each bond number
        tag_bond = dict()
        for i in range(int(min(tag_all['bond'])), int(max(tag_all['bond'])) + 1):
            tag_bond[i] = tag_all[tag_all['bond'] == i].copy()

        # Create bond Summary Dataframe
        tag_bond_sum = pd.DataFrame(columns=tag_all.columns)
        for i in range(int(min(tag_all['bond'])), int(max(tag_all['bond'])) + 1):
            tag_bond_sum.loc[str(i)] = pd.Series(tag_bond[i].sum(min_count=1))
        tag_bond_sum = tag_bond_sum.drop(['bond'], axis=1)

        tag_bond_summary = tag_bond_sum.T

        tag_bond_summary.insert(loc=0, column='GroupNum', value=np.array(species_all.loc['GroupNum']).astype(float))
        tag_bond_summary.insert(loc=0, column='GroupName', value=np.array(species_all.loc['GroupName']))
        tag_bond_summary.insert(loc=0, column='SampleID', value=np.array(species_all.loc['SampleID']))
        ##devided by 3
        tag_bond_summary.iloc[:, 3:] = tag_bond_summary.iloc[:, 3:] / 3

        ##summary bond avg
        tag_bond_summary_drop = tag_bond_summary.drop(['SampleID'], axis=1)

        tag_bsum_avg = tag_sum_avgB.drop(['GroupNum'], axis=1)
        tag_bsum_avg = tag_bsum_avg.T

        tag_bsum_avg['bond'] = tag_bsum_avg.index.str.extract(p, expand=False)
        tag_bsum_avg.bond = tag_bsum_avg.bond.astype(int)
        tag_bond_sum_avg = tag_bsum_avg.groupby(['bond'], as_index=True).sum()
        # tag_bond_sum_avg.index = tag_bond_sum_avg.index.astype(str)
        tag_bond_sum_avg = tag_bond_sum_avg.T
        tag_bond_sum_avg.insert(loc=0, column='GroupNum', value=np.array(tag_sum_avgB['GroupNum']).astype(float))
        # tag_bond_sum_avg['GroupNum'] =  tag_sum_avgB['GroupNum']
        tag_bond_di = tag_bond_sum_avg['GroupNum'].to_dict()

        # summary bond sd
        tag_bond_sum_sd = tag_bond_summary_drop.groupby(['GroupName'], as_index=True).std()
        tag_bond_sum_sd['GroupNum'] = tag_bond_sum_sd.index
        tag_bond_sum_sd['GroupNum'] = tag_bond_sum_sd['GroupNum'].replace(tag_bond_di)
        tag_bond_sum_sd = tag_bond_sum_sd.sort_values(by=['GroupNum'])
        # tag_bond_sum_sd['ExpNum'] = tag_bond_sum_avg['ExpNum']

        # Sum Species
        for i in range(int(min(tag_all['bond'])), int(max(tag_all['bond'])) + 1):
            tag_bond[i].loc[str(i) + 'total'] = pd.Series(tag_bond[i].sum(min_count=1))

        # bond number in sd to str
        tag_bond_sum_avg.columns = tag_bond_sum_avg.columns.astype(str)

        # Write to Excel
        writer = pd.ExcelWriter(file[file.rfind('/') + 1:file.rfind('.')] + '_' + 'TG_bond.xlsx')
        tag_bond_summary.to_excel(writer, 'Summary')
        tag_bond_sum_avg.to_excel(writer, 'SumAvg')
        tag_bond_sum_sd.to_excel(writer, 'SumAvg', startrow=tag_bond_sum_sd.shape[0] + 5, startcol=0)
        for i in range(int(min(tag_all['bond'])), int(max(tag_all['bond'])) + 1):
            tag_bond[i].to_excel(writer, 'TG' + str(i))

        writer.save()
        print('bond data saved')

        # plot avg and sd
        # melt data
        tag_bond_plot = tag_bond_sum_avg.copy()
        tag_bond_plot['GroupName'] = tag_bond_plot.index
        tag_bond_plot = tag_bond_plot.drop(['GroupNum'], axis=1)
        # tag_bond_plot.dropna(axis=1, how='all', inplace=True)
        tag_bond_plot2 = pd.melt(tag_bond_plot, id_vars=["GroupName"], var_name="Bond", value_name="Avg")

        tag_bond_sdplot = tag_bond_sum_sd.copy()
        tag_bond_sdplot['GroupName'] = tag_bond_sdplot.index
        tag_bond_sdplot = tag_bond_sdplot.drop(['GroupNum'], axis=1)
        # tag_bond_sdplot.dropna(axis=1, how='all', inplace=True)
        tag_bond_sdplot2 = pd.melt(tag_bond_sdplot, id_vars=["GroupName"], var_name="Bond", value_name="SD")

        tag_bond_plot3 = pd.merge(tag_bond_plot2, tag_bond_sdplot2, on=['GroupName', 'Bond'])
        tag_bond_plot3['Error'] = tag_bond_plot3['SD'] * 1
        tag_bond_plot3 = tag_bond_plot3.fillna(0)
        tag_bond_plot3['Bond'] = tag_bond_plot3['Bond'].astype(int)
        tag_bond_plot3 = tag_bond_plot3.sort_values(by=['Bond'])

        # plot
        cat = "Bond"
        subcat = "GroupName"
        val = "Avg"
        err = "Error"
        df = tag_bond_plot3
        scat = tag_bond_summary
        grouped_barplot(df, cat, subcat, val, err, scat)
        print('TAG plot saved')

    ######################
    # classtotal barplot #
    ######################
    sns.set_style("ticks")
    expdata = pd.read_excel(file,
                            sheet_name='Class Norm', header=[0], index_col=[0, 4], na_values='.')
    expdata = expdata.sort_values(by=['GroupNum'])
    expdata = expdata.drop(['ExpNum', 'GroupNum', 'SampleNorm', 'NormType'], axis=1)

    expmelt = pd.melt(expdata, id_vars='GroupName', var_name='Class', value_name='ClassTotal')
    g = sns.catplot(data=expmelt, x="GroupName", y="ClassTotal",
                    col="Class", col_wrap=4,
                    sharey=False,  # color='#BDE9F4',
                    facecolor=(1, 1, 1, 0), edgecolor=".2",
                    kind="bar", errwidth=1, capsize=.2,
                    height=4, aspect=1.2, ci='sd')
    g.set_titles(size=30, col_template='{col_name}')
    g.set_xlabels(fontsize=14)
    for axes in g.axes.flat:
        axes.set_xticklabels(axes.get_xticklabels(), rotation=90,
                             fontsize=18,
                             horizontalalignment='center',
                             verticalalignment='top', )
    plt.savefig(file[file.rfind('/') + 1:file.rfind('.')] + '_ClassTotalBar.pdf', bbox_inches='tight')
    plt.close('all')
    print('ClassTotal barplot saved')

    # classtotal boxplot
    sns.set_style("ticks")
    PROPS = {'boxprops': {'facecolor': 'none', 'edgecolor': 'black'},
             'medianprops': {'color': 'black'},
             'whiskerprops': {'color': 'black'},
             'capprops': {'color': 'black'}}
    expdata = pd.read_excel(file,
                            sheet_name='Class Norm', header=[0], index_col=[0, 4], na_values='.')
    expdata = expdata.sort_values(by=['GroupNum'])
    expdata = expdata.drop(['ExpNum', 'GroupNum', 'SampleNorm', 'NormType'], axis=1)

    expmelt = pd.melt(expdata, id_vars='GroupName', var_name='Class', value_name='ClassTotal')
    g = sns.catplot(data=expmelt, x="GroupName", y="ClassTotal",
                    col="Class", col_wrap=4,
                    sharey=False, color='#FFFFFF',
                    linewidth=None,
                    kind="box",
                    height=4, aspect=1.2, fliersize=5,
                    **PROPS
                    )
    g.set_titles(size=30, col_template='{col_name}')
    g.set_xlabels(fontsize=14)
    for axes in g.axes.flat:
        axes.set_xticklabels(axes.get_xticklabels(), rotation=90,
                             fontsize=18,
                             horizontalalignment='center',
                             verticalalignment='top', )
    plt.savefig(file[file.rfind('/') + 1:file.rfind('.')] + '_ClassTotalBox.pdf', bbox_inches='tight')
    plt.close('all')
    print('ClassTotal boxplot saved')

    print("run in %s" % (datetime.datetime.now() - start))
    messagebox.showinfo("Information", "Done")


"""
@author: BaolongSu

Tab5
TAG/Plot function

09/13/2019:
classtotal plot font size changed;
classtotal plot title removed "class="

03/10/2020
classtotal plot ci='sd'

11/23/2020
all plot sort by groupname

02/12/2021
take both TAG or TG as name

03/10/2021
add boxplot of class total
"""
