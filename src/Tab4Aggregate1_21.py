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
# import matplotlib.pyplot as plt
# import matplotlib
# matplotlib.use("Agg")
from tkinter import filedialog
# import glob
import re
# import statistics
import datetime


# from matplotlib.pyplot import cm
# import seaborn as sns


def imp_map(maploc, w_mapsheet, choices_mapsheet, variable_mapsheet):
    # map1 = filedialog.askopenfilename(filetypes=(("excel Files", "*.xlsx"),("all files", "*.*")))
    maploc.configure(state="normal")
    maploc.delete(1.0, END)
    map1 = filedialog.askopenfilename(filetypes=(("excel Files", "*.xlsx"), ("all files", "*.*")))
    maploc.insert(INSERT, map1)
    maploc.configure(state="disabled")
    #clear old options
    variable_mapsheet.set('')
    w_mapsheet['menu'].delete(0, "end")
    # Insert list of new options
    choices_mapsheet = pd.ExcelFile(maploc.get('1.0', 'end-1c')).sheet_names
    for choice in choices_mapsheet:
        w_mapsheet['menu'].add_command(label=choice, 
                                       command=lambda value=choice: variable_mapsheet.set(value))


def imp_method1(method1loc):
    # file1 = filedialog.askopenfilename(filetypes=(("excel Files", "*.xlsx"),("all files", "*.*")))
    method1loc.configure(state="normal")
    method1loc.delete(1.0, END)
    file1 = filedialog.askopenfilename(filetypes=(("excel Files", "*.xlsx"), ("all files", "*.*")))
    method1loc.insert(INSERT, file1)
    method1loc.configure(state="disabled")


def imp_method2(method2loc):
    # file2 = filedialog.askopenfilename(filetypes=(("excel Files", "*.xlsx"),("all files", "*.*")))
    method2loc.configure(state="normal")
    method2loc.delete(1.0, END)
    file2 = filedialog.askopenfilename(filetypes=(("excel Files", "*.xlsx"), ("all files", "*.*")))
    method2loc.insert(INSERT, file2)
    method2loc.configure(state="disabled")


def set_dir(dirloc_aggregate):
    # setdir = filedialog.askdirectory()
    dirloc_aggregate.configure(state="normal")
    dirloc_aggregate.delete(1.0, END)
    setdir = filedialog.askdirectory()
    dirloc_aggregate.insert(INSERT, setdir)
    dirloc_aggregate.configure(state="disabled")


def MergeApp(dirloc_aggregate, proname, method1loc, method2loc, 
             maploc, variable_mapsheet, CheckClustVis):
    start = datetime.datetime.now()
    os.chdir(dirloc_aggregate.get('1.0', 'end-1c'))
    project = proname.get()
    file1 = method1loc.get('1.0', 'end-1c')
    file2 = method2loc.get('1.0', 'end-1c')
    map1 = maploc.get('1.0', 'end-1c')

    # Fix sample index(name) type for proper merge
    def qcname(indname):
        # if 'QC_SPIKE' in str(indname):
        #    return('QC_SPIKE')
        # elif 'QC' in str(indname):
        #    return('QC')
        # elif 'b' in str(indname) or 'm' in str(indname):
        if re.search('[a-zA-Z]', str(indname)):
            return (str(indname))
        else:
            return (int(indname))

    # Import DataFrames from file1
    spequant1 = pd.read_excel(file1, sheet_name='Lipid Species Concentrations', header=0, index_col=0, na_values='.')
    specomp1 = pd.read_excel(file1, sheet_name='Lipid Species Composition', header=0, index_col=0, na_values='.')
    claquant1 = pd.read_excel(file1, sheet_name='Lipid Class Concentration', header=0, index_col=0, na_values='.')
    faquant1 = pd.read_excel(file1, sheet_name='Fatty Acid Concentration', header=0, index_col=0, na_values='.')
    facomp1 = pd.read_excel(file1, sheet_name='Fatty Acid Composition', header=0, index_col=0, na_values='.')

    spequant1.index = list(map(qcname, list(spequant1.index)))
    specomp1.index = list(map(qcname, list(specomp1.index)))
    claquant1.index = list(map(qcname, list(claquant1.index)))
    faquant1.index = list(map(qcname, list(faquant1.index)))
    facomp1.index = list(map(qcname, list(facomp1.index)))

    # Import DataFrames from file2
    if file2 != '':
        spequant2 = pd.read_excel(file2, sheet_name='Lipid Species Concentrations', header=0, index_col=0,
                                  na_values='.')
        specomp2 = pd.read_excel(file2, sheet_name='Lipid Species Composition', header=0, index_col=0, na_values='.')
        claquant2 = pd.read_excel(file2, sheet_name='Lipid Class Concentration', header=0, index_col=0, na_values='.')
        faquant2 = pd.read_excel(file2, sheet_name='Fatty Acid Concentration', header=0, index_col=0, na_values='.')
        facomp2 = pd.read_excel(file2, sheet_name='Fatty Acid Composition', header=0, index_col=0, na_values='.')

        spequant2.index = list(map(qcname, list(spequant2.index)))
        specomp2.index = list(map(qcname, list(specomp2.index)))
        claquant2.index = list(map(qcname, list(claquant2.index)))
        faquant2.index = list(map(qcname, list(faquant2.index)))
        facomp2.index = list(map(qcname, list(facomp2.index)))
    else:
        spequant2 = pd.DataFrame()
        specomp2 = pd.DataFrame()
        claquant2 = pd.DataFrame()
        faquant2 = pd.DataFrame()
        facomp2 = pd.DataFrame()

    # Merge DataFrames
    spequant = pd.concat([spequant1, spequant2], axis=1, sort=False)
    specomp = pd.concat([specomp1, specomp2], axis=1, sort=False)
    claquant = pd.concat([claquant1, claquant2], axis=1, sort=False)
    faquant = pd.concat([faquant1, faquant2], axis=1, sort=False)
    facomp = pd.concat([facomp1, facomp2], axis=1, sort=False)

    # Sort Columns in Merged DataFrames
    spequant = spequant.reindex(sorted(spequant.columns), axis=1)
    specomp = specomp.reindex(sorted(specomp.columns), axis=1)
    claquant = claquant.reindex(sorted(claquant.columns), axis=1)
    clacomp = claquant.apply(lambda x: 100 * x / x.sum(), axis=1)  # get class composit
    faquant = faquant.reindex(sorted(faquant.columns), axis=1)
    facomp = facomp.reindex(sorted(facomp.columns), axis=1)

    # Write Master data sheet
    # master = pd.ExcelWriter(project+'_master.xlsx')
    # spequant.to_excel(master, 'Species Quant')
    # specomp.to_excel(master, 'Species Composit')
    # claquant.to_excel(master, 'Class Quant')
    # clacomp.to_excel(master, 'Class Composit')
    # faquant.to_excel(master, 'FattyAcid Quant')
    # facomp.to_excel(master, 'FattyAcid Composit')
    # master.save()
    # print('master sheet saved')

    # Import Map
    sampinfo = pd.read_excel(map1, sheet_name=variable_mapsheet.get(), header=1, index_col=0, na_values='.')
    # Exp name dict
    expname = dict(zip(sampinfo.ExpNum, sampinfo.ExpName))
    sampinfo = sampinfo.drop(['ExpName'], axis=1)
    sampinfo.index = list(map(qcname, list(sampinfo.index)))
    sampinfo['SampleNorm'] = sampinfo['SampleNorm'].astype('float32')

    # Create Normalized Sheets
    # spenorm = spequant[list(map(lambda x: isinstance(x, int), spequant.index))].copy()
    # #exclude sample with string name

    # clanorm = claquant[list(map(lambda x: isinstance(x, int), claquant.index))].copy()
    # fanorm = faquant[list(map(lambda x: isinstance(x, int), faquant.index))].copy()
    spenorm = spequant.copy()  # inlude all samples
    clanorm = claquant.copy()
    fanorm = faquant.copy()
    spenorm = spenorm.divide(
        40)  # x0.025 to reverse /0.025 in the standard coef. /0.025 is there to simulate LWM result
    clanorm = clanorm.divide(40)
    fanorm = fanorm.divide(40)
    spenorm = spenorm.divide(sampinfo['SampleNorm'], axis='index')
    clanorm = clanorm.divide(sampinfo['SampleNorm'], axis='index')
    fanorm = fanorm.divide(sampinfo['SampleNorm'], axis='index')

    # Fix GroupName. If GroupName and GroupNum doesn't match, change GroupName to match GroupNum
    for i in sampinfo['ExpNum'].unique().astype(int):
        for ii in sampinfo.loc[sampinfo['ExpNum'] == i, 'GroupNum'].unique().astype(int):
            gNamlogic = np.logical_and(sampinfo['GroupNum'] == ii, sampinfo['ExpNum'] == i)
            sampinfo.loc[gNamlogic, "GroupName"] = sampinfo['GroupName'][gNamlogic].reset_index(drop=True)[0]

    # for i in range(1, int(max(sampinfo['ExpNum'])) + 1):
    #     for ii in range(1, int(max(sampinfo.loc[sampinfo['ExpNum'] == i, 'GroupNum'])) + 1):
    #         gNamlogic = np.logical_and(sampinfo['GroupNum'] == ii, sampinfo['ExpNum'] == i)
    #         sampinfo.loc[gNamlogic, "GroupName"] = sampinfo['GroupName'][gNamlogic].reset_index(drop=True)[0]

    # Merge Map, using index (Sample column in map and sample name in raw data)
    spequantin = pd.concat([sampinfo, spequant], axis=1, sort=False, join='inner')
    specompin = pd.concat([sampinfo, specomp], axis=1, sort=False, join='inner')
    claquantin = pd.concat([sampinfo, claquant], axis=1, sort=False, join='inner')
    clacompin = pd.concat([sampinfo, clacomp], axis=1, sort=False, join='inner')
    faquantin = pd.concat([sampinfo, faquant], axis=1, sort=False, join='inner')
    facompin = pd.concat([sampinfo, facomp], axis=1, sort=False, join='inner')
    spenormin = pd.concat([sampinfo, spenorm], axis=1, sort=False, join='inner')
    clanormin = pd.concat([sampinfo, clanorm], axis=1, sort=False, join='inner')
    fanormin = pd.concat([sampinfo, fanorm], axis=1, sort=False, join='inner')

    # Parse Experiments
    ##creat dictionary
    spequantin_exp = dict()
    specompin_exp = dict()
    claquantin_exp = dict()
    clacompin_exp = dict()
    faquantin_exp = dict()
    facompin_exp = dict()
    spenormin_exp = dict()
    clanormin_exp = dict()
    fanormin_exp = dict()

    ##loop
    for i in sampinfo['ExpNum'].unique().astype(int):  # range(1, int(max(spequantin['ExpNum'])) + 1):
        spequantin_exp[i] = spequantin[spequantin['ExpNum'] == i].copy()
        specompin_exp[i] = specompin[specompin['ExpNum'] == i].copy()
        # claquantin_exp[i] = claquantin[claquantin['ExpNum'] == i].copy()
        clacompin_exp[i] = clacompin[clacompin['ExpNum'] == i].copy()
        # faquantin_exp[i] = faquantin[faquantin['ExpNum'] == i].copy()
        facompin_exp[i] = facompin[facompin['ExpNum'] == i].copy()
        spenormin_exp[i] = spenormin[spenormin['ExpNum'] == i].copy()
        clanormin_exp[i] = clanormin[clanormin['ExpNum'] == i].copy()
        fanormin_exp[i] = fanormin[fanormin['ExpNum'] == i].copy()

        # Remove na columns from Experiment
        spequantin_exp[i].dropna(axis=1, how='all', inplace=True)
        specompin_exp[i].dropna(axis=1, how='all', inplace=True)
        # claquantin_exp[i].dropna(axis=1, how='all', inplace=True)
        clacompin_exp[i].dropna(axis=1, how='all', inplace=True)
        # faquantin_exp[i].dropna(axis=1, how='all', inplace=True)
        facompin_exp[i].dropna(axis=1, how='all', inplace=True)
        spenormin_exp[i].dropna(axis=1, how='all', inplace=True)
        clanormin_exp[i].dropna(axis=1, how='all', inplace=True)
        fanormin_exp[i].dropna(axis=1, how='all', inplace=True)

        # Ave and SD
        spenormAve = spenormin_exp[i].groupby(['GroupName'], as_index=True).mean()
        spenormAve = spenormAve.sort_values(by=['GroupNum'])
        spenormAve = spenormAve.drop(['SampleNorm'], axis=1)
        spdict = spenormAve['GroupNum'].to_dict()
        spenormSD = spenormin_exp[i].groupby(['GroupName'], as_index=True).std()
        spenormSD['GroupNum'] = spenormSD.index
        spenormSD['GroupNum'] = spenormSD['GroupNum'].replace(spdict)
        spenormSD = spenormSD.sort_values(by=['GroupNum'])
        spenormSD['ExpNum'] = spenormAve['ExpNum']
        spenormSD = spenormSD.drop(['SampleNorm'], axis=1)
        spenormSD.index.names = ['SD']

        classAve = clanormin_exp[i].groupby(['GroupName'], as_index=True).mean()
        classAve = classAve.sort_values(by=['GroupNum'])
        classAve = classAve.drop(['SampleNorm'], axis=1)
        cladict = classAve['GroupNum'].to_dict()
        classSD = clanormin_exp[i].groupby(['GroupName'], as_index=True).std()
        classSD['GroupNum'] = classSD.index
        classSD['GroupNum'] = classSD['GroupNum'].replace(cladict)
        classSD = classSD.sort_values(by=['GroupNum'])
        classSD['ExpNum'] = classAve['ExpNum']
        classSD = classSD.drop(['SampleNorm'], axis=1)
        classSD.index.names = ['SD']

        # Write Experiment
        exp1 = pd.ExcelWriter(project + '_Exp' + str(i) + '_' + str(expname[i]) + '.xlsx')
        # spequantin_exp[i].to_excel(exp1, 'Species Quant')
        spenormin_exp[i].to_excel(exp1, 'Species Norm')
        fanormin_exp[i].to_excel(exp1, 'FattyAcid Norm')
        clanormin_exp[i].to_excel(exp1, 'Class Norm')
        clacompin_exp[i].to_excel(exp1, 'Class Composit')
        specompin_exp[i].to_excel(exp1, 'Species Composit')
        # claquantin_exp[i].to_excel(exp1, 'Class Quant')
        # faquantin_exp[i].to_excel(exp1, 'FattyAcid Quant')
        facompin_exp[i].to_excel(exp1, 'FattyAcid Composit')
        spenormAve.to_excel(exp1, sheet_name='SpeciesAvg', startrow=0, startcol=0)
        spenormSD.to_excel(exp1, sheet_name='SpeciesAvg', startrow=spenormAve.shape[0] + 5, startcol=0)
        classAve.to_excel(exp1, sheet_name='ClassAvg', startrow=0, startcol=0)
        classSD.to_excel(exp1, sheet_name='ClassAvg', startrow=classAve.shape[0] + 5, startcol=0)
        exp1.save()

        # statcked plot
        # classAve.iloc[:,2:].plot(kind='barh', stacked=True, figsize=(16,8))
        # workbook=exp1.book
        # worksheet=workbook.add_worksheet('SpeciesAve')
        # exp1.sheets['SpeciesAve'] = worksheet

        # Format species normalized data for PCA
        specnormin_exp1PCA = spenormin_exp[i].T
        specnormin_exp1PCA = specnormin_exp1PCA.drop(['ExpNum', 'GroupNum', 'SampleNorm', 'NormType', 'SampleID'],
                                                     axis=0)

        # Write PCA data to CSV file
        if CheckClustVis.get() == 1:
            specnormin_exp1PCA.to_csv(project + '_Exp' + str(i) + '_' + str(expname[i]) + '_PCA.csv')

        print('exp', i)
    print("run in %s" % (datetime.datetime.now() - start))
    messagebox.showinfo("Information", "Done")


"""
@author: baolongsu

Tab4
Merge _m1 and _m2
(if you have only 1 result, always load to m1 and leave m2 blank)

Use Map file(submission form) to get sample info and normalize data
(if a sample in m1/m2 file is not included in the Map, it will simply exclude that sample from analysis)

20191001
changed tab name from Aggregate to Merge

creating of LWM like merged "master" excel file moved to tab 3, muted here

20200716
Added options to export .csv file for clustVis

20210115
using numbers in 01,02...100 for original sample name in Analyst is prefered
letters can be used, but may raise up errors

20210514
force SampleNorm from map to be float32

20210621
Merge tab: add dropdown list to choose map sheet
"""