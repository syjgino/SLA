# -*- coding: utf-8 -*-
"""
Created on Tue Feb  4 14:52:36 2020

@author: BaolongSu

LApack Tab3 Read mzml functions

isotope correction with MAP

Drop species when it's Std intensity<100 or has 3 or more 0s; or it's own intensity has 3 or more 0s

20200706
Merge function moved from previous Merge tab to here

20200828
test on processing only m2
"""
import numpy as np
import pandas as pd
from pyopenms import *
import os
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
from tkinter.messagebox import showinfo
from tkinter import *
from tkinter import filedialog
import glob
import re
import datetime


# set up directory
def set_dir_read(dirloc_read):
    dirloc_read.configure(state="normal")
    dirloc_read.delete(1.0, END)
    setdir = filedialog.askdirectory()
    dirloc_read.insert(INSERT, setdir)
    dirloc_read.configure(state="disabled")


# get spname dictionary location
def get_sp_dict1(sp_dict1_loc):
    sp_dict1_loc.configure(state="normal")
    sp_dict1_loc.delete(1.0, END)
    file1 = filedialog.askopenfilename(filetypes=(("excel Files", "*.xlsx"),
                                                  ("all files", "*.*")))
    sp_dict1_loc.insert(INSERT, file1)
    sp_dict1_loc.configure(state="disabled")


# get isotope correction map location
def get_iso_dict(iso_dict_loc):
    # file1 = filedialog.askopenfilename(filetypes=(("excel Files", "*.xlsx"),("all files", "*.*")))
    iso_dict_loc.configure(state="normal")
    iso_dict_loc.delete(1.0, END)
    file1 = filedialog.askopenfilename(filetypes=(("excel Files", "*.xlsx"),
                                                  ("all files", "*.*")))
    iso_dict_loc.insert(INSERT, file1)
    iso_dict_loc.configure(state="disabled")


# get standard dictionary location
def get_std_dict(std_dict_loc):
    # file1 = filedialog.askopenfilename(filetypes=(("excel Files", "*.xlsx"),("all files", "*.*")))
    std_dict_loc.configure(state="normal")
    std_dict_loc.delete(1.0, END)
    file1 = filedialog.askopenfilename(filetypes=(("excel Files", "*.xlsx"),
                                                  ("all files", "*.*")))
    std_dict_loc.insert(INSERT, file1)
    std_dict_loc.configure(state="disabled")


# readMZML function
def readMZML(dirloc_read, sp_dict1_loc, std_dict_loc, proname3,
             iso_dict_loc, variable_iso, progressBar, variable_dataversion,
             variable_mute):
    progressBar = progressBar
    os.chdir(dirloc_read.get('1.0', 'end-1c'))
    list_of_files = glob.glob('./*.mzML')

    ##average of all intensity scans after dropping 0s
    def centered_average(row):
        mesures = row  # [6:26]
        mesures = mesures[mesures != 0]  # drop all 0s
        if len(mesures) == 0:
            mesures = [0]
        mesures = np.nanmean(mesures)
        return (mesures)

    ##map standard
    def stdNorm(row):
        stdrow = sp_df2['Species'] == std_dict[method]['StdName'][row['Species']]
        return (avgint_list[np.where(stdrow)[0][0]])
        # return(sp_df2['AvgIntensity'][stdrow].item())

    ##map standard coeficient    
    def conCoef(row):
        return (std_dict[method]['Coef'][row['Species']])

    # def spName(row): #not working currently
    #    return(sp_dict[method].loc[(pd.Series(sp_dict[method]['Q1'] == row['Q1']) & pd.Series(sp_dict[method]['Q3'] == row['Q3']))].index[0])

    ##isotope correction, v1, nolonger in use
    def isoadj(row):
        rnum = sp_df2['Species'] == row['Name']
        if row['CT1'] == 'XXX':
            nomin = avgint_list[np.where(rnum)[0][0]]
            B1 = nomin / row['P0']
            avgint_list[np.where(rnum)[0][0]] = B1
        elif row['CT2'] == 'XXX':
            rnum2 = sp_df2['Species'] == row['CT1']
            A1 = avgint_list[np.where(rnum)[0][0]]
            A2 = avgint_list[np.where(rnum2)[0][0]] * row['Pct1']
            nomin = A1 - A2
            B1 = nomin / row['P0']
            avgint_list[np.where(rnum)[0][0]] = B1
        elif row['CT2'] != 'XXX':
            rnum2 = sp_df2['Species'] == row['CT1']
            rnum3 = sp_df2['Species'] == row['CT2']
            A1 = avgint_list[np.where(rnum)[0][0]]
            A2 = avgint_list[np.where(rnum2)[0][0]] * row['Pct1']
            A3 = avgint_list[np.where(rnum3)[0][0]] * row['Pct2']
            nomin = A1 - A2 - A3
            B1 = nomin / row['P0']
            avgint_list[np.where(rnum)[0][0]] = B1

    ##isotope correction V2
    def isoadjv2(row):
        rnum = sp_df2['Species'] == row['Name']
        rnum1 = sp_df2['Species'] == row['CT1']
        rnum2 = sp_df2['Species'] == row['CT2']
        rnum3 = sp_df2['Species'] == row['CT3']
        A = avgint_list[np.where(rnum)[0][0]]
        A1 = avgint_list[np.where(rnum1)[0][0]] * row['Pct1']
        A2 = avgint_list[np.where(rnum2)[0][0]] * row['Pct2']
        A3 = avgint_list[np.where(rnum3)[0][0]] * row['Pct3']
        nomin = A - A1 - A2 - A3
        B1 = max(nomin / row['P0'], 0)
        avgint_list[np.where(rnum)[0][0]] = B1

    ##create variable
    all_df_dict = {'1': {}, '2': {}}
    # out_df2 = pd.DataFrame()
    out_df2 = {'1': pd.DataFrame(), '2': pd.DataFrame()}
    # out_df2_con = pd.DataFrame()
    out_df2_con = {'1': pd.DataFrame(), '2': pd.DataFrame()}
    out_df2_intensity = {'1': pd.DataFrame(), '2': pd.DataFrame()}
    # sp_df3 = {'A':0}

    ##read spname dict
    sp_dict = {}
    sp_dict['1'] = pd.read_excel(sp_dict1_loc.get('1.0', 'end-1c'), sheet_name='1', header=0, index_col=-1,
                                 na_values='.')
    sp_dict['2'] = pd.read_excel(sp_dict1_loc.get('1.0', 'end-1c'), sheet_name='2', header=0, index_col=-1,
                                 na_values='.')

    # read standard dict
    std_dict = {}
    std_dict['1'] = pd.read_excel(std_dict_loc.get('1.0', 'end-1c'), sheet_name='Method1', header=0, index_col=0,
                                  na_values='.').to_dict()
    std_dict['2'] = pd.read_excel(std_dict_loc.get('1.0', 'end-1c'), sheet_name='Method2', header=0, index_col=0,
                                  na_values='.').to_dict()

    # read iso correction dict
    if variable_iso.get() == 'Yes':
        iso_dict = {}
        iso_dict['1'] = pd.read_excel(iso_dict_loc.get('1.0', 'end-1c'),
                                      sheet_name='M1', header=0, index_col=None, na_values='.')
        iso_dict['2'] = pd.read_excel(iso_dict_loc.get('1.0', 'end-1c'),
                                      sheet_name='M2', header=0, index_col=None, na_values='.')

    ##reading mzml loop start
    for sample in range(0, len(list_of_files)):
        ##get method number
        method = re.search('- (.)-', list_of_files[sample]).group(1)
        start = datetime.datetime.now()
        print('method' + method, str(sample))
        # read all files
        exp = MSExperiment()
        MzMLFile().load(list_of_files[sample][2:], exp)
        all_chroms = exp.getChromatograms()

        ##create dataframe with species info
        sp_df = pd.DataFrame({'NativeID': []})
        sp_serie = list()
        for each_chrom in all_chroms:
            NatID = each_chrom.getNativeID()
            if type(NatID) is bytes:
                NatID = NatID.decode("utf-8")
            sp_serie.append(NatID)

        # sp_df['NativeID'] = pd.Series(sp_serie[1:])
        # drop rows may be tic/bic
        ticrows = pd.Series([x for x in list(map(lambda x: len(x), sp_serie))]) > 20
        sp_df['NativeID'] = np.array(sp_serie)[np.flatnonzero(ticrows)]

        sp_df2 = sp_df['NativeID'].apply(lambda x: pd.Series(x.split()))

        if any(sp_df2[0] == '-'):
            a = sp_df2.iloc[0:sum(sp_df2[0] == '-'), 1:9]
            a.columns = range(0, 8)
            sp_df2.iloc[0:sum(sp_df2[0] == '-'), 0:8] = a
            sp_df2 = sp_df2.drop(columns=[0, 1, 8])
        else:
            sp_df2 = sp_df2.drop(columns=[0, 1])

        sp_df2.columns = sp_df2.loc[0].str.extract(r'(.*)=', expand=False)
        sp_df2 = sp_df2.apply(lambda x: x.str.extract(r'=(.*)$', expand=False))
        sp_df2 = sp_df2.astype(float)

        ##get intensity
        intensity_df = list()
        i = 0
        for each_chrom in all_chroms:
            intensity_serie = list()
            for each_in in each_chrom:
                intensity_serie.extend([each_in.getIntensity()])
            intensity_df.append(np.array(intensity_serie))
            i += 1

        intensity_df = pd.DataFrame(intensity_df)
        intensity_df = intensity_df.T
        #####get 20 scans & drop tic/bic in data#####
        intensity_df2 = intensity_df.loc[0:19, ticrows]
        intensity_df2 = intensity_df2.T
        intensity_df2 = intensity_df2.reset_index()
        intensity_df2 = intensity_df2.drop(columns=['index'])

        ##merge spname and intensity
        sp_df2 = sp_df2.merge(intensity_df2, left_index=True, right_index=True)

        ##change experiment1 to negative, add species name
        if max(sp_df2['experiment']) == 2:
            sp_df2['Q1'][sp_df2['experiment'] == 1] = -sp_df2['Q1'][sp_df2['experiment'] == 1]
        sp_df2['Species'] = sp_dict[method].index
        # sp_df2['Species'] = sp_df2.apply(spName, axis=1)

        ##compute avg intesity
        sp_df2['AvgIntensity'] = np.nan
        sp_df2['AvgIntensity'] = sp_df2.iloc[:, 6:26].apply(centered_average, axis=1)  # 20 scans

        ##Isotope correction on intensity with MAP
        if variable_iso.get() == 'Yes':
            if method in ['1', '2']:
                avgint_list = sp_df2.AvgIntensity.values.tolist()
                iso_dict[method].apply(isoadjv2, axis=1)
                sp_df2['AvgIntensity'] = avgint_list

        # compute ratio
        avgint_list = sp_df2.AvgIntensity.values.tolist()
        std_int = sp_df2.apply(stdNorm, axis=1)
        std_int[std_int == 0] = np.nan  # std=0 set to NA
        sp_df2['Ratio'] = sp_df2['AvgIntensity'] / std_int

        # compute concentration (coef = mg/ml / MW * 1E5)
        sp_df2['Concentration'] = sp_df2['Ratio'] * sp_df2.apply(conCoef, axis=1)

        # save standards(or all) intensity, muted
        # sp_df3[sample] = sp_df2[sp_df2['Species'].str[0] == 'd']
        # save all intensity
        # sp_df3[sample] = sp_df2

        # out put intensity, muted
        out_df_intensity = sp_df2[['Species', 'AvgIntensity']].T
        out_df_intensity.columns = out_df_intensity.iloc[0]
        out_df_intensity = out_df_intensity[1:]
        # name of sample(2nd one for analyst, 1st one for lipidyzer)
        # out_df_intensity = out_df_intensity.rename(index={'AvgIntensity': re.search('-[0-9][0-9]* - (.+)[.]', list_of_files[sample]).group(1)} ).astype(float)
        out_df_intensity = out_df_intensity.rename(
            index={'AvgIntensity': re.search('- .-(.+)[.]', list_of_files[sample]).group(1)}).astype(float)
        all_df_dict[method][sample] = sp_df2
        out_df2_intensity[method] = pd.concat([out_df2_intensity[method], out_df_intensity], axis=0, sort=False)

        #####drop species if standard <100 and/or >=3 0s#########
        std_df = sp_df2[sp_df2['Species'].str[0] == 'd']
        dropbothloc = np.logical_or(std_df['AvgIntensity'] < 100, np.sum(std_df.iloc[:, 6:26] == 0, axis=1) >= 3)
        dropstd = std_df.loc[dropbothloc, 'Species']
        #
        std_dict_df = pd.DataFrame.from_dict(std_dict[method]['StdName'], orient='index')
        std_dict_df['SpName'] = std_dict_df.index
        dropstd_indx = std_dict_df[0].apply(lambda x: x in list(dropstd))
        dropstd_sp = std_dict_df.loc[dropstd_indx, 'SpName']
        spdrop_indx = -sp_df2['Species'].apply(lambda x: x in list(dropstd_sp))
        sp_df2 = sp_df2.loc[spdrop_indx, :]

        # drop species if >=3 0s
        sp_df2 = sp_df2[np.sum(sp_df2.iloc[:, 6:26] == 0, axis=1) < 3]  # 20 scans

        # drop standard from output
        sp_df2 = sp_df2[sp_df2['Species'].str[0] != 'd']

        # drop extra 2nd tail scans, for V2
        sp_df2 = sp_df2[sp_df2['Species'].str[-2:] != '_2']

        ############################################
        # drop muted species on spname list 20200730#
        ############################################
        if variable_mute.get() == "Yes":
            mutelist = sp_dict[method][sp_dict[method]["Mute"] == TRUE]
            mutelist = mutelist.index
            sp_df2 = sp_df2[-sp_df2['Species'].str[:].isin(mutelist)]

            # out put ratio
        out_df = sp_df2[['Species', 'Ratio']].T
        out_df.columns = out_df.iloc[0]
        out_df = out_df[1:]

        # name of sample(one for analyst, one for lipidyzer)
        if variable_dataversion.get() == "Analyst":
            out_df = out_df.rename(index={'Ratio': re.search('- .-(.+)[.]', list_of_files[sample]).group(1)}).astype(
                float)  # analyst
        elif variable_dataversion.get() == "LWM":
            out_df = out_df.rename(
                index={'Ratio': re.search('- [0-9][0-9]* - (.+)[.]', list_of_files[sample]).group(1)}).astype(float)

        # all_df_dict[method][sample] = sp_df2
        out_df2[method] = pd.concat([out_df2[method], out_df], axis=0, sort=False)

        # out put concentration
        out_df_con = sp_df2[['Species', 'Concentration']].T
        out_df_con.columns = out_df_con.iloc[0]
        out_df_con = out_df_con[1:]

        # name of sample
        if variable_dataversion.get() == "LWM":
            out_df_con = out_df_con.rename(
                index={'Concentration': re.search('-[0-9][0-9]* - (.+)[.]', list_of_files[sample]).group(1)}).astype(
                float)
        elif variable_dataversion.get() == "Analyst":
            out_df_con = out_df_con.rename(
                index={'Concentration': re.search('- .-(.+)[.]', list_of_files[sample]).group(1)}).astype(float)

        out_df2_con[method] = pd.concat([out_df2_con[method], out_df_con], axis=0, sort=False)

        print("run in %s" % (datetime.datetime.now() - start))
        # progress_var.set((1+sample)/len(list_of_files))
        progressBar["value"] = (1 + sample) / len(list_of_files) * 100
        progressBar.update()
        # root.update_idletasks()

    # Fatty Accid
    #######################
    # Fatty Accid Concentration#
    #######################

    # CE,CER,DCER,FFA,HCER,LCER,LPC,LPE,SM,PA
    FA_con = {}
    if len(out_df2_con['1']) > 0:
        FA_con = {'1': out_df2_con['1'][out_df2['1'].columns[
            -pd.Series(out_df2_con['1'].columns).str.split('(', expand=True)[0].isin(['PE', 'PC', 'PG', 'PI', 'PS'])]]}
    if len(out_df2_con['2']) > 0:
        FA_con['2'] = out_df2_con['2'][
            out_df2['2'].columns[-pd.Series(out_df2_con['2'].columns).str.split('(', expand=True)[0].isin(['DAG'])]]
        FA_con['2'] = FA_con['2'][FA_con['2'].columns[-pd.Series(FA_con['2'].columns).str[0:3].isin(['TAG'])]]

    if len(out_df2_con['1']) > 0:
        # PC
        PC_con = out_df2_con['1'][
            out_df2_con['1'].columns[pd.Series(out_df2_con['1'].columns).str.split('(', expand=True)[0].isin(['PC'])]]
        PC_con_1 = PC_con[:]
        PC_con_1.columns = pd.Series(PC_con_1.columns).apply(lambda st: st[st.find("(") + 1:st.find("/")])
        PC_con_2 = PC_con[:]
        PC_con_2.columns = pd.Series(PC_con_2.columns).apply(lambda st: st[st.find("/") + 1:st.find(")")])
        PC_con_m = pd.concat([PC_con_1, PC_con_2, ], axis=1)
        PC_con_m = PC_con_m.groupby(PC_con_m.columns, axis=1).sum()
        PC_con_m = PC_con_m.replace(0, np.nan)
        PC_con_m.columns = "PC(" + PC_con_m.columns + ")"

        # PE
        PE_con = out_df2_con['1'][
            out_df2_con['1'].columns[pd.Series(out_df2_con['1'].columns).str.split('(', expand=True)[0].isin(['PE'])]]
        PE_con_1 = PE_con[:]
        PE_con_1.columns = pd.Series(PE_con_1.columns).apply(lambda st: st[st.find("(") + 1:st.find("/")])
        PE_con_1.columns = pd.Series(PE_con_1.columns).apply(lambda st: st[st.find("-") + 1:])  # remove O- P-
        PE_con_2 = PE_con[:]
        PE_con_2.columns = pd.Series(PE_con_2.columns).apply(lambda st: st[st.find("/") + 1:st.find(")")])
        PE_con_m = pd.concat([PE_con_1, PE_con_2, ], axis=1)
        PE_con_m = PE_con_m.groupby(PE_con_m.columns, axis=1).sum()
        PE_con_m = PE_con_m.replace(0, np.nan)
        PE_con_m.columns = "PE(" + PE_con_m.columns + ")"

        # PG
        PG_con = out_df2_con['1'][
            out_df2_con['1'].columns[pd.Series(out_df2_con['1'].columns).str.split('(', expand=True)[0].isin(['PG'])]]
        PG_con_1 = PG_con[:]
        PG_con_1.columns = pd.Series(PG_con_1.columns).apply(lambda st: st[st.find("(") + 1:st.find("/")])
        PG_con_2 = PG_con[:]
        PG_con_2.columns = pd.Series(PG_con_2.columns).apply(lambda st: st[st.find("/") + 1:st.find(")")])
        PG_con_m = pd.concat([PG_con_1, PG_con_2, ], axis=1)
        PG_con_m = PG_con_m.groupby(PG_con_m.columns, axis=1).sum()
        PG_con_m = PG_con_m.replace(0, np.nan)
        PG_con_m.columns = "PG(" + PG_con_m.columns + ")"

        # PI
        PI_con = out_df2_con['1'][
            out_df2_con['1'].columns[pd.Series(out_df2_con['1'].columns).str.split('(', expand=True)[0].isin(['PI'])]]
        PI_con_1 = PI_con[:]
        PI_con_1.columns = pd.Series(PI_con_1.columns).apply(lambda st: st[st.find("(") + 1:st.find("/")])
        PI_con_2 = PI_con[:]
        PI_con_2.columns = pd.Series(PI_con_2.columns).apply(lambda st: st[st.find("/") + 1:st.find(")")])
        PI_con_m = pd.concat([PI_con_1, PI_con_2, ], axis=1)
        PI_con_m = PI_con_m.groupby(PI_con_m.columns, axis=1).sum()
        PI_con_m = PI_con_m.replace(0, np.nan)
        PI_con_m.columns = "PI(" + PI_con_m.columns + ")"

        # PS
        PS_con = out_df2_con['1'][
            out_df2_con['1'].columns[pd.Series(out_df2_con['1'].columns).str.split('(', expand=True)[0].isin(['PS'])]]
        PS_con_1 = PS_con[:]
        PS_con_1.columns = pd.Series(PS_con_1.columns).apply(lambda st: st[st.find("(") + 1:st.find("/")])
        PS_con_2 = PS_con[:]
        PS_con_2.columns = pd.Series(PS_con_2.columns).apply(lambda st: st[st.find("/") + 1:st.find(")")])
        PS_con_m = pd.concat([PS_con_1, PS_con_2, ], axis=1)
        PS_con_m = PS_con_m.groupby(PS_con_m.columns, axis=1).sum()
        PS_con_m = PS_con_m.replace(0, np.nan)
        PS_con_m.columns = "PS(" + PS_con_m.columns + ")"

        # DAG
    if len(out_df2_con['2']) > 0:
        DAG_con = out_df2_con['2'][
            out_df2_con['2'].columns[pd.Series(out_df2_con['2'].columns).str.split('(', expand=True)[0].isin(['DAG'])]]
        DAG_con_1 = DAG_con[:]
        DAG_con_1.columns = pd.Series(DAG_con_1.columns).apply(lambda st: st[st.find("(") + 1:st.find("/")])
        DAG_con_2 = DAG_con[:]
        DAG_con_2.columns = pd.Series(DAG_con_2.columns).apply(lambda st: st[st.find("/") + 1:st.find(")")])

        DAG_con_m = pd.concat([DAG_con_1, DAG_con_2, ], axis=1)
        DAG_con_m = DAG_con_m.groupby(DAG_con_m.columns, axis=1).sum()
        DAG_con_m = DAG_con_m.replace(0, np.nan)
        DAG_con_m.columns = "DAG(" + DAG_con_m.columns + ")"

        # TAG
        TAG_con = out_df2_con['2'][out_df2_con['2'].columns[pd.Series(out_df2_con['2'].columns).str[0:3].isin(['TAG'])]]
        TAG_con_m = TAG_con[:]
        TAG_con_m.columns = pd.Series(TAG_con_m.columns).apply(lambda st: st[st.find("FA") + 2:])
        TAG_con_m = TAG_con_m.groupby(TAG_con_m.columns, axis=1).sum()
        TAG_con_m = TAG_con_m.replace(0, np.nan)
        TAG_con_m.columns = "TAG(" + TAG_con_m.columns + ")"

    ##Merge to FA
    if len(out_df2_con['1']) > 0:
        FA_con['1'] = pd.concat([FA_con['1'], PC_con_m, PE_con_m,
                                 PG_con_m, PI_con_m, PS_con_m],
                                axis=1)
    if len(out_df2_con['2']) > 0:
        FA_con['2'] = pd.concat([FA_con['2'], DAG_con_m, TAG_con_m], axis=1)

    # FA_con['1'].to_excel("test.xlsx")
    ######################
    # Fatty Accid Composition#
    ######################
    FA_grp = {}
    if len(out_df2_con['1']) > 0:
        FA_grp = {'1': FA_con['1'].groupby(FA_con['1'].columns.str[0:3], axis=1).sum()}
        FA_grp['1'] = FA_grp['1'].replace(np.inf, np.nan)
        FA_grp['1'] = FA_grp['1'].replace(0, np.nan)
    if len(out_df2_con['2']) > 0:
        FA_grp['2'] = FA_con['2'].groupby(FA_con['2'].columns.str[0:3], axis=1).sum()
        FA_grp['2'] = FA_grp['2'].replace(np.inf, np.nan)
        FA_grp['2'] = FA_grp['2'].replace(0, np.nan)
    FA_comp = {}
    if len(out_df2_con['1']) > 0:
        FA_comp = {'1': FA_con['1'].apply(
            lambda col: 100 * col / FA_grp['1'].iloc[:, FA_grp['1'].columns == col.name[0:3]].iloc[:, 0])}
    if len(out_df2_con['2']) > 0:
        FA_comp['2'] = FA_con['2'].apply(
            lambda col: 100 * col / FA_grp['2'].iloc[:, FA_grp['2'].columns == col.name[0:3]].iloc[:, 0])

    ###################
    # Species Composition#
    ###################
    SP_grp = {}
    if len(out_df2_con['1']) > 0:
        SP_grp = {'1': out_df2_con['1'].groupby(out_df2_con['1'].columns.str[0:3], axis=1).sum()}
        SP_grp['1'] = SP_grp['1'].replace(np.inf, np.nan)
        SP_grp['1'] = SP_grp['1'].replace(0, np.nan)
    if len(out_df2_con['2']) > 0:
        SP_grp['2'] = out_df2_con['2'].groupby(out_df2_con['2'].columns.str[0:3], axis=1).sum()
        SP_grp['2'] = SP_grp['2'].replace(np.inf, np.nan)
        SP_grp['2'] = SP_grp['2'].replace(0, np.nan)
    SP_comp = {}
    if len(out_df2_con['1']) > 0:
        SP_comp = {'1': out_df2_con['1'].apply(
            lambda col: 100 * col / SP_grp['1'].iloc[:, SP_grp['1'].columns == col.name[0:3]].iloc[:, 0])}
    if len(out_df2_con['2']) > 0:
        SP_comp['2'] = out_df2_con['2'].apply(
            lambda col: 100 * col / SP_grp['2'].iloc[:, SP_grp['2'].columns == col.name[0:3]].iloc[:, 0])

    ###########
    # Class Total#
    ###########
    # class consentration total is SP_grp
    if len(out_df2_con['1']) > 0:
        SP_grp['1'].columns = pd.Series(SP_grp['1'].columns).apply(lambda x: x.replace('(', ''))
    if len(out_df2_con['2']) > 0:
        SP_grp['2'] = SP_grp['2'].rename(columns={'DCE': 'DCER', 'HCE': 'HCER', 'LCE': 'LCER'})
        SP_grp['2']['TAG'] = SP_grp['2']['TAG'] / 3  # TAG devided by 3
        SP_grp['2'].columns = pd.Series(SP_grp['2'].columns).apply(lambda x: x.replace('(', ''))

    # class composition
    clas_comp = {}
    if len(out_df2_con['1']) > 0:
        clas_comp = {'1': SP_grp['1'].apply(lambda row: 100 * row / row.sum(), axis=1)}
    if len(out_df2_con['2']) > 0:
        clas_comp['2'] = SP_grp['2'].apply(lambda row: 100 * row / row.sum(), axis=1)

    ######
    # save#
    ######
    # m1
    if len(out_df2_con['1']) > 0:
        master = pd.ExcelWriter(proname3.get() + '_output_m1.xlsx')
        # out_df2['1'] = out_df2['1'].reindex(sorted(out_df2['1'].columns), axis = 1)
        # out_df2['1'].to_excel(master, 'Ratio')
        out_df2_con['1'] = out_df2_con['1'].reindex(sorted(out_df2_con['1'].columns), axis=1)
        out_df2_con['1'].insert(0, 'Name', out_df2_con['1'].index)
        out_df2_con['1'].to_excel(master, 'Lipid Species Concentrations', index=False)
        SP_comp['1'] = SP_comp['1'].reindex(sorted(SP_comp['1'].columns), axis=1)
        SP_comp['1'].insert(0, 'Name', SP_comp['1'].index)
        SP_comp['1'].to_excel(master, 'Lipid Species Composition', index=False)

        SP_grp['1'] = SP_grp['1'].reindex(sorted(SP_grp['1'].columns), axis=1)
        SP_grp['1'].insert(0, 'Name', SP_grp['1'].index)
        SP_grp['1'].to_excel(master, 'Lipid Class Concentration', index=False)
        clas_comp['1'] = clas_comp['1'].reindex(sorted(clas_comp['1'].columns), axis=1)
        clas_comp['1'].insert(0, 'Name', clas_comp['1'].index)
        clas_comp['1'].to_excel(master, 'Lipid Class Composition', index=False)

        FA_con['1'] = FA_con['1'].reindex(sorted(FA_con['1'].columns), axis=1)
        FA_con['1'].insert(0, 'Name', FA_con['1'].index)
        FA_con['1'].to_excel(master, 'Fatty Acid Concentration', index=False)
        FA_comp['1'] = FA_comp['1'].reindex(sorted(FA_comp['1'].columns), axis=1)
        FA_comp['1'].insert(0, 'Name', FA_comp['1'].index)
        FA_comp['1'].to_excel(master, 'Fatty Acid Composition', index=False)

        # add standard dictionary and mrm list(spname) info
        keyinfo_m1 = pd.DataFrame({'Info': [sp_dict1_loc.get('1.0', 'end-1c'),
                                            std_dict_loc.get('1.0', 'end-1c'),
                                            iso_dict_loc.get('1.0', 'end-1c'),
                                            variable_iso.get()]},
                                  index=['SpName/mrm List',
                                         'Standard Dictionary',
                                         'Isotope Correction List',
                                         'isotope Corrected'])
        keyinfo_m1.to_excel(master, 'version info', index=True)

        master.save()

    # m2
    if len(out_df2_con['2']) > 0:
        master2 = pd.ExcelWriter(proname3.get() + '_output_m2.xlsx')
        # out_df2['2'] = out_df2['2'].reindex(sorted(out_df2['2'].columns), axis = 1)
        # out_df2['2'].to_excel(master2, 'Ratio')
        out_df2_con['2'] = out_df2_con['2'].reindex(sorted(out_df2_con['2'].columns), axis=1)
        out_df2_con['2'].insert(0, 'Name', out_df2_con['2'].index)
        out_df2_con['2'].to_excel(master2, 'Lipid Species Concentrations', index=False)
        SP_comp['2'] = SP_comp['2'].reindex(sorted(SP_comp['2'].columns), axis=1)
        SP_comp['2'].insert(0, 'Name', SP_comp['2'].index)
        SP_comp['2'].to_excel(master2, 'Lipid Species Composition', index=False)

        SP_grp['2'] = SP_grp['2'].reindex(sorted(SP_grp['2'].columns), axis=1)
        SP_grp['2'].insert(0, 'Name', SP_grp['2'].index)
        SP_grp['2'].to_excel(master2, 'Lipid Class Concentration', index=False)
        clas_comp['2'] = clas_comp['2'].reindex(sorted(clas_comp['2'].columns), axis=1)
        clas_comp['2'].insert(0, 'Name', clas_comp['2'].index)
        clas_comp['2'].to_excel(master2, 'Lipid Class Composition', index=False)

        FA_con['2'] = FA_con['2'].reindex(sorted(FA_con['2'].columns), axis=1)
        FA_con['2'].insert(0, 'Name', FA_con['2'].index)
        FA_con['2'].to_excel(master2, 'Fatty Acid Concentration', index=False)
        FA_comp['2'] = FA_comp['2'].reindex(sorted(FA_comp['2'].columns), axis=1)
        FA_comp['2'].insert(0, 'Name', FA_comp['2'].index)
        FA_comp['2'].to_excel(master2, 'Fatty Acid Composition', index=False)

        # add standard dictionary and mrm list(spname) info
        keyinfo_m2 = pd.DataFrame({'Info': [sp_dict1_loc.get('1.0', 'end-1c'),
                                            std_dict_loc.get('1.0', 'end-1c'),
                                            iso_dict_loc.get('1.0', 'end-1c'),
                                            variable_iso.get()]},
                                  index=['SpName/mrm List',
                                         'Standard Dictionary',
                                         'Isotope Correction List',
                                         'isotope Corrected'])
        keyinfo_m2.to_excel(master2, 'version info', index=True)

        master2.save()

    # Merge m1 m2 DataFrames
    if (len(out_df2_con['1']) > 0) & (len(out_df2_con['2']) > 0):
        spequant = pd.concat([out_df2_con["1"].iloc[:, 1:], out_df2_con["2"].iloc[:, 1:]],
                             axis=1, sort=False)
        specomp = pd.concat([SP_comp['1'].iloc[:, 1:], SP_comp['2'].iloc[:, 1:]],
                            axis=1, sort=False)
        claquant = pd.concat([SP_grp['1'].iloc[:, 1:], SP_grp['2'].iloc[:, 1:]],
                             axis=1, sort=False)
        faquant = pd.concat([FA_con['1'].iloc[:, 1:], FA_con['2'].iloc[:, 1:]],
                            axis=1, sort=False)
        facomp = pd.concat([FA_comp['1'].iloc[:, 1:], FA_comp['2'].iloc[:, 1:]],
                           axis=1, sort=False)

        # Sort Columns in Merged DataFrames
        spequant = spequant.reindex(sorted(spequant.columns), axis=1)
        specomp = specomp.reindex(sorted(specomp.columns), axis=1)
        claquant = claquant.reindex(sorted(claquant.columns), axis=1)
        clacomp = claquant.apply(lambda x: 100 * x / sum(x), axis=1)  # get class composit
        faquant = faquant.reindex(sorted(faquant.columns), axis=1)
        facomp = facomp.reindex(sorted(facomp.columns), axis=1)

        # Write Master data sheet
        master3 = pd.ExcelWriter(proname3.get() + '_output_merge.xlsx')
        spequant.to_excel(master3, 'Species Quant')
        specomp.to_excel(master3, 'Species Composit')
        claquant.to_excel(master3, 'Class Quant')
        clacomp.to_excel(master3, 'Class Composit')
        faquant.to_excel(master3, 'FattyAcid Quant')
        facomp.to_excel(master3, 'FattyAcid Composit')
        master3.save()

    # save intensity
    masterint = pd.ExcelWriter(proname3.get() + '_intensity.xlsx')
    out_df2_intensity['1'] = out_df2_intensity['1'].reindex(sorted(out_df2_intensity['1'].columns), axis=1)
    out_df2_intensity['1'].to_excel(masterint, 'IntensityM1', index=True)
    if len(out_df2_con['2']) > 0:
        out_df2_intensity['2'] = out_df2_intensity['2'].reindex(sorted(out_df2_intensity['2'].columns), axis=1)
        out_df2_intensity['2'].to_excel(masterint, 'IntensityM2', index=True)
    masterint.save()

    messagebox.showinfo("Information", "Done")
