# -*- coding: utf-8 -*-
"""
Created on Wed May 29 11:41:50 2019

@author: BaolongSu

Tune for LA_V1

"""

import numpy as np
import pandas as pd
from pyopenms import *
import os
import glob
import re
import matplotlib
import matplotlib.pyplot as plt
#from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure
import seaborn as sns          
from scipy import stats
from tkinter import *
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
from tkinter.messagebox import showinfo
from tkinter import filedialog
import datetime

def get_tune1(tunef1): 
    tunef1.configure(state="normal")
    tunef1.delete(1.0, END)
    filedir = filedialog.askopenfilename(filetypes=(("mzML Files", "*.mzML"),("all files", "*.*")))
    tunef1.insert(INSERT, filedir)
    tunef1.configure(state="disabled")

def get_tune2(tunef2):
    #setdir = filedialog.askdirectory()  
    tunef2.configure(state="normal")
    tunef2.delete(1.0, END)
    filedir = filedialog.askopenfilename(filetypes=(("mzML Files", "*.mzML"),("all files", "*.*")))
    tunef2.insert(INSERT, filedir)
    tunef2.configure(state="disabled")
    
def imp_map(maploc_tune):
    #map1 = filedialog.askopenfilename(filetypes=(("excel Files", "*.xlsx"),("all files", "*.*")))
    maploc_tune.configure(state="normal")
    maploc_tune.delete(1.0 ,END)
    map1 = filedialog.askopenfilename(filetypes=(("excel Files", "*.xlsx"),("all files", "*.*")))
    maploc_tune.insert(INSERT, map1)
    maploc_tune.configure(state="disabled")
    
def exportdata(covlist, covlist_dict, maploc_tune):
    #global covlist_dict
    exp_temp_loc=maploc_tune.get('1.0', 'end-1c') #'Tuning_spname_dict'
    neg_temp = pd.read_excel(exp_temp_loc, sheet_name = 'NEG', header=0, index_col=None, na_values='.')
    pos_temp = pd.read_excel(exp_temp_loc, sheet_name = 'POS', header=0, index_col=None, na_values='.')
    sst_temp = pd.read_excel(exp_temp_loc, sheet_name = 'SST', header=0, index_col=None, na_values='.')
    
    covlist_neg = list(neg_temp['Group'].unique()) #+ list(pos_temp['Group'].unique())
    covlist_pos = list(pos_temp['Group'].unique())
    
    #fill in
    for i in range(0, len(covlist_neg)):
        #tvar = covlist_dict[i].get('1.0', 'end-1c')
        tvar = covlist_dict[i].get()
        neg_temp.loc[neg_temp['Group']==covlist_neg[i], 'volt'] = float(tvar)
        #messagebox.showinfo(tvar)
    #auto compute when any LPC/LPE entered 0
    try:
        if any(neg_temp.loc[neg_temp['Group']=='LPC16', 'volt']==0):
            neg_temp.loc[neg_temp['Group']=='LPC16', 'volt']=neg_temp.loc[neg_temp['Group']=='LPC18', 'volt'].iloc[0]-1
        elif any(neg_temp.loc[neg_temp['Group']=='LPC18', 'volt']==0):
            neg_temp.loc[neg_temp['Group']=='LPC18', 'volt']=neg_temp.loc[neg_temp['Group']=='LPC16', 'volt'].iloc[0]+1
    except:
        pass

    try:
        if any(neg_temp.loc[neg_temp['Group']=='LPE16', 'volt']==0):
            neg_temp.loc[neg_temp['Group']=='LPE16', 'volt']=neg_temp.loc[neg_temp['Group']=='LPE18', 'volt'].iloc[0]-1.5
        elif any(neg_temp.loc[neg_temp['Group']=='LPE18', 'volt']==0):
            neg_temp.loc[neg_temp['Group']=='LPE18', 'volt']=neg_temp.loc[neg_temp['Group']=='LPE16', 'volt'].iloc[0]+1.5
    except:
        pass    
    
    for i in range(0, len(covlist_pos)):
        #tvar = covlist_dict[i+len(covlist_neg)].get('1.0', 'end-1c')
        tvar = covlist_dict[i+len(covlist_neg)].get()
        pos_temp.loc[pos_temp['Group']==covlist_pos[i], 'volt'] = float(tvar)
    
    
    #SST result
    for i in sst_temp['Group'].unique():
        sst_temp.loc[sst_temp['Group']==i, 'volt'] = float(covlist_dict[covlist.index(i)].get())
    
    
    #save
    fname = 'TuneResult_' + datetime.datetime.now().strftime("%Y%m%d%H%M") + '.xlsx'
    master = pd.ExcelWriter(fname)
    neg_temp.to_excel(master,'NEG', index = True)
    pos_temp.to_excel(master,'POS', index = True)
    sst_temp.to_excel(master,'SST', index = True)
    master.save()
    os.system(fname)
    #messagebox.showinfo("Information","Done")

def fetchLPC(covlist_dict):
    covlist_dict[3].configure(state=NORMAL)
    covlist_dict[3].delete(0,END)
    try:
        covlist_dict[3].insert(0,float(covlist_dict[2].get())+1)
        covlist_dict[3].configure(state=DISABLED)
    except:
        covlist_dict[3].insert(0,float(0))
        covlist_dict[3].configure(state=DISABLED)
def fetchLPE(covlist_dict):
    covlist_dict[4].configure(state=NORMAL)
    covlist_dict[4].delete(0,END)
    try:
        covlist_dict[4].insert(0,float(covlist_dict[5].get())-1.5)
        covlist_dict[4].configure(state=DISABLED)
    except:
        covlist_dict[4].insert(0,float(0))
        covlist_dict[4].configure(state=DISABLED)    
    
def TuningFun(tunef1, tunef2, maploc_tune, out_text, variable_peaktype, tab1, covlist, covlist_label, covlist_dict):
    out_text.configure(state="normal")
    sp_dict1_loc = maploc_tune.get('1.0', 'end-1c')
    #std_dict_loc = 'C:/Users/baolongsu/Desktop/Projects/StdUnkRatio/standard_dict - LipidizerSimulate_102b_MW_dev2.xlsx'
    file = tunef1.get('1.0', 'end-1c')
    os.chdir(file[0:file.rfind('/')])
    ##get name all files
    list_of_files = [tunef1.get('1.0', 'end-1c'),
                     tunef2.get('1.0', 'end-1c')]
    
    def spName(row):
        return(sp_dict[method].loc[(pd.Series(sp_dict[method]['Q1'] == row['Q1']) & pd.Series(sp_dict[method]['Q3'] == row['Q3']))].index[0])
    
    ##create all variable
    all_df_dict = {'1':{}, '2':{}}
    #out_df2 = pd.DataFrame()
    out_df2 = {'1':pd.DataFrame(), '2': pd.DataFrame()}
    #out_df2_con = pd.DataFrame()
    out_df2_con = {'1':pd.DataFrame(), '2': pd.DataFrame()}
    out_df2_intensity = {'1':pd.DataFrame(), '2': pd.DataFrame()}
    sp_df3 = {'A':0}
    
    ##read dicts
    sp_dict = {}
    sp_dict['1'] = pd.read_excel(sp_dict1_loc, sheet_name = '1', header=0, index_col=-1, na_values='.')
    sp_dict['2'] = pd.read_excel(sp_dict1_loc, sheet_name = '2', header=0, index_col=-1, na_values='.')
    sp_dict['NEG'] = pd.read_excel(sp_dict1_loc, sheet_name = 'NEG', header=0, index_col=None, na_values='.')
    sp_dict['POS'] = pd.read_excel(sp_dict1_loc, sheet_name = 'POS', header=0, index_col=None, na_values='.')
    
        
    ##loop start
    #start0 = datetime.datetime.now()
    for sample in range(0, len(list_of_files)):
        ##get method number
        #method = re.search('- (.)-', list_of_files[sample]).group(1)
        method = str(sample + 1)
        
        #start = datetime.datetime.now()
        #print('method'+method, str(sample))
        #read all files
        exp=MSExperiment()
        MzMLFile().load(list_of_files[sample][:],exp) 
        #[sample][2:] if get listoffiles with glob.glob
        all_chroms=exp.getChromatograms()
        
        
        ##create dataframe with species info
        sp_df = pd.DataFrame({'NativeID':[]})
        sp_serie = list()
        for each_chrom in all_chroms:
            NatID = each_chrom.getNativeID()
            if type(NatID) is bytes:
                NatID = NatID.decode("utf-8")
            sp_serie.append(NatID)
            #sp_serie.append(each_chrom.getNativeID().decode("utf-8"))
        #######################drop rows may be tic/bic#################
        ticrows = pd.Series([x for x in list(map(lambda x: len(x), sp_serie))]) > 20
        sp_df['NativeID'] = np.array(sp_serie)[np.flatnonzero(ticrows)]
        #below is assuming all tic/bic are on top
        #ticrows = len([x for x in list(map(lambda x: len(x), sp_serie)) if x < 10])
        #sp_df['NativeID'] = pd.Series(sp_serie[ticrows:])
        
        
        sp_df2 = sp_df['NativeID'].apply(lambda x: pd.Series(x.split()))
    
        if any(sp_df2[0]=='-'):
            a = sp_df2.iloc[0:sum(sp_df2[0]=='-'),1:9]
            a.columns = range(0,8)
            sp_df2.iloc[0:sum(sp_df2[0]=='-'),0:8] = a
            sp_df2 = sp_df2.drop(columns=[0,1,8])
        else:
            sp_df2 = sp_df2.drop(columns=[0,1])
            
        sp_df2.columns = sp_df2.loc[0].str.extract(r'(.*)=', expand=False)
        sp_df2 = sp_df2.apply(lambda x: x.str.extract(r'=(.*)$', expand=False))
        sp_df2 = sp_df2.astype(float)
        
        ##get intensity
        #intensity_df = pd.DataFrame()
        intensity_df = list()
        i=0
        for each_chrom in all_chroms:
            intensity_serie = list()
            for each_in in each_chrom:
                intensity_serie.extend([each_in.getIntensity()])
            #intensity_df[i] = pd.Series(intensity_serie)
            intensity_df.append(np.array(intensity_serie))
            i+=1
    
        intensity_df = pd.DataFrame(intensity_df)
        intensity_df = intensity_df.T
        ####################drop tic/bic in data#############################
        intensity_df2 = intensity_df.loc[:, ticrows]
        #intensity_df2 = intensity_df.iloc[:, ticrows:len(intensity_df.loc[1])]
        intensity_df2 = intensity_df2.T
        intensity_df2 = intensity_df2.reset_index()
        intensity_df2 = intensity_df2.drop(columns = ['index'])
        
        ##merge sp and intensity
        sp_df2 = sp_df2.merge(intensity_df2, left_index = True, right_index=True)
        
        ##change experiment1 to negative, add species nam
        ##not used here
        if max(sp_df2['experiment']) == 2:
            sp_df2['Q1'][sp_df2['experiment'] == 1] = -sp_df2['Q1'][sp_df2['experiment'] == 1]
        #sp_df2['Species'] = sp_dict[method].index
        sp_df2['Species'] = ''
        sp_df2['Species'] = sp_df2.apply(spName, axis=1)
        sp_df3[sample] = sp_df2
    
    #get peaks
    #current version only takes 1neg and 1pos file.
    peaks_pos = []
    peaks_neg = []
    if variable_peaktype.get()=='Max1':
        out_text.insert(END, 'Get Max Value:\n')
        #print SM
        for i in range(0,1):
            avar = sp_df3[i].iloc[0,6:-1].astype(float)
            bvar = avar.idxmax()*0.101-5
            tvar = sp_df3[i].Species[0] + '  ' + str(round(bvar,3))
            peaks_pos = peaks_pos + [round(bvar,3)]
            out_text.insert(END, tvar+'\n')
        #print other
        for j in range(0,len(sp_df3[1])): #0-6 for v1, 0-7 for v2
            for i in range (1,2):
                avar = sp_df3[i].iloc[j,6:-1].astype(float)
                bvar = avar.idxmax()*0.101-20
                tvar = sp_df3[i].Species[j] + '  ' + str(round(bvar,3))
            peaks_neg = peaks_neg + [round(bvar,3)]
            out_text.insert(END, tvar+'\n')
    
    elif variable_peaktype.get()=='Max3':
        out_text.insert(END, 'Get Median of Top 3 Values:\n')
        #print SM
        for i in range(0,1):
            avar = sp_df3[i].iloc[0,6:-1].astype(float)
            bvar = avar.nlargest(3)
            bmid = sorted(bvar.index)[1]*0.101-5
            tvar = sp_df3[i].Species[0] + '  ' + str(round(bmid,3))
            peaks_pos = peaks_pos + [round(bmid,3)]
            out_text.insert(END, tvar+'\n')
        #print other
        for j in range(0,len(sp_df3[1])): #0-6 for v1, 0-7 for v2
            for i in range (1,2):
                avar = sp_df3[i].iloc[j,6:-1].astype(float)
                bvar = avar.nlargest(3)
                bmid = sorted(bvar.index)[1]*0.101-20
                tvar = sp_df3[i].Species[j] + '  ' + str(round(bmid,3))
            peaks_neg = peaks_neg + [round(bmid,3)]
            out_text.insert(END, tvar+'\n')
            
    elif variable_peaktype.get()=='Con5':
        out_text.insert(END, 'Get Median of Max Consecutive 5 Values:\n')
        #print SM
        for i in range(0,1):
            avar = sp_df3[i].iloc[0,6:-1].astype(float)
            bvar = avar.idxmax()
            c5 = 0
            idc5 = 0
            for j in range(bvar-15,bvar+12):
                if c5 < sum(avar[j:j+5]):
                    c5=sum(avar[j:j+5])
                    idc5 = j+2
            tvar = sp_df3[i].Species[0] + '  ' + str(round(idc5*0.101-5,3))
            peaks_pos = peaks_pos + [round(idc5*0.101-5,3)]
            out_text.insert(END, tvar+'\n')
        #print other
        for j in range(0,len(sp_df3[1])): #0-6 for v1, 0-7 for v2
            for i in range (1,2):
                avar = sp_df3[i].iloc[j,6:-1].astype(float)
                bvar = avar.idxmax()
                c5 = 0
                idc5 = 0
                for jj in range(bvar-15,bvar+12):
                    if c5 < sum(avar[jj:jj+5]):
                        c5=sum(avar[jj:jj+5])
                        idc5 = jj+2
                tvar = sp_df3[i].Species[j] + '  ' + str(round(idc5*0.101-20,3))
            peaks_neg = peaks_neg + [round(idc5*0.101-20,3)]
            out_text.insert(END, tvar+'\n')

#%%
    elif variable_peaktype.get()=='Con9':
        out_text.insert(END, 'Get Median of Max Consecutive 9 Values:\n')
        #print SM
        for i in range(0,1):
            avar = sp_df3[i].iloc[0,6:-1].astype(float)
            bvar = avar.idxmax()
            c9 = 0
            idc9 = 0
            for j in range(bvar-15,bvar+12):
                if c9 < sum(avar[j:j+9]):
                    c9=sum(avar[j:j+9])
                    idc9 = j+4
            tvar = sp_df3[i].Species[0] + '  ' + str(round(idc9*0.101-5,3))
            peaks_pos = peaks_pos + [round(idc9*0.101-5,3)]
            out_text.insert(END, tvar+'\n')
        #print other
        for j in range(0,len(sp_df3[1])): #0-6 for v1, 0-7 for v2
            for i in range (1,2):
                avar = sp_df3[i].iloc[j,6:-1].astype(float)
                bvar = avar.idxmax()
                c9 = 0
                idc9 = 0
                for jj in range(bvar-15,bvar+12):
                    if c9 < sum(avar[jj:jj+9]):
                        c9=sum(avar[jj:jj+9])
                        idc9 = jj+4
                tvar = sp_df3[i].Species[j] + '  ' + str(round(idc9*0.101-20,3))
            peaks_neg = peaks_neg + [round(idc9*0.101-20,3)]
            out_text.insert(END, tvar+'\n')
            
    out_text.insert(END, 'LPC/LPE16 cover C12-16\nLPC/LPE18 cover C17-24\n', "blue")
    out_text.insert(END, 'LPC16+1.0=LPC18\nLPE16+1.5=LPE18\n', "blue")
    

#%%    
    out_text.insert(END, ' '+'\n')
    out_text.see(END)
    out_text.configure(state="disabled")
    
    ##Add entry box according to template
    covlist[:] = list(sp_dict['NEG']['Group'].unique()) + list(sp_dict['POS']['Group'].unique())
    
    for i in covlist_label:
        #try:
        covlist_label[i].grid_forget() #remove lable from last run
    
    for i in covlist_dict:
        covlist_dict[i].grid_forget() #remove entry from last run
    
    #covlist_dict = dict() #reset covlist_dict
    
    for i in range(0,len(covlist)):
        covlist_label[i] = ttk.Label(tab1, text=covlist[i])
        covlist_label[i].grid(row=5+i,column=4, pady=1)
        covlist_dict[i]=ttk.Entry(tab1, width=10, state=NORMAL)
        covlist_dict[i].grid(row=5+i, column=5, columnspan = 1, 
                             sticky='w',padx=5,pady=1)
       
    #bind LPC1/2 LPE1/2
    #covlist_dict[covlist.index('LPC1')].bind('<Any-KeyRelease>', (lambda x: fetchLPC(covlist_dict)))
    #covlist_dict[covlist.index('LPE2')].bind('<Any-KeyRelease>', (lambda x: fetchLPE(covlist_dict)))

    #plot
    matplotlib.use('TkAgg')
    sns.set_style("whitegrid")
    pltdata = sp_df3[1].iloc[:,list(range(6,len(sp_df3[1].columns)))]
    pltdata_norm = pltdata.copy()
    #pltdata_norm = pltdata.drop([0,3], axis=0)
    pltdata_norm.loc[:,0:(len(pltdata_norm.columns)-1)] = pltdata_norm.iloc[:,0:(len(pltdata_norm.columns)-1)].div(pltdata_norm.iloc[:,0:(len(pltdata_norm.columns)-1)].max(axis=1), axis=0)
    pltdata1 = pd.melt(pltdata_norm, id_vars=["Species"], var_name="COV", value_name="Intensity") #use pltdata to plot real intensity
    pltdata1.COV = pltdata1.COV.astype(float)*0.101-20
    
    pltdata_sm = sp_df3[0].iloc[:,list(range(6,155))+[len(sp_df3[0].columns)-1]]
    pltdata_sm1 = pd.melt(pltdata_sm, id_vars=["Species"], var_name="COV", value_name="Intensity")
    pltdata_sm1.COV = pltdata_sm1.COV.astype(float)*0.101-5
    
    #pltdata1 = pd.concat([pltdata1, pltdata_sm1])    
    
    fig1 = plt.figure(num=None, figsize=(9, 5), dpi=120, facecolor='w', edgecolor='k')
    ax = sns.lineplot(x="COV", y="Intensity", hue="Species",
                  data=pltdata1)
    for i in range(0,len(peaks_neg)): #add peak lines [1,2,4,5] for V1
        ax.axvline(x=peaks_neg[i], linewidth=1)
    fig1.show()
    
    fig2 = plt.figure(num=None, figsize=(9, 5), dpi=120, facecolor='w', edgecolor='k')
    ax = sns.lineplot(x="COV", y="Intensity", hue="Species",
                  data=pltdata_sm1)
    for i in [0]: #add peak lines
        ax.axvline(x=peaks_pos[i], linewidth=1)
    fig2.show()
    
    #plot_canvas = FigureCanvasTkAgg(fig)
    #plot_canvas.insert(ax)
    #plt.savefig('peakplot.png')

    #messagebox.showinfo("Information","Done")

