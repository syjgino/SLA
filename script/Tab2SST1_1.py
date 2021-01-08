# -*- coding: utf-8 -*-
"""
Created on Thu Jan 16 16:58:09 2020

@author: Gino

for LA_V1
get SST result

"""

import numpy as np
import pandas as pd       
import os
import glob
import re

from scipy import stats
from pyopenms import *
from tkinter import *
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
from tkinter.messagebox import showinfo
from tkinter import filedialog
import datetime

#%%
#def set_dir(dirloc):
#    #setdir = filedialog.askdirectory()  
#    dirloc.configure(state="normal")
#    dirloc.delete(1.0, END)
#    setdir = filedialog.askdirectory()
#    dirloc.insert(INSERT, setdir)
#    dirloc.configure(state="disabled")
    
def imp_map2(maploc2):
    #map1 = filedialog.askopenfilename(filetypes=(("excel Files", "*.xlsx"),("all files", "*.*")))
    maploc2.configure(state="normal")
    maploc2.delete(1.0 ,END)
    map1 = filedialog.askopenfilename(filetypes=(("excel Files", "*.xlsx"),("all files", "*.*")))
    maploc2.insert(INSERT, map1)
    maploc2.configure(state="disabled")
   
def imp_mzml(mzmlloc):
    #map1 = filedialog.askopenfilename(filetypes=(("excel Files", "*.xlsx"),("all files", "*.*")))
    mzmlloc.configure(state="normal")
    mzmlloc.delete(1.0 ,END)
    mzml1 = filedialog.askopenfilename(filetypes=(("mzML Files", "*.mzML"),("all files", "*.*")))
    mzmlloc.insert(INSERT, mzml1)
    mzmlloc.configure(state="disabled")


def SSTFun(mzmlloc, maploc2, text):
    text.configure(state="normal")
    ##get name all files
    sp_dict1_loc = maploc2.get('1.0', 'end-1c')
    #std_dict_loc = 'C:/Users/baolongsu/Desktop/Projects/StdUnkRatio/standard_dict - LipidizerSimulate_102b_MW_dev2.xlsx'
    mzml_loc = mzmlloc.get('1.0', 'end-1c')
    #change dir to where mzml file is
    os.chdir(os.path.dirname(mzml_loc))
    #list_of_files = glob.glob('./*.mzML')
    
    def spName(row):
        return(sp_dict[method].loc[(pd.Series(sp_dict[method]['Q1'] == row['Q1']) & pd.Series(sp_dict[method]['Q3'] == row['Q3']))].index[0])
    
    ##create all variable
    all_df_dict = {'1':{}, '2':{}}
    #out_df2 = pd.DataFrame()
    out_df2 = {'1':pd.DataFrame(), '2': pd.DataFrame()}
    #out_df2_con = pd.DataFrame()
    out_df2_con = {'1':pd.DataFrame(), '2': pd.DataFrame()}
    out_df2_intensity = {'1':pd.DataFrame(), '2': pd.DataFrame()}
    #sp_df3 = {'A':0}
    
    ##read dicts
    sp_dict = {}
    sp_dict['1'] = pd.read_excel(sp_dict1_loc, sheet_name = 'SST', header=0, index_col=-2, na_values='.')
    #sp_dict['2'] = pd.read_excel(sp_dict1_loc, sheet_name = '2', header=0, index_col=-1, na_values='.')
    
    ##loop start
    #start0 = datetime.datetime.now()
    #for sample in range(0, len(list_of_files)):
    ##get method number
    method = '1'
    #start = datetime.datetime.now()
    #print('method'+method, str(sample))
    #read all files
    exp=MSExperiment()
    MzMLFile().load(mzml_loc,exp)
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
    
    sp_df['NativeID'] = pd.Series(sp_serie[2:]) #first 2 are TIC and BPC
    
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
    intensity_df2 = intensity_df.iloc[:, 2:len(intensity_df.loc[1])]  #first 2 are place holder. in result and tuning, only first 1 is holder
    intensity_df2 = intensity_df2.T
    intensity_df2 = intensity_df2.reset_index()
    intensity_df2 = intensity_df2.drop(columns = ['index'])
    intensity_df2 = intensity_df2.iloc[:, 0:100]  #100 measures
    
    ##merge sp and intensity
    sp_df2 = sp_df2.merge(intensity_df2, left_index = True, right_index=True)
    
    ##change experiment1 to negative, add species name
    if max(sp_df2['experiment']) == 2:
        sp_df2['Q1'][sp_df2['experiment'] == 1] = -sp_df2['Q1'][sp_df2['experiment'] == 1]
    #sp_df2['Species'] = sp_dict[method].index
    sp_df2['Species'] = ''
    sp_df2['Species'] = sp_df2.apply(spName, axis=1)
    sp_df2['AverageIntensity'] = sp_df2.iloc[:, 6:106].mean(axis=1)
    sp_df3 = sp_df2.loc[:,['Species','AverageIntensity']]
    noblank = -sp_df3['Species'].apply(lambda x: "BLANK" in x)
    sp_df3 = sp_df3.loc[noblank,:]
    sp_df3 = sp_df3.sort_values('Species')
    
    text.insert(END, 'Average Intensity:\n')
    text.insert(END, sp_df3.to_string(index=False)+'\n\n')
    text.see(END)
    text.configure(state="disabled")
    
    #save
    fname = 'SSTResult_' + datetime.datetime.now().strftime("%Y%m%d%H%M") + '.xlsx'
    master = pd.ExcelWriter(fname)
    sp_df3.to_excel(master,'result', index = True)
    master.save()
    os.system(fname)
    
    #messagebox.showinfo("Information","Done")



