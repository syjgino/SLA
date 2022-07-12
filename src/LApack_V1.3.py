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



import os
from tkinter import *
import tkinter as tk
from tkinter import ttk
from ttkthemes import ThemedTk

os.chdir(os.path.dirname(os.path.realpath(__file__)))

from Tab1Tune1_1 import *
from Tab2SST1_1 import *
# from Tab3Readmzml1_1 import *  # using LWM lipid name format
from Tab3Readmzml1_21 import *  # new name and FA analysis version
# from Tab3Readmzml1_2_drop import *  # custom drop points among 20 scans
from Tab4Aggregate1_21 import *
from Tab5TAGplot1_2 import *

covlist = list()  # list of COV
covlist_label = dict()  # dictionary of COV lavbel
covlist_dict = dict()  # dictionary of COV entry box

# TK
root = ThemedTk(theme="clearlooks")
root.title("ShotgunLipidomicsAssistant V1.21")
root.iconbitmap('SLAicon_256.ico')

note = ttk.Notebook(root)
tab1_container = ttk.Frame(note)
tab2 = ttk.Frame(note)
tab3 = ttk.Frame(note)
tab4 = ttk.Frame(note)
tab5 = ttk.Frame(note)
note.add(tab1_container, text="Tuning")
note.add(tab2, text="SST")
note.add(tab3, text="Read mzml")
note.add(tab4, text="Merge")
note.add(tab5, text="ClassTotal/TAG Analysis")
note.pack()

########
##Tab1##
########
# choices for peak selecting method
tab1_canvas = tk.Canvas(tab1_container)
tab1scroll = ttk.Scrollbar(tab1_container, command=tab1_canvas.yview)
tab1_canvas.config(yscrollcommand=tab1scroll.set,
                   height='13.7c')  # scrollregion=(0,0,0,1000))
tab1scroll.grid(column=1, row=0, sticky='ns', rowspan=100)
tab1_canvas.grid(column=0, row=0, rowspan=100, sticky='ns')
# tab1scroll.grid(column=1,row=0, sticky='ns', rowspan=100)
tab1 = ttk.Frame(tab1_canvas)
tab1_canvas.create_window((0, 0), window=tab1, anchor="nw")
tab1.bind("<Configure>",
          lambda event,
                 tab1_canvas=tab1_canvas: tab1_canvas.configure(scrollregion=tab1_canvas.bbox("all"),
                                                                width=event.width))

choices_peaktype = ['Max1', 'Max3', 'Con5', 'Con9']
variable_peaktype = StringVar(tab1)
w_peakm = ttk.OptionMenu(tab1, variable_peaktype, choices_peaktype[3], *choices_peaktype)
w_peakm.grid(row=3, column=2, sticky='w')

# text boxes
tunef1 = Text(tab1, width=50, height=2, state=DISABLED)  # mzml directory location
tunef1.grid(row=0, column=1, columnspan=2, sticky='w', padx=1, pady=1)
tunef2 = Text(tab1, width=50, height=2, state=DISABLED)  # mzml directory location
tunef2.grid(row=1, column=1, columnspan=2, sticky='w', padx=1, pady=1)
maploc_tune = Text(tab1, width=50, height=2, state=DISABLED)  # tuning_spname_dict file location
maploc_tune.grid(row=2, column=1, columnspan=2, sticky='w', padx=1, pady=1)

scroll = ttk.Scrollbar(tab1)  # scroll bar for output text box
scroll.grid(row=5, column=3, sticky=N + S + W, rowspan=12)
out_text = Text(tab1, width=50, height=18.5, state=DISABLED)
out_text.grid(row=5, column=0, columnspan=3, rowspan=12, sticky='wnse', pady=5, padx=5)
out_text.tag_config("blue", foreground="blue")
out_text.config(yscrollcommand=scroll.set)
scroll.config(command=out_text.yview)

# labels
ttk.Label(tab1, text='Peak Method').grid(row=3, column=1, sticky='e')
ttk.Label(tab1, text='').grid(row=3, column=7, columnspan=1, pady=2, padx=2)
ttk.Label(tab1, text='Result:\nLPC/LPE16 cover C12-16 | LPC/LPE18 cover C17-24\nLPC16+1.0=LPC18 | LPE16+1.5=LPE18',
          font=('Helvetica', 10, 'bold')).grid(row=4, column=0, columnspan=3, pady=2, sticky='w')
ttk.Label(tab1, text='Export:', font=('Helvetica', 10, 'bold')).grid(row=4, column=4, pady=2, sticky='w')

# buttons
ttk.Button(tab1, text='Import POS', command=lambda: get_tune1(tunef1)).grid(row=0, column=0)
ttk.Button(tab1, text='Import NEG', command=lambda: get_tune2(tunef2)).grid(row=1, column=0)
ttk.Button(tab1, text='Import Tune Dict', command=lambda: imp_tunekey(maploc_tune)).grid(row=2, column=0, padx=1)
ttk.Button(tab1, text='Run',
           command=lambda: TuningFun(tunef1, tunef2, maploc_tune, out_text, variable_peaktype, tab1, covlist,
                                     covlist_label, covlist_dict)).grid(row=3, column=0, padx=1)
ttk.Button(tab1, text='ExportResult', command=lambda: exportdata(covlist, covlist_dict, maploc_tune)).grid(row=3,
                                                                                                           column=5,
                                                                                                           padx=0,
                                                                                                           pady=1)

########
##Tab2##
########
# text boxes
scroll = ttk.Scrollbar(tab2)
scroll.grid(row=3, column=2, sticky=N + S + W, rowspan=1, pady=5)
text = Text(tab2, width=50, height=25, state=DISABLED)  # output text
text.grid(row=3, column=0, columnspan=2, rowspan=1, sticky='nsew', padx=5, pady=5)
text.config(yscrollcommand=scroll.set)
scroll.config(command=text.yview)
maploc2 = Text(tab2, width=50, height=2, state=DISABLED)  # tuning_spname_dict location, same as maploc_tune
maploc2.grid(row=1, column=1, columnspan=1, sticky='w', padx=1, pady=1)
mzmlloc = Text(tab2, width=50, height=2, state=DISABLED)
mzmlloc.grid(row=0, column=1, columnspan=1, sticky='w', padx=1, pady=1)

# buttons
ttk.Button(tab2, text='Import Tune Dict', command=lambda: imp_tunekey2(maploc2)).grid(row=1, column=0, padx=1)
ttk.Button(tab2, text='Import SST', command=lambda: imp_mzml(mzmlloc)).grid(row=0, column=0)
ttk.Button(tab2, text='Run', command=lambda: SSTFun(mzmlloc, maploc2, text)).grid(row=2, column=0)

#################
##Tab3 readMZml##
#################
ttk.Button(tab3, text='Set Directory', command=lambda: set_dir_read(dirloc_read)).grid(row=0, column=0)
ttk.Button(tab3, text='Import Standard Dict', command=lambda: get_std_dict(std_dict_loc)).grid(row=1, column=0)
ttk.Button(tab3, text='Import SpName Dict', command=lambda: get_sp_dict1(sp_dict1_loc)).grid(row=2, column=0)
ttk.Button(tab3, text='Import Iso Dict', command=lambda: get_iso_dict(iso_dict_loc)).grid(row=3, column=0)
ttk.Button(tab3, text='Read MZML', command=lambda: readMZML(dirloc_read,
                                                            sp_dict1_loc, std_dict_loc,
                                                            proname3, iso_dict_loc, variable_iso, variable_std0cut,
                                                            progressBar, variable_dataversion, variable_mute
                                                            )).grid(row=6, column=2)

ttk.Label(tab3, text='Project Name').grid(row=4, column=0, pady=10)
ttk.Label(tab3, text='Max 0s for Standard').grid(row=5, column=0, pady=10)
ttk.Label(tab3, text='Isotope Correction').grid(row=6, column=0, pady=10)
ttk.Label(tab3, text='mzml Version').grid(row=7, column=0, pady=8)  # choice V, dataversion
ttk.Label(tab3, text='Mute Species').grid(row=8, column=0, pady=8)  # choice V2, mute species

dirloc_read = Text(tab3, width=50, height=2, state=DISABLED)  # work directory location
dirloc_read.grid(row=0, column=1, columnspan=2, sticky='w', padx=1, pady=1)
std_dict_loc = Text(tab3, width=50, height=2, state=DISABLED)  # standard dictionary location
std_dict_loc.grid(row=1, column=1, columnspan=2, sticky='w', padx=1, pady=1)
sp_dict1_loc = Text(tab3, width=50, height=2, state=DISABLED)  # species name/mrm list location
sp_dict1_loc.grid(row=2, column=1, columnspan=2, sticky='w', padx=1, pady=1)
iso_dict_loc = Text(tab3, width=50, height=2, state=DISABLED)  # isotope correction file location
iso_dict_loc.grid(row=3, column=1, columnspan=2, sticky='w', padx=1, pady=1)
proname3 = ttk.Entry(tab3, width=30)
proname3.grid(row=4, column=1, columnspan=1, sticky='w')

# choice for number of 0s cut
choices_std0cut = range(0,21)
variable_std0cut = IntVar(tab3)
w_std0cut = ttk.OptionMenu(tab3, variable_std0cut, choices_std0cut[2], *choices_std0cut)
w_std0cut.grid(row=5, column=1, sticky='w')


# choice for isotope correction
choices_iso = ['Yes', 'No']
variable_iso = StringVar(tab3)
# variable_iso.set('No')
w_iso = ttk.OptionMenu(tab3, variable_iso, choices_iso[0], *choices_iso)
w_iso.grid(row=6, column=1, sticky='w')

# choice for data file version
choices_dataversion = ['Analyst', 'LWM']
variable_dataversion = StringVar(tab3)
# data_version.set('Analyst')
V = ttk.OptionMenu(tab3, variable_dataversion, choices_dataversion[0], *choices_dataversion)
V.grid(row=7, column=1, sticky='w')

# choice for species version, muted or full list
choices_mute = ['Yes', 'No']
variable_mute = StringVar(tab3)
# variable_version.set('NewClass')
V2 = ttk.OptionMenu(tab3, variable_mute, choices_mute[0], *choices_mute)
V2.grid(row=8, column=1, sticky='w')

# progressbar
progressBar = ttk.Progressbar(tab3, orient=HORIZONTAL,
                              length=100, mode='determinate')
progressBar.grid(row=9, column=0, columnspan=3, sticky=EW, padx=10, pady=3)

########################
##Tab4 Aggregate/merge##
########################
ttk.Label(tab4, text='Project Name').grid(row=4, column=0, pady=10)
ttk.Label(tab4, text='Map Sheet').grid(row=1, column=2, sticky='w', padx=1)

dirloc_aggregate = Text(tab4, width=50, height=2, state=DISABLED)  # directory location
dirloc_aggregate.grid(row=0, column=1, columnspan=1, sticky='w', pady=3, padx=1)
maploc = Text(tab4, width=50, height=2, state=DISABLED)  # sample submition form/map location
maploc.grid(row=1, column=1, columnspan=1, sticky='w', pady=3, padx=1)
method1loc = Text(tab4, width=50, height=2, state=DISABLED)  # m1 output file
method1loc.grid(row=2, column=1, columnspan=1, sticky='w', pady=3, padx=1)
method2loc = Text(tab4, width=50, height=2, state=DISABLED)  # m2 output file
method2loc.grid(row=3, column=1, columnspan=1, sticky='w', pady=3, padx=1)
proname = ttk.Entry(tab4, width=30)  # project name
proname.grid(row=4, column=1, columnspan=1, sticky='w', pady=3, padx=1)

ttk.Button(tab4, text='Set Directory', command=lambda: set_dir(dirloc_aggregate)).grid(row=0, column=0, padx=1)
ttk.Button(tab4, text='Import Map', 
           command=lambda: imp_map(maploc, 
                                   w_mapsheet, 
                                   choices_mapsheet,
                                   variable_mapsheet)).grid(row=1, column=0, padx=1)
ttk.Button(tab4, text='Import Method1', command=lambda: imp_method1(method1loc)).grid(row=2, column=0, padx=1)
ttk.Button(tab4, text='Import Method2', command=lambda: imp_method2(method2loc)).grid(row=3, column=0, padx=1)
ttk.Button(tab4, text='Run Merge',
           command=lambda: MergeApp(dirloc_aggregate, proname, 
                                    method1loc, method2loc, 
                                    maploc, variable_mapsheet,
                                    CheckClustVis)).grid(row=5, column=1, sticky='e')

# option to output .csv file for ClustVis
CheckClustVis = IntVar()
ttk.Checkbutton(tab4, text="Generate .csv file for ClustVis", variable=CheckClustVis,
                onvalue=1, offvalue=0,  # height=0,
                width=0).grid(row=5, column=1, sticky=W, pady=0)

# dropdown to choose map sheet
choices_mapsheet = ['Sheet1']
variable_mapsheet = StringVar(tab4)
# data_version.set('Analyst')
w_mapsheet = ttk.OptionMenu(tab4, variable_mapsheet, 
                            choices_mapsheet[0], *choices_mapsheet)
w_mapsheet.grid(row=1, column=3, sticky='w')


############
##Tab5 TAG##
############
exploc = Text(tab5, width=50, height=2, state=DISABLED)
exploc.grid(row=0, column=1, columnspan=2, sticky='w', pady=5, padx=1)

CheckVar1 = IntVar()
CheckVar2 = IntVar()
ttk.Checkbutton(tab5, text="Mute Error Bar", variable=CheckVar1,
                onvalue=1, offvalue=0,  # height=0, standard Checkbutton only, ttk.Checkbutton doesn't alow height
                width=0).grid(row=2, column=1, sticky=W, pady=0)
ttk.Checkbutton(tab5, text="Show Points(TG plot)", variable=CheckVar2,
                onvalue=1, offvalue=0,  # height=0,
                width=0).grid(row=2, column=2, sticky=W, pady=0)
ttk.Button(tab5, text='Choose File(_exp)',
           command=lambda: imp_exp(exploc)).grid(row=0, column=0, pady=0, padx=1)
ttk.Button(tab5, text='Run',
           command=lambda: tagplot(exploc, CheckVar1, CheckVar2)).grid(row=3, column=2, sticky='e')

root.mainloop()




"""
@author: BaolongSu

SLA main
all imported Tab files should be in the same location as this file

v1.21 
add dropdown list to choose the max number of 0s allowed

v1.21 20210920
fix map_loc function name duplication

v1.3 20220712
add chol as m3
add option to choose number of scans
add option to choose max number of 0s in unknows allowed
"""