<!-- PROJECT SHIELDS -->
[![MIT License][license-shield]][license-url]


<!-- PROJECT TITLE -->
<p align="center">
 <h3 align="center">Shotgun Lipidomics Assistant</h3>
  <p align="center">
    A python project to read and analysis Sciex Analyst output data
    <br />
    <a href="https://www.uclalipidomics.net/"><strong>UCLA Lipidomics Core</strong></a>
    <br />
  </p>
</p>



<!-- TABLE OF CONTENTS -->
## Table of Contents

* [Getting Started](#getting-started)
* [Tuning](#Tuning)
  * [Instrument Setup](#Instrument-Setup)
  * [Run](#Run)
  * [Data Analysis](#Data-Analysis)
* [Suitability Test](#Suitability-Test)
  * [Setup](#Setup)
  * [Running](#Running)
  * [Analysis](#Analysis)
* [Sample Run Setup](#Sample-Run-Setup)
* [Sample Run](#Sample-Run)
* [Data Analysis Setup](#Data-Analysis-Setup)
* [Read MZML](#Read-MZML)
* [Merge data with sample map](#Merge-data-with-sample-map)
* [Further Analysis](#Further-Analysis)



<!-- Getting Started -->
## Getting Started
1. Download SLA_V1.12 from Github and extract to fold of your choice. If desired, create
Exe shortcut to desktop. Run SLA_V1.12.exe.
2. If it does not open properly, or if it collapses while running, try to open it from
command prompt to check the error message.
3. To convert wiff file to mzml file, you need to download and isntall MSconvertGUI from
Proteowizard. (Note: We have been using version 3.0.19082-ade61137d. Some updated
versions may not transfer the data or file name correctly.)
4. A spreadsheet program must be installed on the control computer (Excel or LibreOffice).



<!-- Tuning – Instrument Setup -->
## Tuning

### Instrument Setup(#Tuning-Instrument-Setup)
1. This protocol assumes that users are utilizing a Sciex Lipidyzer with the 100ul sample
loop installed in the Shimadzu Autosampler.
2. Prior to Tuning/Sample Run, make sure that the instrument is topped off with
appropriate solvents and solutions. The pump system should be supplied with running
buffer (50/50% Methanol/Dichloromethane with 10mM ammonium acetate). The
autosampler wash should be topped off with 2-propanol. The SelexION should be
supplied with 1-propanol as modifier.
3. Tuning solution is prepared by diluting EquiSPLASH™ LIPIDOMIX® (Avanti, 330731-1EA) 1
to 20 with running buffer. The syringe pump should be loaded with .5-1mL of this
tuning solution and connected to the electrode (the autosampler output is disconnected
from the mass spec for this step).
4. Load a sample vial containing running buffer in Position 1 in sample rack. This will be
utilized in the Tuning and Suitability steps.

### Run
1. It is recommended that you create a new project or subproject for each experiment in
order to avoid confusion in appropriately tuned methods. In Analyst: Tools → Project →
Create Project
2. Copy “Neg Infusion COV 3500_SPLASH”, “Pos Infusion COV 3500_SPLASH”, “SST v3”,
“Method 1 v3” and “Method 2 v3” from provided method package to the current
project “Acquisition Method” folder.
3. “Hardware Config” → Activate “Lipidyzer” profile
4. Acquire → Ready Instrument
5. Purge Modifier and wait 30min for instrument to warm up
6. While waiting, set up batches. Negative and positive infusion batches should each
contain three repeats of the same respective method.
7. For Positive Batch: Build Acquisition Batch → add batch name(e.g. “Tune – *Date* – 1”)
and add 3 samples named “POS1/2/3” (all drawing from vial position 1) → submit
8. For Negative Batch: Build Acquisition Batch → add batch name(e.g. “Tune – *Date* – 2”)
and add 3 samples named “NEG1/2/3” (all drawing from vial position 1) → submit
9. Start syringe pump 3-5min before sample run start

### Data Analysis:
1. Following the run, find and copy both sets of .WIFF files to a working folder in a location
of your choice. Convert to mzml with MSConvertGUI.
2. Read Tuning data with SLA Tuning tab.
3. Click Import POS/Import NEG and choose the corresponding mzml files. (We
recommend using the last one among the 3 replicates.)
4. Click Import Tune Dict to import the Tuning_spname_dict_xxx.xlsx file.
5. Choose peak finding method from drop down list. We recommend “Con9”.
6. Hit the “Run” button, the auto selected peak results will be printed to the Result
window and plots will pop out. (The Group column in the POS and NEG tab in
Tuning_spname_dict_xxx.xlsx file are corresponding to the items listed under the Export
area. You can customize it by editing the Group column.)
7. Fill in COV values in the Export area. You can copy and paste the recommended values
from results to Export boxes. Alternatively, you can manually choose COV values. You
can use your mouse to point at a different peaks in the plots. The x shown on the
bottom right corner of the plot is the COV value. Note that most tuning mixes have a
single LPC and LPE species and yet these subclasses have two different COV values for
shorter and longer chain species. The calculation for the offset between these is
provided on the tab. Copy the one COV value and calculate the other COV from the
first. Hit ExportResult button when finished. Result will be saved to a xlsx file under the
same directory with Mzml files. (The excel file will be open automatically. A command
prompt may pop-up. You can close it after the excel file is open.)

![tuneshot1][tuneshot1](https://example.com)
[![tuneshot2][tuneshot2]](https://example.com)
[![tuneshot3][tuneshot3]](https://example.com)

8. Copy/paste the volt column to the corresponding Analyst method COV column. The
“POS” and “NEG” tabs correspond to the respective experiments in “Method 1”. The
SST tab contains COV values for both positive and negative experiments in the SST
method. Save the respective files after this modification.

[![tuneshot4][tuneshot4]](https://example.com)



<!-- Suitability Test -->
## Suitability Test

### Setup
(Note: The Suitability Test is not required but is recommended to ensure that the instrument is
in good working order and that the previous tuning step was carried out properly. The
Suitability Test mix should be prepared fresh on a regular basis. Baseline results for the
suitability test should be established and significant decreases should be troubleshot
appropriately.)

1. The suitability test mix is prepared from two working stocks: the Lipidyzer System
Suitability Kit (Sciex 5040407) and the PG/PI/PS mix from Avanti (described below).
2. The PG/PI/PS mix is prepared from 17:0-18:1 PI-d5, 17:0-18:1 PG-d5 and 17:0-18:1 PS-
d5 (Avanti 850111L-500ug, 858133L-1mg, 858151L-1mg). A 100ug/mL stock of these
three compounds should be prepared by diluting in 50/50 DCM/Methanol. Several mLs
can be prepared at once and aliquoted into vials or ampules. As an example, the
following formulation makes 1mL (100uL PG, 200uL PI, 100uL PS, 300uL Methanol,
300uL DCM).
3. The Suitability Test mix is prepared by combining 10ul of Lipidyzer System Suitability
Mix with 10ul of PG/PI/PS mix in 980ul of running buffer.
4. Prepare Suitability Test mix in a robovial and place in the 105 position in the sample
rack. Make sure there is sufficient volume of running buffer in the position 1 vial.
5. Reconnect the autosampler output to the source.

### Analysis
1. Find and copy .WIFF files to a working folder. Extract to Mzml format with
MSConverterGUI.
2. Read the Suitability Test results with the SLA using the SST tab.
3. Import the Mzml file with the SST result (the “LOD” file above).
4. Import the Tuning_spname_dict_xxx.xlsx file
5. Hit Run and the result will be saved to an xlsx file under the same directory with the
Mzml file
6. Compare results with previous suitability tests.

[![SSTshot1][SSTshot1]](https://example.com)






[license-shield]: https://img.shields.io/github/license/othneildrew/Best-README-Template.svg?style=flat-square
[license-url]: https://github.com/syjgino/LA_V1/blob/master/LICENSE
[tuneshot1]: screeshot/Tune1.PNG
[tuneshot2]: screeshot/Tune2.PNG
[tuneshot3]: screeshot/Tune3.PNG
[tuneshot4]: screeshot/Tune4.PNG
[SSTshot1]: screeshot/SST1.PNG
