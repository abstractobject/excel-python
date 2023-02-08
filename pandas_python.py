import pandas as pd
import tkinter as tk
import numpy as np
from tkinter import filedialog

#required before we can ask for input file
root = tk.Tk()
root.withdraw()

#asks for input file
excel_file = filedialog.askopenfilename()

#asks for save location
output_directory = filedialog.askdirectory()

##Multi 21 sheet

#read the excel file's first sheet, set line 1 (2nd line) as header for column names
df = pd.read_excel(excel_file, sheet_name=0, header=1)
#get rid of the top line of garbage
df = df[1:]

#rename column "ITEM.1" to "QTY"
df.rename(columns = {'ITEM.1':'QTY'}, inplace=True)


#####Angle order#####

#filter out everyhing but angle only
dfAngle = df[df['PART DESCRIPTION'].str.contains("Angle*", na=False, case=False)]
#filter out specifically flat bar. some slipped through that were 
dfAngle = dfAngle[~dfAngle['PART DESCRIPTION'].str.contains("Flat*", na=False, case=False)]
#sort by column MATERIAL DESCRIPTION
dfAngle = dfAngle.sort_values('MATERIAL DESCRIPTION')
#round up angles over half a stock length to a whole stock piece
dfAngle.loc[dfAngle['LENGTH.1'] >240, 'LENGTH.1'] = 480
#column sum = (total qty) x (length in inches)
dfAngle['SUM'] = dfAngle.apply(lambda row:(row['TOTAL'] * row['LENGTH.1']),axis=1)
#save to new excel file
dfAngle.to_excel(output_directory + "//DEBUGMultiAngle.xlsx", sheet_name="Sheet 1")
#add all of each material together
dfAngleGroup = dfAngle.groupby('MATERIAL DESCRIPTION').sum(numeric_only=True)
#delete the irrelevant columns that also got summed
dfAngleGroup = dfAngleGroup.drop('REV', axis=1)
dfAngleGroup = dfAngleGroup.drop('ITEM', axis=1)
dfAngleGroup = dfAngleGroup.drop('WEIGHT', axis=1)
#add STOCK column that divides sum by 480
dfAngleGroup['STOCK'] = dfAngleGroup.apply(lambda row:(row['SUM'] / 480),axis=1)
#add ROUND column that rounds up STOCK column
dfAngleGroup['ROUND'] = dfAngleGroup['STOCK'].apply(np.ceil)
#add +10% column that adds 10% to ROUND column
dfAngleGroup['+10%'] = dfAngleGroup.apply(lambda row:(row['ROUND'] * 1.1),axis=1)
#add ORDER coumn that rounds up +10% column
dfAngleGroup['ORDER'] = dfAngleGroup['+10%'].apply(np.ceil)
#save final product to different excel file
dfAngleGroup.to_excel(output_directory + "//DEBUGMultiAngleSum.xlsx", sheet_name="Sheet 1")
#delete the math columns so you get a clean copy-paste to the order form
dfAngleGroup = dfAngleGroup.drop('SUM', axis=1)
dfAngleGroup = dfAngleGroup.drop('STOCK', axis=1)
dfAngleGroup = dfAngleGroup.drop('ROUND', axis=1)
dfAngleGroup = dfAngleGroup.drop('+10%', axis=1)
#save the final order to a different excel file
dfAngleGroup.to_excel(output_directory + "//MultiAngleOrder.xlsx", sheet_name="Sheet 1")


#####Flat Bar order#####

#filter out everything but flat bar only
dfFlatBar = df[df['PART DESCRIPTION'].str.contains("Flat*", na=False, case=False)]
#sort by column MATERIAL DESCRIPTION
dfFlatBar = dfFlatBar.sort_values('MATERIAL DESCRIPTION')
#round up flat bar over half a stock length to a whole stock piece
dfFlatBar.loc[dfFlatBar['LENGTH.1'] >120, 'LENGTH.1'] = 240
#column sum = (total qty) x (length in inches)
dfFlatBar['SUM'] = dfFlatBar.apply(lambda row:(row['TOTAL'] * row['LENGTH.1']),axis=1)
#save to new excel file
dfFlatBar.to_excel(output_directory + "//DEBUGMultiFlatBar.xlsx", sheet_name="Sheet 1")
#add all of each material together
dfFlatBarGroup = dfFlatBar.groupby('MATERIAL DESCRIPTION').sum(numeric_only=True)
#delete the irrelevant columns that also got summed
dfFlatBarGroup = dfFlatBarGroup.drop('REV', axis=1)
dfFlatBarGroup = dfFlatBarGroup.drop('ITEM', axis=1)
dfFlatBarGroup = dfFlatBarGroup.drop('WEIGHT', axis=1)
#add STOCK column that divides sum by 240
dfFlatBarGroup['STOCK'] = dfFlatBarGroup.apply(lambda row:(row['SUM'] / 240),axis=1)
#add ROUND column that rounds up STOCK column
dfFlatBarGroup['ROUND'] = dfFlatBarGroup['STOCK'].apply(np.ceil)
#add +10% column that adds 10% to ROUND column
dfFlatBarGroup['+10%'] = dfFlatBarGroup.apply(lambda row:(row['ROUND'] * 1.1),axis=1)
#add ORDER coumn that rounds up +10% column
dfFlatBarGroup['ORDER'] = dfFlatBarGroup['+10%'].apply(np.ceil)
#save final product to different excel file
dfFlatBarGroup.to_excel(output_directory + "//DEBUGMultiFlatBarSum.xlsx", sheet_name="Sheet 1")
#delete the math columns so you get a clean copy-paste to the order form
dfFlatBarGroup = dfFlatBarGroup.drop('SUM', axis=1)
dfFlatBarGroup = dfFlatBarGroup.drop('STOCK', axis=1)
dfFlatBarGroup = dfFlatBarGroup.drop('ROUND', axis=1)
dfFlatBarGroup = dfFlatBarGroup.drop('+10%', axis=1)
#save the final order to a different excel file
dfFlatBarGroup.to_excel(output_directory + "//MultiFlatBarOrder.xlsx", sheet_name="Sheet 1")

#####Misc Material#####

#filter out everyhing but misc linear only
dfMisc = df[df['PART DESCRIPTION'].str.contains("w-beam*|s-beam*|pipe*|tube*", na=False, case=False)]
#sort by column MATERIAL DESCRIPTION
dfMisc = dfMisc.sort_values('MATERIAL DESCRIPTION')
#column sum = (total qty) x (length in inches)
dfMisc['SUM'] = dfMisc.apply(lambda row:(row['TOTAL'] * row['LENGTH.1']),axis=1)
#save to new excel file
dfMisc.to_excel(output_directory + "//MiscMaterial.xlsx", sheet_name="Sheet 1")

#####NUTS AND BOLTS#####

#filter out everyhing but nuts, bolts, and washers only
dfNutsAndBolts = df[df['PART DESCRIPTION'].str.contains("nut*|bolt*|washer*", na=False, case=False)].copy(deep=True)
#sort by column MATERIAL DESCRIPTION
#dfNutsAndBolts = dfNutsAndBolts.sort_values('MATERIAL DESCRIPTION')
#explodes entries with multiple stations to one line per station
dfNutsAndBolts = dfNutsAndBolts.assign(STRUCTURES=dfNutsAndBolts['STRUCTURES'].str.strip().str.split("|")).explode('STRUCTURES').reset_index(drop=True)
dfNutsAndBolts = dfNutsAndBolts.assign(STRUCTURES=dfNutsAndBolts['STRUCTURES'].str.strip())
#deleting assy and total columns to avoid confusion now structures are one pre line
dfNutsAndBolts = dfNutsAndBolts.drop('ASSY.', axis=1)
dfNutsAndBolts = dfNutsAndBolts.drop('TOTAL', axis=1)
#deleting columns that don't help ordering nust and bolts
dfNutsAndBolts = dfNutsAndBolts.drop('ITEM', axis=1)
dfNutsAndBolts = dfNutsAndBolts.drop('REV', axis=1)
dfNutsAndBolts = dfNutsAndBolts.drop('WEIGHT', axis=1)

#verification hardware
dfNutsAndBoltsVerif = dfNutsAndBolts.copy(deep=True)
#delete irrelevant columns
dfNutsAndBoltsVerif = dfNutsAndBoltsVerif.drop('DRAWING', axis=1)
dfNutsAndBoltsVerif = dfNutsAndBoltsVerif.drop('QTY', axis=1)
dfNutsAndBoltsVerif = dfNutsAndBoltsVerif.drop('STRUCTURES', axis=1)
#only one row per type of bolt
dfNutsAndBoltsVerif['GRADE'] = dfNutsAndBoltsVerif['GRADE'].fillna("N/A")
dfNutsAndBoltsVerif = dfNutsAndBoltsVerif.groupby(['PROJECT','MATERIAL DESCRIPTION','GRADE'], dropna=False).sum(numeric_only=True).reset_index()
#order 3 bolts per type
dfNutsAndBoltsVerif['TOTAL QTY'] = 3
#add column noting these as verification bolts
dfNutsAndBoltsVerif['USE'] = "Samples"
#save to new excel file
dfNutsAndBoltsVerif.to_excel(output_directory + "//Verification Hardware Nuts&Bolts ORDER.xlsx", sheet_name="Sheet 1")

# OLD shop bolts
#filter for shop bolts and field bolts. filter is whether sheet name contains an E
#dfShopBolts = dfNutsAndBolts[~dfNutsAndBolts['DRAWING'].str.contains("E", na=False, case=False)].copy(deep=True)
#dfShopBolts = dfShopBolts[~dfShopBolts['DRAWING'].str.contains("CA", na=False, case=False)]
#get a sum of bolts by type and station
#dfShopBolts.groupby(['PROJECT','MATERIAL DESCRIPTION','GRADE','DRAWING','STRUCTURES', 'QTY'], dropna=False).sum(numeric_only=True).reset_index(inplace=True)
#add 8% or +5 to shop bolts, whichever is more
#dfShopBolts['ORDER'] = dfShopBolts['QTY'].apply(lambda row:(row*1.08) if row>62 else (row+5))
#round up
#dfShopBolts['ORDER'] = dfShopBolts['ORDER'].apply(np.ceil)
#save to separate excel file
#dfShopBolts.to_excel(output_directory + "//DEBUGShopNuts&Bolts.xlsx", sheet_name="Sheet 1")
#add sheet name to station name column'
#dfShopBoltsCheck = dfShopBolts.copy(deep=True)
#dfShopBoltsOrder = dfShopBolts.copy(deep=True)
#dfShopBoltsOrder['STRUCTURES'] = dfShopBoltsOrder['DRAWING'] + ' | ' + dfShopBoltsOrder['STRUCTURES']
#delete sheet name column
#dfShopBoltsOrder = dfShopBoltsOrder.drop('DRAWING', axis=1)
#delete qty column
#dfShopBoltsOrder = dfShopBoltsOrder.drop('QTY', axis=1)
#pivot data to match nuts and bolts order form
#dfShopBoltsOrder['GRADE'] = dfShopBoltsOrder['GRADE'].fillna("N/A")
#dfShopBoltsOrder = pd.pivot_table(dfShopBoltsOrder, values='ORDER', index=['PROJECT','MATERIAL DESCRIPTION', 'GRADE'], columns='STRUCTURES', aggfunc=np.sum, fill_value=0)
#add total qty column adding together each bolt/nut/washer type
#dfShopBoltsOrder['TOTAL QTY'] = dfShopBoltsOrder.sum(axis=1)
#add column labeling all as "ASSY" so bolt order gets marked correctly
#dfShopBoltsOrder['USE'] = "ASSY"
#save to excel file
#dfShopBoltsOrder.to_excel(output_directory + "//Assy Nuts&Bolts ORDER.xlsx", sheet_name="Sheet 1")

#NEW shop bolts
#filter for shop bolts and field bolts. filter is whether sheet name contains an E
dfShopBolts2 = dfNutsAndBolts[~dfNutsAndBolts['DRAWING'].str.contains("E", na=False, case=False)].copy(deep=True)
dfShopBolts2 = dfShopBolts2[~dfShopBolts2['DRAWING'].str.contains("CA", na=False, case=False)]
#get a sum of bolts by type and station
dfShopBolts2.sort_values(by=['DRAWING', 'STRUCTURES', 'PART DESCRIPTION'], inplace=True)
#add sheet name to station name column'
dfShopBolts2['STRUCTURES'] = dfShopBolts2['DRAWING'] + ' | ' + dfShopBolts2['STRUCTURES']
#delete unnecessary columns
dfShopBolts2 = dfShopBolts2.drop('DRAWING', axis=1)
dfShopBolts2 = dfShopBolts2.drop('SHEET', axis=1)
dfShopBolts2 = dfShopBolts2.drop('MAIN NUMBER', axis=1)
dfShopBolts2 = dfShopBolts2.drop('PART NUMBER', axis=1)
dfShopBolts2 = dfShopBolts2.drop('WIDTH', axis=1)
dfShopBolts2 = dfShopBolts2.drop('WIDTH.1', axis=1)
dfShopBolts2 = dfShopBolts2.drop('LENGTH', axis=1)
dfShopBolts2 = dfShopBolts2.drop('LENGTH.1', axis=1)
dfShopBolts2.loc[dfShopBolts2['PART DESCRIPTION'] == 'Washer', 'GRADE'] = ' '
dfShopBolts2.loc[dfShopBolts2['PART DESCRIPTION'] == 'Nut', 'GRADE'] = ' '
#add 8% or +5 to shop bolts, whichever is more
dfShopBolts2['ORDER'] = dfShopBolts2['QTY'].apply(lambda row:(row*1.08) if row>62 else (row+5))
#round up
dfShopBolts2['ORDER'] = dfShopBolts2['ORDER'].apply(np.ceil)
#delete unnecessary columns
dfShopBolts2 = dfShopBolts2.drop('PART DESCRIPTION', axis=1)
dfShopBolts2 = dfShopBolts2.drop('QTY', axis=1)
dfShopBolts2['USE'] = "ASSY"
#save to separate excel file
dfShopBolts2.to_excel(output_directory + "//DEBUG-NEW-ShopNuts&Bolts.xlsx", sheet_name="Sheet 1")
dfShopBolts3 = pd.DataFrame([[''] * len(dfShopBolts2.columns)], columns=dfShopBolts2.columns)
# For each grouping Apply insert headers
dfShopBolts4 = (dfShopBolts2.groupby('STRUCTURES', group_keys=False)
        .apply(lambda d: d.append(dfShopBolts3))
        .iloc[:-2]
        .reset_index(drop=True))
dfShopBolts4.to_excel(output_directory + "//Assy Hardware Order.xlsx.xlsx", sheet_name="Sheet 1")


#NEW col assy bolts
#filter for shop bolts and field bolts. filter is whether sheet name contains "CA"
dfColAssyBolts = dfNutsAndBolts[dfNutsAndBolts['DRAWING'].str.contains("CA", na=False, case=False)].copy(deep=True)
#get a sum of bolts by type and station
dfColAssyBolts.sort_values(by=['DRAWING', 'STRUCTURES', 'PART DESCRIPTION'], inplace=True)
#add sheet name to station name column'
dfColAssyBolts['STRUCTURES'] = dfColAssyBolts['DRAWING'] + ' | ' + dfColAssyBolts['STRUCTURES']
#delete unnecessary columns
dfColAssyBolts = dfColAssyBolts.drop('DRAWING', axis=1)
dfColAssyBolts = dfColAssyBolts.drop('SHEET', axis=1)
dfColAssyBolts = dfColAssyBolts.drop('MAIN NUMBER', axis=1)
dfColAssyBolts = dfColAssyBolts.drop('PART NUMBER', axis=1)
dfColAssyBolts = dfColAssyBolts.drop('WIDTH', axis=1)
dfColAssyBolts = dfColAssyBolts.drop('WIDTH.1', axis=1)
dfColAssyBolts = dfColAssyBolts.drop('LENGTH', axis=1)
dfColAssyBolts = dfColAssyBolts.drop('LENGTH.1', axis=1)
dfColAssyBolts.loc[dfColAssyBolts['PART DESCRIPTION'] == 'Washer', 'GRADE'] = ' '
dfColAssyBolts.loc[dfColAssyBolts['PART DESCRIPTION'] == 'Nut', 'GRADE'] = ' '
#add 8% or +5 to shop bolts, whichever is more
dfColAssyBolts['ORDER'] = dfColAssyBolts['QTY'].apply(lambda row:(row*1.08) if row>62 else (row+5))
#round up
dfColAssyBolts['ORDER'] = dfColAssyBolts['ORDER'].apply(np.ceil)
#delete unnecessary columns
dfColAssyBolts = dfColAssyBolts.drop('PART DESCRIPTION', axis=1)
dfColAssyBolts = dfColAssyBolts.drop('QTY', axis=1)
dfColAssyBolts['USE'] = "COLUMN ASSY"
#save to separate excel file
dfColAssyBolts.to_excel(output_directory + "//DEBUG-NEW-ColAssyNuts&Bolts.xlsx", sheet_name="Sheet 1")
dfColAssyBolts2 = pd.DataFrame([[''] * len(dfColAssyBolts.columns)], columns=dfColAssyBolts.columns)
# For each grouping Apply insert headers
dfColAssyBolts3 = (dfColAssyBolts.groupby('STRUCTURES', group_keys=False)
        .apply(lambda d: d.append(dfShopBolts2))
        .iloc[:-2]
        .reset_index(drop=True))
dfColAssyBolts3.to_excel(output_directory + "//Col Assy Hardware Order.xlsx", sheet_name="Sheet 1")

# #column assy bolts
# #filter for shop bolts and field bolts. filter is whether sheet name contains an E
# dfColAssyBolts = dfNutsAndBolts[dfNutsAndBolts['DRAWING'].str.contains("CA", na=False, case=False)].copy(deep=True)
# #add 8% or +5 to shop bolts, whichever is more
# dfColAssyBolts['ORDER'] = dfColAssyBolts['QTY'].apply(lambda row:(row*1.08) if row>62 else (row+5))
# #round up
# dfColAssyBolts['ORDER'] = dfColAssyBolts['ORDER'].apply(np.ceil)
# #get a sum of bolts by type and station
# dfColAssyBolts.groupby(['PROJECT','MATERIAL DESCRIPTION','GRADE','DRAWING','STRUCTURES', 'QTY', 'ORDER'], dropna=False).sum(numeric_only=True).reset_index(inplace=True)
# #save to new excel file
# dfColAssyBolts.to_excel(output_directory + "//DEBUGColAssyNuts&Bolts.xlsx", sheet_name="Sheet 1")
# #add e-sheet name to station name column
# dfColAssyBoltsCheck = dfColAssyBolts.copy(deep=True)
# dfColAssyBoltsOrder = dfColAssyBolts.copy(deep=True)
# dfColAssyBoltsOrder['STRUCTURES'] = dfColAssyBoltsOrder['DRAWING'] + ' | ' + dfColAssyBoltsOrder['STRUCTURES']
# #delete e-sheet name column
# dfColAssyBoltsOrder = dfColAssyBoltsOrder.drop('DRAWING', axis=1)
# #delete qty column
# dfColAssyBoltsOrder = dfColAssyBoltsOrder.drop('QTY', axis=1)
# #pivot data to match nuts and bolts order form
# dfColAssyBoltsOrder['GRADE'] = dfColAssyBoltsOrder['GRADE'].fillna("N/A")
# dfColAssyBoltsOrder = pd.pivot_table(dfColAssyBoltsOrder, values='ORDER', index=['PROJECT','MATERIAL DESCRIPTION', 'GRADE'], columns='STRUCTURES', aggfunc=np.sum, fill_value=0)
# #add total qty column adding together each bolt/nut/washer type
# dfColAssyBoltsOrder['TOTAL QTY'] = dfColAssyBoltsOrder.sum(axis=1)
# #add column labeling all as "SHIP LOOSE" so bolt order gets marked correctly
# dfColAssyBoltsOrder['USE'] = "COLUMN ASSY"
# #save to excel file
# dfColAssyBoltsOrder.to_excel(output_directory + "//Column Assy Nuts&Bolts ORDER.xlsx", sheet_name="Sheet 1")

#NEW field bolts
#filter for shop bolts and field bolts. filter is whether sheet name contains "CA"
dfFieldBolts = dfNutsAndBolts[dfNutsAndBolts['DRAWING'].str.contains("E", na=False, case=False)].copy(deep=True)
#get a sum of bolts by type and station
dfFieldBolts.sort_values(by=['DRAWING', 'STRUCTURES', 'PART DESCRIPTION'], inplace=True)
#add sheet name to station name column'
dfFieldBolts['STRUCTURES'] = dfFieldBolts['DRAWING'] + ' | ' + dfFieldBolts['STRUCTURES']
#delete unnecessary columns
dfFieldBolts = dfFieldBolts.drop('DRAWING', axis=1)
dfFieldBolts = dfFieldBolts.drop('SHEET', axis=1)
dfFieldBolts = dfFieldBolts.drop('MAIN NUMBER', axis=1)
dfFieldBolts = dfFieldBolts.drop('PART NUMBER', axis=1)
dfFieldBolts = dfFieldBolts.drop('WIDTH', axis=1)
dfFieldBolts = dfFieldBolts.drop('WIDTH.1', axis=1)
dfFieldBolts = dfFieldBolts.drop('LENGTH', axis=1)
dfFieldBolts = dfFieldBolts.drop('LENGTH.1', axis=1)
dfFieldBolts.loc[dfFieldBolts['PART DESCRIPTION'] == 'Washer', 'GRADE'] = ' '
dfFieldBolts.loc[dfFieldBolts['PART DESCRIPTION'] == 'Nut', 'GRADE'] = ' '
#add 2 to each bolt count
dfFieldBolts['ORDER'] = dfFieldBolts.apply(lambda row:(row['QTY'] + 2),axis=1)
#delete unnecessary columns
dfFieldBolts = dfFieldBolts.drop('PART DESCRIPTION', axis=1)
dfFieldBolts = dfFieldBolts.drop('QTY', axis=1)
dfFieldBolts['USE'] = "SHIP LOOSE"
#save to separate excel file
dfFieldBolts.to_excel(output_directory + "//DEBUG-NEW-ShipLooseNuts&Bolts.xlsx", sheet_name="Sheet 1")
dfFieldBolts2 = pd.DataFrame([[''] * len(dfFieldBolts.columns)], columns=dfFieldBolts.columns)
# For each grouping Apply insert headers
dfFieldBolts3 = (dfFieldBolts.groupby('STRUCTURES', group_keys=False)
        .apply(lambda d: d.append(dfFieldBolts2))
        .iloc[:-2]
        .reset_index(drop=True))
dfFieldBolts3.to_excel(output_directory + "//Ship Loose Hardware Order.xlsx", sheet_name="Sheet 1")

# #field bolts
# #filter for shop bolts and field bolts. filter is whether sheet name contains an E
# dfFieldBolts = dfNutsAndBolts[dfNutsAndBolts['DRAWING'].str.contains("E", na=False, case=False)].copy(deep=True)
# #add 2 to each bolt count
# dfFieldBolts['ORDER'] = dfFieldBolts.apply(lambda row:(row['QTY'] + 2),axis=1)
# #get a sum of bolts by type and station
# dfFieldBolts.groupby(['PROJECT','MATERIAL DESCRIPTION','GRADE','DRAWING','STRUCTURES', 'QTY', 'ORDER'], dropna=False).sum(numeric_only=True).reset_index(inplace=True)
# #save to new excel file
# dfFieldBolts.to_excel(output_directory + "//DEBUGFieldNuts&Bolts.xlsx", sheet_name="Sheet 1")
# #add e-sheet name to station name column
# dfFieldBoltsOrder = dfFieldBolts
# dfFieldBoltsOrder['STRUCTURES'] = dfFieldBoltsOrder['DRAWING'] + ' | ' + dfFieldBoltsOrder['STRUCTURES']
# #delete e-sheet name column
# dfFieldBoltsOrder = dfFieldBoltsOrder.drop('DRAWING', axis=1)
# #delete qty column
# dfFieldBoltsOrder = dfFieldBoltsOrder.drop('QTY', axis=1)
# #pivot data to match nuts and bolts order form
# dfFieldBoltsOrder['GRADE'] = dfFieldBoltsOrder['GRADE'].fillna("N/A")
# dfFieldBoltsOrder = pd.pivot_table(dfFieldBoltsOrder, values='ORDER', index=['PROJECT','MATERIAL DESCRIPTION', 'GRADE'], columns='STRUCTURES', aggfunc=np.sum, fill_value=0)
# #add total qty column adding together each bolt/nut/washer type
# dfFieldBoltsOrder['TOTAL QTY'] = dfFieldBoltsOrder.sum(axis=1)
# #add column labeling all as "SHIP LOOSE" so bolt order gets marked correctly
# dfFieldBoltsOrder['USE'] = "SHIP LOOSE"
# #save to excel file
# dfFieldBoltsOrder.to_excel(output_directory + "//Ship Loose Nuts&Bolts ORDER.xlsx", sheet_name="Sheet 1")

#function for checking old nuts and bolts orders
#adds assy bolts and col assy bolts
#dfNutsAndBoltsCheck = pd.concat([dfShopBoltsCheck, dfColAssyBoltsCheck])
#delete drawing column
#dfNutsAndBoltsCheck = dfNutsAndBoltsCheck.drop('DRAWING', axis=1)
#delete qty column
#dfNutsAndBoltsCheck = dfNutsAndBoltsCheck.drop('QTY', axis=1)
#adds together quantities of similar nuts and bolts
#dfNutsAndBoltsCheck['GRADE'] = dfNutsAndBoltsCheck['GRADE'].fillna("N/A")
#dfNutsAndBoltsCheck.groupby(['PROJECT','MATERIAL DESCRIPTION','GRADE','STRUCTURES', 'ORDER'], dropna=False).sum(numeric_only=True).reset_index(inplace=True)
#pivots info to match old nuts and bolts order form
#dfNutsAndBoltsCheck= pd.pivot_table(dfNutsAndBoltsCheck, values='ORDER', index=['PROJECT','MATERIAL DESCRIPTION', 'GRADE'], columns='STRUCTURES', aggfunc=np.sum, fill_value=0)
#adds sum column to end
#dfNutsAndBoltsCheck['TOTAL QTY'] = dfNutsAndBoltsCheck.sum(axis=1)
#saves to new excel form
#dfNutsAndBoltsCheck.to_excel(output_directory + "//CHECK OLD Nuts&Bolts ORDER.xlsx", sheet_name="Sheet 1")


#####Misc Hardware#####

#filter out everyhing already covered
dfRemain = df[~df['PART DESCRIPTION'].str.contains("angle*|flat*|plate*|beam*|pipe*|tube*|screw*|bolt*|washer*|nut*|weld*", na=False, case=False)]
#sort by column MATERIAL DESCRIPTION
dfRemain = dfRemain.sort_values('MATERIAL DESCRIPTION')
#save to new excel file
dfRemain.to_excel(output_directory + "//MiscHardware.xlsx", sheet_name="Sheet 1")


#filter out everything but clamp plates
dfClampPl = df[df['PART NUMBER'].str.contains("CPS*", na=False, case=False)].copy(deep=True)
#sory be part number column
dfClampPl = dfClampPl.sort_values('PART NUMBER')
#delete unnecessary columns
dfClampPl['TOTAL'] = dfClampPl.apply(lambda row:(row['QTY'] * row['ASSY.']),axis=1)
dfClampPl = dfClampPl.drop('REV', axis=1)
dfClampPl = dfClampPl.drop('ITEM', axis=1)
dfClampPl = dfClampPl.drop('WEIGHT', axis=1)
dfClampPl = dfClampPl.drop('SHEET', axis=1)
dfClampPl = dfClampPl.drop('MAIN NUMBER', axis=1)
dfClampPl = dfClampPl.drop('WIDTH', axis=1)
dfClampPl = dfClampPl.drop('WIDTH.1', axis=1)
dfClampPl = dfClampPl.drop('DRAWING', axis=1)
dfClampPl = dfClampPl.drop('LENGTH.1', axis=1)
dfClampPl = dfClampPl.drop('QTY', axis=1)
dfClampPl = dfClampPl.drop('ASSY.', axis=1)
dfClampPl = dfClampPl.drop('STRUCTURES', axis=1)
dfClampPl = dfClampPl.drop('GRADE', axis=1)
#add together clamp plates of the same name
dfClampPl = dfClampPl.groupby(['PROJECT', 'PART NUMBER', 'PART DESCRIPTION', 'MATERIAL DESCRIPTION', 'LENGTH'],dropna=False).sum(numeric_only=True).reset_index()
#save to new excel file
dfClampPl.to_excel(output_directory + "//ClampPlates.xlsx", sheet_name="Sheet 1")

