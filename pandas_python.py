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
#column sum = (total qty) x (length in inches)
dfAngle['SUM'] = dfAngle.apply(lambda row:(row['TOTAL'] * row['LENGTH.1']),axis=1)
#save to new excel file
dfAngle.to_excel(output_directory + "//MultiAngle.xlsx", sheet_name="Sheet 1")
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
dfAngleGroup.to_excel(output_directory + "//MultiAngleSum.xlsx", sheet_name="Sheet 1")
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
#column sum = (total qty) x (length in inches)
dfFlatBar['SUM'] = dfFlatBar.apply(lambda row:(row['TOTAL'] * row['LENGTH.1']),axis=1)
#save to new excel file
dfFlatBar.to_excel(output_directory + "//MultiFlatBar.xlsx", sheet_name="Sheet 1")
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
dfFlatBarGroup.to_excel(output_directory + "//MultiFlatBarSum.xlsx", sheet_name="Sheet 1")
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
dfNutsAndBolts = df[df['PART DESCRIPTION'].str.contains("nut*|bolt*|washer*", na=False, case=False)]
#sort by column MATERIAL DESCRIPTION
dfNutsAndBolts = dfNutsAndBolts.sort_values('MATERIAL DESCRIPTION')
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

#shop bolts
#filter for shop bolts and field bolts. filter is whether sheet name contains an E
dfShopBolts = dfNutsAndBolts[~dfNutsAndBolts['DRAWING'].str.contains("E", na=False, case=False)]
#get a sum of bolts by type and station
dfShopBolts = dfShopBolts.groupby(['MATERIAL DESCRIPTION','GRADE','DRAWING','STRUCTURES','QTY']).sum(numeric_only=True).reset_index()
#add 8% or +5 to shop bolts, whichever is more
dfShopBolts['ORDER'] = dfShopBolts['QTY'].apply(lambda row:(row*1.08) if row>62 else (row+5))
#round up
dfShopBolts['ORDER'] = dfShopBolts['ORDER'].apply(np.ceil)
#save to separate excel file
dfShopBolts.to_excel(output_directory + "//ShopNuts&Bolts.xlsx", sheet_name="Sheet 1")

#column assy bolts
#filter for shop bolts and field bolts. filter is whether sheet name contains an E
dfColAssyBolts = dfNutsAndBolts[dfNutsAndBolts['DRAWING'].str.contains("CA", na=False, case=False)]
#add 8% or +5 to shop bolts, whichever is more
#need to fix this so it doesn't give me warnings every time
dfColAssyBolts['ORDER'] = dfColAssyBolts['QTY'].apply(lambda row:(row*1.08) if row>62 else (row+5))
#round up
dfColAssyBolts['ORDER'] = dfColAssyBolts['ORDER'].apply(np.ceil)
#get a sum of bolts by type and station
dfColAssyBolts = dfColAssyBolts.groupby(['PROJECT','MATERIAL DESCRIPTION','GRADE','DRAWING','STRUCTURES', 'QTY', 'ORDER']).sum(numeric_only=True).reset_index()
#save to new excel file
dfColAssyBolts.to_excel(output_directory + "//FieldNuts&Bolts.xlsx", sheet_name="Sheet 1")
#add e-sheet name to station name column
dfColAssyBolts['STRUCTURES'] = dfColAssyBolts['DRAWING'] + ' | ' + dfColAssyBolts['STRUCTURES']
#delete e-sheet name column
dfColAssyBolts = dfColAssyBolts.drop('DRAWING', axis=1)
#delete qty column
dfColAssyBolts = dfColAssyBolts.drop('QTY', axis=1)
#pivot data to match nuts and bolts order form
dfColAssyBolts = pd.pivot_table(dfColAssyBolts, values='ORDER', index=['PROJECT','MATERIAL DESCRIPTION', 'GRADE'], columns='STRUCTURES', fill_value=0)
#add total qty column adding together each bolt/nut/washer type
dfColAssyBolts['TOTAL QTY'] = dfColAssyBolts.sum(axis=1)
#add column labeling all as "SHIP LOOSE" so bolt order gets marked correctly
dfColAssyBolts['USE'] = "COLUMN ASSY"
#save to excel file
dfColAssyBolts.to_excel(output_directory + "//Column Assy Nuts&Bolts ORDER.xlsx", sheet_name="Sheet 1")

#field bolts
#filter for shop bolts and field bolts. filter is whether sheet name contains an E
dfFieldBolts = dfNutsAndBolts[dfNutsAndBolts['DRAWING'].str.contains("E", na=False, case=False)]
#add 2 to each bolt count
#need to fix this so it doesn't give me warnings every time
dfFieldBolts['ORDER'] = dfFieldBolts.apply(lambda row:(row['QTY'] + 2),axis=1)
#get a sum of bolts by type and station
dfFieldBolts = dfFieldBolts.groupby(['PROJECT','MATERIAL DESCRIPTION','GRADE','DRAWING','STRUCTURES', 'QTY', 'ORDER']).sum(numeric_only=True).reset_index()
#save to new excel file
dfFieldBolts.to_excel(output_directory + "//FieldNuts&Bolts.xlsx", sheet_name="Sheet 1")
#add e-sheet name to station name column
dfFieldBolts['STRUCTURES'] = dfFieldBolts['DRAWING'] + ' | ' + dfFieldBolts['STRUCTURES']
#delete e-sheet name column
dfFieldBolts = dfFieldBolts.drop('DRAWING', axis=1)
#delete qty column
dfFieldBolts = dfFieldBolts.drop('QTY', axis=1)
#pivot data to match nuts and bolts order form
dfFieldBoltsOrder = pd.pivot_table(dfFieldBolts, values='ORDER', index=['PROJECT','MATERIAL DESCRIPTION', 'GRADE'], columns='STRUCTURES', fill_value=0)
#add total qty column adding together each bolt/nut/washer type
dfFieldBoltsOrder['TOTAL QTY'] = dfFieldBoltsOrder.sum(axis=1)
#add column labeling all as "SHIP LOOSE" so bolt order gets marked correctly
dfFieldBoltsOrder['USE'] = "SHIP LOOSE"
#save to excel file
dfFieldBoltsOrder.to_excel(output_directory + "//Ship Loose Nuts&Bolts ORDER.xlsx", sheet_name="Sheet 1")

#####Misc Hardware#####

#filter out everyhing already covered
dfRemain = df[~df['PART DESCRIPTION'].str.contains("angle*|flat*|plate*|beam*|pipe*|tube*|screw*|bolt*|washer*|nut*|weld*", na=False, case=False)]
#sort by column MATERIAL DESCRIPTION
dfRemain = dfRemain.sort_values('MATERIAL DESCRIPTION')
#save to new excel file
dfRemain.to_excel(output_directory + "//MiscHardware.xlsx", sheet_name="Sheet 1")