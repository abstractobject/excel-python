import pandas as pd
import tkinter as tk
import numpy as np
from tkinter import filedialog
from ortools.linear_solver import pywraplp

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

#get project name
projectName = df.loc[2]['PROJECT']

#####Angle order#####

#filter out everyhing but angle only
dfAngle = df[df['PART DESCRIPTION'].str.contains("Angle*", na=False, case=False)]
#filter out specifically flat bar. some slipped through that were 
dfAngle = dfAngle[~dfAngle['PART DESCRIPTION'].str.contains("Flat*", na=False, case=False)]
#sort by column MATERIAL DESCRIPTION
dfAngle = dfAngle.sort_values('MATERIAL DESCRIPTION')
#round up angles over half a stock length to a whole stock piece
dfAngleRound = dfAngle.copy(deep=True)
dfAngleRound.loc[dfAngleRound['LENGTH.1'] >240, 'LENGTH.1'] = 480
#column sum = (total qty) x (length in inches)
dfAngleSum = dfAngleRound
dfAngleSum['SUM'] = dfAngleSum.apply(lambda row:(row['TOTAL'] * row['LENGTH.1']),axis=1)
#save to new excel file
#dfAngle.to_excel(output_directory + "//DEBUGMultiAngle.xlsx", sheet_name="Sheet 1")
#add all of each material together
#dfNutsAndBoltsVerif = dfNutsAndBoltsVerif.groupby(['PROJECT','MATERIAL DESCRIPTION','GRADE'], dropna=False).sum(numeric_only=True).reset_index()
dfAngleGroup = dfAngleSum.groupby(['PROJECT','MATERIAL DESCRIPTION'],dropna=False).sum(numeric_only=True)
#dfFlatBarGroup = dfFlatBar.groupby('MATERIAL DESCRIPTION').sum(numeric_only=True)
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
dfAngleGroup.to_excel(output_directory + "//" + projectName + " DEBUGMultiAngleSum.xlsx", sheet_name="Sheet 1")
#delete the math columns so you get a clean copy-paste to the order form
dfAngleGroup = dfAngleGroup.drop('SUM', axis=1)
dfAngleGroup = dfAngleGroup.drop('STOCK', axis=1)
dfAngleGroup = dfAngleGroup.drop('ROUND', axis=1)
dfAngleGroup = dfAngleGroup.drop('+10%', axis=1)
#save the final order to a different excel file
#dfAngleGroup.to_excel(output_directory + "//MultiAngleOrder.xlsx", sheet_name="Sheet 1")

#prepping data for angle nesting
dfAngleNest = dfAngle.copy(deep=True)
dfAngleNest = dfAngleNest.assign(STRUCTURES=dfAngleNest['STRUCTURES'].str.strip().str.split("|")).explode('STRUCTURES').reset_index(drop=True)
dfAngleNest = dfAngleNest.assign(STRUCTURES=dfAngleNest['STRUCTURES'].str.strip())
dfAngleNest = dfAngleNest.drop('ASSY.', axis=1)
dfAngleNest = dfAngleNest.drop('TOTAL', axis=1)
dfAngleNest = dfAngleNest.loc[dfAngleNest.index.repeat(dfAngleNest['QTY'])].reset_index(drop=True)
dfAngleNest['QTY'] = 1
dfAngleNest = dfAngleNest.drop('REV', axis=1)
dfAngleNest = dfAngleNest.drop('SHEET', axis=1)
dfAngleNest = dfAngleNest.drop('MAIN NUMBER', axis=1)
#dfAngleNest = dfAngleNest.drop('ITEM', axis=1)
dfAngleNest = dfAngleNest.drop('PART DESCRIPTION', axis=1)
dfAngleNest = dfAngleNest.drop('WIDTH', axis=1)
dfAngleNest = dfAngleNest.drop('WIDTH.1', axis=1)
dfAngleNest = dfAngleNest.drop('GRADE', axis=1)
dfAngleNest = dfAngleNest.drop('WEIGHT', axis=1)
dfAngleNest['LENGTH.1'] = dfAngleNest['LENGTH.1'].apply(lambda x: x*10000)
dfAngleNest['LENGTH.1'] = dfAngleNest['LENGTH.1'].apply(lambda x:(x+1250) if x<4800000 else x)
dfAngleNest.to_excel(output_directory + "//" + projectName + " DEBUGMultiAngleNest.xlsx", sheet_name="Sheet 1")

#prepping excel sheet for angle order after nesting

AngleCutTicketWorksetDataFrame = []
AngleNestWorksetDataFrame = []

def create_data_model_angle():
      data = {}
      data['weights'] = dfAngleType['LENGTH.1'].values.tolist()
      data['items'] = list(range(len(data['weights'])))
      data['bins'] = data['items']
      data['bin_capacity'] = 4800000
      data['material'] = dfAngleType.iloc[0,5]
      data['structures'] = dfAngleType.iloc[0,8]
      data['drawing'] = dfAngleType.iloc[0,2]
      return data

#angle nesting fuction
for group, dfAngleType in dfAngleNest.groupby(['DRAWING', 'MATERIAL DESCRIPTION', 'STRUCTURES']):    
    
    angleMaterial = dfAngleType.iloc[0,5]

    data = create_data_model_angle()

        # Create the mip solver with the SCIP backend.
    solver = pywraplp.Solver.CreateSolver('CP-SAT')
    #solver.set_time_limit = 60000
   
        # Variables
        # x[i, j] = 1 if item i is packed in bin j.
    x = {}
    for i in data['items']:
        for j in data['bins']:
            x[(i, j)] = solver.IntVar(0, 1, 'x_%i_%i' % (i, j))

        # y[j] = 1 if bin j is used.
    y = {}
    for j in data['bins']:
        y[j] = solver.IntVar(0, 1, 'y[%i]' % j)

        # Constraints
        # Each item must be in exactly one bin.
    for i in data['items']:
        solver.Add(sum(x[i, j] for j in data['bins']) == 1)

        # The amount packed in each bin cannot exceed its capacity.
    for j in data['bins']:
        solver.Add(
            sum(x[(i, j)] * data['weights'][i] for i in data['items']) <= y[j] *
            data['bin_capacity'])

        # Objective: minimize the number of bins used.
    solver.Minimize(solver.Sum([y[j] for j in data['bins']]))

    status = solver.Solve()

    if status == pywraplp.Solver.OPTIMAL or status == pywraplp.Solver.FEASIBLE:
        num_bins = 0
        bin_usage = 0
        for j in data['bins']:
            if y[j].solution_value() == 1:
                bin_items = []
                bin_weight = 0
                for i in data['items']:
                    if x[i, j].solution_value() > 0:
                        bin_items.append(i)
                        bin_weight += data['weights'][i]
                if bin_items:
                    num_bins += 1
                    if bin_weight/4800000 < 0.75 and bin_weight/4800000 > 0.25:
                        bin_usage += round(bin_weight/4800000, 2)
                    elif bin_weight/4800000 > 0.75:
                        bin_usage += 1
                    else:
                        bin_usage += 0.25
                    # print('Stick number', j)
                    # print('  Items nested:', '\n',  dfAngleType.iloc[bin_items,2], '\n', dfAngleType.iloc[bin_items,3])
                    # print('  Total length:', bin_weight/4800000)
                    # print('  Usage:', bin_usage)
                    # print()
        #print(dfAngleType.iloc[bin_items,3])
        #print('Number of sticks used:', num_bins)
        #print('Time = ', solver.WallTime(), ' milliseconds')
        AngleNestDictionary = {'PROJECT': projectName, 'DRAWING': data['drawing'], 'MATERIAL DESCRIPTION': data['material'], 'ORDER':num_bins, 'USAGE':bin_usage, 'STRUCTURES': data['structures']}
        AngleNestDictionaryDataFrame = pd.DataFrame(data=AngleNestDictionary, index=[0])
        AngleNestWorksetDataFrame.append(AngleNestDictionaryDataFrame)
        dfAngleTypeSum = dfAngleType.groupby(['PROJECT', 'DRAWING', 'ITEM', 'PART NUMBER', 'MATERIAL DESCRIPTION', 'LENGTH', 'STRUCTURES'])['QTY'].sum(numeric_only=True).reset_index()
        dfAngleTypeSum['ORDER'] = num_bins
        dfAngleTypeSum['USAGE'] = bin_usage
        AngleCutTicketWorksetDataFrame.append(dfAngleTypeSum)
        solver.Clear()
    else:
          print('The problem does not have an optimal or feasible solution.')

        
#saving angle nesting results        
AngleCutTicketDataFrame = pd.concat(AngleCutTicketWorksetDataFrame, ignore_index=True)
AngleCutTicketDataFrame.to_excel(output_directory + "//" + projectName + " DEBUGAngleCutTicket.xlsx", sheet_name="Sheet 1")

writerCutTicket = pd.ExcelWriter(output_directory + "//" + projectName + " DEBUGCutTicket.xlsx")
for group, dfCutTicket in AngleCutTicketDataFrame.groupby(['DRAWING', 'STRUCTURES']): 
    dfCutTicket = dfCutTicket.sort_values(by='ITEM')
    dfCutTicket = dfCutTicket.sort_values(by='MATERIAL DESCRIPTION')
    dfCutTicket.to_excel(writerCutTicket, sheet_name=dfCutTicket.iloc[0,1] + " | " + dfCutTicket.iloc[0,6])

writerCutTicket.close()

writer = pd.ExcelWriter(output_directory + "//" + projectName + " DEBUGNestAngleOrder.xlsx")
AnglePoseNestDataFrame = pd.concat(AngleNestWorksetDataFrame, ignore_index=True)
AnglePoseNestDataFrame.to_excel(output_directory + "//" + projectName + " DEBUGPostNestAngle.xlsx", sheet_name="Sheet 1")
AnglePoseNestDataFrame = AnglePoseNestDataFrame.drop('STRUCTURES', axis=1)
AnglePoseNestDataFrame = AnglePoseNestDataFrame.drop('DRAWING', axis=1)
AnglePoseNestDataFrameSUM = AnglePoseNestDataFrame.groupby('MATERIAL DESCRIPTION').sum(numeric_only=True).reset_index()
#print(AnglePostNestDCutTicketSUM)
AnglePoseNestDataFrameSUM.to_excel(writer)
writer.close()

#####Flat Bar order#####

#filter out everything but flat bar only
dfFlatBar = df[df['PART DESCRIPTION'].str.contains("Flat*", na=False, case=False)]
#sort by column MATERIAL DESCRIPTION
dfFlatBar = dfFlatBar.sort_values('MATERIAL DESCRIPTION')
#round up flat bar over half a stock length to a whole stock piece
dfFlatBarRound = dfFlatBar.copy(deep=True)
dfFlatBarRound.loc[dfFlatBarRound['LENGTH.1'] >120, 'LENGTH.1'] = 240
#column sum = (total qty) x (length in inches)
dfFlatBarSum = dfFlatBarRound
dfFlatBarSum['SUM'] = dfFlatBarSum.apply(lambda row:(row['TOTAL'] * row['LENGTH.1']),axis=1)
#save to new excel file
#dfFlatBar.to_excel(output_directory + "//DEBUGMultiFlatBar.xlsx", sheet_name="Sheet 1")
#add all of each material together
dfFlatBarGroup= dfFlatBarSum.groupby(['PROJECT','MATERIAL DESCRIPTION'],dropna=False).sum(numeric_only=True)
#dfFlatBarGroup = dfFlatBar.groupby('MATERIAL DESCRIPTION').sum(numeric_only=True)
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
dfFlatBarGroup.to_excel(output_directory + "//" + projectName + " DEBUGMultiFlatBarSum.xlsx", sheet_name="Sheet 1")
#delete the math columns so you get a clean copy-paste to the order form
dfFlatBarGroup = dfFlatBarGroup.drop('SUM', axis=1)
dfFlatBarGroup = dfFlatBarGroup.drop('STOCK', axis=1)
dfFlatBarGroup = dfFlatBarGroup.drop('ROUND', axis=1)
dfFlatBarGroup = dfFlatBarGroup.drop('+10%', axis=1)
#save the final order to a different excel file
#dfFlatBarGroup.to_excel(output_directory + "//MultiFlatBarOrder.xlsx", sheet_name="Sheet 1")

#prepping data for flat bar nesting
dfFlatBarNest = dfFlatBar.copy(deep=True)
dfFlatBarNest = dfFlatBarNest.assign(STRUCTURES=dfFlatBarNest['STRUCTURES'].str.strip().str.split("|")).explode('STRUCTURES').reset_index(drop=True)
dfFlatBarNest = dfFlatBarNest.assign(STRUCTURES=dfFlatBarNest['STRUCTURES'].str.strip())
dfFlatBarNest = dfFlatBarNest.drop('ASSY.', axis=1)
dfFlatBarNest = dfFlatBarNest.drop('TOTAL', axis=1)
dfFlatBarNest = dfFlatBarNest.loc[dfFlatBarNest.index.repeat(dfFlatBarNest['QTY'])].reset_index(drop=True)
dfFlatBarNest = dfFlatBarNest.drop('QTY', axis=1)
dfFlatBarNest = dfFlatBarNest.drop('REV', axis=1)
dfFlatBarNest = dfFlatBarNest.drop('SHEET', axis=1)
dfFlatBarNest = dfFlatBarNest.drop('MAIN NUMBER', axis=1)
#dfFlatBarNest = dfFlatBarNest.drop('ITEM', axis=1)
dfFlatBarNest = dfFlatBarNest.drop('PART DESCRIPTION', axis=1)
dfFlatBarNest = dfFlatBarNest.drop('WIDTH', axis=1)
dfFlatBarNest = dfFlatBarNest.drop('WIDTH.1', axis=1)
dfFlatBarNest = dfFlatBarNest.drop('GRADE', axis=1)
dfFlatBarNest = dfFlatBarNest.drop('WEIGHT', axis=1)
dfFlatBarNest['LENGTH.1'] = dfFlatBarNest['LENGTH.1'].apply(lambda x: x*10000)
dfFlatBarNest['LENGTH.1'] = dfFlatBarNest['LENGTH.1'].apply(lambda x:(x+1250) if x<2400000 else x)
dfFlatBarNest.to_excel(output_directory + "//" + projectName + " DEBUGMultiFlatBarNest.xlsx", sheet_name="Sheet 1")

#prepping excel sheet for FlatBar order after nesting
writer = pd.ExcelWriter(output_directory + "//" + projectName + " DEBUGNestFlatBarOrder.xlsx")
FlatBarNestWorksetDataFrame = []
text_file = open("FlatBarNestingDebugOutput.txt", "w")

def create_data_model_FlatBar():
      data = {}
      data['weights'] = dfFlatBarType['LENGTH.1'].values.tolist()
      #data['items'] = dfFlatBarNest['PART NUMBER'].values.tolist()
      data['items'] = list(range(len(data['weights'])))
      data['bins'] = data['items']
      data['bin_capacity'] = 2400000
      data['material'] = dfFlatBarType.iloc[0,4]
      return data

#FlatBar nesting fuction
for group, dfFlatBarType in dfFlatBarNest.groupby(['MATERIAL DESCRIPTION', 'STRUCTURES']):    

    flatBarMaterial = dfFlatBarType.iloc[0,3]

    data = create_data_model_FlatBar()

        # Create the mip solver with the SCIP backend.
    solver = pywraplp.Solver.CreateSolver('CP-SAT')
    #solver.set_time_limit = 60000
   

        # Variables
        # x[i, j] = 1 if item i is packed in bin j.
    x = {}
    for i in data['items']:
        for j in data['bins']:
            x[(i, j)] = solver.IntVar(0, 1, 'x_%i_%i' % (i, j))

        # y[j] = 1 if bin j is used.
    y = {}
    for j in data['bins']:
        y[j] = solver.IntVar(0, 1, 'y[%i]' % j)

        # Constraints
        # Each item must be in exactly one bin.
    for i in data['items']:
        solver.Add(sum(x[i, j] for j in data['bins']) == 1)

        # The amount packed in each bin cannot exceed its capacity.
    for j in data['bins']:
        solver.Add(
            sum(x[(i, j)] * data['weights'][i] for i in data['items']) <= y[j] *
            data['bin_capacity'])

        # Objective: minimize the number of bins used.
    solver.Minimize(solver.Sum([y[j] for j in data['bins']]))

    status = solver.Solve()

    if status == pywraplp.Solver.OPTIMAL or status == pywraplp.Solver.FEASIBLE:
        num_bins = 0
        bin_usage = 0
        for j in data['bins']:
            if y[j].solution_value() == 1:
                bin_items = []
                bin_weight = 0
                for i in data['items']:
                    if x[i, j].solution_value() > 0:
                        bin_items.append(i)
                        bin_weight += data['weights'][i]
                if bin_items:
                    num_bins += 1
                    if bin_weight/2400000 < 0.85 and bin_weight/2400000 > 0.25:
                        bin_usage += round(bin_weight/2400000, 2)
                    elif bin_weight/2400000 > 0.85:
                        bin_usage += 1
                    else:
                        bin_usage += 0.25
                #     print('Stick number', j)
                #     print('  Items nested:', '\n',  dfFlatBarType.iloc[bin_items,2], '\n', dfFlatBarType.iloc[bin_items,3])
                #     print('  Total length:', bin_weight/10000)
                #     print('  Usage:', bin_usage)
                #     print()
        #text_file.write(dfFlatBarType.iloc[bin_items,3])
        #text_file.write('Number of sticks used:', num_bins)
        #text_file.write('Time = ', solver.WallTime(), ' milliseconds')
        FlatBarNestDictionary = {'PROJECT': projectName, 'MATERIAL DESCRIPTION': data['material'], 'ORDER':num_bins}
        FlatBarNestDictionaryDataFrame = pd.DataFrame(data=FlatBarNestDictionary, index=[0])
        FlatBarNestWorksetDataFrame.append(FlatBarNestDictionaryDataFrame)
        solver.Clear()
    else:
          print('The problem does not have an optimal or feasible solution.')

        
#saving FlatBar nesting results   
text_file.close()     
FlatBarPostNestDataFrame = pd.concat(FlatBarNestWorksetDataFrame, ignore_index=True)
FlatBarPostNestDataFrameSUM= FlatBarPostNestDataFrame.groupby('MATERIAL DESCRIPTION').sum(numeric_only=True).reset_index()
#print(FlatBarPostNestDataFrameSUM)
FlatBarPostNestDataFrameSUM.to_excel(writer)
writer.close()

#combined anglematic nested order
dfAnglematicNestedInput = [AnglePoseNestDataFrameSUM,FlatBarPostNestDataFrameSUM]
dfAnglematicNested = pd.concat(dfAnglematicNestedInput)
dfAnglematicNested['HEAT #'] = None
dfAnglematicNested.to_excel(output_directory + "//" + projectName + " Anglematic Order Nested.xlsx", sheet_name="Sheet 1")

#Combined Anglematic Order#

dfAnglematicInput = [dfAngleGroup,dfFlatBarGroup]
dfAnglematic = pd.concat(dfAnglematicInput)
dfAnglematic['HEAT #'] = None
dfAnglematic.to_excel(output_directory + "//" + projectName + " Anglematic Order.xlsx", sheet_name="Sheet 1")

#####Misc Material#####

#filter out everyhing but misc linear only
dfMisc = df[df['PART DESCRIPTION'].str.contains("w-beam*|s-beam*|pipe*|tube*", na=False, case=False)]
#sort by column MATERIAL DESCRIPTION
dfMisc = dfMisc.sort_values('MATERIAL DESCRIPTION')
#column sum = (total qty) x (length in inches)
dfMisc['SUM'] = dfMisc.apply(lambda row:(row['TOTAL'] * row['LENGTH.1']),axis=1)
#save to new excel file
dfMisc.to_excel(output_directory + "//" + projectName + " Misc Material.xlsx", sheet_name="Sheet 1")

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
dfNutsAndBoltsVerif.to_excel(output_directory + "//" + projectName + " Verification Hardware Order.xlsx", sheet_name="Sheet 1")


#NEW shop bolts
#filter for shop bolts and field bolts. filter is whether sheet name contains an E
dfShopBolts2 = dfNutsAndBolts[~dfNutsAndBolts['DRAWING'].str.contains("E", na=False, case=False)].copy(deep=True)
dfShopBolts2 = dfShopBolts2[~dfShopBolts2['DRAWING'].str.contains("CA", na=False, case=False)]
#function for sorting bolts by dia
dfShopBolts2['DIA'] = np.where(dfShopBolts2['MATERIAL DESCRIPTION'].str.contains('1/2"ø'), 0.5,
                   np.where(dfShopBolts2['MATERIAL DESCRIPTION'].str.contains('5/8"ø'), 0.625,
                   np.where(dfShopBolts2['MATERIAL DESCRIPTION'].str.contains('3/4"ø'), .75, "OTHER")))
#get a sum of bolts by type and station
dfShopBolts2.sort_values(by=['DRAWING', 'STRUCTURES','DIA', 'PART DESCRIPTION'], inplace=True)
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
#dfShopBolts2.to_excel(output_directory + "//DEBUG-NEW-ShopNuts&Bolts.xlsx", sheet_name="Sheet 1")
#delete unnecessary column
dfShopBolts2 = dfShopBolts2.drop('DIA', axis=1)
dfShopBolts3 = pd.DataFrame([[''] * len(dfShopBolts2.columns)], columns=dfShopBolts2.columns)
# For each grouping Apply insert headers
dfShopBolts4 = (dfShopBolts2.groupby('STRUCTURES', group_keys=False)
        .apply(lambda d: d.append(dfShopBolts3))
        .iloc[:-2]
        .reset_index(drop=True))
dfShopBolts4.to_excel(output_directory + "//" + projectName + " Assy Hardware Order.xlsx", sheet_name="Sheet 1")


#NEW col assy bolts
#filter for shop bolts and field bolts. filter is whether sheet name contains "CA"
dfColAssyBolts = dfNutsAndBolts[dfNutsAndBolts['DRAWING'].str.contains("CA", na=False, case=False)].copy(deep=True)
#function for sorting bolts by dia
dfColAssyBolts['DIA'] = np.where(dfColAssyBolts['MATERIAL DESCRIPTION'].str.contains('1/2"ø'), 0.5,
                   np.where(dfColAssyBolts['MATERIAL DESCRIPTION'].str.contains('5/8"ø'), 0.625,
                   np.where(dfColAssyBolts['MATERIAL DESCRIPTION'].str.contains('3/4"ø'), .75, "OTHER")))
#get a sum of bolts by type and station
dfColAssyBolts.sort_values(by=['DRAWING', 'STRUCTURES','DIA', 'PART DESCRIPTION'], inplace=True)
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
#dfColAssyBolts.to_excel(output_directory + "//DEBUG-NEW-ColAssyNuts&Bolts.xlsx", sheet_name="Sheet 1")
#delete unnecessary column
dfColAssyBolts = dfColAssyBolts.drop('DIA', axis=1)
dfColAssyBolts2 = pd.DataFrame([[''] * len(dfColAssyBolts.columns)], columns=dfColAssyBolts.columns)
# For each grouping Apply insert headers
dfColAssyBolts3 = (dfColAssyBolts.groupby('STRUCTURES', group_keys=False)
        .apply(lambda d: d.append(dfColAssyBolts2))
        .iloc[:-2]
        .reset_index(drop=True))
dfColAssyBolts3.to_excel(output_directory + "//" + projectName + " Col Assy Hardware Order.xlsx", sheet_name="Sheet 1")


#NEW field bolts
#filter for shop bolts and field bolts. filter is whether sheet name contains "CA"
dfFieldBolts = dfNutsAndBolts[dfNutsAndBolts['DRAWING'].str.contains("E", na=False, case=False)].copy(deep=True)
#get a sum of bolts by type and station
dfFieldBolts['DIA'] = "OTHER"
#function for sorting bolts by dia
dfFieldBolts['DIA'] = np.where(dfFieldBolts['MATERIAL DESCRIPTION'].str.contains('1/2"ø'), 0.5,
                   np.where(dfFieldBolts['MATERIAL DESCRIPTION'].str.contains('5/8"ø'), 0.625,
                   np.where(dfFieldBolts['MATERIAL DESCRIPTION'].str.contains('3/4"ø'), .75, "OTHER")))
dfFieldBolts.sort_values(by=['DRAWING', 'STRUCTURES', 'DIA', 'PART DESCRIPTION'], inplace=True)
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
#dfFieldBolts.to_excel(output_directory + "//DEBUG-NEW-ShipLooseNuts&Bolts.xlsx", sheet_name="Sheet 1")
#delete unnecessary column
dfFieldBolts = dfFieldBolts.drop('DIA', axis=1)
#function for adding a blank line after every sheet/station
dfFieldBolts2 = pd.DataFrame([[''] * len(dfFieldBolts.columns)], columns=dfFieldBolts.columns)
# For each grouping Apply insert headers
dfFieldBolts3 = (dfFieldBolts.groupby('STRUCTURES', group_keys=False)
        .apply(lambda d: d.append(dfFieldBolts2))
        .iloc[:-2]
        .reset_index(drop=True))
dfFieldBolts3.to_excel(output_directory + "//" + projectName + " Ship Loose Hardware Order.xlsx", sheet_name="Sheet 1")


#####Misc Hardware#####

#filter out everyhing already covered
dfRemain = df[~df['PART DESCRIPTION'].str.contains("angle*|flat*|plate*|beam*|pipe*|tube*|screw*|bolt*|washer*|nut*|weld*", na=False, case=False)]
#sort by column MATERIAL DESCRIPTION
dfRemain = dfRemain.sort_values('MATERIAL DESCRIPTION')
#save to new excel file
dfRemain.to_excel(output_directory + "//" + projectName + " Misc Hardware.xlsx", sheet_name="Sheet 1")


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
dfClampPl.to_excel(output_directory + "//" + projectName + " Clamp Plates.xlsx", sheet_name="Sheet 1")

