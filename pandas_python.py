import pandas as pd
import tkinter as tk
import numpy as np
import pyexcel as p
import sys
import math
from tkinter import *
from tkinter import ttk
from tkinter import filedialog
from fractions import Fraction
from ortools.sat.python import cp_model
import re

#required before we can ask for input file
root = tk.Tk()
root.withdraw()

gui = Tk()
gui.geometry("400x150")
gui.title("SS Mat Order")
gui.columnconfigure(0, weight=3)
gui.columnconfigure(1, weight=1)

folderPath = StringVar()
filePath = StringVar()

class FolderSelect(Frame):
    def __init__(self,parent=None,folderDescription="",**kw):
        Frame.__init__(self,master=parent,**kw)
        self.folderPath = StringVar()
        self.lblName = Label(self, text=folderDescription)
        self.lblName.grid(row=0,column=0, sticky="ew", pady=1)
        self.entPath = Entry(self, textvariable=self.folderPath)
        self.entPath.grid(row=1,column=0, sticky="ew", pady=1)
        self.btnFind = ttk.Button(self, text="Select Folder",command=self.setFolderPath)
        self.btnFind.grid(row=1,column=1, pady=1)
    def setFolderPath(self):
        folder_selected = filedialog.askdirectory()
        self.folderPath.set(folder_selected)
        self.entPath.insert(0,folder_selected)
    @property
    def folder_path(self):
        self.entPath.update()
        return self.folderPath.get()
    
class FileSelect(Frame):
    def __init__(self,parent=None,folderDescription="",**kw):
        Frame.__init__(self,master=parent,**kw)
        self.filePath = StringVar()
        self.lblName = Label(self, text=folderDescription)
        self.lblName.grid(row=0,column=0, sticky="ew", pady=1)
        self.entPath = Entry(self, textvariable=self.filePath)
        self.entPath.grid(row=1,column=0, sticky="ew", pady=1)
        self.btnFind = ttk.Button(self, text="Select File",command=self.setFilePath)
        self.btnFind.grid(row=1,column=1, pady=1)
    def setFilePath(self):
        file_selected = filedialog.askopenfilename()
        self.filePath.set(file_selected)
        self.entPath.insert(0,file_selected)
    @property
    def file_path(self):
        self.entPath.update()
        return self.filePath.get()
        

def doStuff():
    global excel_file
    global output_directory
    excel_file = file1Select.file_path
    output_directory = directory1Select.folder_path
    root.quit()

def endProgram():
    sys.exit()


file1Select = FileSelect(gui,"Excel BOM File:")
file1Select.grid(row=0)

directory1Select = FolderSelect(gui,"Order Files Output Folder:")
directory1Select.grid(row=1)

c = ttk.Button(gui, text="RUN", command=doStuff)
c.grid(row=4,column=0, pady=1)
e = ttk.Button(gui, text="EXIT", command=endProgram)
e.grid(row=4,column=1, pady=1)
gui.mainloop()

if excel_file[len(excel_file)-1] == "s":
        p.save_book_as(file_name=excel_file,
               dest_file_name=excel_file + "x")
        excel_file = excel_file + "x"

#read the excel file's first sheet, set line 1 (2nd line) as header for column names
df = pd.read_excel(excel_file, sheet_name=0, header=[1], skiprows=[2], dtype_backend="pyarrow")

#rename column "ITEM.1" to "QTY"
df.rename(columns = {'ITEM.1':'QTY'}, inplace=True)

#get project name
projectName = df.loc[2]['PROJECT']
#####Angle order#####

#adding some data sanitization. excel will not a '/' in a sheet name
df['STRUCTURES'] = df['STRUCTURES'].str.replace('/','&')

#filter out everyhing but angle only
dfAngle = df[df['PART DESCRIPTION'].str.contains("Angle*", na=False, case=False)]
#filter out specifically flat bar. some slipped through that were 
dfAngle = dfAngle[~dfAngle['PART DESCRIPTION'].str.contains("Flat*", na=False, case=False)]
AnglePoseNestDataFrameSUM = None       

#check if pandas dataframe dfAngle is empty
if not dfAngle.empty:

    #sort by column MATERIAL DESCRIPTION
    dfAngle = dfAngle.sort_values('MATERIAL DESCRIPTION')
    #splitting by structure, "qty req'd" is no longer relevant
    dfAngle = dfAngle.assign(STRUCTURES=dfAngle['STRUCTURES'].astype(str).str.strip().str.split("|")).explode('STRUCTURES').reset_index(drop=True)
    dfAngle = dfAngle.assign(STRUCTURES=dfAngle['STRUCTURES'].astype(str).str.strip())
    #dropping assy and totat. not needed after splitting by structure
    dfAngle = dfAngle.drop('ASSY.', axis=1)
    dfAngle = dfAngle.drop('TOTAL', axis=1)
    #one line per part, 10 qty = 10 lines
    dfAngle = dfAngle.loc[dfAngle.index.repeat(dfAngle['QTY'])].reset_index(drop=True)
    #setting all qty to 1
    dfAngle['QTY'] = 1
    #deleting unnecessary/irrelevant columns
    dfAngle = dfAngle.drop('REV', axis=1)
    dfAngle = dfAngle.drop('SHEET', axis=1)
    dfAngle = dfAngle.drop('MAIN NUMBER', axis=1)
    dfAngle = dfAngle.drop('PART DESCRIPTION', axis=1)
    dfAngle = dfAngle.drop('WIDTH', axis=1)
    dfAngle = dfAngle.drop('WIDTH.1', axis=1)
    dfAngle = dfAngle.drop('GRADE', axis=1)
    dfAngle = dfAngle.drop('WEIGHT', axis=1)
    #making length an interger, makes computer sweat less
    dfAngle['LENGTH.1'] = dfAngle['LENGTH.1'].apply(lambda x: round(x, 3))
    dfAngle['LENGTH.1'] = dfAngle['LENGTH.1'].apply(lambda x: x*1000)
    #adding kerf unless the part is a whole stick
    dfAngle['LENGTH.1'] = dfAngle['LENGTH.1'].apply(lambda x:(x+125) if x<480000 else x)
    #saving to excel file

    AngleCutTicketWorksetDataFrame = []
    AngleNestWorksetDataFrame = []

    def create_data_model_angle():
        data = {}
        # part lengths
        data['weights'] = dfAngleType['LENGTH.1'].astype(int).values.tolist()
        data['items'] = list(range(len(data['weights'])))
        data['bins'] = data['items']
        # stick size
        data['bin_capacity'] = 480000
        data['material'] = dfAngleType.iloc[0, 5]
        data['structures'] = dfAngleType.iloc[0, 8]
        data['drawing'] = dfAngleType.iloc[0, 1]
        return data

    # angle nesting function
    for group, dfAngleType in dfAngle.groupby(['DRAWING', 'MATERIAL DESCRIPTION', 'STRUCTURES']):
        data = create_data_model_angle()

        # Create the CP-SAT model.
        model = cp_model.CpModel()

        # Variables
        # x[i, j] = 1 if item i is packed in bin j.
        x = {}
        for i in data['items']:
            for j in data['bins']:
                x[(i, j)] = model.NewIntVar(0, 1, 'x_%i_%i' % (i, j))

        # y[j] = 1 if bin j is used.
        y = {}
        for j in data['bins']:
            y[j] = model.NewIntVar(0, 1, 'y[%i]' % j)

        # Constraints
        # Each item must be in exactly one bin.
        for i in data['items']:
            model.Add(sum(x[i, j] for j in data['bins']) == 1)

        # The amount packed in each bin cannot exceed its capacity.
        for j in data['bins']:
            model.Add(
                sum(x[(i, j)] * data['weights'][i] for i in data['items']) <= y[j] *
                data['bin_capacity'])

        # Objective: minimize the number of bins used.
        model.Minimize(sum(y[j] for j in data['bins']))

        solver = cp_model.CpSolver()
        solver.parameters.max_time_in_seconds = 180.0
        status = solver.Solve(model)

        if status == cp_model.OPTIMAL or status == cp_model.FEASIBLE:
            num_bins = 0
            bin_usage = 0
            for j in data['bins']:
                if solver.Value(y[j]) == 1:
                    bin_items = []
                    bin_weight = 0
                    for i in data['items']:
                        if solver.Value(x[i, j]) > 0:
                            bin_items.append(i)
                            # stick usage
                            bin_weight += data['weights'][i]
                    if bin_items:
                        # counting number of sticks pulled
                        num_bins += 1
                        # estimating material usage
                        if bin_weight < 360000 and bin_weight > 120000:
                            bin_usage += round(bin_weight / 480000, 2)
                        elif bin_weight > 360000:
                            bin_usage += 1
                        else:
                            bin_usage += 0.25
            # make list of parts
            AngleNestDictionary = {'PROJECT': projectName, 'DRAWING': data['drawing'], 'MATERIAL DESCRIPTION': data['material'], 'ORDER': num_bins, 'USAGE': bin_usage, 'STRUCTURES': data['structures']}
            # list to dataframe
            AngleNestDictionaryDataFrame = pd.DataFrame(data=AngleNestDictionary, index=[0])
            # add parts to overall list
            AngleNestWorksetDataFrame.append(AngleNestDictionaryDataFrame)
            dfAngleTypeSum = dfAngleType.groupby(['PROJECT', 'DRAWING', 'ITEM', 'PART NUMBER', 'MATERIAL DESCRIPTION', 'LENGTH', 'STRUCTURES'])['QTY'].sum(numeric_only=True).reset_index()
            dfAngleTypeSum['ORDER'] = num_bins
            dfAngleTypeSum['USAGE'] = bin_usage
            AngleCutTicketWorksetDataFrame.append(dfAngleTypeSum)
        else:
            # there's either a fatal problem, or there's too many "good" solutions
            print('Angle nesting problem does not have an optimal or feasible solution.')

            
    writerCutTicket = pd.ExcelWriter(output_directory + "//" + projectName + " Anglematic Cut Ticket Data.xlsx")
    #saving angle nesting results 
    if AngleCutTicketWorksetDataFrame:
        AngleCutTicketDataFrame = pd.concat(AngleCutTicketWorksetDataFrame, ignore_index=True)

        for group, dfAngleCutTicket in AngleCutTicketDataFrame.groupby(['DRAWING', 'STRUCTURES']): 
            #sorting by BOM item number first
            dfAngleCutTicket = dfAngleCutTicket.sort_values(by='ITEM')
            #then by material type
            dfAngleCutTicket = dfAngleCutTicket.sort_values(by='MATERIAL DESCRIPTION')
            #filling out cut ticket info, stick size is 40'
            dfAngleCutTicket['SIZE'] = "40'"
            #adding blank column so output can be copy-pasted to cut ticket template
            dfAngleCutTicket['INVENTORY ID'] = None
            #re-sorting columns in correct order
            dfAngleCutTicket = dfAngleCutTicket[['ITEM', 'DRAWING', 'PART NUMBER', 'LENGTH', 'QTY','INVENTORY ID', 'MATERIAL DESCRIPTION', 'USAGE', 'SIZE', 'ORDER', 'STRUCTURES']]
            #adding to excel file, tab name is "sheet name | station"
            dfAngleCutTicket.to_excel(writerCutTicket, sheet_name=dfAngleCutTicket.iloc[0,1] + " | " + dfAngleCutTicket.iloc[0,10])

        #new excel file
        AnglePostNestDataFrame = pd.concat(AngleNestWorksetDataFrame, ignore_index=True)
        #deleting unnessary/irrelevant columns
        AnglePostNestDataFrame = AnglePostNestDataFrame.drop('STRUCTURES', axis=1)
        AnglePostNestDataFrame = AnglePostNestDataFrame.drop('DRAWING', axis=1)
        #combing by material type
        AnglePoseNestDataFrameSUM = AnglePostNestDataFrame.groupby('MATERIAL DESCRIPTION').sum(numeric_only=True).reset_index()
else:
    print("No angle material found in BOM")
#####Flat Bar order#####

#filter out everything but flat bar only
dfFlatBar = df[df['PART DESCRIPTION'].str.contains("Flat*", na=False, case=False)]
FlatBarPostNestDataFrameSUM = None

#check if pandas dataframe dfFlatBar is empty
if not dfFlatBar.empty:

    #sort by column MATERIAL DESCRIPTION
    dfFlatBar = dfFlatBar.sort_values('MATERIAL DESCRIPTION')
    #splitting by structure, "qty req'd" is no longer relevant
    dfFlatBar = dfFlatBar.assign(STRUCTURES=dfFlatBar['STRUCTURES'].astype(str).str.strip().str.split("|")).explode('STRUCTURES').reset_index(drop=True)
    dfFlatBar = dfFlatBar.assign(STRUCTURES=dfFlatBar['STRUCTURES'].astype(str).str.strip())
    #dropping assy and totat. not needed after splitting by structure
    dfFlatBar = dfFlatBar.drop('ASSY.', axis=1)
    dfFlatBar = dfFlatBar.drop('TOTAL', axis=1)
    #one line per part, 10 qty = 10 lines
    dfFlatBar = dfFlatBar.loc[dfFlatBar.index.repeat(dfFlatBar['QTY'])].reset_index(drop=True)
    #setting all qty to 1
    dfFlatBar['QTY'] = 1
    #deleting unnessary/irrelevant columns
    dfFlatBar = dfFlatBar.drop('REV', axis=1)
    dfFlatBar = dfFlatBar.drop('SHEET', axis=1)
    dfFlatBar = dfFlatBar.drop('MAIN NUMBER', axis=1)
    dfFlatBar = dfFlatBar.drop('PART DESCRIPTION', axis=1)
    dfFlatBar = dfFlatBar.drop('WIDTH', axis=1)
    dfFlatBar = dfFlatBar.drop('WIDTH.1', axis=1)
    dfFlatBar = dfFlatBar.drop('GRADE', axis=1)
    dfFlatBar = dfFlatBar.drop('WEIGHT', axis=1)
    #making length an interger, makes computer sweat less
    dfFlatBar['LENGTH.1'] = dfFlatBar['LENGTH.1'].apply(lambda x: round(x, 3))
    dfFlatBar['LENGTH.1'] = dfFlatBar['LENGTH.1'].apply(lambda x: x*1000)
    #adding kerf unless the part is a whole stick (should not happen on flat bar anyways)
    dfFlatBar['LENGTH.1'] = dfFlatBar['LENGTH.1'].apply(lambda x:(x+125) if x<240000 else x)

    #prepping excel sheet for FlatBar order after nesting
    FlatBarCutTicketWorksetDataFrame = []
    FlatBarNestWorksetDataFrame = []

    def create_data_model_FlatBar():
        data = {}
        #part lengths
        data['weights'] = dfFlatBarType['LENGTH.1'].astype(int).values.tolist()
        data['items'] = list(range(len(data['weights'])))
        data['bins'] = data['items']
        #stick size
        data['bin_capacity'] = 240000
        data['material'] = dfFlatBarType.iloc[0,5]
        data['structures'] = dfFlatBarType.iloc[0,8]
        data['drawing'] = dfFlatBarType.iloc[0,2]
        return data

    #FlatBar nesting fuction
    for group, dfFlatBarType in dfFlatBar.groupby(['DRAWING', 'MATERIAL DESCRIPTION', 'STRUCTURES']):    
        
        data = create_data_model_FlatBar()

        # Create the CP-SAT model.
        model = cp_model.CpModel()

        # Variables
        # x[i, j] = 1 if item i is packed in bin j.
        x = {}
        for i in data['items']:
            for j in data['bins']:
                x[(i, j)] = model.NewIntVar(0, 1, 'x_%i_%i' % (i, j))

        # y[j] = 1 if bin j is used.
        y = {}
        for j in data['bins']:
            y[j] = model.NewIntVar(0, 1, 'y[%i]' % j)

        # Constraints
        # Each item must be in exactly one bin.
        for i in data['items']:
            model.Add(sum(x[i, j] for j in data['bins']) == 1)

        # The amount packed in each bin cannot exceed its capacity.
        for j in data['bins']:
            model.Add(
                sum(x[(i, j)] * data['weights'][i] for i in data['items']) <= y[j] *
                data['bin_capacity'])

        # Objective: minimize the number of bins used.
        model.Minimize(sum(y[j] for j in data['bins']))

        # Create the solver and solve the model.
        solver = cp_model.CpSolver()
        solver.parameters.max_time_in_seconds = 180.0
        status = solver.Solve(model)

        #letting the solver give us either a perfect solution or if there's multiple good solutions, just giving one of those
        if status == cp_model.OPTIMAL or status == cp_model.FEASIBLE:
            #zero out to start
            num_bins = 0
            bin_usage = 0
            for j in data['bins']:
                if solver.Value(y[j]) == 1:
                    bin_items = []
                    bin_weight = 0
                    for i in data['items']:
                        if solver.Value(x[i, j]) > 0:
                            bin_items.append(i)
                            #stick usage
                            bin_weight += data['weights'][i]
                    if bin_items:
                        #counting number of sticks pulled
                        num_bins += 1
                        #estimating material usage
                        if bin_weight < 180000 and bin_weight > 60000:
                            bin_usage += round(bin_weight/240000, 2)
                        elif bin_weight > 180000:
                            bin_usage += 1
                        else:
                            bin_usage += 0.25
            #make list of parts
            FlatBarNestDictionary = {'PROJECT': projectName, 'DRAWING': data['drawing'], 'MATERIAL DESCRIPTION': data['material'], 'ORDER':num_bins, 'USAGE':bin_usage, 'STRUCTURES': data['structures']}
            #list to dataframe
            FlatBarNestDictionaryDataFrame = pd.DataFrame(data=FlatBarNestDictionary, index=[0])
            #add parts to overall list
            FlatBarNestWorksetDataFrame.append(FlatBarNestDictionaryDataFrame)
            dfFlatBarTypeSum = dfFlatBarType.groupby(['PROJECT', 'DRAWING', 'ITEM', 'PART NUMBER', 'MATERIAL DESCRIPTION', 'LENGTH', 'STRUCTURES'])['QTY'].sum(numeric_only=True).reset_index()
            dfFlatBarTypeSum['ORDER'] = num_bins
            dfFlatBarTypeSum['USAGE'] = bin_usage
            FlatBarCutTicketWorksetDataFrame.append(dfFlatBarTypeSum)
        else:
            #there's either a fatal problem, or there's too many "good" solutions
            print('Flat bar nesting problem does not have an optimal or feasible solution.')

    if FlatBarCutTicketWorksetDataFrame:
        #saving FlatBar nesting results   
        FlatBarCutTicketDataFrame = pd.concat(FlatBarCutTicketWorksetDataFrame, ignore_index=True)

        #each page of cut ticket being written to excel file
        for group, dfFlatBarCutTicket in FlatBarCutTicketDataFrame.groupby(['DRAWING', 'STRUCTURES']): 
            #sorting by BOM item number first
            dfFlatBarCutTicket = dfFlatBarCutTicket.sort_values(by='ITEM')
            #then by material type
            dfFlatBarCutTicket = dfFlatBarCutTicket.sort_values(by='MATERIAL DESCRIPTION')
            #filling out cut ticket info, stick size is 20'
            dfFlatBarCutTicket['SIZE'] = "20'"
            #adding blank column so output can be copy-pasted to cut ticket template
            dfFlatBarCutTicket['INVENTORY ID'] = None
            #re-sorting columns in correct order
            dfFlatBarCutTicket = dfFlatBarCutTicket[['ITEM', 'DRAWING', 'PART NUMBER', 'LENGTH', 'QTY','INVENTORY ID', 'MATERIAL DESCRIPTION', 'USAGE', 'SIZE', 'ORDER', 'STRUCTURES']]
            #adding to excel file, tab name is "sheet name | station"
            dfFlatBarCutTicket.to_excel(writerCutTicket, sheet_name=dfFlatBarCutTicket.iloc[0,1] + " | " + dfFlatBarCutTicket.iloc[0,10])

        #saving excel file
        writerCutTicket.close()

        FlatBarPostNestDataFrame = pd.concat(FlatBarNestWorksetDataFrame, ignore_index=True)
        #deleting unnessary/irrelevant columns
        FlatBarPostNestDataFrame = FlatBarPostNestDataFrame.drop('STRUCTURES', axis=1)
        FlatBarPostNestDataFrame = FlatBarPostNestDataFrame.drop('DRAWING', axis=1)
        #combining by material type
        FlatBarPostNestDataFrameSUM= FlatBarPostNestDataFrame.groupby('MATERIAL DESCRIPTION').sum(numeric_only=True).reset_index()
else:
    print("No flat bar material found in BOM")

#combined anglematic nested order
if AnglePoseNestDataFrameSUM is not None and FlatBarPostNestDataFrameSUM is not None:
    
    dfAnglematicNestedInput = [AnglePoseNestDataFrameSUM,FlatBarPostNestDataFrameSUM]
    dfAnglematicNested = pd.concat(dfAnglematicNestedInput)
    #saving to excel file
    dfAnglematicNested.to_excel(output_directory + "//" + projectName + " Anglematic Order Nested.xlsx", sheet_name="Sheet 1")

elif AnglePoseNestDataFrameSUM is not None:
    
    #saving to excel file
    AnglePoseNestDataFrameSUM.to_excel(output_directory + "//" + projectName + " Anglematic Order Nested.xlsx", sheet_name="Sheet 1")

elif FlatBarPostNestDataFrameSUM is not None:
    
    #saving to excel file
    FlatBarPostNestDataFrameSUM.to_excel(output_directory + "//" + projectName + " Anglematic Order Nested.xlsx", sheet_name="Sheet 1")


#####Misc Material#####

#filter out everyhing but misc linear only
dfMisc = df[df['PART DESCRIPTION'].str.contains("w-beam*|s-beam*|pipe*|tube*|s-tee*", na=False, case=False)]

if not dfMisc.empty:
    #sort by column MATERIAL DESCRIPTION
    dfMisc = dfMisc.sort_values('MATERIAL DESCRIPTION')
    #column sum = (total qty) x (length in inches)
    dfMisc['SUM'] = dfMisc.apply(lambda row:(row['TOTAL'] * row['LENGTH.1']),axis=1)
    #save to new excel file
    dfMisc.to_excel(output_directory + "//" + projectName + " Misc Material.xlsx", sheet_name="Sheet 1")
else:
    print("No misc material found in BOM")

#prepping data for sign bracket nesting
#grabbing anything that includes "w-beam" or "s-beam" in the part description and has SB in the part name
dfSignBracketNest = dfMisc[dfMisc['PART DESCRIPTION'].str.contains("w-beam*|s-beam*", na=False, case=False)]
dfSignBracketNest = dfSignBracketNest[dfSignBracketNest['PART NUMBER'].str.contains("SB*", na=False, case=False)]

if not dfSignBracketNest.empty:
    #splitting by structure, "qty req'd" is no longer relevant
    dfSignBracketNest = dfSignBracketNest.assign(STRUCTURES=dfSignBracketNest['STRUCTURES'].astype(str).str.strip().str.split("|")).explode('STRUCTURES').reset_index(drop=True)
    dfSignBracketNest = dfSignBracketNest.assign(STRUCTURES=dfSignBracketNest['STRUCTURES'].astype(str).str.strip())
    #dropping assy and totat. not needed after splitting by structure
    dfSignBracketNest = dfSignBracketNest.drop('ASSY.', axis=1)
    dfSignBracketNest = dfSignBracketNest.drop('TOTAL', axis=1)
    #making length an interger, makes computer sweat less
    dfSignBracketNest['LENGTH.1'] = dfSignBracketNest['LENGTH.1'].apply(lambda x: round(x, 3))
    dfSignBracketNest['LENGTH.1'] = dfSignBracketNest['LENGTH.1'].apply(lambda x: x*1000)
    #adding kerf unless the part is a whole stick (should not happen on sign brackets anyways)
    dfSignBracketNest['LENGTH.1'] = dfSignBracketNest['LENGTH.1'].apply(lambda x:(x+125) if x<480000 else x)
    #one line per part, 10 qty = 10 lines
    dfSignBracketNest = dfSignBracketNest.loc[dfSignBracketNest.index.repeat(dfSignBracketNest['QTY'])].reset_index(drop=True)
    #setting all qty to 1
    dfSignBracketNest['QTY'] = 1
    #deleting unnecessary/irrelevant columns
    dfSignBracketNest = dfSignBracketNest.drop('WIDTH', axis=1)
    dfSignBracketNest = dfSignBracketNest.drop('WIDTH.1', axis=1)
    dfSignBracketNest = dfSignBracketNest.drop('WEIGHT', axis=1)
    dfSignBracketNest = dfSignBracketNest.drop('REV', axis=1)
    dfSignBracketNest = dfSignBracketNest.drop('SHEET', axis=1)

    #prepping excel sheet for FlatBar order after nesting
    SignBracketCutTicketWorksetDataFrame = []
    SignBracketNestWorksetDataFrame = []

    def create_data_model_sign_bracket():
        data = {}
        #part lengths
        data['weights'] = dfSignBracketType['LENGTH.1'].astype(int).values.tolist()
        data['items'] = list(range(len(data['weights'])))
        data['bins'] = data['items']
        #stick size
        data['bin_capacity'] = 480000
        data['material'] = dfSignBracketType.iloc[0,7]
        data['structures'] = dfSignBracketType.iloc[0,11]
        data['drawing'] = dfSignBracketType.iloc[0,1]
        return data

    #angle nesting fuction
    for group, dfSignBracketType in dfSignBracketNest.groupby(['PROJECT', 'MATERIAL DESCRIPTION']):
        data = create_data_model_sign_bracket()

        # Create the CP-SAT model.
        model = cp_model.CpModel()

        # Variables
        # x[i, j] = 1 if item i is packed in bin j.
        x = {}
        for i in data['items']:
            for j in data['bins']:
                x[(i, j)] = model.NewIntVar(0, 1, 'x_%i_%i' % (i, j))

        # y[j] = 1 if bin j is used.
        y = {}
        for j in data['bins']:
            y[j] = model.NewIntVar(0, 1, 'y[%i]' % j)

        # Constraints
        # Each item must be in exactly one bin.
        for i in data['items']:
            model.Add(sum(x[i, j] for j in data['bins']) == 1)

        # The amount packed in each bin cannot exceed its capacity.
        for j in data['bins']:
            model.Add(
                sum(x[(i, j)] * data['weights'][i] for i in data['items']) <= y[j] *
                data['bin_capacity'])

        # Objective: minimize the number of bins used.
        model.Minimize(sum(y[j] for j in data['bins']))

        # Create the solver and solve the model.
        solver = cp_model.CpSolver()
        solver.parameters.max_time_in_seconds = 180.0
        status = solver.Solve(model)

        # Letting the solver give us either a perfect solution or if there's multiple good solutions, just giving one of those
        if status == cp_model.OPTIMAL or status == cp_model.FEASIBLE:
            # Zero out to start
            num_bins = 0
            bin_usage = 0
            for j in data['bins']:
                if solver.Value(y[j]) == 1:
                    bin_items = []
                    bin_weight = 0
                    for i in data['items']:
                        if solver.Value(x[i, j]) > 0:
                            bin_items.append(i)
                            # Stick usage
                            bin_weight += data['weights'][i]
                            SignBracketNestDictionary = {'PROJECT': projectName, 'PART': dfSignBracketType.iloc[i,4], 'QTY': 1, 'GRADE': dfSignBracketType.iloc[i,10], 'MATERIAL DESCRIPTION': data['material'], 'LENGTH': dfSignBracketType.iloc[i,8], 'STATION': data['structures'], 'NESTED LENGTH': (data['weights'][i])/10000, 'STICK': j}
                            # List of parts to dataframe
                            SignBracketNestDictionaryDataFrame = pd.DataFrame(data=SignBracketNestDictionary, index=[0])
                            # Add the parts to the overall list
                            SignBracketNestWorksetDataFrame.append(SignBracketNestDictionaryDataFrame)
                    if bin_items:
                        # Counting number of sticks pulled
                        num_bins += 1
                        # Estimating material usage
                        if bin_weight < 360000 and bin_weight > 120000:
                            bin_usage += round(bin_weight/480000, 2)
                        elif bin_weight > 360000:
                            bin_usage += 1
                        else:
                            bin_usage += 0.25
        else:
            # There's either a fatal problem, or there's too many "good" solutions
            print('Sign bracket nesting problem does not have an optimal or feasible solution.')

    SignBracketPoseNestDataFrame = pd.concat(SignBracketNestWorksetDataFrame, ignore_index=True)
    #combining multiple quantities of the same part on the same stick
    SignBracketPoseNestDataFrame = SignBracketPoseNestDataFrame.groupby(['PROJECT', 'PART', 'GRADE', 'MATERIAL DESCRIPTION', 'LENGTH', 'STATION', 'NESTED LENGTH', 'STICK'])['QTY'].sum(numeric_only=True).reset_index()
    #adding cutting instruction for cut ticket
    SignBracketPoseNestDataFrame['SHOP NOTES'] = "CUT " + SignBracketPoseNestDataFrame['QTY'].apply(str) + " PCS @ " + SignBracketPoseNestDataFrame['LENGTH']
    #sorting by what stick the part is nested on
    SignBracketPoseNestDataFrame.sort_values(by=['MATERIAL DESCRIPTION', 'STICK'], inplace=True)
    #adding blank columns so output can be copy-pasted to cut ticket template
    SignBracketPoseNestDataFrame['STOCK CODE'] = None
    SignBracketPoseNestDataFrame['RAW MAT QTY'] = None
    SignBracketPoseNestDataFrame['HEAT NUMBER'] = None
    SignBracketPoseNestDataFrame['LOCATION'] = None
    #sorting columns in correct order
    SignBracketPoseNestDataFrame = SignBracketPoseNestDataFrame[['PROJECT', 'PART', 'QTY', 'STOCK CODE', 'GRADE', 'MATERIAL DESCRIPTION', 'RAW MAT QTY', 'HEAT NUMBER', 'LOCATION', 'SHOP NOTES', 'LENGTH', 'STATION', 'NESTED LENGTH', 'STICK']]
    #save to excel file
    SignBracketPoseNestDataFrame.to_excel(output_directory + "//" + projectName + " Sign Brackets Nested.xlsx", sheet_name="Sheet 1")

else: 
    print("No sign brackets found in BOM")

#prepping data for s-tee nesting
#grabbing anything that includes "w-beam" or "s-beam" in the part description and has SB in the part name
dfSteeNest = dfMisc[dfMisc['PART DESCRIPTION'].str.contains("s-tee*", na=False, case=False)]
dfSteeNest = dfSteeNest[dfSteeNest['PART NUMBER'].str.contains("SB*", na=False, case=False)]

if not dfSteeNest.empty:
    #splitting by structure, "qty req'd" is no longer relevant
    dfSteeNest = dfSteeNest.assign(STRUCTURES=dfSteeNest['STRUCTURES'].astype(str).str.strip().str.split("|")).explode('STRUCTURES').reset_index(drop=True)
    dfSteeNest = dfSteeNest.assign(STRUCTURES=dfSteeNest['STRUCTURES'].astype(str).str.strip())
    #dropping assy and totat. not needed after splitting by structure
    dfSteeNest = dfSteeNest.drop('ASSY.', axis=1)
    dfSteeNest = dfSteeNest.drop('TOTAL', axis=1)
    #making length an interger, makes computer sweat less
    dfSteeNest['LENGTH.1'] = dfSteeNest['LENGTH.1'].apply(lambda x: round(x, 3))
    dfSteeNest['LENGTH.1'] = dfSteeNest['LENGTH.1'].apply(lambda x: x*1000)
    #S-Tees get 2 per length on s-beams
    dfSteeNest['LENGTH.1'] = dfSteeNest['LENGTH.1'].apply(lambda x:(x/2))
    #adding kerf unless the part is a whole stick (should not happen on s-tees anyways)
    dfSteeNest['LENGTH.1'] = dfSteeNest['LENGTH.1'].apply(lambda x:(x+125) if x<480000 else x)
    #one line per part, 10 qty = 10 lines
    dfSteeNest = dfSteeNest.loc[dfSteeNest.index.repeat(dfSteeNest['QTY'])].reset_index(drop=True)
    #setting all qty to 1
    dfSteeNest['QTY'] = 1
    #deleting unnecessary/irrelevant columns
    dfSteeNest = dfSteeNest.drop('WIDTH', axis=1)
    dfSteeNest = dfSteeNest.drop('WIDTH.1', axis=1)
    dfSteeNest = dfSteeNest.drop('WEIGHT', axis=1)
    dfSteeNest = dfSteeNest.drop('REV', axis=1)
    dfSteeNest = dfSteeNest.drop('SHEET', axis=1)

    #prepping excel sheet for FlatBar order after nesting
    SteeCutTicketWorksetDataFrame = []
    SteeNestWorksetDataFrame = []

    def create_data_model_s_tee():
        data = {}
        #part lengths
        data['weights'] = dfSteeType['LENGTH.1'].astype(int).values.tolist()
        data['items'] = list(range(len(data['weights'])))
        data['bins'] = data['items']
        #stick size
        data['bin_capacity'] = 480000
        data['material'] = dfSteeType.iloc[0,7]
        data['structures'] = dfSteeType.iloc[0,11]
        data['drawing'] = dfSteeType.iloc[0,1]
        return data

    #angle nesting fuction
    for group, dfSteeType in dfSteeNest.groupby(['PROJECT', 'LENGTH']):    
        
        data = create_data_model_s_tee()

        # Create the CP-SAT model.
        model = cp_model.CpModel()

        # Variables
        # x[i, j] = 1 if item i is packed in bin j.
        x = {}
        for i in data['items']:
            for j in data['bins']:
                x[(i, j)] = model.NewIntVar(0, 1, 'x_%i_%i' % (i, j))

        # y[j] = 1 if bin j is used.
        y = {}
        for j in data['bins']:
            y[j] = model.NewIntVar(0, 1, 'y[%i]' % j)

        # Constraints
        # Each item must be in exactly one bin.
        for i in data['items']:
            model.Add(sum(x[i, j] for j in data['bins']) == 1)

        # The amount packed in each bin cannot exceed its capacity.
        for j in data['bins']:
            model.Add(
                sum(x[(i, j)] * data['weights'][i] for i in data['items']) <= y[j] *
                data['bin_capacity'])

        # Objective: minimize the number of bins used.
        model.Minimize(sum(y[j] for j in data['bins']))

        # Create the solver and solve the model.
        solver = cp_model.CpSolver()
        solver.parameters.max_time_in_seconds = 180.0
        status = solver.Solve(model)

        #letting the solver give us either a perfect solution or if there's multiple good solutions, just giving one of those
        if status == cp_model.OPTIMAL or status == cp_model.FEASIBLE:
            #zero out to start
            num_bins = 0
            bin_usage = 0
            for j in data['bins']:
                if solver.Value(y[j]) == 1:
                    bin_items = []
                    bin_weight = 0
                    for i in data['items']:
                        if solver.Value(x[i, j]) > 0:
                            bin_items.append(i)
                            #stick usage
                            bin_weight += data['weights'][i]
                            SteeNestDictionary = {'PROJECT': projectName, 'PART': dfSteeType.iloc[i,4], 'QTY': 1, 'GRADE': dfSteeType.iloc[i,10], 'MATERIAL DESCRIPTION': data['material'], 'LENGTH': dfSteeType.iloc[i,8], 'STATION': data['structures'], 'NESTED LENGTH': (data['weights'][i])/10000, 'STICK': j}
                            #list of parts to dataframe
                            SteeNestDictionaryDataFrame = pd.DataFrame(data=SteeNestDictionary, index=[0])
                            #add the parts to the overall list
                            SteeNestWorksetDataFrame.append(SteeNestDictionaryDataFrame)
                    if bin_items:
                        #counting number of sticks pulled
                        num_bins += 1
                        #estimating material usage
                        if bin_weight < 360000 and bin_weight > 120000:
                            bin_usage += round(bin_weight/480000, 2)
                        elif bin_weight > 360000:
                            bin_usage += 1
                        else:
                            bin_usage += 0.25
        else:
            #there's either a fatal problem, or there's too many "good" solutions
            print('S-Tee nesting problem does not have an optimal or feasible solution.')

    SteePostNestDataFrame = pd.concat(SteeNestWorksetDataFrame, ignore_index=True)
    #combining multiple quantities of the same part on the same stick
    SteePostNestDataFrame = SteePostNestDataFrame.groupby(['PROJECT', 'PART', 'GRADE', 'MATERIAL DESCRIPTION', 'LENGTH', 'STATION', 'NESTED LENGTH', 'STICK'])['QTY'].sum(numeric_only=True).reset_index()
    #adding cutting instruction for cut ticket
    SteePostNestDataFrame['SHOP NOTES'] = SteePostNestDataFrame['QTY'].apply((lambda row:(math.ceil(row/2))))
    SteePostNestDataFrame['SHOP NOTES'] = "CUT " + SteePostNestDataFrame['SHOP NOTES'].apply(str) + " PCS @ " + SteePostNestDataFrame['LENGTH'] + " SPLIT IN HALF TO GET " + (SteePostNestDataFrame['SHOP NOTES']*2).apply(str)
    #sorting by what stick the part is nested on
    SteePostNestDataFrame.sort_values(by=['MATERIAL DESCRIPTION', 'STICK'], inplace=True)
    #adding blank columns so output can be copy-pasted to cut ticket template
    SteePostNestDataFrame['STOCK CODE'] = None
    SteePostNestDataFrame['RAW MAT QTY'] = None
    SteePostNestDataFrame['HEAT NUMBER'] = None
    SteePostNestDataFrame['LOCATION'] = None
    #sorting columns in correct order
    SteePostNestDataFrame = SteePostNestDataFrame[['PROJECT', 'PART', 'QTY', 'STOCK CODE', 'GRADE', 'MATERIAL DESCRIPTION', 'RAW MAT QTY', 'HEAT NUMBER', 'LOCATION', 'SHOP NOTES', 'LENGTH', 'STATION', 'NESTED LENGTH', 'STICK']]
    #save to excel file
    SteePostNestDataFrame.to_excel(output_directory + "//" + projectName + " S-Tees Nested.xlsx", sheet_name="Sheet 1")

else:
    print("No s-tees found in BOM")

#####NUTS AND BOLTS#####

#filter out everyhing but nuts, bolts, and washers only
dfNutsAndBolts = df[df['PART DESCRIPTION'].str.contains("nut*|bolt*|washer*", na=False, case=False)].copy(deep=True)
#explodes entries with multiple stations to one line per station
dfNutsAndBolts = dfNutsAndBolts.assign(STRUCTURES=dfNutsAndBolts['STRUCTURES'].astype(str).str.strip().str.split("|")).explode('STRUCTURES').reset_index(drop=True)
dfNutsAndBolts = dfNutsAndBolts.assign(STRUCTURES=dfNutsAndBolts['STRUCTURES'].astype(str).str.strip())
#deleting assy and total columns to avoid confusion now structures are one per line
dfNutsAndBolts = dfNutsAndBolts.drop('ASSY.', axis=1)
dfNutsAndBolts = dfNutsAndBolts.drop('TOTAL', axis=1)
#deleting columns that don't help ordering nust and bolts
dfNutsAndBolts = dfNutsAndBolts.drop('ITEM', axis=1)
dfNutsAndBolts = dfNutsAndBolts.drop('REV', axis=1)
dfNutsAndBolts = dfNutsAndBolts.drop('WEIGHT', axis=1)

#verification hardware
dfNutsAndBoltsVerif = dfNutsAndBolts.copy(deep=True)
#deleting 3/8" dia and 1/2" dia hardware. we don't need to order these for verification
dfNutsAndBoltsVerif = dfNutsAndBoltsVerif[~dfNutsAndBoltsVerif['MATERIAL DESCRIPTION'].str.contains('1/2"ø', na=False, case=False)].copy(deep=True)
dfNutsAndBoltsVerif = dfNutsAndBoltsVerif[~dfNutsAndBoltsVerif['MATERIAL DESCRIPTION'].str.contains('3/8"ø', na=False, case=False)].copy(deep=True)
dfNutsAndBoltsVerif = dfNutsAndBoltsVerif[~dfNutsAndBoltsVerif['MATERIAL DESCRIPTION'].str.contains('USS', na=False, case=False)].copy(deep=True)
dfNutsAndBoltsVerif = dfNutsAndBoltsVerif[~dfNutsAndBoltsVerif['MATERIAL DESCRIPTION'].str.contains('NEOPRENE', na=False, case=False)].copy(deep=True)
dfNutsAndBoltsVerif = dfNutsAndBoltsVerif[~dfNutsAndBoltsVerif['MATERIAL DESCRIPTION'].str.contains('OUTSOURCED', na=False, case=False)].copy(deep=True)
#delete irrelevant columns
dfNutsAndBoltsVerif = dfNutsAndBoltsVerif.drop('DRAWING', axis=1)
dfNutsAndBoltsVerif = dfNutsAndBoltsVerif.drop('QTY', axis=1)
dfNutsAndBoltsVerif = dfNutsAndBoltsVerif.drop('STRUCTURES', axis=1)
dfNutsAndBoltsVerif = dfNutsAndBoltsVerif.drop('WIDTH.1', axis=1)
dfNutsAndBoltsVerif = dfNutsAndBoltsVerif.drop('LENGTH.1', axis=1)
#deleting washer and nut grades because our drafters don't care they're wrong
dfNutsAndBoltsVerif.loc[dfNutsAndBoltsVerif['PART DESCRIPTION'] == 'Washer', 'GRADE'] = ' '
dfNutsAndBoltsVerif.loc[dfNutsAndBoltsVerif['PART DESCRIPTION'] == 'Nut', 'GRADE'] = ' '
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
#filter for shop bolts and field bolts. filter is whether sheet name contains an E or CA
dfShopBolts2 = dfNutsAndBolts[~dfNutsAndBolts['DRAWING'].str.contains("E", na=False, case=False)].copy(deep=True)
dfShopBolts2 = dfShopBolts2[~dfShopBolts2['DRAWING'].str.contains("CA", na=False, case=False)]
#if there's no shop bolts, skip this section
if not dfShopBolts2.empty:
    #function for sorting bolts by dia
    dfShopBolts2['DIA'] = np.where(dfShopBolts2['MATERIAL DESCRIPTION'].str.contains('1/2"ø'), 0.5,
                    np.where(dfShopBolts2['MATERIAL DESCRIPTION'].str.contains('5/8"ø'), 0.625,
                    np.where(dfShopBolts2['MATERIAL DESCRIPTION'].str.contains('3/4"ø'), .75, "OTHER")))
    #get a sum of bolts by type and station
    dfShopBolts2.sort_values(by=['DRAWING', 'STRUCTURES','DIA', 'PART DESCRIPTION'], inplace=True)
    #add sheet name to station name column'
    dfShopBolts2['STRUCTURES'] = dfShopBolts2['DRAWING'].astype(str) + ' | ' + dfShopBolts2['STRUCTURES'].astype(str)
    #delete unnecessary columns
    dfShopBolts2 = dfShopBolts2.drop('DRAWING', axis=1)
    dfShopBolts2 = dfShopBolts2.drop('SHEET', axis=1)
    dfShopBolts2 = dfShopBolts2.drop('MAIN NUMBER', axis=1)
    dfShopBolts2 = dfShopBolts2.drop('PART NUMBER', axis=1)
    dfShopBolts2 = dfShopBolts2.drop('WIDTH', axis=1)
    dfShopBolts2 = dfShopBolts2.drop('WIDTH.1', axis=1)
    dfShopBolts2 = dfShopBolts2.drop('LENGTH', axis=1)
    dfShopBolts2 = dfShopBolts2.drop('LENGTH.1', axis=1)
    #deleting washer and nut grades because our drafters don't care they're wrong
    dfShopBolts2.loc[dfShopBolts2['PART DESCRIPTION'] == 'Washer', 'GRADE'] = ' '
    dfShopBolts2.loc[dfShopBolts2['PART DESCRIPTION'] == 'Nut', 'GRADE'] = ' '
    #add 8% or +5 to shop bolts, whichever is more
    dfShopBolts2['ORDER'] = dfShopBolts2['QTY'].apply(lambda row:(row*1.08) if row>62 else (row+5))
    #round up
    dfShopBolts2['ORDER'] = dfShopBolts2['ORDER'].apply(np.ceil)
    #delete unnecessary columns
    dfShopBolts2 = dfShopBolts2.drop('PART DESCRIPTION', axis=1)
    dfShopBolts2 = dfShopBolts2.drop('QTY', axis=1)
    #adding "use" column so vendor can mark buckets accordingly
    dfShopBolts2['USE'] = "ASSY"
    #delete unnecessary column
    dfShopBolts2 = dfShopBolts2.drop('DIA', axis=1)
    #function for adding blank lines after every station/sheet
    dfShopBolts3 = pd.DataFrame([[''] * len(dfShopBolts2.columns)], columns=dfShopBolts2.columns)
    # For each grouping Apply insert headers
    dfShopBolts4 = (dfShopBolts2.groupby('STRUCTURES', group_keys=False)
            .apply(lambda d: pd.concat([d, dfShopBolts3]))
            .iloc[:-2]
            .reset_index(drop=True))
    #adding last line back on, not sure why it gets deleted
    dfShopBolts4 = pd.concat([dfShopBolts4, dfShopBolts2.tail(1)], ignore_index=True)
    #saving to excel file
    dfShopBolts4.to_excel(output_directory + "//" + projectName + " Assy Hardware Order.xlsx", sheet_name="Sheet 1")
else:
    print("No shop bolts found in BOM")


#NEW col assy bolts
#filter for shop bolts and field bolts. filter is whether sheet name contains "CA"
dfColAssyBolts = dfNutsAndBolts[dfNutsAndBolts['DRAWING'].str.contains("CA", na=False, case=False)].copy(deep=True)
#if there's no col assy bolts, skip this section
if not dfColAssyBolts.empty:
    #function for sorting bolts by dia
    dfColAssyBolts['DIA'] = np.where(dfColAssyBolts['MATERIAL DESCRIPTION'].str.contains('1/2"ø'), 0.5,
                    np.where(dfColAssyBolts['MATERIAL DESCRIPTION'].str.contains('5/8"ø'), 0.625,
                    np.where(dfColAssyBolts['MATERIAL DESCRIPTION'].str.contains('3/4"ø'), .75, "OTHER")))
    #get a sum of bolts by type and station
    dfColAssyBolts.sort_values(by=['DRAWING', 'STRUCTURES','DIA', 'PART DESCRIPTION'], inplace=True)
    #add sheet name to station name column'
    dfColAssyBolts['STRUCTURES'] = dfColAssyBolts['DRAWING'].astype(str) + ' | ' + dfColAssyBolts['STRUCTURES'].astype(str)
    #delete unnecessary columns
    dfColAssyBolts = dfColAssyBolts.drop('DRAWING', axis=1)
    dfColAssyBolts = dfColAssyBolts.drop('SHEET', axis=1)
    dfColAssyBolts = dfColAssyBolts.drop('MAIN NUMBER', axis=1)
    dfColAssyBolts = dfColAssyBolts.drop('PART NUMBER', axis=1)
    dfColAssyBolts = dfColAssyBolts.drop('WIDTH', axis=1)
    dfColAssyBolts = dfColAssyBolts.drop('WIDTH.1', axis=1)
    dfColAssyBolts = dfColAssyBolts.drop('LENGTH', axis=1)
    dfColAssyBolts = dfColAssyBolts.drop('LENGTH.1', axis=1)
    #deleting washer and nut grades because our drafters don't care they're wrong
    dfColAssyBolts.loc[dfColAssyBolts['PART DESCRIPTION'] == 'Washer', 'GRADE'] = ' '
    dfColAssyBolts.loc[dfColAssyBolts['PART DESCRIPTION'] == 'Nut', 'GRADE'] = ' '
    #add 8% or +5 to shop bolts, whichever is more
    dfColAssyBolts['ORDER'] = dfColAssyBolts['QTY'].apply(lambda row:(row*1.08) if row>62 else (row+5))
    #round up
    dfColAssyBolts['ORDER'] = dfColAssyBolts['ORDER'].apply(np.ceil)
    #delete unnecessary columns
    dfColAssyBolts = dfColAssyBolts.drop('PART DESCRIPTION', axis=1)
    dfColAssyBolts = dfColAssyBolts.drop('QTY', axis=1)
    #adding "use" column so vendor can mark buckets accordingly
    dfColAssyBolts['USE'] = "COLUMN ASSY"
    #delete unnecessary column
    dfColAssyBolts = dfColAssyBolts.drop('DIA', axis=1)
    #function for adding blank lines after every sheet/station
    dfColAssyBolts2 = pd.DataFrame([[''] * len(dfColAssyBolts.columns)], columns=dfColAssyBolts.columns)
    dfColAssyBolts3 = (dfColAssyBolts.groupby('STRUCTURES', group_keys=False)
            .apply(lambda d: pd.concat([d, dfColAssyBolts2]))
            .iloc[:-2]
            .reset_index(drop=True))
    #adding last line back on, not sure why it gets deleted
    dfColAssyBolts3 = pd.concat([dfColAssyBolts3, dfColAssyBolts.tail(1)], ignore_index=True)
    #saving to excel file
    dfColAssyBolts3.to_excel(output_directory + "//" + projectName + " Col Assy Hardware Order.xlsx", sheet_name="Sheet 1")
else:
    print("No column assembly bolts found in BOM")


#NEW field bolts
#filter for shop bolts and field bolts. filter is whether sheet name contains "CA"
dfFieldBolts = dfNutsAndBolts[dfNutsAndBolts['DRAWING'].str.contains("E", na=False, case=False)].copy(deep=True)
#if there's no field bolts, skip this section
if not dfFieldBolts.empty:
    #get a sum of bolts by type and station
    dfFieldBolts['DIA'] = "OTHER"
    #function for sorting bolts by dia
    dfFieldBolts['DIA'] = np.where(dfFieldBolts['MATERIAL DESCRIPTION'].str.contains('1/2"ø'), 0.5,
                    np.where(dfFieldBolts['MATERIAL DESCRIPTION'].str.contains('5/8"ø'), 0.625,
                    np.where(dfFieldBolts['MATERIAL DESCRIPTION'].str.contains('3/4"ø'), .75, "OTHER")))
    dfFieldBolts.sort_values(by=['DRAWING', 'STRUCTURES', 'DIA', 'PART DESCRIPTION'], inplace=True)
    #delete unnecessary columns
    dfFieldBolts = dfFieldBolts.drop('SHEET', axis=1)
    dfFieldBolts = dfFieldBolts.drop('MAIN NUMBER', axis=1)
    dfFieldBolts = dfFieldBolts.drop('PART NUMBER', axis=1)
    dfFieldBolts = dfFieldBolts.drop('WIDTH', axis=1)
    dfFieldBolts = dfFieldBolts.drop('WIDTH.1', axis=1)
    dfFieldBolts = dfFieldBolts.drop('LENGTH', axis=1)
    dfFieldBolts = dfFieldBolts.drop('LENGTH.1', axis=1)
    #deleting washer and nut grades because our drafters don't care they're wrong
    dfFieldBolts.loc[dfFieldBolts['PART DESCRIPTION'] == 'Washer', 'GRADE'] = ' '
    dfFieldBolts.loc[dfFieldBolts['PART DESCRIPTION'] == 'Nut', 'GRADE'] = ' '
    #add 2 to each bolt count
    dfFieldBolts['ORDER'] = dfFieldBolts.apply(lambda row:(row['QTY'] + 2),axis=1)
    #delete unnecessary columns
    dfFieldBolts = dfFieldBolts.drop('PART DESCRIPTION', axis=1)
    dfFieldBolts = dfFieldBolts.drop('QTY', axis=1)
    #adding "use" column so vendor can mark buckets accordingly
    dfFieldBolts['USE'] = "SHIP LOOSE"
    #delete unnecessary column
    dfFieldBolts = dfFieldBolts.drop('DIA', axis=1)
    #function for adding a blank line after every sheet/station
    dfFieldBolts2 = pd.DataFrame([[''] * len(dfFieldBolts.columns)], columns=dfFieldBolts.columns)
    # For each grouping Apply insert headers
    dfFieldBolts3 = (dfFieldBolts.groupby('STRUCTURES', group_keys=False)
            .apply(lambda d: pd.concat([d,dfFieldBolts2]))
            .iloc[:-2]
            .reset_index(drop=True))
    #adding last line back on, not sure why it gets deleted
    dfFieldBolts3 = pd.concat([dfFieldBolts3, dfFieldBolts.tail(1)], ignore_index=True)
    dfFieldBolts3['STRUCTURES'] = dfFieldBolts3['DRAWING'].astype(str) + ' | ' + dfFieldBolts3['STRUCTURES'].astype(str)
    dfFieldBolts3 = dfFieldBolts3.drop('DRAWING', axis=1)
    #saving to excel file
    dfFieldBolts3.to_excel(output_directory + "//" + projectName + " Ship Loose Hardware Order.xlsx", sheet_name="Sheet 1")
else:
    print("No field bolts found in BOM")


#####Misc Hardware#####

#filter out everyhing already covered
dfRemain = df[~df['PART DESCRIPTION'].str.contains("angle*|flat*|plate*|beam*|pipe*|tube*|screw*|bolt*|washer*|nut*|weld*|s-tee*", na=False, case=False)]
#if dfRemain is empty, skip this section
if not dfRemain.empty:
    dfRemain = dfRemain.dropna(subset=['MATERIAL DESCRIPTION'])
    #sort by column MATERIAL DESCRIPTION
    dfRemain = dfRemain.sort_values('MATERIAL DESCRIPTION')
    #save to new excel file
    dfRemain.to_excel(output_directory + "//" + projectName + " Misc Hardware.xlsx", sheet_name="Sheet 1")
else:
    print("Nothing remaning found in BOM not caught by other filters")

dodecCheat = pd.DataFrame(np.array([
    ['PL 1/4"', 15.4375, 48, 60],
    ['PL 1/4"', 19, 48, 72],
    ['PL 1/4"', 23.1875, 48, 84],
    ['PL 1/4"', 29, 48, 96],
    ['PL 5/16"', 15.4375, 48, 60],
    ['PL 5/16"', 19, 48, 72],
    ['PL 5/16"', 23.1875, 48, 84],
    ['PL 5/16"', 29, 48, 96],
    ['PL 3/8"', 17.375, 48, 60],
    ['PL 3/8"', 29, 48, 96],
    ['PL 7/16"', 29, 48, 96],
    ['PL 1/2"', 23, 48, 84],
    ['PL 1/2"', 29, 48, 96],]),
    columns=['MATERIAL DESCRIPTION', 'WIDTH.1', 'HALF STRAP', 'FULL STRAP'])

dfDodec = df[df['PART DESCRIPTION'].str.contains("Plate*", na=False, case=False)].copy(deep=True)
if not dfDodec.empty:
    
    dfDodec = dfDodec[dfDodec['PART NUMBER'].str.contains("CA*a", na=False, case=False)]
    dfDodec = dfDodec[dfDodec['GRADE'].str.contains("A572 GR 50", na=False, case=False)]
    dfDodec = dfDodec[dfDodec['ITEM'] == 1]
    dfDodec['WIDTH.1'] = dfDodec['WIDTH.1'].astype(float)
    dodecCheat['WIDTH.1'] = dodecCheat['WIDTH.1'].astype(float)
    dfDodec = pd.merge(left=dfDodec, right=dodecCheat, how='left', left_on=['MATERIAL DESCRIPTION', 'WIDTH.1'], right_on=['MATERIAL DESCRIPTION', 'WIDTH.1'])
    dfDodec['PLATEWIDTH'] = dfDodec['FULL STRAP'].astype(str) + " | " + dfDodec['HALF STRAP'].astype(str)
    dfDodec = dfDodec.assign(PLATEWIDTH=dfDodec['PLATEWIDTH'].astype(str).str.strip().str.split("|")).explode('PLATEWIDTH').reset_index(drop=True)
    dfDodec = dfDodec.assign(PLATEWIDTH=dfDodec['PLATEWIDTH'].astype(str).str.strip())
    dfDodec = dfDodec.drop('HALF STRAP', axis=1)
    dfDodec = dfDodec.drop('FULL STRAP', axis=1)
    dfDodec.loc[(dfDodec['PLATEWIDTH'].str.contains("48", na=False, case=False)), 'PART NUMBER'] = "OPTIONAL"
    dfDodec.loc[(dfDodec['PLATEWIDTH'].str.contains("48", na=False, case=False)), 'TOTAL'] = dfDodec['TOTAL'].astype(int) * 2
    dfDodec = dfDodec.drop('DRAWING', axis=1)
    dfDodec = dfDodec.drop('REV', axis=1)
    dfDodec = dfDodec.drop('MAIN NUMBER', axis=1)
    dfDodec = dfDodec.drop('ITEM', axis=1)
    dfDodec = dfDodec.drop('SHEET', axis=1)
    dfDodec = dfDodec.drop('WIDTH', axis=1)
    dfDodec = dfDodec.drop('LENGTH', axis=1)
    dfDodec = dfDodec.drop('QTY', axis=1)
    dfDodec = dfDodec.drop('ASSY.', axis=1)
    dfDodec = dfDodec.drop('PART DESCRIPTION', axis=1)
    dfDodec.rename(columns = {'PLATEWIDTH':'WIDTH'}, inplace=True)
    dfDodec['LENGTH.1'] = np.ceil(dfDodec['LENGTH.1']) + 5
    dfDodec.rename(columns = {'LENGTH.1':'LENGTH'}, inplace=True)
    dfDodec = dfDodec.drop('WIDTH.1', axis=1)
    dfDodec = dfDodec.drop('STRUCTURES', axis=1)
    dfDodec.rename(columns = {'TOTAL':'QUANTITY'}, inplace=True)
    dfDodec.rename(columns = {'MATERIAL DESCRIPTION':'THICKNESS'}, inplace=True)
    # Extract fraction from string in df['THICKNESS'] column
    dfDodec['THICKNESS'] = dfDodec['THICKNESS'].apply(lambda x: re.findall(r'\d+/\d+', x)[0] if re.findall(r'\d+/\d+', x) else x)
    # dfDodec['THICKNESS'] = dfDodec.apply(lambda row:(row.astype(str).split(" ",1)[1].split('"',1)[0]),axis=1)
    # dfDodec['WEIGHT EACH'] = dfDodec['']
    # dfDodec = dfDodec['PROJECT', 'QUANTITY', 'PART NUMBER', 'GRADE', 'THICKNESS', 'WIDTH', 'LENGTH']
    dfDodec.to_excel(output_directory + "//" + projectName + " Dodec Plates.xlsx", sheet_name="Sheet 1")
    

#filter out everything but clamp plates
dfClampPl = df[df['PART NUMBER'].str.contains("CPS*", na=False, case=False)].copy(deep=True)
if not dfClampPl.empty:
    #adding offset to clamp plate length, done by CPS name
    dfClampPl['LENGTH.1'] = dfClampPl.apply(lambda row:(int((row['PART NUMBER'])[-2:])/16)+row['LENGTH.1'],axis=1)
    #splitting by structure, "qty req'd" is no longer relevant
    dfClampPl = dfClampPl.assign(STRUCTURES=dfClampPl['STRUCTURES'].astype(str).str.strip().str.split("|")).explode('STRUCTURES').reset_index(drop=True)
    dfClampPl = dfClampPl.assign(STRUCTURES=dfClampPl['STRUCTURES'].astype(str).str.strip())
    #one line per part, 10 qty = 10 lines
    dfClampPl = dfClampPl.loc[dfClampPl.index.repeat(dfClampPl['QTY'])].reset_index(drop=True)
    #setting all quantities to 1
    dfClampPl['QTY'] = 1
    #making length an interger, makes computer sweat less
    dfClampPl['LENGTH.1'] = dfClampPl['LENGTH.1'].apply(lambda x: round(x, 3))
    dfClampPl['LENGTH.1'] = dfClampPl['LENGTH.1'].apply(lambda x: x*1000)
    #adding kerf unless the part is a whole stick (should not happen on clamp plates anyways)
    dfClampPl['LENGTH.1'] = dfClampPl['LENGTH.1'].apply(lambda x:(x+125) if x<240000 else x)
    #sort by part number column
    dfClampPl = dfClampPl.sort_values('PART NUMBER')
    #delete unnecessary columns
    dfClampPl = dfClampPl.drop('REV', axis=1)
    dfClampPl = dfClampPl.drop('WEIGHT', axis=1)
    dfClampPl = dfClampPl.drop('SHEET', axis=1)
    dfClampPl = dfClampPl.drop('WIDTH', axis=1)
    dfClampPl = dfClampPl.drop('WIDTH.1', axis=1)
    dfClampPl = dfClampPl.drop('ASSY.', axis=1)
    dfClampPl = dfClampPl.drop('TOTAL', axis=1)
    dfClampPl = dfClampPl.drop('GRADE', axis=1)
    #save to new excel file
    dfClampPl.to_excel(output_directory + "//" + projectName + " Clamp Plates.xlsx", sheet_name="Sheet 1")

    #prepping excel sheet for FlatBar order after nesting
    ClampPlatetCutTicketWorksetDataFrame = []
    ClampPlateNestWorksetDataFrame = []

    def create_data_model_clamp_pl():
        data = {}
        #part lengths
        data['weights'] = dfClampPlateType['LENGTH.1'].astype(int).values.tolist()
        data['items'] = list(range(len(data['weights'])))
        data['bins'] = data['items']
        #stick size
        data['bin_capacity'] = 240000
        data['material'] = dfClampPlateType.iloc[0,7]
        data['structures'] = dfClampPlateType.iloc[0,10]
        data['drawing'] = dfClampPlateType.iloc[0,1]
        return data

    #angle nesting fuction
    for group, dfClampPlateType in dfClampPl.groupby(['PROJECT', 'MATERIAL DESCRIPTION']):    
        
        data = create_data_model_clamp_pl()

        # Create the CP-SAT model.
        model = cp_model.CpModel()
    
        # Variables
        # x[i, j] = 1 if item i is packed in bin j.
        x = {}
        for i in data['items']:
            for j in data['bins']:
                x[(i, j)] = model.NewIntVar(0, 1, 'x_%i_%i' % (i, j))

        # y[j] = 1 if bin j is used.
        y = {}
        for j in data['bins']:
            y[j] = model.NewIntVar(0, 1, 'y[%i]' % j)

        # Constraints
        # Each item must be in exactly one bin.
        for i in data['items']:
            model.Add(sum(x[i, j] for j in data['bins']) == 1)

        # The amount packed in each bin cannot exceed its capacity.
        for j in data['bins']:
            model.Add(
                sum(x[(i, j)] * data['weights'][i] for i in data['items']) <= y[j] *
                data['bin_capacity'])

        # Objective: minimize the number of bins used.
        model.Minimize(sum(y[j] for j in data['bins']))

        solver = cp_model.CpSolver()
        solver.parameters.max_time_in_seconds = 180.0
        status = solver.Solve(model)

        #letting the solver give us either a perfect solution or if there's multiple good solutions, just giving one of those
        if status == cp_model.OPTIMAL or status == cp_model.FEASIBLE:
            #zero out to start
            num_bins = 0
            bin_usage = 0
            for j in data['bins']:
                if solver.Value(y[j]) == 1:
                    bin_items = []
                    bin_weight = 0
                    for i in data['items']:
                        if solver.Value(x[i, j]) > 0:
                            bin_items.append(i)
                            #stick usage
                            bin_weight += data['weights'][i]
                            #make list of parts
                            ClampPlateNestDictionary = {'PROJECT': projectName, 'PART': dfClampPlateType.iloc[i,4], 'MATERIAL DESCRIPTION': data['material'], 'LENGTH': (data['weights'][i])/10000, 'QTY': 1, 'STICK': j}
                            #list of parts to dataframe
                            ClampPlateNestDictionaryDataFrame = pd.DataFrame(data=ClampPlateNestDictionary, index=[0])
                            #add the parts to the overall list
                            ClampPlateNestWorksetDataFrame.append(ClampPlateNestDictionaryDataFrame)
                    if bin_items:
                        #counting number of sticks pulled
                        num_bins += 1
                        #estimating material usage
                        if bin_weight < 180000 and bin_weight > 60000:
                            bin_usage += round(bin_weight/240000, 2)
                        elif bin_weight > 180000:
                            bin_usage += 1
                        else:
                            bin_usage += 0.25
        else:
            #there's either a fatal problem, or there's too many "good" solutions
            print('Clamp plate nesting problem does not have an optimal or feasible solution.')

    ClampPlatePoseNestDataFrame = pd.concat(ClampPlateNestWorksetDataFrame, ignore_index=True)
    #combining same parts on same stick
    ClampPlatePoseNestDataFrame = ClampPlatePoseNestDataFrame.groupby(['PROJECT', 'PART', 'MATERIAL DESCRIPTION', 'LENGTH', 'STICK'])['QTY'].sum(numeric_only=True).reset_index()
    #sorting to group by stick
    ClampPlatePoseNestDataFrame.sort_values(by=['MATERIAL DESCRIPTION', 'STICK'], inplace=True)
    #saving to excel file
    ClampPlatePoseNestDataFrame.to_excel(output_directory + "//" + projectName + " Clamp Plates Nested.xlsx", sheet_name="Sheet 1")

else:
    print("No clamp plates found in BOM")


#creating bill of lading for galvanizer
dfGalvBOL = df[~df['PART DESCRIPTION'].str.contains("nut*|bolt*|washer*|stainless*|aluminum*", na=False, case=False)].copy(deep=True)
dfGalvBOL = dfGalvBOL[~dfGalvBOL['GRADE'].str.contains("durometer*", na=False, case=False)].copy(deep=True)
dfGalvBOL = dfGalvBOL.assign(STRUCTURES=dfGalvBOL['STRUCTURES'].astype(str).str.strip().str.split("|")).explode('STRUCTURES').reset_index(drop=True)
dfGalvBOL = dfGalvBOL.assign(STRUCTURES=dfGalvBOL['STRUCTURES'].astype(str).str.strip())
dfGalvBOL = dfGalvBOL.drop('ASSY.', axis=1)
dfGalvBOL = dfGalvBOL.drop('TOTAL', axis=1)
dfGalvBOL = dfGalvBOL.dropna(subset=['PART NUMBER'])
#supposed to catch all column weldments and call them column weldments
dfGalvBOL.loc[(dfGalvBOL['MAIN NUMBER'].str.contains("CA*", na=False, case=False)) & (dfGalvBOL['PART NUMBER'].str.contains("CA.*[aAB]", na=False)), 'MATERIAL DESCRIPTION'] = "COLUMN WELDMENT"
#supposed to catch all WF parts that need to be galvanized seperately without catching dodec baseplates in that net
dfGalvBOL.loc[(dfGalvBOL['MAIN NUMBER'].str.contains("CA*", na=False, case=False)) & (dfGalvBOL['MAIN NUMBER'].str[-1].str.contains("[0-9]", na=False)) & (dfGalvBOL['PART NUMBER'].str[-1].str.contains("[A-Z]", na=False)) & (~dfGalvBOL['PART NUMBER'].str[0:2].str.contains("bp", na=False)), 'MAIN NUMBER'] = dfGalvBOL['PART NUMBER']
#supposed to catch all sign bracket weldments and name them instead of a main number that doesnt match to anything
dfGalvBOL.loc[(dfGalvBOL['MAIN NUMBER'].str.contains("SB*", na=False, case=False)) & (dfGalvBOL['MAIN NUMBER'].str.strip().str[-1].str.contains("[0-9]", na=False)) & (dfGalvBOL['PART NUMBER'].str.contains("SB*", na=False, case=False)), 'MAIN NUMBER'] = dfGalvBOL['PART NUMBER']
#supposed to catch all the truss parts that need to be galvanized seperately without catching strut weldment parts
dfGalvBOL.loc[(dfGalvBOL['MAIN NUMBER'].str.contains("TA*", na=False, case=False)) & (dfGalvBOL['MAIN NUMBER'].str.strip().str[-1].str.contains("[0-9]", na=False)) & (dfGalvBOL['PART NUMBER'].str.contains("T*", na=False, case=False)) & (dfGalvBOL['PART NUMBER'].str.strip().str[-1].str.contains("[A-Z]", na=False)), 'MAIN NUMBER'] = dfGalvBOL['PART NUMBER']
dfGalvBOL['MATERIAL DESCRIPTION'] = dfGalvBOL['MATERIAL DESCRIPTION'].astype(str) + ' x ' + dfGalvBOL['LENGTH'].astype(str)
#supposed to catch hand hole covers and name them accordingly
dfGalvBOL.loc[dfGalvBOL['MATERIAL DESCRIPTION'].eq("PL 1/8\" x 0'-7 1/2\"") & (dfGalvBOL['PART NUMBER'].str.contains("CA*c*", na=False, case=False)), 'MATERIAL DESCRIPTION'] = "HAND HOLE COVER"

dfGalBOLWorkset = []

for group, dfMainBOL in dfGalvBOL.groupby(['PROJECT', 'MAIN NUMBER', 'STRUCTURES']): 

    if (dfMainBOL['MATERIAL DESCRIPTION'].str.contains("weldment*", na=False, case=False)).any():
        #make weight of each part of the weldment the weight of the sum of all parts
        dfMainBOL['WEIGHT'] = (dfMainBOL['WEIGHT'].sum()) / dfMainBOL['QTY']
        #supposed to catch hand hole covers and change the weight to 1 qty
        dfMainBOL.loc[(dfMainBOL['MATERIAL DESCRIPTION'].str.contains("HAND", na=False, case=False)) & (dfMainBOL['PART NUMBER'].str.contains("CA*c*", na=False, case=False)), 'WEIGHT'] = 1.5
        #drops all lines that do not contain "weldment" or "hand"
        dfMainBOL = dfMainBOL[(dfMainBOL['MATERIAL DESCRIPTION'].str.contains("weldment*|hand*", na=False, case=False))]
        
        #if there's no hand hole cover, drop any line that contains "column weldment" but part number does not match main number
        if ~(dfMainBOL['MATERIAL DESCRIPTION'].str.contains("hand*", na=False, case=False)).any():
            dfMainBOL = dfMainBOL[~(((dfMainBOL['MATERIAL DESCRIPTION'].str.contains("column weldment*", na=False, case=False))) & ((dfMainBOL['PART NUMBER'].astype(str) != dfMainBOL['MAIN NUMBER'].astype(str))))]
   
    else:
        dfMainBOL['WEIGHT'] = dfMainBOL['WEIGHT'].copy(deep=True) / dfMainBOL['QTY'].copy(deep=True)
   
    dfGalBOLWorkset.append(dfMainBOL)

writerGalvBOL = pd.ExcelWriter(output_directory + "//" + projectName + " Galv BOL.xlsx")

if dfGalBOLWorkset:
    dfGalBOLWorksetOutput = pd.concat(dfGalBOLWorkset, ignore_index=True)        
    for group, dfStationBOL in dfGalBOLWorksetOutput.groupby(['PROJECT', 'STRUCTURES']): 
        #re-sorting columns in correct order
        dfStationBOL = dfStationBOL[['MAIN NUMBER', 'QTY', 'PART NUMBER', 'MATERIAL DESCRIPTION', 'WEIGHT', 'STRUCTURES']]
        dfStationBOL.to_excel(writerGalvBOL, sheet_name=dfStationBOL.iloc[0,5])

writerGalvBOL.close()


dfShipTicket = df[~df['PART DESCRIPTION'].str.contains("nut*|bolt*|washer*", na=False, case=False)].copy(deep=True)
dfShipTicket = dfShipTicket.assign(STRUCTURES=dfShipTicket['STRUCTURES'].astype(str).str.strip().str.split("|")).explode('STRUCTURES').reset_index(drop=True)
dfShipTicket = dfShipTicket.assign(STRUCTURES=dfShipTicket['STRUCTURES'].astype(str).str.strip())
dfShipTicket = dfShipTicket.drop('ASSY.', axis=1)
dfShipTicket = dfShipTicket.drop('TOTAL', axis=1)
dfShipTicket = dfShipTicket.dropna(subset=['PART NUMBER'])

#supposed to catch all column weldments and call them column weldments
dfShipTicket.loc[(dfShipTicket['MAIN NUMBER'].str.contains("CA*", na=False, case=False)) & (dfShipTicket['PART NUMBER'].str.contains("CA.*[aAB]", na=False)), 'MATERIAL DESCRIPTION'] = "COLUMN WELDMENT"
dfShipTicket.loc[(dfShipTicket['MAIN NUMBER'].str.contains("CA*", na=False, case=False)) & (dfShipTicket['PART NUMBER'].str.contains("CA.*a", na=False)) & (dfShipTicket['GRADE'].str.contains("A572 GR 50", na=False, case=False)) & (dfShipTicket['PART DESCRIPTION'].str.contains("Plate", na=False, case=False)), 'PART NUMBER'] = dfShipTicket['MAIN NUMBER'] + "A"
#supposed to catch all sign bracket weldments and name them instead of a main number that doesnt match to anything
dfShipTicket.loc[(dfShipTicket['MAIN NUMBER'].str.contains("SB*", na=False, case=False)) & (dfShipTicket['MAIN NUMBER'].str.strip().str[-1].str.contains("[0-9]", na=False)) & (dfShipTicket['PART NUMBER'].str.contains("SB*", na=False, case=False)), 'MAIN NUMBER'] = dfShipTicket['PART NUMBER']
dfShipTicket.loc[(dfShipTicket['MAIN NUMBER'].str.contains("TA*", na=False, case=False)) & (dfShipTicket['MAIN NUMBER'].str.strip().str[-1].str.contains("[0-9]", na=False)) & (dfShipTicket['MATERIAL DESCRIPTION'].str.contains("STRUT*", na=False, case=False)), 'MAIN NUMBER'] = dfShipTicket['PART NUMBER']
dfShipTicket['MATERIAL DESCRIPTION'] = dfShipTicket['MATERIAL DESCRIPTION'].astype(str) + ' x ' + dfShipTicket['LENGTH'].astype(str)

dfShipTicketWorkset = []

for group, dfShip in dfShipTicket.groupby(['PROJECT', 'MAIN NUMBER', 'STRUCTURES']): 
    if (dfShip['MATERIAL DESCRIPTION'].str.contains("weldment*", na=False, case=False)).any():
        #make weight of each part of the weldment the weight of the sum of all parts
        dfShip['WEIGHT'] = (dfShip['WEIGHT'].sum()) / dfShip['QTY']
        #drops all lines that do not contain "weldment"
        dfShip = dfShip[(dfShip['MATERIAL DESCRIPTION'].str.contains("weldment*", na=False, case=False))]
        #if there's no hand hole cover, drop any line that contains "column weldment" but part number does not match main number
        if (dfShip['MATERIAL DESCRIPTION'].str.contains("column weldment*", na=False, case=False)).any():
            dfShip = dfShip[(((dfShip['MATERIAL DESCRIPTION'].str.contains("column weldment*", na=False, case=False))) & ((dfShip['PART NUMBER'].str.contains("CA.*A", na=False))))]
    
    elif (dfShip['MAIN NUMBER'].str.contains("CA.*[A-Z]", na=False, case=False)).any():
        dfShip = dfShip[(((dfShip['MATERIAL DESCRIPTION'].str.contains("column weldment*", na=False, case=False))) & ((dfShip['PART NUMBER'].str.contains("CA.*A", na=False))))]
    
    elif (dfShip['MAIN NUMBER'].str.contains("TA*", na=False, case=False)).any():
        dfShip = dfShip[(dfShip['PART NUMBER'].str.contains("T*", na=False, case=False)) & (dfShip['GRADE'].str.contains("A572 GR 50", na=False, case=False)) & (dfShip['PART NUMBER'].str.strip().str[-2:].str.contains("A|AR|S|T", na=False))]
        dfShip.loc[(dfShip['PART NUMBER'].str.strip().str[-2:].str.contains("AR?", na=False)), 'MATERIAL DESCRIPTION'] = "TRUSS ASSY x " + dfShip['LENGTH'].astype(str) + ' x XX" SQ'
        dfShip.loc[(dfShip['PART NUMBER'].str.strip().str[-1].str.contains("S", na=False)) & (dfShip['MATERIAL DESCRIPTION'].str.contains("L*", na=False)), 'MATERIAL DESCRIPTION'] = "SPLICE ANGLE x " + dfShip['LENGTH'].astype(str)
        dfShip.loc[(dfShip['PART NUMBER'].str.strip().str[-1].str.contains("S", na=False)) & (dfShip['MATERIAL DESCRIPTION'].str.contains("FB*", na=False)), 'MATERIAL DESCRIPTION'] = "SPLICE PLATE x " + dfShip['LENGTH'].astype(str)
        dfShip.loc[(dfShip['PART NUMBER'].str.strip().str[-1].str.contains("S", na=False)) & (dfShip['MATERIAL DESCRIPTION'].str.contains("PL*", na=False)), 'MATERIAL DESCRIPTION'] = "SPLICE PLATE x " + dfShip['LENGTH'].astype(str)
        dfShip.loc[(dfShip['PART NUMBER'].str.strip().str[-1].str.contains("T", na=False)) & (dfShip['MATERIAL DESCRIPTION'].str.contains("FB*", na=False)), 'MATERIAL DESCRIPTION'] = "SPLICE PLATE x " + dfShip['LENGTH'].astype(str)
        dfShip.loc[(dfShip['PART NUMBER'].str.strip().str[-1].str.contains("T", na=False)) & (dfShip['MATERIAL DESCRIPTION'].str.contains("PL*", na=False)), 'MATERIAL DESCRIPTION'] = "SPLICE PLATE x " + dfShip['LENGTH'].astype(str)
        dfShip.loc[(dfShip['MATERIAL DESCRIPTION'].str.contains("TRUSS ASSY*", na=False, case=False)), 'WEIGHT'] = ((((dfShip['WEIGHT']*4)*1.26) + (dfShip['WEIGHT']*4)) * 0.06) + (dfShip['WEIGHT']*4) + ((dfShip['WEIGHT']*4)*1.26) + 400 
        dfShip.loc[:, 'WEIGHT'] = dfShip['WEIGHT'].copy(deep=True) / dfShip['QTY'].copy(deep=True)

    elif (dfShip['MAIN NUMBER'].str.contains("CPS*", na=False, case=False)).any():
        dfShip['OFFSET'] = dfShip['PART NUMBER'].str.strip().str[-2:].astype(int)
        dfShip.loc[dfShip['PART NUMBER'].str.contains("CPS*", na=False, case=False), 'MATERIAL DESCRIPTION'] = dfShip.apply(lambda row: "CLAMP PLATE, " + str(Fraction(int(row['OFFSET']),16)) + '" OFFSET', axis=1)    
    else:
        dfShip.loc[:, 'WEIGHT'] = dfShip['WEIGHT'].copy(deep=True) / dfShip['QTY'].copy(deep=True)
   
    dfShipTicketWorkset.append(dfShip)

writerShipTicket = pd.ExcelWriter(output_directory + "//" + projectName + " Shipping Ticket.xlsx")

if dfShipTicketWorkset:
    dfShipTicketWorksetOutput = pd.concat(dfShipTicketWorkset, ignore_index=True)
    dfShipTicketWorksetOutput = dfShipTicketWorksetOutput[['PROJECT', 'MAIN NUMBER', 'QTY', 'PART NUMBER', 'MATERIAL DESCRIPTION', 'WEIGHT', 'STRUCTURES', 'GRADE']]
    for group, dfStationBOL in dfShipTicketWorksetOutput.groupby(['PROJECT', 'STRUCTURES']): 
       
        for group, dfFieldBoltStation in dfFieldBolts.groupby(['PROJECT', 'STRUCTURES']):
            dfFieldBoltStation.rename(columns = {'ORDER':'QTY'}, inplace=True)
            dfFieldBoltStation = dfFieldBoltStation.drop('USE', axis=1)
           
            if dfFieldBoltStation.iloc[0,4] in dfStationBOL.iloc[0,6]:
                dfStationBOL = pd.concat([dfStationBOL, dfFieldBoltStation], ignore_index=True)
        dfStationBOL.loc[(dfStationBOL['MATERIAL DESCRIPTION'].str.contains("BOLT*", na=False)), 'MATERIAL DESCRIPTION'] = dfStationBOL['MATERIAL DESCRIPTION'].astype(str) + " " + dfStationBOL['GRADE'].astype(str)
        dfStationBOL = dfStationBOL[['PROJECT', 'MAIN NUMBER', 'QTY', 'PART NUMBER', 'MATERIAL DESCRIPTION', 'WEIGHT', 'STRUCTURES']]
        dfStationBOL.to_excel(writerShipTicket, sheet_name=dfStationBOL.iloc[0,6])

writerShipTicket.close()