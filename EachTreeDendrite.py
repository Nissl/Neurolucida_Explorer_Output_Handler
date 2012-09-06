# This script produces a spreadsheet listing the following for each cell 
# analyzed:
# Total number of branches
# Total dendrite length
# Average length of a branch
# Average length of a dendritic tree (all branches from a primary dendrite)
# The length of all dendrites where spines were analyzed
# The numbers of each of the 6 classes of spines (thin, stubby, mushroom, 
# filopodial, branched, detached) in Neurolucida Explorer
# The densities of each of the 6 classes of spines in Neurolucida Explorer

# It takes the Neurolucida Explorer spreadsheet "each tree dendrite" as input.

##############################################################################
# Section 1: setting up your computer to run this program
# This program is written in Python 2.6-2.7. 

# Step 1: Install Python 2.7.3 from here: 
# http://www.python.org/download/releases/2.7.3/
# On the page, select the Windows MSI installer (x86 if you have 32-bit 
# Windows installed, x86-64 if you have 64-bit Windows installed.)
# I suggest using the default option, which will install Python to c:/Python27

# Step 2: Copy this program into the c:/Python27 directory
# You can also put it into another directory that is added to the correct 
# PATH.

##############################################################################
# Section 2: file preparation for this program (and related scripts)

# Open Neurolucida Explorer 
# Open Neurolucida file for each cell of interest
# For each cell, select all relevant dendrites and the soma 
# Select all possible analyses that return spreadsheet data.
# Right click and save each spreadsheet as text file with format:
# "casename region cellnumber analysisname.txt"
# For example, "M32-09 lateral 1 cell bodies.txt"

##############################################################################
# Section 3: prepare program settings

# The directory where your saved .txt files are stored
directory = "C:\Documents and Settings\Administrator\Desktop\NewGolgiTest"

# List case names, as well as the number of cells analyzed for each case
cases = [["M31-09", 10], ["M32-09", 10], ["M38084", 10], ["R2-11", 10], 
         ["R3-11", 10], ["R4-11", 10], ["R5-11", 10]]

# List the brain regions examined (this can be a single region)
regions = ["lateral", "central"] 

# List the name of the source file, as saved
analysisname = "each tree dendrite"

# Name the output file, this will be placed in the same directory as your input files
outputfile = "eachtreedendritetestCSV.txt"

##############################################################################
# Program begins here

# Import module to handle a tab-delimited text file
import csv


def create_filelist(cases, regions, analysisname):
    """function to create list of files"""
    filelist = []
    for case in cases:
        for region in regions:
            cellcount = case[1]
            for cell in range(1, cellcount + 1):
                fileappend = (case[0] + " " + region + " " + str(cell) + " " + 
                              analysisname)
                filelist.append(fileappend)            
    return filelist


def spine_count(spinelist, count_location, density_location, output_line):
    """module to do spine calculations"""
    spine_count = 0
    for entry in spinelist:
        spine_count = spine_count + float(entry)
    output_line.append(spine_count)
    return


def spine_density(spinelist, count_location, density_location, output_line):
    """module to do spine calculations"""
    spine_count = 0
    for entry in spinelist:
        spine_count = spine_count + float(entry)
    output_line.append((spine_count / spinebranchlength))
    return


# generate file list
filelist = create_filelist(cases, regions, analysisname)

# create output sheet and add captions
out_path = directory + "\\" + outputfile
output_writer = csv.writer(open(out_path, 'w'), delimiter='\t', quotechar='|',
                           quoting=csv.QUOTE_MINIMAL)
captions = ["Filename", "Total Branch Number", "Dendrite Length Total", 
            "Branch Length Average", "Dendrite Length Average", 
            "Total Dendrite Volume", "Spiny Dendrite Length", 
            "Total Spine Number", "Thin Spine Number", "Stubby Spine Number",
            "Mushroom Spine Number", "Filopodial Spine Number", 
            "Branched Spine Number", "Detached Spine Number", 
            "Spine Density, spines/um", "Thin Spine Density", 
            "Stubby Spine Density", "Mushroom Spine Density", 
            "Filopodial Spine Density", "Branched Spine Density", 
            "Detached Spine Density"]
output_writer.writerow(captions)

# create lists to write incoming data to
for filename in filelist:
    path = directory + "\\" + filename + ".txt"
    input_list = [[], [], [], [], [], [], [], [], [], [], []]
    try:
        myfileobj = open(path, "r") 
        csv_read = csv.reader(myfileobj, dialect=csv.excel_tab)      
        for line in csv_read:
            input_list[0].append(line[0])
            input_list[1].append(line[2])
            input_list[2].append(line[3])
            input_list[3].append(line[9])
            input_list[4].append(line[17])
            input_list[5].append(line[18])
            input_list[6].append(line[19])
            input_list[7].append(line[20])
            input_list[8].append(line[21])
            input_list[9].append(line[22])
            input_list[10].append(line[23])
        input1 = input_list[0][1:]
        input2 = input_list[1][1:]
        input3 = input_list[2][1:]
        input9 = input_list[3][1:]
        inputspines = input_list[4][1:]
        inputthin = input_list[5][1:]
        inputstub = input_list[6][1:]
        inputmushroom = input_list[7][1:]
        inputfilo = input_list[8][1:]
        inputbranched = input_list[9][1:]
        inputdetached = input_list[10][1:]
    
        output_line = []   
    # branchcount
        branchcount = 0
        for branch in input2:
            branchcount = branchcount + float(branch)
        output_line.append(filename)
        output_line.append(branchcount)
    
    # branch average
        dendritelength = 0
        for branch in input3:
            dendritelength = dendritelength + float(branch)
        branchaverage = dendritelength / branchcount    
        output_line.append(dendritelength)
        output_line.append(branchaverage)
    
    # count base dendrite number
        dendritenumber = 1
        storebranch = 1
        for branch in input1:
            currbranch = float(branch)      
            if currbranch > storebranch:
                dendritenumber += 1
            storebranch = currbranch
        dendriteaverage = dendritelength / dendritenumber
        output_line.append(dendriteaverage)
    
    # sum dendrite volumes
        denvol = 0
        for branchvol in input9:
            denvol = denvol + float(branchvol)
        output_line.append(denvol)
    
    # pull all dendrite data when >5 total spines observed
    # part 1: identify dendrites with 5+ spines on them
        testpool = len(inputspines)
        testrun = 0
        listcount = 0
        markedlist = []
        while testrun < testpool:
            validation = float(inputspines[testrun])
            if validation > 5:
                treenumber = input1[testrun]
                if markedlist == []:
                    markedlist.append(treenumber)
                    listcount = 0
                else:
                    repeatcheck = markedlist[listcount]
                    if treenumber == repeatcheck:
                        pass
                    else:
                        markedlist.append(treenumber)
                        listcount = listcount+1
            else:
                pass
            testrun += 1
    
    # pull branch length data for marked branches, output spinebranchlength
        spinebranchlength = 0
        testrun = 0
        while testrun < testpool:
                test = int(input1[testrun])
                comparetest = 0
                comparelength = len(markedlist)
                while comparetest < comparelength:
                    comparemark = int(markedlist[comparetest])
                    if comparemark == test:
                        spinebranchlength = spinebranchlength + float(input3[testrun])
                    comparetest += 1
                testrun += 1
        output_line.append(spinebranchlength)
    
    # do all spine calculations and write to sheet using function spine_calcs
        spine_count(inputspines, 7, 14, output_line)
        spine_count(inputthin, 8, 15, output_line)
        spine_count(inputstub, 9, 16, output_line)
        spine_count(inputmushroom, 10, 17, output_line)
        spine_count(inputfilo, 11, 18, output_line)
        spine_count(inputbranched, 12, 19, output_line)
        spine_count(inputdetached, 13, 20, output_line)
        spine_density(inputspines, 7, 14, output_line)
        spine_density(inputthin, 8, 15, output_line)
        spine_density(inputstub, 9, 16, output_line)
        spine_density(inputmushroom, 10, 17, output_line)
        spine_density(inputfilo, 11, 18, output_line)
        spine_density(inputbranched, 12, 19, output_line)
        spine_density(inputdetached, 13, 20, output_line)
    
        output_writer.writerow(output_line)
    except:
        print filename + ".txt" + " not found"       

    

