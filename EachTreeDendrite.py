##This script produces a spreadsheet listing the following for each cell analyzed:
##Total number of branches
##Total dendrite length
##Average length of a branch
##Average length of a dendritic tree (all branches from a primary dendrite)
##The length of all dendrites where spines were analyzed
##The numbers of each of the 6 classes of spines (thin, stubby, mushroom, filopodial,
##branched, detached) in Neurolucida Explorer
##The densities of each of the 6 classes of spines in Neurolucida Explorer

##It takes the Neurolucida Explorer spreadsheet "each tree dendrite" as input.

####################################################################
##Section 1: setting up your computer to run this program
##This program is written in Python 2.6-2.7. It also uses the xlwt addon library to make Excel spreadsheets.

##Step 1: Install Python 2.7.3 from here: http://www.python.org/download/releases/2.7.3/
##On the page, select the Windows MSI installer (x86 if you have 32-bit Windows installed,
##x86-64 if you have 64-bit Windows installed.)
##I suggest using the default option, which will install Python to c:/Python27

##Step 2: Install the xlwt library from here: http://pypi.python.org/pypi/xlwt/
##Use the program WinRAR to unzip the files to a directory
##Go to "run" in the start menu and type cmd
##Type cd c:\directory_where_xlwt_was_unzipped_to
##Type setup.py install

##Step 3: Copy this program into the c:/Python27 directory
##You can also put it into another directory that is added to the correct PATH.

####################################################################
##Section 2: file preparation for this program (and related scripts)

##Open Neurolucida Explorer 
##Open Neurolucida file for each cell of interest
##For each cell, select all relevant dendrites and the soma 
##Select all possible analyses that return spreadsheet data.
##Right click and save each spreadsheet as text file with format:
##"casename region cellnumber analysisname.txt"
##For example, "M32-09 lateral 1 cell bodies.txt"

####################################################################
##Section 3: prepare program settings

##The directory where your saved .txt files are stored
directory = "C:\Documents and Settings\Administrator\Desktop\NewGolgiTest"

##List case names, as well as the number of cells analyzed for each case
cases = [["M31-09", 10], ["M32-09", 10], ["M38084", 10], ["R2-11", 10], ["R3-11", 10], ["R4-11", 10], ["R5-11", 10]]

##List the brain regions examined (this can be a single region)
regions = ["lateral", "central"] 

##List the name of the source file, as saved
analysisname = "each tree dendrite"

##Name the output file, this will be placed in the same directory as your input files
outputfile = "eachtreedendritetest.xls"

####################################################################
##Program begins here

#Import modules to handle a tab-delimited text file and produce .xls output
import csv
import xlwt

#function to create list of files
def create_filelist(cases, regions, analysisname):
    filelist = []
    for case in cases:
        for region in regions:
            cellcount = case[1]
            for cell in range(1, cellcount + 1):
                fileappend = case[0] + " " + region + " " + str(cell) + " " + analysisname
                filelist.append(fileappend)            
    return filelist

#module to do spine calculations
def spine_calcs(spinelist, count_location, density_location):
    spine_count = 0
    for entry in spinelist:
        spine_count = spine_count + float(entry)
    sheet1.write(filecount, count_location, spine_count)
    sheet1.write(filecount, density_location, (spine_count / spinebranchlength))
    return

#generate file list
filelist = create_filelist(cases, regions, analysisname)

#set up worksheet to write to
book = xlwt.Workbook(encoding="utf-8")
sheet1 = book.add_sheet("Python Sheet 1")

#create lists to write incoming data to

filecount=1
for filename in filelist:
    path = directory + "\\" + filename + ".txt"
    input_list = [[], [], [], [], [], [], [], [], [], [], []]
    try:
        myfileobj = open(path, "r") 
        csv_read = csv.reader(myfileobj, dialect = csv.excel_tab)      
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
       
    #branchcount
        branchcount = 0
        for branch in input2:
            branchcount = branchcount + float(branch)
        sheet1.write(filecount, 0, filename)
        sheet1.write(filecount, 1, branchcount)
    
    #branch average
        dendritelength = 0
        for branch in input3:
            dendritelength = dendritelength + float(branch)
        branchaverage = dendritelength / branchcount    
        sheet1.write(filecount, 2, dendritelength)
        sheet1.write(filecount, 3, branchaverage)
    
    #count base dendrite number
        dendritenumber = 1
        storebranch = 1
        for branch in input1:
            currbranch = float(branch)      
            if currbranch > storebranch:
                dendritenumber += 1
            storebranch = currbranch
        dendriteaverage = dendritelength / dendritenumber
        sheet1.write(filecount, 4, dendriteaverage)
    
    #sum dendrite volumes
        denvol = 0
        for branchvol in input9:
            denvol = denvol + float(branchvol)
        sheet1.write(filecount, 5, denvol)
    
    #pull all dendrite data when >5 total spines observed
    #part 1: identify dendrites with 5+ spines on them
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
    
    #pull branch length data for marked branches, output spinebranchlength
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
        sheet1.write(filecount, 6, spinebranchlength)
    
    #do all spine calculations and write to sheet using function spine_calcs
        spine_calcs(inputspines, 7, 14)
        spine_calcs(inputthin, 8, 15)
        spine_calcs(inputstub, 9, 16)
        spine_calcs(inputmushroom, 10, 17)
        spine_calcs(inputfilo, 11, 18)
        spine_calcs(inputbranched, 12, 19)
        spine_calcs(inputdetached, 13, 20)
    
    except:
        print filename + ".txt" + " not found"       
    filecount += 1
    
caption0 = "Filename"
sheet1.write(0, 0, caption0)
caption1 = "Total Branch Number"
sheet1.write(0, 1, caption1)
caption2 = "Dendrite Length Total"
sheet1.write(0, 2, caption2)
caption3 = "Branch Length Average"
sheet1.write(0, 3, caption3)
caption4 = "Dendrite Length Average"
sheet1.write(0, 4, caption4)
caption5="Total Dendrite Volume"
sheet1.write(0, 5, caption5)
caption6 = "Spiny Dendrite Length"
sheet1.write(0, 6, caption6)
caption7 = "Total Spine Number"
sheet1.write(0, 7, caption7)
caption8 = "Thin Spine Number"
sheet1.write(0, 8, caption8)
caption9 = "Stubby Spine Number"
sheet1.write(0, 9, caption9)
caption10 = "Mushroom Spine Number"
sheet1.write(0, 10, caption10)
caption11 = "Filopodial Spine Number"
sheet1.write(0, 11, caption11)
caption12 = "Branched Spine Number"
sheet1.write(0, 12, caption12)
caption13 = "Detached Spine Number"
sheet1.write(0, 13, caption13)
caption14 = "Spine Density, spines/um"
sheet1.write(0, 14, caption14)
caption15 = "Thin Spine Density"
sheet1.write(0, 15, caption15)
caption16 = "Stubby Spine Density"
sheet1.write(0, 16, caption16)
caption17 = "Mushroom Spine Density"
sheet1.write(0, 17, caption17)
caption18 = "Filopodial Spine Density"
sheet1.write(0, 18, caption18)
caption19 = "Branched Spine Density"
sheet1.write(0, 19, caption19)
caption20 = "Detached Spine Density"
sheet1.write(0, 20, caption20)

savepath = directory + "\\" + outputfile
book.save(savepath)

