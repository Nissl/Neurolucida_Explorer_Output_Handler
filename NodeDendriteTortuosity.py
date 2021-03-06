# This script produces a spreadsheet listing dendrite tortuosity for each
# cell analyzed. Tortuousity is taken for each branch and weighted by branch 
# length. It takes the Neurolucida Explorer spreadsheet "node dendrite" as 
# input.

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
analysisname = "node dendrite"

# Name the output file, this will be placed in the same directory as your 
# input files
outputfile = "nodedendritetestCSV.txt"

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


# generate file list
filelist = create_filelist(cases, regions, analysisname)

# create output sheet and add captions
out_path = directory + "\\" + outputfile
output_writer = csv.writer(open(out_path, 'w'), delimiter='\t', quotechar='|',
                           quoting=csv.QUOTE_MINIMAL)
captions = ["Filename", "Dendrite Tortuosity"]
output_writer.writerow(captions)

for filename in filelist:
    path = directory + "\\" + filename + ".txt"
    try: 
        myfileobj = open(path, "r")     
        csv_read = csv.reader(myfileobj, dialect=csv.excel_tab) 
        tortuosity = []
        distance = []
        for line in csv_read:
            distance.append(line[2])
            tortuosity.append(line[4])
        distance = distance[1:]
        tortuosity = tortuosity[1:]    
        dist_sum = 0
        for branch in distance:
            dist_sum = dist_sum + float(branch)      
        branchtotal = len(distance)
        branchnumber = 0
        weightsum = 0
        while branchnumber < branchtotal:
            weight = (float(distance[branchnumber]) * 
                      float(tortuosity[branchnumber]) / dist_sum)
            weightsum = weightsum + weight
            branchnumber += 1
        output_writer.writerow([filename, weightsum])
    except:
        print filename + ".txt" + " not found"    