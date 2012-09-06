# This script produces a Sholl analysis of spine number at distance intervals
# from the cell body. It is the same function as ShollSpineCount, but it takes 
# the Neurolucida Explorer spreadsheet "spine details" as input. To get spine 
# density at distances from a cell body (e.g. as in several papers by Elston 
# et al.): Divide the output of this file by the output from a 
# ShollDendriteLength.py run, which took input at the same interval as this 
# file, and with *only* trees counted for spines highlighted. You need to set 
# the sholl_interval value in this file to the interval you specify in 
# Neurolucida Explorer.

##############################################################################
# Section 1: setting up your computer to run this program
# This program is written in Python 2.6-2.7. 

# Step 1: Install Python 2.7.3 from here: 
# http://www.python.org/download/releases/2.7.3/
# On the page, select the Windows MSI installer (x86 if you have 32-bit 
# Windows installed, x86-64 if you have 64-bit Windows installed.)
# I suggest using the default option, which will install Python to c:/Python27

# Step 3: Copy this program into the c:/Python27 directory
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
cases = [["M31-09", 1], ["M32-09", 10], ["M38084", 10], ["R2-11", 10], 
         ["R3-11", 10], ["R4-11", 10], ["R5-11", 10]]

# List the brain regions examined (this can be a single region)
regions = ["lateral", "central"] 

# Set the size of your Sholl circles, **you must also do this in Neurolucida 
# Explorer**:
sholl_interval = 25

# Set the number of Sholl circles to analyze:
sholl_num = 200

# List the name of the source file, as saved
analysisname = "spine details"

# Name the output file, this will be placed in the same directory as your 
# input files
outputfile = "spinedetailssholldistCSV.txt"

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
captions = ["Cell"]
for entry in range(0, sholl_num): 
    captions.append("Sholl Sp# " + str(entry * sholl_interval) + " um")
output_writer.writerow(captions)

for filename in filelist:
    path = directory + "\\" + filename + ".txt"
    try:
        myfileobj = open(path, "r") 
        csv_read = csv.reader(myfileobj, dialect=csv.excel_tab)   
        sholldist = []
        for line in csv_read:
            sholldist.append(line[6])
        sholldist=sholldist[1:]
        sholllength=len(sholldist)
        #remove extra blanks from end of this filetype 
        sholldist=sholldist[:(sholllength-2)]
        corrected_sholldist = []
        output_row = [filename]
        for spine in sholldist:
            if len(spine) > 0:
                corrected_sholldist.append(float(spine) / sholl_interval)
        for distance in range(0,sholl_num):
            spinesum = 0
            for corrected_spine in corrected_sholldist:
                if corrected_spine < distance + 1 and corrected_spine > distance:
                    spinesum += 1
            output_row.append(spinesum)
        output_writer.writerow(output_row)
    except:
        print filename + ".txt" + " not found"            
