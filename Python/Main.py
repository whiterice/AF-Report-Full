# Imported Modules
import sys
import os
import csv
import xlwt
import datetime as DT
from xlrd import open_workbook
import argparse
import CopyFiles
import subprocess

sys.path.insert(0, './EDSA')
import EDSA



def main():

    parser = argparse.ArgumentParser(description='Creates Arc Flash Reports')
    parser.add_argument('PCE_Rep', help='PowerCore Representative (ie. roman)')    
    parser.add_argument('Job_Num', help='Job Number (ie. S2876)')
    parser.add_argument('Customer_Comp', help='Customer')
    parser.add_argument('Customer_Build', help='Location')
    parser.add_argument('Customer_Add', help='Address')
    parser.add_argument('Working_Dir', help='working directory')
    args = parser.parse_args()

    #Arc Heat Table
    EDSA.ArcheatTable(args.Job_Num, args.Customer_Comp, args.Customer_Build, args.Customer_Add, args.Working_Dir)

    #Copy Report Template
    ReportFolderName = '{!s}-AF-Report[{:%Y-%m-%d_%H%M%S}]'.format(args.Job_Num, DT.datetime.now())

    Report_Path = '{!s}{!s}'.format(args.Working_Dir, ReportFolderName)

    CopyFiles.copyanything('c:\Users\Scott\Dropbox\Scripts\Python\AF-Report-full\Latex', Report_Path)
    print '\n', ReportFolderName, ' Generated', '\n'

    #Generate Main and Results Page for Latex Report
    AF_Path = '{!s}/{!s}'.format(Report_Path, "Arc Flash")
    os.chdir(AF_Path)
    subprocess.call(["perl", "/c/Users/Scott/Dropbox/Scripts/Python/AF-Report-full/Perl/CreateAF.pl", args.Job_Num, args.Customer_Comp, args.Customer_Build, args.Customer_Add, args.Working_Dir])
    subprocess.call(["perl", "/c/Users/Scott/Dropbox/Scripts/Python/AF-Report-full/Perl/CreateResults.pl", args.PCE_Rep, args.Working_Dir, AF_Path])
    

    
    os.chdir(args.Working_Dir)


   

if __name__ == '__main__':
    main()
    sys.exit()

