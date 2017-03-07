# Imported Modules
import sys
import os
import csv
import xlwt
import datetime as DT
from xlrd import open_workbook
import argparse
import subprocess

sys.path.insert(0, './EDSA')
import EDSA

sys.path.insert(0, './CreateLabel')
import CopyFilesComplete

def main():

    parser = argparse.ArgumentParser(description='Creates Arc Flash Reports')

    parser.add_argument('Report_Type', choices=['AF', 'SCC', 'FULL'], help='Report Type Choice of just AF, SCC or FULL (AF, PDC and SCC)')
    parser.add_argument('PCE_Rep', help='PowerCore Representative (ie. roman)')    
    parser.add_argument('Job_Num', help='Job Number (ie. S2876)')
    parser.add_argument('Customer_Comp', help='Customer')
    parser.add_argument('Customer_Build', help='Location',)
    parser.add_argument('Customer_Add', help='Address')
    parser.add_argument('Working_Dir', help='working directory')
    args = parser.parse_args()

    

    #Copy Report Templates
    AFReportFolderName = '{!s}-{!s}-Report[{:%Y-%m-%d_%H%M%S}]'.format(args.Job_Num, "AF", DT.datetime.now())
    PDCReportFolderName = '{!s}-{!s}-Report[{:%Y-%m-%d_%H%M%S}]'.format(args.Job_Num, "PDC", DT.datetime.now())
    SCCReportFolderName = '{!s}-{!s}-Report[{:%Y-%m-%d_%H%M%S}]'.format(args.Job_Num, "SCC", DT.datetime.now())
    AFReport_Path = '{!s}/{!s}'.format(args.Working_Dir, AFReportFolderName)
    PDCReport_Path = '{!s}/{!s}'.format(args.Working_Dir, PDCReportFolderName)
    SCCReport_Path = '{!s}/{!s}'.format(args.Working_Dir, SCCReportFolderName)

    if args.Report_Type == 'FULL':
        CopyFilesComplete.copyanything('z:\Source-Code\ArcFlash\AF-Report-full\Latex\ArcFlash', AFReport_Path)
        print '\n', AFReportFolderName, ' Generated', '\n'
        AF_Path = '{!s}/{!s}'.format(AFReport_Path, "Arc Flash")
        os.chdir(AF_Path)
        subprocess.call(["perl", "/z/Source-Code/ArcFlash/AF-Report-full/Perl/CreateAF.pl", args.PCE_Rep, args.Job_Num, args.Customer_Comp, args.Customer_Build, args.Customer_Add, args.Working_Dir, args.Report_Type])
        subprocess.call(["perl", "/z/Source-Code/ArcFlash/AF-Report-full/Perl/CreateResults.pl", args.PCE_Rep, args.Working_Dir, AF_Path, args.Report_Type])
		
		
        CopyFilesComplete.copyanything('z:\Source-Code\ArcFlash\AF-Report-full\Latex\PDC', PDCReport_Path)
	print '\n', PDCReportFolderName, ' Generated', '\n'
        #Generate PDC Latex Report
        PDC_Path = '{!s}/{!s}'.format(PDCReport_Path, "PDC")
        os.chdir(PDC_Path)
        subprocess.call(["perl", "/z/Source-Code/ArcFlash/AF-Report-full/Perl/CreatePDC.pl", args.PCE_Rep, args.Job_Num, args.Customer_Comp, args.Customer_Build, args.Customer_Add, args.Working_Dir, args.Report_Type])

    elif args.Report_Type == 'PDC':
        CopyFilesComplete.copyanything('z:\Source-Code\ArcFlash\AF-Report-full\Latex\PDC', PDCReport_Path)
	print '\n', PDCReportFolderName, ' Generated', '\n'
        #Generate PDC Latex Report
        PDC_Path = '{!s}/{!s}'.format(PDCReport_Path, "PDC")
        os.chdir(PDC_Path)
        subprocess.call(["perl", "/z/Source-Code/ArcFlash/AF-Report-full/Perl/CreatePDC.pl", args.PCE_Rep, args.Job_Num, args.Customer_Comp, args.Customer_Build, args.Customer_Add, args.Working_Dir, args.Report_Type])


    elif args.Report_Type == 'SCC':
        CopyFilesComplete.copyanything('z:\Source-Code\ArcFlash\AF-Report-full\Latex\SCC', SCCReport_Path)
	print '\n', SCCReportFolderName, ' Generated', '\n'
        #Generate SCC Latex Report
        SCC_Path = '{!s}/{!s}'.format(SCCReport_Path, "SCC")
        os.chdir(SCC_Path)
        subprocess.call(["perl", "/z/Source-Code/ArcFlash/AF-Report-full/Perl/CreateSCC.pl", args.PCE_Rep, args.Job_Num, args.Customer_Comp, args.Customer_Build, args.Customer_Add, args.Working_Dir, args.Report_Type])

    else:
        CopyFilesComplete.copyanything('z:\Source-Code\ArcFlash\AF-Report-full\Latex\ArcFlash', AFReport_Path)
        print '\n', AFReportFolderName, ' Generated', '\n'
        #Generate AF Latex Report
        AF_Path = '{!s}/{!s}'.format(AFReport_Path, "Arc Flash")
        os.chdir(AF_Path)
        subprocess.call(["perl", "/z/Source-Code/ArcFlash/AF-Report-full/Perl/CreateAF.pl", args.PCE_Rep, args.Job_Num, args.Customer_Comp, args.Customer_Build, args.Customer_Add, args.Working_Dir, args.Report_Type])
        subprocess.call(["perl", "/z/Source-Code/ArcFlash/AF-Report-full/Perl/CreateResults.pl", args.PCE_Rep, args.Working_Dir, AF_Path, args.Report_Type])

    if (args.Report_Type == 'FULL') or (args.Report_Type == 'AF'):
        #Arc Heat Table
        EDSA.ArcheatTable(args.Job_Num, args.Customer_Comp, args.Customer_Build, args.Customer_Add, args.Working_Dir, AFReport_Path)   
    else:
        pass
        
    os.chdir(args.Working_Dir)


   

if __name__ == '__main__':
    main()
    sys.exit()

