#Perl
#!/usr/bin/perl
use strict;
use warnings;

my $StudyTitle = " ";
my $PCERepName = " ";
use Data::Dumper qw(Dumper);

my $PCERep = $ARGV[0];
my $JobNum = $ARGV[1];
my $Customer = $ARGV[2];
my $Building  = $ARGV[3];
my $Address   = $ARGV[4];
my $WorkDir   = $ARGV[5];
my $ReportType   = $ARGV[6];

print "Create AF: ", $PCERep, $JobNum, " ", $Customer, " ", $Building, " ", $Address, " ", $WorkDir, "\n",;

#PCERep Signature
if (($PCERep eq "Vince")||($PCERep eq "vince"))
{
$PCERepName = "Vince Klingenberger";

}
elsif (($PCERep eq "Scott")||($PCERep eq "scott"))
{
$PCERepName = "Scott Vermeire, P.Eng.";
	
}
else
{
$PCERepName = "";
}

#Title Selection
if ($ReportType eq "AF")
{
	$StudyTitle = "ARC FLASH HAZARD ANALYSIS";
}
elsif ($ReportType eq "FULL")
{
	$StudyTitle = "ARC FLASH HAZARD ANALYSIS";
}
else
{
	$StudyTitle = "ARC FLASH HAZARD ANALYSIS";
}
	
#Latex
open(my $fh, '>', 'Arc Flash.tex');

print $fh <<"END_OF_REPORT";

\\documentclass{article}
\\usepackage{graphicx}
\\usepackage[pdftex,bookmarks,bookmarksnumbered,colorlinks]{hyperref}
\\usepackage{pdfpages}
\\usepackage{fancyhdr}
\\usepackage[english]{babel}
\\usepackage{multicol}
\\usepackage[margin=1.1in]{geometry}
\\usepackage{textcomp}
\\usepackage{fmtcount}
\\usepackage[utf8]{inputenc}
\\usepackage{hyperref}

%Prevents Tables from being repositioned
\\usepackage{float}
\\restylefloat{table}

% Table spacing
%Top space
\\newcommand{\\Toprule}{\\rule{0pt}{3.0ex}}
%Bottom space
\\newcommand{\\Botrule}{\\rule[-1.2ex]{0pt}{0pt}}

%Constants
\\newcommand{\\DocTitle}{$StudyTitle}
\\newcommand{\\Customer}{$Customer}
\\newcommand{\\Target}{$Building}
\\newcommand{\\Address}{$Address}
\\newcommand{\\JobNum}{$JobNum}
\\newcommand{\\PCERep}{$PCERepName}

%hyperlinks
\\hypersetup{
pdftoolbar=true,        								% show Acrobat’s toolbar?
pdfmenubar=true,        								% show Acrobat’s menu?
pdffitwindow=false,     								% window fit to page when opened
pdfstartview={FitH},    								% fits the width of the page to the window
pdftitle={\\DocTitle},    								% title
pdfauthor={PowerCore Engineering},    					% author
pdfsubject={Report},   									% subject of the document
pdfcreator={PowerCore Engineering},   					% creator of the document
pdfproducer={PowerCore Engineering}, 					% producer of the document
pdfnewwindow=true,      								% links in new window
colorlinks=true,       									% false: boxed links; true: colored links
linkcolor=black,          								% color of internal links (change box color with linkbordercolor)
citecolor=blue,        									% color of links to bibliography
filecolor=blue,      									% color of file links
urlcolor=blue}           								% color of external links


%Header and Footer
\\renewcommand{\\sectionmark}[1]{\\markboth{#1}{}}
\\renewcommand{\\footrulewidth}{0.4pt}
\\fancyhead[R]{\\leftmark} % 1. sectionname
\\fancyfoot[C]{\\thepage}
\\fancyfoot[L]{
PowerCore Engineering\\\\
London, Ontario, Canada\\\\
Tel: (519) 474-1175\\\\
\\url{www.powercore.ca}}
\\cfoot{ }

\\fancyfoot[R]{
\\DocTitle \\\\ \\vspace{12pt}  -\\thepage -}
\\fancypagestyle{plain}{%
  \\fancyhf{}%
  \\renewcommand{\\headrulewidth}{0pt}%
}
\\fancyhead[L]{ % right
   \\includegraphics[height=0.4in]{../Images/BackSignLogo.jpg}
}

\\begin{document}

\\pagenumbering{Alph}
\\input{../CoverPage.tex}
\\pagenumbering{arabic}
\\pagestyle{fancy}

%TOC
\\tableofcontents
\\pagebreak

\\include{Introduction}
\\include{Results}
\\include{Objectives}
\\include{Procedures}
\\include{Observations}
%\\include{Modifications}
\\include{Bibliography}

\\end{document}

END_OF_REPORT


close $fh;

print "Main LaTex Arc Flash file Generated\n"