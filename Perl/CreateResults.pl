#Perl
#!/usr/bin/perl
use strict;
use warnings;
use Data::Dumper qw(Dumper);

#Variable Declaration
my $fh;
my $signature = " ";
my $PCERep = $ARGV[0];



#PCERep Signature
if (($PCERep eq "Vince")||($PCERep eq "vince"))
{
$signature = "
%Vince Signature
%\\begin{comment}
\\begin{multicols}{2}
\\centering
\\includegraphics[height=0.5in, keepaspectratio=true]{../Images/Roman_signature.jpg} \\\\
Roman Bulla, P. Eng. \\Power Systems Engineer \\\\
\\includegraphics[height=0.5in, keepaspectratio=true]{../Images/Vince_signature.jpg} \\\\
Vince Klingenberger \\\\Electrical Engineering Technologist \\\\
\\end{multicols}
%\\end{comment}";

}
	
elsif (($PCERep eq "Scott")||($PCERep eq "scott"))
{
$signature = "
%Scott Signature
%\\begin{comment}
\\begin{multicols}{2}
\\centering
\\includegraphics[height=0.5in, keepaspectratio=true]{../Images/Roman_signature.jpg} \\\\
Roman Bulla, P. Eng. \\Power Systems Engineer \\\\
\\includegraphics[height=0.5in, keepaspectratio=true]{../Images/Scott_signature.jpg} \\\\
Scott Vermeire \\\\Engineering Intern \\\\
\\end{multicols}
%\\end{comment}";
	
}
else
{
$signature = "
\\beging{flushleft}
\\includegraphics[height=0.5in, keepaspectratio=true]{../Images/Roman_signature.jpg} \\\\
Roman Bulla, P. Eng. \\Power Systems Engineer \\\\
\\end{flushleft}
%\\end{comment}";
}

#Latex
open($fh, '>', 'Results.tex');

print $fh <<"END_OF_REPORT";

%Arc Flash Study Results & Recommendations

\\section{Results}
\\label{af:results}

\\subsection{Arc Flash Study Results and Recommendations}
\\label{af:results:afrr}

The Arc Flash Hazard Evaluation Study has shown that:
\\begin{enumerate}
\\item WDMB01 is Arc Hazard Class 4
\\item WDMB02-SW is Arc Hazard Class 4
\\item WDMB03-SW is Arc Hazard Class 4


\\end{enumerate}

\\pagebreak

\\subsection{Arc Flash Study Details}
\\label{af:results:afsd}

The following issues have been encountered and are of note:

\\begin{itemize} 
\\item The risk category for all downstream equipment from the last bus or power panel listed on the summary sheet and shown on the drawings shall be treated as the last indicated Hazard Category, unless indicated otherwise.

\\item Arc Flash Hazard Identification labels for all equipment within the scope of this study will be affixed in a visible area.

\\item Please note that the operation of fused switches or starters with cover on or enclosure doors closed does not require arc flash rated PPE Hazard, based on CSA Z462.

\\item Section 4.3.5.1 of the CSA Z462-2012 standard states. "An arc flash hazard analysis shall determine the (a) arc flash boundary; (b) incident energy at the working distance; and (c) PPE that personnel within the arc flash boundary shall use.\cite{CSA}

\\item The analysis shall be updated when a major modification or renovation takes place. It shall be reviewed periodically, not to exceed 5 years, to account for changes in the electrical distribution system that could affect the results of the analysis.\cite{CSA}


\\item The analysis shall take into consideration the design of the overcurrent protective device and its opening time, including its condition of maintenance. Incident energy need not be determined if the requirements of Clauses 4.3.7.3.15 and 4.3.7.3.16 are met."\cite{CSA}
 
\\item\\textbf{Equipment below 240 V need not be considered (for Arc Flash Hazard Analysis) unless it involves at least one 125 kVA or larger low impedance transformer in its immediate power supply.}\\cite{IEEE}	 
\\end{itemize}
\\pagebreak
\\subsection{Study Recommendations}
\\label{af:results:afsr}

Based on the Arc Flash evaluation Study performed, we recommend the following:
\\begin{itemize}
\\item We recommend that all personnel authorized to operate any electrical equipment be properly educated on Arc Flash Hazard and trained  in use of arc flash PPE (Personal Protective Equipment)

\\item	Provide PPE of appropriate level to all personnel that will be operating electrical equipment.

\\item	We recommend that a formal written Electrical Equipment Operation Procedures Manual as well as an Arc Flash Hazard Program that meets the regulations noted in CSA Z462-15 and IEEE-1584, be developed and implemented.
\\end{itemize}
\\vspace{10mm}
\\noindent Thank you for this opportunity to be of service to you.  If you have any questions regarding the recommendations in this report or any other matter, please contact our London Engineering Services office at (519) 474-1175. \newline
\\vspace{5mm}
\\\\
\\noindent Sincerely,\\newline

\\vspace{5mm}
\\noindent\\textbf{PowerCore Engineering}\\newline

$signature

END_OF_REPORT


close $fh;

print "Results Page for LaTex Arc Flash Report Generated"