

%macro updatePageBreaks;
ods output properties=documentProps; 
  proc document name=docout.chdrReport;
  dir \LISTING;
  list / levels=all details;
  run; 
ods output close;


data _null_;
  set documentProps;
  where upcase(type)='TABLE';
  command=cats('%removepagebreak','(doc=docout.chdrReport,obj=',path,')');
  call execute(command);
run;

%mend updatePageBreaks;


%macro removepagebreak(doc=, obj=);
   proc document name=docout.chdrReport label=" ";
      obpage &obj / delete; /* delete pagebreak BEFORE listing */
      obpage &obj / after;  /* add pagebreak AFTER listing */
   quit; 
%mend removepagebreak;




proc template;
	define style styles.listingOCS;
        parent=styles.Monospace;
		class usertext / 
			font_size = 8pt
			font_face = "Courier New"
          ;
	  style paragraph from usertext;
	  Style Data from Data /
             Just=right;

		class body /
			topmargin = 0.79in
			leftmargin = 0.79in
			rightmargin = 0.79in
			;
		class contenttitle /
		fontweight = bold
		just=left
		font_size = 24pt
		font_face = "Arial";
	end;
run;




%MACRO header(label=,pagebreak=no);
	ods proclabel " ";
	proc odstext contents="&label" pagebreak=&pagebreak;
	/* adds a HEADER1 to word document */
/*	p "{\pard\s1\b\ul &label \par}" / style=[fontsize=14pt fontweight=bold fontfamily=Arial];  */
/*	p "Header - &label";*/
	p "{\pard\s1\b\ul &label \par}" / style=[fontsize=14pt fontweight=bold fontfamily=Arial];  /* adds a HEADER1 to word document */
	run;
%MEND header;



%MACRO TOC(pagebreak=no);
    ODS proclabel " ";

	PROC ODSTEXT contents=" " pagebreak=&pagebreak;
	p "{\pard\b\fs28\l\ulth {Table of Contents} \par}"; 
	RUN;
	
	ODS TEXT="{\field{\*\fldinst {\\TOC \\f \\h} }  }";
	run; 

%MEND TOC;



%MACRO PROTOCOL(label=);

data header;
     length column1 column2 $200;
     column1="CHDR protocol ID";
     column2="CHDR1834";
     output;
     column1="Short title";
     column2="&STit";
     output;
     column1="Sponsor protocal ID";
     column2=&SpID;
     output;
     column1="File name";
     column2="CHDR_PD_IP.DOCX";
     output;
     column1="Date";
     column2=put(today(),yymmddD10.);
     output;
     run;




ODS proclabel " ";
proc report data=header nofs  nowd contents= " " 
style(report)={rules=none   cellspacing=0} ;
      column column1 column2;
      define column1 / "" width=100 style(column)=[just=left width=65mm LEFTmargin=1mm cellpadding=1mm];
      define column2 / "" width=100 style(column)=[just=left width=120mm]; 
run;


%MEND PROTOCOL;




%MACRO COVERPAGE_2;

ODS TEXT='^S={preimage="c:\home\01_projects\CHDR\chdr.jpg" leftmargin=0cm}';
ODS TEXT='^{newline 1}';
ODS TEXT='^S={LEFTMARGIN=3.10in font_size=20pt font_face=Arial }Statistical appendix:  PD analysis output CHDR1834';
ODS TEXT='^S={LEFTMARGIN=3.10in font_size=16pt font_face=Arial font_weight=bold }A randomized, double-blind, double-dummy, placebo-controlled,';
ODS TEXT='^S={LEFTMARGIN=3.10in font_size=16pt font_face=Arial font_weight=bold }three-way crossover study to investigate the effects of IV ';
ODS TEXT='^S={LEFTMARGIN=3.10in font_size=16pt font_face=Arial font_weight=bold }lidocaine and oral lacosamide on nerve excitability and evoked pain tests in healthy subjects '; 

%MEND COVERPAGE_2;


%MACRO coverpage;
     ODS TEXT='^S={preimage="c:\home\01_projects\CHDR\chdr.jpg" leftmargin=0cm}';
	ODS TEXT='^{newline 1}';
	ODS TEXT='^S={LEFTMARGIN=0cm font_size=20pt font_face=Arial }Statistical appendix:  PD analysis output CHDR1834';
	ODS TEXT='^S={LEFTMARGIN=0cm font_size=16pt font_face=Arial font_weight=bold }A randomized, double-blind, double-dummy, placebo-controlled,';
	ODS TEXT='^S={LEFTMARGIN=0cm font_size=16pt font_face=Arial font_weight=bold }three-way crossover study to investigate the effects of IV lidocaine';
	ODS TEXT='^S={LEFTMARGIN=0cm font_size=16pt font_face=Arial font_weight=bold }and oral lacosamide on nerve excitability and evoked pain tests in healthy subjects '; 

	PROC ODSTEXT contents=" "; /* end with a a dummy proc: to add pagebreak */
		p '';
	RUN;
%MEND coverpage;




/* The PROC always forces a pagebreak if needed */

%MACRO PAGEBREAK;
	PROC ODSTEXT pagebreak=yes contents=" ";
		p '';
	RUN;
%MEND;




%MACRO ODSPD_Init;
%global ODSPD_currentLibrary;
%global ODSPD_currentCatalog;
%MEND ODSPD_Init;

%MACRO ODSPD_Open(library=, catalog=, mode=update);
%let ODSPD_currentLibrary=&library;
%let ODSPD_currentCatalog=&catalog;
ods document name=&library..&catalog(&mode);
%MEND ODSPD_Open;

%MACRO ODSPD_Path(path=);
 ods document dir=(path=&path);
%MEND ODSPD_Path;

%MACRO ODSPD_Close;
ods document close;
%MEND ODSPD_Close;

%MACRO ODSPD_List(library=, catalog=);

proc document name=&library..&catalog;
list / levels=all;
run;
quit;

%MEND ODSPD_List;

%MACRO ODSPD_Replay(library=, catalog=, path=);
   proc document name=&library..&catalog;
        dir &path;
 	replay;
   run;
   quit;
%MEND ODSPD_Replay;


