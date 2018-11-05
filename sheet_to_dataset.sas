/*Upgrade to SAS 9.4 (TS1M2) and use XLSX files then you can use the XLSX libname engine.*/
libname x xlsx 'c:\downloads\test.xlsx';
proc sql noprint;
select memname into : dslist separated by ' '
from dictionary.members
where libname='X'
;
quit;

/*If you have an older version of SAS but you are using XLSX files you can read the metadata out of the XLSX file to generate the list of member names.
Here is some SAS code to do that. */
%let filename='c:\downloads\test.xlsx';
%let sheets=work.sheets;
/**----------------------------------------------------------------------;
* Generate XMLMAP to read the sheetnames from xl/workbook.xml ;
*----------------------------------------------------------------------;*/
filename _wbmap temp;
data _null_;
  file _wbmap;
  put '<SXLEMAP version="2.1"><TABLE name="Sheets">'
    / '<TABLE-PATH>/workbook/sheets/sheet</TABLE-PATH>'
    / '<COLUMN name="Sheet"><TYPE>character</TYPE>'
    / '<DESCRIPTION>Sheet Name</DESCRIPTION>'
    / '<PATH>/workbook/sheets/sheet/@name</PATH>'
    / '<DATATYPE>string</DATATYPE><LENGTH>32</LENGTH>'
    / '</COLUMN>'
    / '<COLUMN name="State"><TYPE>character</TYPE>'
    / '<DESCRIPTION>Sheet State</DESCRIPTION>'
    / '<PATH>/workbook/sheets/sheet/@state</PATH>'
    / '<DATATYPE>string</DATATYPE><LENGTH>20</LENGTH>'
    / '</COLUMN>'
    / '</TABLE></SXLEMAP>'
  ;
run;

/*----------------------------------------------------------------------;
Copy xl/workbook.xml from XLSX file to physical file ;
Note: Cannot use ZIP filename engine with XMLV2 libname engine ;
filename _wb pipe %sysfunc(quote(unzip -p &filename xl/workbook.xml)); 
----------------------------------------------------------------------;*/
filename _wbzip ZIP &filename member='xl/workbook.xml';
filename _wb temp ;
data _null_;
  infile _wbzip lrecl=30000;
  file _wb lrecl=30000;
  input;
  put _infile_;
run;

/*----------------------------------------------------------------------;

* Generate LIBNAME pointing to copy of xl/workbook.xml from XLSX file ;

*----------------------------------------------------------------------;*/
libname _wb xmlv2 xmlmap=_wbmap ;
/**----------------------------------------------------------------------;

* Read sheet names from XLSX file into a SAS dataset. ;

* Create valid SAS dataset name from sheetname or sheetnumber. ;

*----------------------------------------------------------------------;*/
filename extract temp;
data &sheets ;
  if eof then call symputx('nsheets',_n_-1);
  length Number 8;
  set _wb.sheets end=eof;
  number+1;
  length Memname $32 Filename $256 ;
  label number='Sheet Number' memname='Mapped SAS Memname' filename='Source Filename' ;
  filename = &filename ;
  if ^nvalid(compress(sheet),'v7') then memname = cats('Sheet',number);
  else memname = translate(trim(compbl(sheet)),'_',' ');
run;
%let sheets=&syslast;

/*----------------------------------------------------------------------;

* Clear the libname and filenames used in reading the sheetnames. ;

*----------------------------------------------------------------------;*/

libname _wb clear ;
filename _wb clear ;
filename _wbzip clear ;
filename _wbmap clear ;
