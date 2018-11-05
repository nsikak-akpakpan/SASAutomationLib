filename in url "http://www.LotsOfData.org/data/data.xlsx"
  /* PROXY= is important for going outside firewall, if you have one */
  /* proxy="http://yourProxy.company.com" */
  ;
filename out "c:\temp\data.xlsx";
data _null_;
 length filein 8 fileid 8;
 filein = fopen('in','I',1,'B');
 fileid = fopen('out','O',1,'B');
 rec = '20'x;
 do while(fread(filein)=0);
  rc = fget(filein,rec,1);
  rc = fput(fileid, rec);
  rc =fwrite(fileid);
 end;
 rc = fclose(filein);
 rc = fclose(fileid);
run;
 
/* Works on 32-bit Windows */
/* If using 64-bit SAS, you must use DBMS=EXCELCS */
PROC IMPORT OUT= WORK.test 
  DATAFILE = out /* the downloaded copy */
  DBMS=EXCEL REPLACE;
  SHEET="FirstSheet";
  SCANTEXT=YES;
  USEDATE=YES;
  SCANTIME=YES;
  GETNAMES=YES; /* not supported for EXCELCS */
  MIXED=NO; /* not supported for EXCELCS */
RUN;
 
filename in clear;
filename out clear;
