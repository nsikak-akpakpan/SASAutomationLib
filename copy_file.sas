/* these IN and OUT filerefs can point to anything */
filename in "c:\dataIn\input.xlsx"; 
filename out "c:\dataOut\output.xlsx"; 
 
/* copy the file byte-for-byte  */
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
 
filename in clear;
filename out clear;
