/*Using SAS to add PivotTables to your Excel workbook  15
By Chevell Parker on SAS Users March 27, 2015 Programming Tips
SAS Technical Support Problem SolversIn Microsoft Excel, a PivotTable can help you to create an interactive view of summarized data. Within a PivotTable, it’s easy to adjust the dimensions (columns and rows) and calculated measures to suit your ad-hoc reporting needs. You can also create a PivotChart – similar in concept to a PivotTable, but using a visualization technique such as a bar chart instead of a table.

SAS provides a special ODS tagset that can add a PivotTable or a PivotChart to your Microsoft Excel workbook automatically. It’s called the TableEditor tagset, and you can download it for free from the SAS support site.

In the example in this post, we’ll use the ODS EXCEL destination to create a native Open Office XML file (XLSX file) for Excel to read. Then we’ll use the TableEditor tagset to update the workbook to add a PivotTable to this worksheet.

How to access the TableEditor tagset
You can automate PivotTable creation by using the downloadable TableEditor tagset on the Windows operating system. If your network allows you to access the Web from within your SAS session, you can even use %INCLUDE to access the tagset directly within your program:*/

/* reference the tagset from support.sas.com */
Filename tpl url 
"http://support.sas.com/rnd/base/ods/odsmarkup/tableeditor/tableeditor.tpl";

/* insert the tagset into the search path for ODS templates */
Ods path(Prepend) work.templat(update);
%include tpl;
/*If you cannot use FILENAME URL, then simply download the TPL file to a local folder on your PC, and change the %INCLUDE statement to reference the file from that location.

Creating the Excel workbook
This latest version of the tagset allows you add a PivotTable to any data source that Excel can read, regardless of how that Excel file was created and without having to generate an intermediate HTML file. SAS allows you to create Excel content using several methods, including:

ODS CSV (or simple DATA step) to create a comma-separated value representation of data as a source for an Excel worksheet
ODS tagsets.ExcelXP, which creates an XML representation of a workbook that Excel can read.
PROC EXPORT with DBMS=EXCEL, EXCELCS, or XLSX (requires SAS/ACCESS to PC Files)
ODS EXCEL (new in SAS 9.4 and still labeled “experimental” as of SAS 9.4 Maintenance 2)
Here’s program to create the Excel content from the SHOES sample data. Change the “temp” file path as needed for your own system.*/

ods excel file='/folders/myshortcuts/myfolder/temp/temp.xlsx' options(sheet_name="shoe_report");

proc print data=sashelp.shoes;
run;

ods excel close;
/*Here’s a sample of the output Excel workbook.

Excel workbook created with SAS ODS

Working with TableEditor tagset options for PivotTables
The TableEditor tagset has options to control the various drop zones of an Excel PivotTable such as:

Page area -  uses one or more fields  to subset or filter the data .
Data area— a field that contains  values to be summarized.
Column area—a field to assign to a column orientation in the PivotTable.
Row area—a field that you assign to a row orientation which is used to categorize the data.
The options to control these drop zones in the PivotTable are the PIVOTPAGE=, PIVOTROW=, PIVOTCOL= and PIVOTDATA= options. Each of these options can specify a single column or multiple columns (each separated with a comma).

The UPDATE_TARGET= option contains the name of the workbook to update, while the SHEET_NAME= option which specifies which sheet should be used as the source for the PivotTable. The OUTPUT_TYPE= option is set to “Script”, which tells the tagset to create a JavaScript output file with the Excel commands. Other options can control formatting, appearance, and the summarizations of the PivotTable.

NOTE: when specifying a file path in the UPDATE_TARGET= option, you must “escape” each backslash by using an additional backslash. The backslash character has a special meaning in JavaScript that turns special characters into string characters.

Here are some notes on the remaining options:

PIVOT_SHEET_NAME= (new option) allows you to name the PivotTable to be named different from the source worksheet name. (The default is to append “_Pivot” to the source worksheet name.)
PIVOT_TITLE= (new option) adds a title to the PivotTable.
PIVOTDATA_FMT= specifies the numeric display format
PIVOT_FORMAT= specifies one of the Excel table formats (found on the formatting style ribbon).
Creating the PivotTable
We’ll use a two-step technique to add a PivotTable to our sample workbook:

Use ODS tagsets.TableEditor and special PIVOT options to create a script file that contains instructions for the PivotTable that we want.
Use the X command to execute that script file, which will automate Microsoft Excel to add the PivotTable content.
Here’s a program that generates the script and executes it.  The NOXSYNC and NOXWAIT options allow control to return to the SAS session as the script is run.*/

*options noxsync noxwait;
ods tagsets.tableeditor file='/folders/myshortcuts/myfolder/temp/PivotTable.js'                                                                                                                                      
/* remember to escape the backslashes */
  options(update_target='/folders/myshortcuts/myfolder/temp/temp.xlsx' doc="help"                                                                                                                                 
    output_type="script"                                                                                                                                           
    sheet_name="shoe_report"  
    pivot_sheet_name="Profit Analysis" 
    pivotrow="region"                                                                                                                                              
    pivotcol="product"                                                                                                                                             
    pivotdata="sales"  
    pivotdata_fmt="$#,###" 
    pivot_format="light1"
    pivot_title="Pivot Analysis for XXX" 
);                                                                                                                                             

/* dummy output to trigger the file creation */                                                                                                                                                                                          
data _null_;                                                                                                                                                                          
 file print; 
 put "test";                                                                                                                                                                             
run;                                                                                                                                                                            
                                                                                                                                                                  
ods tagsets.tableeditor close; 
x '/folders/myshortcuts/myfolder/temp/PivotTable.js';    
/*Here is a sample of the PivotTable output.

pivot2

Creating PivotCharts
What’s a good PivotTable without a PivotChart?  PivotCharts can be added to an existing workbook as well. Simply add the PIVOTCHARTS=”yes” option along with the CHART_TYPE option.
*/
*options noxsync noxwait;
ods tagsets.tableeditor file='/folders/myshortcuts/myfolder/temp/PivotChart.js'     
  /* remember to escape the backslashes */ 
  options(update_target='/folders/myshortcuts/myfolder/temp/temp.xlsx'                                                                                                                                  
    output_type="script"                                                                                                                                 
    sheet_name="shoe_report"  
    pivot_sheet_name="Profit Analysis Chart" 
    pivotrow="region"                                                                                                                                        
    pivotcol="product"                                                                                                                                          
    pivotdata="sales"  
    pivotdata_fmt="$#,###" 
    pivot_title="Pivot Analysis for XXX"
    pivot_format="light1"
    pivotcharts="yes"  
    pivot_chart_name="Profit Analysis Charts"
    chart_type="columnclustered" 
);                                                                                                                                  
 
/* dummy output to trigger the file creation */ 
data _null_;                                                                                                                                                                    
 file print;                                                                                                                                                            
 put "test";                                                                                                                                                              
run;                                                                                                                                                                  
                                                                                                                                                                 
ods tagsets.tableeditor close;
x '/folders/myshortcuts/myfolder/temp/PivotChart.js';
*Here is a sample of the PivotChart output.
