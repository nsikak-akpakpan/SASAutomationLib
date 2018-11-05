/* simple script to send email */
FILENAME Mailbox EMAIL 'Erik.Tilanus@planet.nl'
 Subject='Test Mail message';
DATA _NULL_;
FILE Mailbox;
PUT "Hello";
PUT "This is a message from the DATA step";
RUN;

/*Program 2: As Program 1, but with all mail data moved from FILENAME to FILE statement*/
FILENAME Mailbox EMAIL;
DATA _NULL_;
FILE Mailbox TO='Erik_Tilanus@compuserve.com'
 SUBJECT='Test Mail message';
PUT "Hello";
PUT "This is a message from the DATA step";
RUN;

FILE MailBox TO=('Erik <Erik_Tilanus@compuserve.com>'
 'Myself <Erik.Tilanus@planet.nl>')
 CC=('Synchrona@planet.nl'
 'Erik_Tilanus@hotmail.com')
 BCC='Anneke.Tilanus@planet.nl'
 FROM='My Business <Synchrona@planet.nl>'
 REPLYTO='My anonymous address <villavinkeveen@hotmail.com>'
 SUBJECT='Test Mail message to group and other options';
 
/* Program 3: Sending more mails simultaneously */
FILENAME MailBox1 EMAIL;
FILENAME MailBox2 EMAIL;
FILENAME MailBox3 EMAIL;
DATA _NULL_;
FILE MailBox1 TO='Erik <Erik_Tilanus@compuserve.com>'
 SUBJECT='First addressee';
PUT "Hello Mailbox1";
FILE MailBox2 TO='erik.tilanus@planet.nl'
 SUBJECT='Second addressee';
PUT "Hello Mailbox2 ";
FILE MailBox3 TO='synchrona@planet.nl'
 SUBJECT='Third addressee';
PUT "Hello Mailbox3 ";
RUN;

/*SENDING ATTACHMENTS
Including attachments is also rather straightforward: specify the attachment in either the FILENAME or FILE
statement with the keyword ATTACH, like: */
FILENAME MailBox EMAIL ATTACH='C:\SASUtil\Tips.doc';
/*or */
FILE Mailbox ATTACH='C:\SASUtil\Tips.doc';
/*If you want to specify the attachment under program control, you can use the special option in the PUT statement: */
PUT '!EM_ATTACH!' AttachInfo;

/* Program 4: Creating a report in a DATA step and sending it as attachment */
FILENAME MailBox EMAIL 'Erik.Tilanus@planet.nl'
 SUBJECT='Mail message with RTF attachment';
FILENAME rtffile 'e:\testlist.rtf';
ODS LISTING CLOSE;
ODS RTF FILE=rtffile;
DATA _NULL_;
FILE PRINT;
PUT "This is my report";
RUN;
ODS RTF CLOSE;
ODS LISTING;
DATA _NULL_;
attachment=FINFO(FOPEN('rtffile'),'File Name');
FILE MailBox;
PUT "attached you find the report";
PUT '!EM_ATTACH!' attachment;
RUN;

/*Program 5: Sending procedure output directly by mail (HTML or RTF)*/
FILENAME mail EMAIL TO="erik.tilanus@planet.nl"
 SUBJECT="HTML OUTPUT" CONTENT_TYPE="text/html";
ODS LISTING CLOSE;
ODS HTML BODY=mail;
PROC PRINT DATA=mydata.Sales;
RUN;
ODS HTML CLOSE;
ODS LISTING;

/*Program 6: Sending bulk mail */
FILENAME mail EMAIL;
DATA _NULL_;
SET forum.Officelist END=eof ;
FILE mail;
PUT '!EM_TO!' mailaddress;
PUT '!EM_SUBJECT!' 'New price information for ' office;
PUT "These are the fares as valid on &SYSDATE";
PUT 'Selection: ' location;
DO n=1 to NOBS;
 SET forum.fares POINT=n NOBS=nobs;
 IF (location = 'XXX' and package = 'Special') OR
 (location ne 'XXX' and package ne 'Special' and
 location eq destination)
 THEN PUT destination package price;
END;
PUT '!EM_SEND!' / '!EM_NEWMSG!';
IF eof THEN PUT '!EM_ABORT!';
RUN;
