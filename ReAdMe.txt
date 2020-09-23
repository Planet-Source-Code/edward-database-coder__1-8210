Project:	vbDatabaseCoder Add-In and EXE
Version:	1.3.1 (26/May/2000)
Copyright:	© 1999-2000 qbd software ltd
Author:		edward moth

=====================================================================

WHAT'S NEW:

Version 1.3.1 (26/May/2000)
1.	Improved handling of Default field values.  Earlier versions
	could cause error in Appending fields or TableDefs.
	(Thanks: Romell E. Avendaño)

Version 1.3 (20/May/2000)
1.	Improved Code Production
	Changed how the final code is written.  The output module can
	be 50-60% smaller than original version.  Also broke up some
	of the tasks involved to avoid the final compiled procedure
	breaking VB's 64K maximum. (Thanks: Renzo Bagorda)

2.	Improved Index/Relation handling
	Not sure how I missed this one but I did.  The original was
	indexing the Foreign Fields in a Relation as Indexes within
	their own table.

3.	QueryDef Support Added
	The Coder now provides support for QueryDefs.  The Coder
	looks at the SQL, converts CR and LF characters to the
	constant vbCrLf and quotes (") to Chr$(34).  The SQL is
	placed in a string variable before being used in the
	CreateTableDef to reduce the possibility of error for big SQL
	statements. (Pleased with that coding ... hehehe)

NOTE:	I have not found a way of distinguishing between System
	Queries and user defined queries.  If anyone has any idea
	please let me know.  They usually appear with '~' symbols
	in their name.	
=====================================================================

PLEASE READ THE BORING WARRANTY AND LICENSE INFORMATION - FAILURE TO
DO SO WILL DEEM YOU 'A BIT NAUGHTY' UNDER THE PROVISIONS OF SECTION
118.12(A) OF THE PROTECTION OF EDWARD MOTH ACT (1999) - HM GOVERNMENT
(UNITED KINGDOM) PLEASE ALSO NOTE THAT UNDER SECTION 112.3 OF SAID
ACT, ALL PAYMENTS TO EDWARD SHOULD BE MADE IN CASH IN BROWN ENVELOPES - BET NO ONE READS THIS - HEHEHE - THEY NEVER DO.

=====================================================================

A.	PURPOSE
B.	REQUIREMENTS
C.	INSTRUCTIONS FOR SETTING UP AND USE
D.	WARRANTY AND LICENSE
E.	CONTACT

=====================================================================

A PURPOSE:

DatabaseCoder is a VB Add-in that will analyse an Access97 database
and make the code in VB required to create an empty replica of the
database from scratch.  It does not copy records, it only creates the
structure.  You can view the information the add-in has retrieved
from the database,  Tables, Fields, Indexes and Relationships.  It 
now supports Querys.  I am still working on a version that can
incorporate LookUp Tables and that would allow selection tables/
fields etc. to be included.  I am also working on a version that can
create a working project built around the database although work may
want that to be commercial.

=====================================================================

B REQUIREMENTS:

It's kinda handy if you have a copy of Visual Basic (version 5 or 6).
MS Access 97 may also prove to be a useful addition although
technically not essential

The Project uses:
	MS Common Dialog (v.6.00.8169)
	MS Windows Common Controls 5 (v.6.00.8022)
	MS DAO 3.51 Object Library

The program was written in VB5 (Service Pack 3)

ASSUMES:	The walls have ears
RETURNS:	Only with valid receipt signed by edward moth

=====================================================================

C INSTRUCTIONS FOR USE:

SETTING UP THE PROJECT (ADD-IN - Ignore for EXE):

1. Load the Project dbCoderAddIn.vbp
2. Open the Immediate Window from the 'View' Menu
3. In the Immediate window, type: AddToINI <Press Enter>
4. Select 'Make dbCoderAddIn.dll...' from the 'File' Menu.
5. Optional (but recommended, unless you plan on changing the
   inteface): Under the Projects Menu, choose DatabaseCoder 
   Properties, choose the Component Tab, and set the project to 
   binary compatibilty with dbCoderAddIn.dll
6. Use regsvr32 to register dbCoderAddin.dll
7. Select 'Add-In Manager' from the 'Add-Ins' Menu
8.a. VB5 Check 'qbd Database Coder' then 'Okay'
  b. VB6 Check 'Loaded/Unloaded' in the 'Load Behaviour' then 'Okay'
9. The Add-In will be available from the Add-Ins Menu
   (listed as 'VB Database Coder')
	
USING THE PROGRAM (ADD-IN and EXE):
Click Open and select an Access97 database file (.mdb).  The database
structure will appear in the left window.  Clicking on an item will
show the Attributes for that item in the right window.

CREATING THE CODE (ADD-IN):
Click the 'Insert Code' button.  The Code to create the database will
be put in the Active Project in a new Standard Module called
'mdbCreator' (if the module exists a procedure will be added to it).

CREATING THE CODE (EXE):
Click the 'Copy Code' to place the code on the Clipboard.

Click the 'Save Module' button to save the code in a standard vb module
file (.bas).


HOW TO USE THE CODE:
You can create the database by adding the code:

bOkay = Database_Create(sFilename)

Where:
sFilename	is the path and name of the Database to be created
bOkay		is the return value (True - creation successful / 
		False - unsuccessful)

ASSUMES:	The database doesn't already exist
		Reference to DAO 3.5x Object Library exists

RETURNS:	True if successful
		False if unsuccessful
		
=====================================================================

D BLAH ... THE BORING BITS

WARRANTIES:
All code is provided 'as is', without warranties of any kind whatsoever
no matter who you say your dad is, even if it is expressed or implied
(the warranties that is, not your dad).

LIABILTY:
qbd software ltd and edward moth accept no liability whatsoever even
if you or your partner get up the duff.  By using this code you accept
that Manchester United are the greatest football team of all time
(okay ... if you're not happy with that, I'll let you off).

LICENSE:
You are free to use and modify any of the code but it would be nice
if you left references to qbd software limited, qbd and edward moth in
place :-) obviously if you nick it and make out it's your own coding
the management and employees of qbd software ltd will be forced to
take serious retaliatory action such as tuttting loudly and calling
you names behind your back.

You can freely distribute the zip and the code but give us credit
(Amex Platinum for preference but 'Thanx to edward moth and qbd
software ltd' would suffice).  If you wish to contant us then please
see the instructions below.

=====================================================================

E CONTACT US:

If you wish to contact edward moth or qbd software then
please follow these instructions:

Categorise your mail:

Mail Type One:	
		* Begging Letters
		* Spam
		* Complaint
		* Abusive
		* Marriage proposal
		* Advertising
		* Porn related
		* Pyramid selling
		* Stupid
		* Contains a virus (we particularly dislike common colds)
		* Contains blank lines because you sent it before you
		  had written it
		* To brag about how much better you are a PSX games
		  than edward (highly unlikely)
		
	or - 	* If you have any contagious medical condition
		* You have an obsession that encompasses your whole life
		* You are William Hague MP

	Send your Mail to: 	trash@qbdsoftware.co.uk
				It will be treated with the utmost
				respect and dealt with promptly

Mail Type Two:

		* Praise
		* Contracts for Tender
		* Job offers (minimum salaries apply - make sure it's
		  good or we'll laugh)
		* Funny
		* Worth reading
		* Exciting Ideas that will make us rich (and that does
		  not include selling perfume)

	or:	* You are Bill Gates regarding 'those share options'
		* You are edward's mom
			

	Send your Mail to:	edward@qbdsoftware.co.uk


