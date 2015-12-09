<% 

strsql="SELECT main.RecordID, main.dataset_id, main.id, main.ExportName, main.Title, File_Attribute_status.fa_number_rows, File_Attribute_status.fa_number_records, File_Attribute_status.fa_update, File_Attribute_status.fa_version, main.File_Attribute_Type, File_Attribute_table.File_Attribute_Type, File_Attribute_table.File_type FROM (File_Attribute_status RIGHT JOIN COMBINED ON File_Attribute_status.dataset_id = main.dataset_id) LEFT JOIN File_Attribute_table ON main.File_Attribute_Type = File_Attribute_table.File_Attribute_Type;"


'strsql="SELECT main.dataset_id FROM COMBINED LEFT JOIN File_Attribute_table ON main.dataset_id=File_Attribute_table.dataset_id"

strsql="SELECT * FROM COMBINED LEFT JOIN File_Attribute_table ON main.RecordID = File_Attribute_table.dataset_id"


' MAIN CONNECTION

strsql="SELECT main.RecordID, main.dataset_id, main.id, main.ExportName, main.Title,  main.File_Attribute_Type, File_Attribute_table.File_Attribute_Type, File_Attribute_table.File_type, File_Attribute_status.fa_update FROM ((COMBINED main LEFT JOIN File_Attribute_table ON main.dataset_id = File_Attribute_table.dataset_id) LEFT JOIN File_Attribute_status ON main.dataset_id=File_Attribute_status.dataset_id)"

' STATUS FILE CONECTION (TO GET ACCESS RULES, # COLUMNS, REVISION DATE, EC

strsql2="SELECT main2.dataset_id, accesspermissions.principal FROM COMBINED main2 LEFT JOIN accesspermissions ON main2.dataset_id=accesspermissions.dataset_id"

set conn=Server.CreateObject("ADODB.Connection")
conn.Provider="Microsoft.Jet.OLEDB.4.0"
conn.Open "c:/inetpub/wwwroot/attribute_level_emlmaker/esri2bes_int_osrs-search_jwalsh.mdb"


'OPEN MAIN DATASET
set rs = Server.CreateObject("ADODB.recordset")
rs.Open strSQL, conn

'OPEN STATUS DATASET
set rs2 = Server.CreateObject("ADODB.recordset")
rs2.Open strSQL2, conn


rs.movefirst

emergencystop=0

do while not rs.EOF and emergencystop<2000

	emergencystop=emergencystop+1
	response.write "Dataset_id= " & rs("dataset_id") & " * " 
	response.write ", emergencystop = " & emergencystop & " * "

	response.write "Record id= " & rs("RecordID")  & " * " 
	response.write ", FA_version= " & rs("fa_update") & " * " 
	response.write ", File Type= " & rs("file_type") & " * "   
	response.write "EOF 1 " & rs.EOF & " EOF 2 " & rs2.EOF
	response.write  " <br/>" 


' Notes 2013-06-01 - With the statements below commented out, it does not crash, BUT with a two record sample datafile, it iterates 5x, NOT 2x.

	'if rs2.EOF then
	'	'nothing
	'else
	'	m_dataset_id=rs2("dataset_id")
	'	do while rs2("dataset_id")=m_dataset_id
	'		response.write " ***  Current new record RS2 dataset_id=***" & rs2("dataset_id") & "****  "
	'		response.write ",principal = " & rs2("principal")
	'		response.write  " <br/>" 
	'		rs2.movenext
	'	loop
	'end if
		

	
	
	if rs.EOF then
		response.write "EOF ########################################"
	end if
	rs.movenext



'response.write("1111")
'response.write("2222")



loop

%> 

'Example from http://www.w3schools.com/ado/met_rs_open.asp
'<    
'set conn=Server.CreateObject("ADODB.Connection")
'conn.Provider="Microsoft.Jet.OLEDB.4.0"
'conn.Open "c:/webdata/northwind.mdb"
'set rs = Server.CreateObject("ADODB.recordset")
'rs.Open "Select * from Customers", conn
'%> 

'Open connection to database
accessdb="esri2bes_int_osrs-search_jwalsh.mdb"                 ' Just setting a variable to the name
db=Server.MapPath(accessdb)                    ' JW enhanced variables
'dbGenericPath = "/emlmaker_attribute_level_eml/"
dbConn = "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & db  'my  way


'Open main data recordset
'dbRs = "COMBINED" 'Name of table.
strConn = dbConn
strTable = dbRs
' Open Connection to the database
set xConn = Server.CreateObject("ADODB.Connection")
xConn.Open strConn


''Open Second data recordset
'dbRs2 = "COMBINED" 'name of table
'strTable2 = dbRs2


'######################################################################  TESTING SQL STATEMENT
'######################################################################  TESTING SQL STATEMENT
'######################################################################  TESTING SQL STATEMENT
strsql= "SELECT main.RecordID, main.dataset_id, main.id, main.ExportName, main.Title, File_Attribute_status.fa_number_rows, File_Attribute_status.fa_number_records, File_Attribute_status.fa_update, File_Attribute_status.fa_version FROM File_Attribute_status LEFT JOIN COMBINED ON File_Attribute_status.dataset_id = main.dataset_id;"

strsql= "SELECT main.RecordID, main.dataset_id, main.id, main.ExportName, main.Title, File_Attribute_status.fa_number_rows, File_Attribute_status.fa_number_records, File_Attribute_status.fa_update, File_Attribute_status.fa_version FROM File_Attribute_status RIGHT JOIN COMBINED ON File_Attribute_status.dataset_id = main.dataset_id;"

strsql="SELECT main.RecordID, main.dataset_id, main.id, main.ExportName, main.Title, File_Attribute_status.fa_number_rows, File_Attribute_status.fa_number_records, File_Attribute_status.fa_update, File_Attribute_status.fa_version, main.File_Attribute_Type, File_Attribute_table.File_Attribute_Type, File_Attribute_table.File_type FROM (File_Attribute_status RIGHT JOIN COMBINED ON File_Attribute_status.dataset_id = main.dataset_id) LEFT JOIN File_Attribute_table ON main.File_Attribute_Type = File_Attribute_table.File_Attribute_Type;"

'strsql="SELECT main.RecordID, main.dataset_id, main.id, main.ExportName, main.Title main.File_Attribute_Type FROM COMBINED LEFT JOIN File_Attribute_table ON main.File_Attribute_Type = File_Attribute_table.File_Attribute_Type;"

strsql2=strsql

'######################################################################  TESTING SQL STATEMENT
'######################################################################  TESTING SQL STATEMENT
'######################################################################  TESTING SQL STATEMENT


response.write "Test sql statement in effect!!!! <br />"
response.write "strsql=" & strsql & "<br />"
response.write "strsql2=" & strsql2 & "<br />"


set xrs = Server.CreateObject("ADODB.Recordset")
xrs.Open strsql, xConn, 1, 2  'The 1 and 2 refer to index elements like in line 3 lines above order by xxx + xxx + xxx - WATCH FOR TYPE MISMATCHES!
xrs.movefirst

set xrs2 = Server.CreateObject("ADODB.Recordset")
xrs2.Open strsql2, xConn, 1, 2  'The 1 and 2 refer to index elements like in line 3 lines above order by xxx + xxx + xxx - WATCH FOR TYPE MISMATCHES!
xrs2.movefirst

response.write "#####################File should be open now##################"


dim objFSO
Set objFSO = server.CreateObject("Scripting.FileSystemObject")
dim objTextFile 'There are many of these, these will be the xml files we write
dim objTextFile2 'This will be the file harvestlist.xml
dim objTextFile3 'This will be the html page the public will see.
dim objTextFile4 'There are many of these, these will be the html files we write for the full metadata records
Set objTextFile = objFSO.CreateTextFile(server.mappath("emloutput.txt"), true)
Set objTextFile2 = objFSO.CreateTextFile("c:\inetpub\wwwroot\metacat_harvest_attribute_level_eml\harvestlist.xml", true)
Set objTextFile3 = objFSO.CreateTextFile("c:\inetpub\wwwroot\metacat_harvest_attribute_level_eml\frame7-page_1_auto.asp", true)
objTextFile2.WriteLine "<?xml version=""1.0"" encoding=""UTF-8"" ?>"

objTextFile2.WriteLine "<hrv:harvestList xmlns:hrv=""eml://ecoinformatics.org/harvestList"" >"



Const Filename = "/metacat_harvest_attribute_level_eml/lter_revision_no.txt"	' file to read
Const ForReading = 1, ForWriting = 2, ForAppending = 3
Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0




xrs.requery
xxx= xrs.eof
response.write " EOF EQUALS:"&xxx&"***"
xrs.movefirst
fname=0
mRecordID = 0
emergencystop=0

do while not xrs.EOF 'and emergencystop<20

	emergencystop=emergencystop+1
	'fname=fname+1
	mRecordID=xrs("RecordID")
	response.write "Record id= " & xrs("RecordID")
	response.write ", Dataset_id= " & xrs("dataset_id")
	response.write ", FA_version= " & xrs("fa_update")
	response.write ", File Type= " & xrs("file_type") & " <br/>" 

	xrs.movenext

'	response.write "** END OF RECORD" & "<br>" & Vbcrlf & "<br>" & Vbcrlf
	
loop


'************************************************
'************************************************
'************************************************
'************************************************ END LOOP TO CREATE XML FILES FOR COMBINED DATA



set xConn = Nothing
set xrs = Nothing
objTextFile.Close
Set objTextFile = Nothing
Set objFSO = Nothing


objTextFile2.Close
Set objTextFile2 = Nothing



Set fname = Nothing
Set mRecordID = Nothing
Set cname = Nothing
Set tfname = Nothing
Set harvestlistname = Nothing
Set emergencystop = Nothing



Set accessdb = Nothing
Set db = Nothing
Set dbGenericPath = Nothing
Set dbConn = Nothing
Set dbRs = Nothing
Set strConn = Nothing
Set strTable = Nothing
Set xConn = Nothing




%>



