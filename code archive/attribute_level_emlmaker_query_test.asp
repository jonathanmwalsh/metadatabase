<% 

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
strsql= "SELECT COMBINED.RecordID, COMBINED.dataset_id, COMBINED.id, COMBINED.ExportName, COMBINED.Title, File_Attribute_status.fa_number_rows, File_Attribute_status.fa_number_records, File_Attribute_status.fa_update, File_Attribute_status.fa_version FROM File_Attribute_status LEFT JOIN COMBINED ON File_Attribute_status.dataset_id = COMBINED.dataset_id;"

strsql= "SELECT COMBINED.RecordID, COMBINED.dataset_id, COMBINED.id, COMBINED.ExportName, COMBINED.Title, File_Attribute_status.fa_number_rows, File_Attribute_status.fa_number_records, File_Attribute_status.fa_update, File_Attribute_status.fa_version FROM File_Attribute_status RIGHT JOIN COMBINED ON File_Attribute_status.dataset_id = COMBINED.dataset_id;"

strsql="SELECT COMBINED.RecordID, COMBINED.dataset_id, COMBINED.id, COMBINED.ExportName, COMBINED.Title, File_Attribute_status.fa_number_rows, File_Attribute_status.fa_number_records, File_Attribute_status.fa_update, File_Attribute_status.fa_version, COMBINED.File_Attribute_Type, File_Attribute_table.File_Attribute_Type, File_Attribute_table.File_type FROM (File_Attribute_status RIGHT JOIN COMBINED ON File_Attribute_status.dataset_id = COMBINED.dataset_id) LEFT JOIN File_Attribute_table ON COMBINED.File_Attribute_Type = File_Attribute_table.File_Attribute_Type;"

strsql="SELECT COMBINED.RecordID, COMBINED.dataset_id, COMBINED.id, COMBINED.ExportName, COMBINED.Title COMBINED.File_Attribute_Type FROM COMBINED LEFT JOIN File_Attribute_table ON COMBINED.File_Attribute_Type = File_Attribute_table.File_Attribute_Type"

strsql2=strsql

'######################################################################  TESTING SQL STATEMENT
'######################################################################  TESTING SQL STATEMENT
'######################################################################  TESTING SQL STATEMENT


response.write "Test sql statement in effect!!!! <br />"
response.write "strsql=" & strsql & "<br />"
response.write "strsql2=" & strsql2 & "<br />"


set xrs = Server.CreateObject("ADODB.Recordset")
xrs.Open strsql, xConn, 3, 4  'The 1 and 2 refer to index elements like in line 3 lines above order by xxx + xxx + xxx - WATCH FOR TYPE MISMATCHES!
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



