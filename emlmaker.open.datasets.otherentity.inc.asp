<% 

strsql="SELECT COMBINED.RecordID, COMBINED.dataset_id, COMBINED.edition, COMBINED.id, COMBINED.ExportName, COMBINED.Title, COMBINED.File_Attribute_Type, COMBINED.creatorid, COMBINED.orgname, COMBINED.surname, COMBINED.givenname, COMBINED.themekeywords, COMBINED.placekeywords, COMBINED.abstract, COMBINED.publicationdate, COMBINED.datacred, COMBINED.onlinelinkage, COMBINED.onlineinfolink, COMBINED.geographicdescription, COMBINED.west, COMBINED.east, COMBINED.north, COMBINED.south, COMBINED.temporalbegin, COMBINED.temporalend, COMBINED.publisher, COMBINED.datacred, COMBINED.filename, COMBINED.filesizeunit, COMBINED.filesize, COMBINED.characterencoding, COMBINED.dataformat, COMBINED.numheaderlines, COMBINED.numfooterlines, COMBINED.recorddelimiter, COMBINED.linesperrecord, COMBINED.fielddelimiter, COMBINED.quotecharacter, COMBINED.orientation, COMBINED.onlinedescription, COMBINED.pastaview, COMBINED.dataformat, COMBINED.filetype, File_Attribute_table.File_Attribute_Type, File_Attribute_table.file_type, File_Attribute_status.fa_update FROM ((COMBINED LEFT JOIN File_Attribute_table ON COMBINED.dataset_id = File_Attribute_table.dataset_id) LEFT JOIN File_Attribute_status ON COMBINED.dataset_id = File_Attribute_status.dataset_id) WHERE pastaview=1 and DataEntityType=2 ORDER BY COMBINED.dataset_id"

strSQLnames="SELECT n.orgorindividual, n.givenname, n.surname, n.orgname, n.addr1, n.addr2, n.addr3, n.city, n.state, n.zip, n.country  FROM creatornames n"

' MAIN CONNECTION

'works works works strsql="SELECT main.RecordID, main.dataset_id, main.id, main.ExportName, main.Title,  main.File_Attribute_Type, File_Attribute_table.File_Attribute_Type, File_Attribute_table.File_type, File_Attribute_status.fa_update FROM ((COMBINED main LEFT JOIN File_Attribute_table ON main.dataset_id = File_Attribute_table.dataset_id) LEFT JOIN File_Attribute_status ON main.dataset_id=File_Attribute_status.dataset_id)"

'works strsql="SELECT COMBINED.RecordID, COMBINED.dataset_id, COMBINED.id, COMBINED.ExportName, COMBINED.Title,  COMBINED.File_Attribute_Type, File_Attribute_table.File_Attribute_Type, File_Attribute_table.File_type, File_Attribute_status.fa_update FROM ((COMBINED LEFT JOIN File_Attribute_table ON COMBINED.dataset_id = File_Attribute_table.dataset_id) LEFT JOIN File_Attribute_status ON COMBINED.dataset_id=File_Attribute_status.dataset_id)"

' STATUS FILE CONECTION (TO GET ACCESS RULES, # COLUMNS, REVISION DATE, EC

strsql2="SELECT main2.dataset_id, accesspermissions.principal FROM COMBINED main2 LEFT JOIN accesspermissions ON main2.dataset_id=accesspermissions.dataset_id"

set conn=Server.CreateObject("ADODB.Connection")
conn.Provider="Microsoft.Jet.OLEDB.4.0"
conn.Open "c:/inetpub/wwwroot/attribute_level_emlmaker/metadatabase.mdb"


'OPEN MAIN DATASET
set rs = Server.CreateObject("ADODB.recordset")
rs.Open strSQL, conn

'OPEN STATUS DATASET
set rs2 = Server.CreateObject("ADODB.recordset")
rs2.Open strSQL2, conn

'OPEN NAMES DATASET
set rsnames = Server.CreateObject("ADODB.recordset")
rsnames.Open strSQLnames, conn

'OPEN METHODS AND METHOD LINK DATASETS
'strSQLmethods="SELECT m.ID, m.methodid, m.methodname, m.methoddescription FROM methods m"
'set rsmethods =  Server.CreateObject("ADODB.recordset")
'rsmethods.Open strSQLmethods, conn

'strSQLmethodlink="SELECT ml.methodid, ml.datasetid FROM Methodlink ml"
'set rsmethodlink =  Server.CreateObject("ADODB.recordset")
'rsmethodlink.Open strSQLmethodlink, conn







'response.write "dataset open <br />"



%>



