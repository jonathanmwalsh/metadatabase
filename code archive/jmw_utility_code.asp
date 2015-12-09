<% 

strsql="SELECT COMBINED.RecordID, COMBINED.dataset_id, COMBINED.child_of_multi_id_dataset, COMBINED.part_of_multi_id_dataset, COMBINED.child_of_multi_id_dataset, COMBINED.edition, COMBINED.id, COMBINED.ExportName, COMBINED.Title, COMBINED.File_Attribute_Type, COMBINED.creatorid, COMBINED.orgname, COMBINED.surname, COMBINED.givenname, COMBINED.themekeywords, COMBINED.placekeywords, COMBINED.abstract, COMBINED.publicationdate, COMBINED.datacred, COMBINED.onlinelinkage, COMBINED.onlineinfolink, COMBINED.geographicdescription, COMBINED.west, COMBINED.east, COMBINED.north, COMBINED.south, COMBINED.temporalbegin, COMBINED.temporalend, COMBINED.publisher, COMBINED.datacred, COMBINED.filename, COMBINED.filesizeunit, COMBINED.filesize, COMBINED.characterencoding, COMBINED.dataformat, COMBINED.numheaderlines, COMBINED.numfooterlines, COMBINED.recorddelimiter, COMBINED.linesperrecord, COMBINED.fielddelimiter, COMBINED.quotecharacter, COMBINED.orientation, COMBINED.onlinedescription, COMBINED.pastaview, File_Attribute_table.File_Attribute_Type, File_Attribute_table.file_type, File_Attribute_status.fa_update FROM ((COMBINED LEFT JOIN File_Attribute_table ON COMBINED.dataset_id = File_Attribute_table.dataset_id) LEFT JOIN File_Attribute_status ON COMBINED.dataset_id = File_Attribute_status.dataset_id) WHERE pastaview=1 and DataEntityType=1 and  child_of_multi_id_dataset<1 ORDER BY COMBINED.dataset_id"
'The "child of's" are in multi-id records, and we set it to zero for just thej first one so we dojnt skip past eof in the sub loop.

response.write strsql

'strSQLnames="SELECT n.orgorindividual, n.givenname, n.surname, n.orgname, n.addr1, n.addr2, n.addr3, n.city, n.state, n.zip, n.country  FROM creatornames n"


set conn=Server.CreateObject("ADODB.Connection")
conn.Provider="Microsoft.Jet.OLEDB.4.0"
conn.Open "c:/inetpub/wwwroot/attribute_level_emlmaker/esri2bes_int_osrs-search_jwalsh.mdb"


'OPEN MAIN DATASET
'set rs = Server.CreateObject("ADODB.recordset")
'rs.Open strSQL, conn

'OPEN STATUS DATASET
'set rs2 = Server.CreateObject("ADODB.recordset")
'rs2.Open strSQL2, conn

'OPEN NAMES DATASET
'set rsnames = Server.CreateObject("ADODB.recordset")
'rsnames.Open strSQLnames, conn


'updatenamessql="UPDATE COMBINED SET COMBINED.publisher = [COMBINED].[givenname] & ' ' & [COMBINED].[surname]"
'set rsupdatenames = Server.CreateObject("ADODB.recordset")
'rsupdatenames.Open updatenamessql, conn

'select * into COMBINED_TEST from COMBINED_FULL_OLD_STRUCTURE 
'WHERE dataset_id="BES_589" 'doesn't work with SELECT *

INSERT INTO c:\inetpub\wwwroot\attribute_level_emlmaker\esri2bes_int_osrs-search_jwalsh.mdb.COMBINED_TEST (SELECT * FROM c:\inetpub\wwwroot\attribute_level_emlmaker\esri2bes_int_osrs-search_jwalsh.mdb.COMBINED_FULL_OLD_STRUCTURE WHERE 2+2=4)



response.write "<br />Done <br />"




%>



