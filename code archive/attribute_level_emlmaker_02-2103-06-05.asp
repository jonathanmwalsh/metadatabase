<% 

strsql="SELECT main.RecordID, main.dataset_id, main.id, main.ExportName, main.Title, File_Attribute_status.fa_number_rows, File_Attribute_status.fa_number_records, File_Attribute_status.fa_update, File_Attribute_status.fa_version, main.File_Attribute_Type, File_Attribute_table.File_Attribute_Type, File_Attribute_table.File_type FROM (File_Attribute_status RIGHT JOIN COMBINED ON File_Attribute_status.dataset_id = main.dataset_id) LEFT JOIN File_Attribute_table ON main.File_Attribute_Type = File_Attribute_table.File_Attribute_Type;"


'strsql="SELECT main.dataset_id FROM COMBINED LEFT JOIN File_Attribute_table ON main.dataset_id=File_Attribute_table.dataset_id"

strsql="SELECT COMBINED.dataset_id, COMBINED.RecordID, File_Attribute_table.file_type FROM COMBINED LEFT JOIN File_Attribute_table ON COMBINED.dataset_id = File_Attribute_table.dataset_id"   'WORKS WORKS

strsql="SELECT COMBINED.RecordID, COMBINED.dataset_id, COMBINED.id, COMBINED.ExportName, COMBINED.Title, COMBINED.File_Attribute_Type, File_Attribute_table.File_Attribute_Type, File_Attribute_table.file_type, File_Attribute_status.fa_update FROM ((COMBINED LEFT JOIN File_Attribute_table ON COMBINED.dataset_id = File_Attribute_table.dataset_id) LEFT JOIN File_Attribute_status ON COMBINED.dataset_id = File_Attribute_status.dataset_id)"

' MAIN CONNECTION

'works works works strsql="SELECT main.RecordID, main.dataset_id, main.id, main.ExportName, main.Title,  main.File_Attribute_Type, File_Attribute_table.File_Attribute_Type, File_Attribute_table.File_type, File_Attribute_status.fa_update FROM ((COMBINED main LEFT JOIN File_Attribute_table ON main.dataset_id = File_Attribute_table.dataset_id) LEFT JOIN File_Attribute_status ON main.dataset_id=File_Attribute_status.dataset_id)"

'works strsql="SELECT COMBINED.RecordID, COMBINED.dataset_id, COMBINED.id, COMBINED.ExportName, COMBINED.Title,  COMBINED.File_Attribute_Type, File_Attribute_table.File_Attribute_Type, File_Attribute_table.File_type, File_Attribute_status.fa_update FROM ((COMBINED LEFT JOIN File_Attribute_table ON COMBINED.dataset_id = File_Attribute_table.dataset_id) LEFT JOIN File_Attribute_status ON COMBINED.dataset_id=File_Attribute_status.dataset_id)"

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

response.write "dataset open <br />"

rs.movefirst

emergencystop=0


'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXX                                                                                  XXXX
'XXXX                                                                                  XXXX
'XXXX                                                                                  XXXX
'XXXX        A L L     C O D E     W O R K S                                           XXXX
'XXXX                                                                                  XXXX
'XXXX                                                                                  XXXX
'XXXX                                                                                  XXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX




do while not rs.EOF and emergencystop<2000

	emergencystop=emergencystop+1
	response.write "Dataset_id= " & rs("dataset_id") & " * " 
	'response.write "      ***   here  **      "
	response.write ", Iteration = " & emergencystop & " * "

	response.write "Record id= " & rs("RecordID")  & " * " 
	'response.write ", FA_version= " & rs("fa_update") & " * " 
	response.write ", File Type= " & rs("file_type") & " * "   
	'response.write "EOF 1 " & rs.EOF & " EOF 2 " & rs2.EOF
	response.write  " <br/>" 

	'Now step through accesspermissions and stuff
	m_dataset_id=rs("dataset_id")
	strsql3="select ap.dataset_id, ap.allowdeny, ap.principal, ap.permission FROM accesspermissions ap WHERE ap.dataset_id='" & m_dataset_id & "'"
	'response.write strsql3
	set rs3 = Server.CreateObject("ADODB.recordset")
	rs3.Open strsql3, conn
	'response.write "dataset accesspermissions open <br />"
	emergencystop2=0
	do while not rs3.EOF and emergencystop2<2000
		emergencystop2=emergencystop2+1
		response.write "sub-iteration: " & emergencystop2 & " *  "
		response.write rs3("dataset_id") & ", " & rs3("allowdeny") & ", " & rs3("principal") & ", " & rs3("permission") & "<br />"
		rs3.movenext
	loop

	
	
	rs.movenext
	if rs.EOF then
		response.write "EOF ########################################"
	end if





loop

%>



