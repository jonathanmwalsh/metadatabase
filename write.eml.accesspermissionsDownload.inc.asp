<%

	' Module to write access permissions for when dataset is downloaded
	
	'Now step through accesspermissions and stuff
	m_dataset_id=rs("dataset_id")
	strsqldownload="select ap.dataset_id, ap.allowdeny, ap.principal, ap.permission FROM accesspermissions ap WHERE ap.dataset_id='" & m_dataset_id & "'"
	'objTextFile.WriteLine strsql3
	set rsdownload = Server.CreateObject("ADODB.recordset")
	rsdownload.Open strsqldownload, conn
	'objTextFile.WriteLine "dataset accesspermissions open <br />"

	'  <access authSystem="https://pasta.lternet.edu/authentication"
    'order="allowFirst" scope="document" system="https://pasta.lternet.edu">
	
	objTextFile.WriteLine indent8 & "  <access authSystem=""knb"" order=""allowFirst"" scope=""document"">" 
'	objTextFile.WriteLine indent8 & "<access authSystem=""https://pasta.lternet.edu/authentication"" "  & "order=""allowFirst"" scope=""document"" system=""https://pasta.lternet.edu"">" 
'	objTextFile.WriteLine indent8 & "<access authSystem=""https://pasta.lternet.edu/authentication"" order=""allowFirst"" scope=""document"" system=""BES"">"
	
	emergencystopdnld=0
	do while not rsdownload.EOF and emergencystopdnld<2000
		emergencystopdnld=emergencystopdnld+1
		'objTextFile.WriteLine "  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"

		'objTextFile.WriteLine "sub-iteration: " & emergencystop2 & " *  "
		objTextFile.WriteLine indent9 & "<" & rsdownload("allowdeny") & ">" 
		'objTextFile.WriteLine rs3("dataset_id") 
		objTextFile.WriteLine indent10 & "<principal>" & rsdownload("principal") & "</principal>" 
		objTextFile.WriteLine indent10 & "<permission>" & rsdownload("permission") & "</permission>" 
		objTextFile.WriteLine indent9 & "</" & rsdownload("allowdeny") & ">" 

		rsdownload.movenext
	loop


	' close tag
	objTextFile.WriteLine indent8 & "</access>" 



%>

