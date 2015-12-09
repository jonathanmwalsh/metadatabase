<%

	' Module to write access permissions
	
	'Now step through accesspermissions and stuff
	m_dataset_id=rs("dataset_id")
	strsql3="select ap.dataset_id, ap.allowdeny, ap.principal, ap.permission FROM accesspermissions ap WHERE ap.dataset_id='" & m_dataset_id & "'"
	'objTextFile.WriteLine strsql3
	set rs3 = Server.CreateObject("ADODB.recordset")
	rs3.Open strsql3, conn
	'objTextFile.WriteLine "dataset accesspermissions open <br />"

	'  <access authSystem="https://pasta.lternet.edu/authentication"
    'order="allowFirst" scope="document" system="https://pasta.lternet.edu">
	
'	objTextFile.WriteLine indent2 & "  <access authSystem=""knb"" order=""allowFirst"" scope=""document"">" 
'	objTextFile.WriteLine indent2 & "<access authSystem=""https://pasta.lternet.edu/authentication"" " & "order=""allowFirst"" scope=""document"" system=""https://pasta.lternet.edu"">" 
	objTextFile.WriteLine indent2 & "<access authSystem=""https://pasta.lternet.edu/authentication"" order=""allowFirst"" scope=""document"" system=""BES"">"
	emergencystop2=0
	do while not rs3.EOF and emergencystop2<2000
		emergencystop2=emergencystop2+1
		'objTextFile.WriteLine "  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"

		'objTextFile.WriteLine "sub-iteration: " & emergencystop2 & " *  "
		objTextFile.WriteLine indent3 & "<" & rs3("allowdeny") & ">" 
		'objTextFile.WriteLine rs3("dataset_id") 
		objTextFile.WriteLine indent4 & "<principal>" & rs3("principal") & "</principal>" 
		objTextFile.WriteLine indent4 & "<permission>" & rs3("permission") & "</permission>" 
		objTextFile.WriteLine indent3 & "</" & rs3("allowdeny") & ">" 

		rs3.movenext
	loop


	' close tag
	objTextFile.WriteLine indent2 & "</access>" 

%>