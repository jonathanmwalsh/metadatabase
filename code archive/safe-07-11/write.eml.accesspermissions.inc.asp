<%

	' Module to write access permissions
	
	'Now step through accesspermissions and stuff
	m_dataset_id=rs("dataset_id")
	strsql3="select ap.dataset_id, ap.allowdeny, ap.principal, ap.permission FROM accesspermissions ap WHERE ap.dataset_id='" & m_dataset_id & "'"
	'response.write strsql3
	set rs3 = Server.CreateObject("ADODB.recordset")
	rs3.Open strsql3, conn
	'response.write "dataset accesspermissions open <br />"

	'  <access authSystem="https://pasta.lternet.edu/authentication"
    'order="allowFirst" scope="document" system="https://pasta.lternet.edu">
	
	response.write indent2 & "  &LT;access authSystem=&quot;https://pasta.lternet.edu/authentication&quot;" & "<br />" & indent2 & "order=&quot;allowFirst&quot; scope=&quot;document&quot; system=&quot;https://pasta.lternet.edu&quot;&GT;" & "<br />"
	
	emergencystop2=0
	do while not rs3.EOF and emergencystop2<2000
		emergencystop2=emergencystop2+1
		'response.write "  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"

		'response.write "sub-iteration: " & emergencystop2 & " *  "
		response.write indent3 & "&LT;" & rs3("allowdeny") & "&GT;" & "<br />"
		'response.write rs3("dataset_id") & "<br />"
		response.write indent4 & "&LT;" & rs3("principal") & "&GT;" & "<br />"
		response.write indent4 & "&LT;" & rs3("permission") & "&GT;" & "<br />"
		response.write indent3 & "&LT;/" & rs3("allowdeny") & "&GT;" & "<br />"

		rs3.movenext
	loop


	' close tag
	response.write indent2 & "&LT;access&GT;" & "<br />"

%>