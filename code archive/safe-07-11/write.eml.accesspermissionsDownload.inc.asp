<%

	' Module to write access permissions for when dataset is downloaded
	
	'Now step through accesspermissions and stuff
	m_dataset_id=rs("dataset_id")
	strsqldownload="select ap.dataset_id, ap.allowdeny, ap.principal, ap.permission FROM accesspermissions ap WHERE ap.dataset_id='" & m_dataset_id & "'"
	'response.write strsql3
	set rsdownload = Server.CreateObject("ADODB.recordset")
	rsdownload.Open strsqldownload, conn
	'response.write "dataset accesspermissions open <br />"

	'  <access authSystem="https://pasta.lternet.edu/authentication"
    'order="allowFirst" scope="document" system="https://pasta.lternet.edu">
	
	response.write indent8 & "  &LT;access authSystem=&quot;knb&quot; order=&quot;allowFirst&quot; scope=&quot;document&quot;&GT;" & "<br />"
	
	emergencystopdnld=0
	do while not rsdownload.EOF and emergencystopdnld<2000
		emergencystopdnld=emergencystopdnld+1
		'response.write "  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"

		'response.write "sub-iteration: " & emergencystop2 & " *  "
		response.write indent9 & "&LT;" & rsdownload("allowdeny") & "&GT;" & "<br />"
		'response.write rs3("dataset_id") & "<br />"
		response.write indent10 & "&LT;" & rsdownload("principal") & "&GT;" & "<br />"
		response.write indent10 & "&LT;" & rsdownload("permission") & "&GT;" & "<br />"
		response.write indent9 & "&LT;/" & rsdownload("allowdeny") & "&GT;" & "<br />"

		rsdownload.movenext
	loop


	' close tag
	response.write indent8 & "&LT;/access&GT;" & "<br />"
	response.write indent7 & "&LT;/distribution&GT;" & "<br />"
	response.write indent6 & "&LT;/physical&GT;" & "<br />"




%>

