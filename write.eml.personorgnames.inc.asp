<% 
	
	'1 - DO organization name
	strSQLnames="SELECT n.orgorindividual, n.givenname, n.surname, n.orgname, n.addr1, n.addr2, n.addr3, n.city, n.state, n.zip, n.country , n.orgurl, n.email, n.indid, n.orgid FROM creatornames n WHERE n.orgname=""" & rs("orgname") & """"
	'response.write strSQLnames 

	set rsnames = Server.CreateObject("ADODB.recordset")
	rsnames.Open strSQLnames, conn

	objTextFile.WriteLine indent3 & "<creator id=""" & trim(rsnames("orgid"))  & """>" 
	objTextFile.WriteLine indent3 & "<organizationName>" & rs("orgname") & "</organizationName>" 
	objTextFile.WriteLine indent4 & "<address>" 
	if not trim(rsnames("addr1")) & "VBA SUXORS" = "VBA SUXORS" then ' VBA suxors because I can't simply say if = 0
		objTextFile.WriteLine indent5 & "<deliveryPoint>" & rsnames("addr1") & "</deliveryPoint>" 
	end if
	if not trim(rsnames("addr2")) & "VBA SUXORS" = "VBA SUXORS" then
		objTextFile.WriteLine indent5 & "<deliveryPoint>" & rsnames("addr2") & "</deliveryPoint>" 
	end if
	if not trim(rsnames("addr3")) & "VBA SUXORS" = "VBA SUXORS" then
		objTextFile.WriteLine indent5 & "<deliveryPoint>" & rsnames("addr3") & "</deliveryPoint>" 
	end if
	objTextFile.WriteLine indent5 & "<city>" & rsnames("city") & "</city>" 
	objTextFile.WriteLine indent5 & "<administrativeArea>" & rsnames("state") & "</administrativeArea>" 
	objTextFile.WriteLine indent5 & "<postalCode>" & rsnames("zip") & "</postalCode>" 
	objTextFile.WriteLine indent5 & "<country>" & rsnames("country") & "</country>" 
	objTextFile.WriteLine indent4 & "</address>" 
	if not trim(rsnames("orgurl")) & "VBA SUXORS" = "VBA SUXORS" then
		objTextFile.WriteLine indent4 & "<onlineUrl>" & rsnames("orgurl") & "</onlineUrl>" 
	end if
	objTextFile.WriteLine indent3 & "</creator>" 




	'2 - do personal name
	'rsnames.close
	'objTextFile.WriteLine strSQLnames & "hi<br />"

	strSQLnames="SELECT n.orgorindividual, n.givenname, n.surname, n.orgname, n.addr1, n.addr2, n.addr3, n.city, n.state, n.zip, n.country, n.orgurl, n.email, n.indid, n.orgid  FROM creatornames n WHERE n.givenname=""" & rs("givenname") & """" & " AND n.surname=""" & rs("surname") & """"
	'objTextFile.WriteLine strSQLnames & ";lk;lk;lk<br />"

	set rsnames = Server.CreateObject("ADODB.recordset")
	rsnames.Open strSQLnames, conn
	'objTextFile.WriteLine "made it to here <br />"
	if rsnames.EOF then
		response.write "EOF ######################################## in a bad place: We are in module: write.eml.personorgnames.inc.asp "  & rs("givenname") & ", " & rs("surname")
	end if


	objTextFile.WriteLine indent3 & "<creator id=""" & trim(rsnames("indid")) & """>" 
	objTextFile.WriteLine indent3 & "<organizationName>" & rs("orgname") & "</organizationName>" 
	objTextFile.WriteLine indent4 & "<individualName>" 
	objTextFile.WriteLine indent5 & "<givenName>" & rs("givenname") & "</givenName>" 
	objTextFile.WriteLine indent5 & "<surName>" & rs("surname") & "</surName>" 
	objTextFile.WriteLine indent4 & "</individualName>" 

	objTextFile.WriteLine indent4 & "<address>" 
	if not trim(rsnames("addr1")) & "VBA SUXORS" = "VBA SUXORS" then ' VBA suxors because I can't simply say if = 0
		objTextFile.WriteLine indent5 & "<deliveryPoint>" & rsnames("addr1") & "</deliveryPoint>" 
	end if
	if not trim(rsnames("addr2")) & "VBA SUXORS" = "VBA SUXORS" then
		objTextFile.WriteLine indent5 & "<deliveryPoint>" & rsnames("addr2") & "</deliveryPoint>" 
	end if
	if not trim(rsnames("addr3")) & "VBA SUXORS" = "VBA SUXORS" then
		objTextFile.WriteLine indent5 & "<deliveryPoint>" & rsnames("addr3") & "</deliveryPoint>" 
	end if
	objTextFile.WriteLine indent5 & "<city>" & rsnames("city") & "</city>" 
	objTextFile.WriteLine indent5 & "<administrativeArea>" & rsnames("state") & "</administrativeArea>" 
	objTextFile.WriteLine indent5 & "<postalCode>" & rsnames("zip") & "</postalCode>" 
	objTextFile.WriteLine indent5 & "<country>" & rsnames("country") & "</country>" 
	objTextFile.WriteLine indent4 & "</address>" 
	objTextFile.WriteLine indent4 & "<electronicMailAddress>" & rsnames("email") & "</electronicMailAddress>" 
	'response.write indent4 & "MADE IT THROUGH THE SUBMODULE PERSONORGNAMES" 
	objTextFile.WriteLine indent3 & "</creator>" 


%>



