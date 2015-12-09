<% 
	'1 - DO organization name
	strSQLnames="SELECT n.orgorindividual, n.givenname, n.surname, n.orgname, n.addr1, n.addr2, n.addr3, n.city, n.state, n.zip, n.country , n.orgurl, n.email FROM creatornames n WHERE n.orgname=""" & rs("orgname") & """"
	response.write strSQLnames & "<br />"

	set rsnames = Server.CreateObject("ADODB.recordset")
	rsnames.Open strSQLnames, conn

	response.write indent4 & "&LT;address>" & "<br />"
	response.write indent5 & "&LT;deliverypoint>" & rsnames("addr1") & "&LT;/deliverypoint>" & "<br />"
	response.write indent5 & "&LT;deliverypoint>" & rsnames("addr2") & "&LT;/deliverypoint>" & "<br />"
	response.write indent5 & "&LT;deliverypoint>" & rsnames("addr3") & "&LT;/deliverypoint>" & "<br />"
	response.write indent5 & "&LT;city>" & rsnames("city") & "&LT;/city>" & "<br />"
	response.write indent5 & "&LT;administrativeArea>" & rsnames("state") & "&LT;/administrativeArea>" & "<br />"
	response.write indent5 & "&LT;postalCode>" & rsnames("zip") & "&LT;/postalCode>" & "<br />"
	response.write indent5 & "&LT;country>" & rsnames("country") & "&LT;/country>" & "<br />"
	response.write indent4 & "&LT;/address>" & "<br />"
	response.write indent4 & "&LT;onlineUrl>" & rsnames("orgurl") & "&LT;/onlineUrl>" & "<br />"

	'2 - do personal name
	'rsnames.close
	'response.write strSQLnames & "hi<br />"

	strSQLnames="SELECT n.orgorindividual, n.givenname, n.surname, n.orgname, n.addr1, n.addr2, n.addr3, n.city, n.state, n.zip, n.country, n.orgurl, n.email  FROM creatornames n WHERE n.givenname=""" & rs("givenname") & """" & " AND n.surname=""" & rs("surname") & """"
	'response.write strSQLnames & ";lk;lk;lk<br />"

	set rsnames = Server.CreateObject("ADODB.recordset")
	rsnames.Open strSQLnames, conn
	'response.write "made it to here <br />"
	if rsnames.EOF then
		response.write "EOF ########################################"
	end if


	response.write indent4 & "&LT;individualName>" & "<br />"
	response.write indent5 & "&LT;givenName>" & rs("givenname") & "&LT;/givenName>" & "<br />"
	response.write indent5 & "&LT;surName>" & rs("surname") & "&LT;/surName>" & "<br />"
	response.write indent4 & "&LT;/individualName>" & "<br />"
	'response.write "hello"

	response.write indent4 & "&LT;address>" & "<br />"
	response.write indent5 & "&LT;deliverypoint>" & rsnames("addr1") & "&LT;/deliverypoint>" & "<br />"
	response.write indent5 & "&LT;deliverypoint>" & rsnames("addr2") & "&LT;/deliverypoint>" & "<br />"
	response.write indent5 & "&LT;deliverypoint>" & rsnames("addr3") & "&LT;/deliverypoint>" & "<br />"
	response.write indent5 & "&LT;city>" & rsnames("city") & "&LT;/city>" & "<br />"
	response.write indent5 & "&LT;administrativeArea>" & rsnames("state") & "&LT;/administrativeArea>" & "<br />"
	response.write indent5 & "&LT;postalCode>" & rsnames("zip") & "&LT;/postalCode>" & "<br />"
	response.write indent5 & "&LT;country>" & rsnames("country") & "&LT;/country>" & "<br />"
	response.write indent4 & "&LT;/address>" & "<br />"
	response.write indent4 & "&LT;electronicMailAddress>" & rsnames("email") & "&LT;/electronicMailAddress>" & "<br />"
	'response.write indent4 & "MADE IT THROUGH THE SUBMODULE PERSONORGNAMES" & "<br />"


%>



