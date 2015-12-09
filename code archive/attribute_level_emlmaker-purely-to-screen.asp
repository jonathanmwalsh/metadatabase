<% 

indent1="&nbsp;&nbsp;"
indent2=indent1 & "&nbsp;&nbsp;"
indent3=indent2 & "&nbsp;&nbsp;"
indent4=indent3 & "&nbsp;&nbsp;"
indent5=indent4 & "&nbsp;&nbsp;"
indent6=indent5 & "&nbsp;&nbsp;"
indent7=indent6 & "&nbsp;&nbsp;"
indent8=indent7 & "&nbsp;&nbsp;"
indent9=indent8 & "&nbsp;&nbsp;"
indent10=indent9 & "&nbsp;&nbsp;"
indent11=indent10 & "&nbsp;&nbsp;"
indent12=indent11 & "&nbsp;&nbsp;"

'Make database connections and
'Open main datasets
	%>
	<!--#include file="emlmaker.open.datasets.inc.asp"-->	
	<%


'rs.movelast 'LOL No such command
'find last dataset for file positioning purposes
do while not rs.eof
	mdatasetid=rs("dataset_id")
	rs.movenext
loop
response.write "Last dataset ID in recordset is: " & mdatasetid
rs.movefirst

emergencystop=0


'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXX                                                                                  XXXX
'XXXX                                                                                  XXXX
'XXXX                                                                                  XXXX
'XXXX        A L L     C O D E                                                         XXXX
'XXXX                                                                                  XXXX
'XXXX                                                                                  XXXX
'XXXX                                                                                  XXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX



do while not rs.EOF and emergencystop<2000

	emergencystop=emergencystop+1

	response.write "<br />&nbsp;<br />&nbsp;"


	'Write EML top of page - Header
	%>
	<!--#include file="write.eml.header.inc.asp"-->	
	<%
	
	
	
	'HELPS DEBUGGING - FIGURING OUT WHERE YOU ARE IN THE RECORD SET
	response.write "********************************************************************************************************** Dataset_id= " & rs("dataset_id") & " * " 
	response.write  " <br/>" 


	response.write "Record id= " & rs("RecordID")  & " * " 
	response.write ", File Type= " & rs("file_type") & " * "   
	'response.write "EOF 1 " & rs.EOF & " EOF 2 " & rs2.EOF

	response.write  " <br/>" 
		response.write ", FA_version= " & rs("fa_update") & " * " 
		response.write  " <br/>" 




	'Now step through accesspermissions and stuff
	%>
	<!--#include file="write.eml.accesspermissions.inc.asp"-->	
	<%
	
	

	
	'Open <dataset> tag
	response.write indent2 & "&LT;dataset scope=""document"">" & "<br />"  'I don't know why using a scope identifier (scope = document) but I see it on examples I am using

	'Title of dataset
	response.write indent3 & "&LT;title>" & rs("title") & "&LT;/title>" & "<br />"



	'creator nodes
	response.write indent3 & "&LT;creator id=""" & rs("creatorid") & ">" & "<br />"

	response.write indent3 & "&LT;organizationName>" & rs("orgname") & "&LT;/organizationName>" & "<br />"
	
	'OPEN NAMES DATASET, do personal name, do organisations name, 
	%>
	<!--#include file="write.eml.personorgnames.inc.asp"-->	
	<%

	'PUBDATE  
	'#########################################
	'###### Be aware pubdate for me is yyyy/mm/dd and eml calls for yyyy-mm-dd 
	'######  might have validation problems
	'###########################################

	mdate=rs("publicationdate")
	datepartday=day(mdate)
	datepartmonth=month(mdate)
	datepartyear=year(mdate)
	if 2>len(datepartday) then
		datepartdayformat="0" & datepartday
	else
		datepartdayformat=datepartday
	end if

	if 2>len(datepartmonth) then
		datepartmonthformat="0" & datepartmonth
	else
		datepartmonthformat=dateparmonth
	end if

	 publicationdate= datepartyear &  "-" & _
	 datepartmonthformat & _
	"-" & datepartdayformat  

	response.write indent4 & "&LT;pubdate>" & publicationdate & "&LT;/pubdate>" & "<br />"


	'DO KEYWORDS - parse keyword fields, themekeywords, placekeywords, sift out hte comma separated words 
	%>
	<!--#include file="write.eml.keywords.inc.asp"-->	
	<%

	'Write ABSTRACT 
	%>
	<!--#include file="write.eml.abstract.inc.asp"-->	
	<%


	'Write intellectual rights information
	response.write indent4 & "&lt;intellectualRights>" & "<br />"
	response.write indent5 & "&lt;para>" & rs("datacred") & "&lt;/para>" & "<br />"
	response.write indent4 & "&lt;/intellectualRights>" & "<br />"


	'Write the url of the dataset package
	response.write indent4 & "&lt;distribution>" & "<br />"
	response.write indent5 & "&lt;url function=""information"">" & rs("onlinelinkage") & "&lt;/url>" & "<br />"
	response.write indent4 & "&lt;/distribution>" & "<br />"

	
	'Write geographic coverage and temporal coverage
	%>
	<!--#include file="write.eml.coverage.inc.asp"-->	
	<%

	
	'Write contact information
	response.write indent4 & "&LT;contact>&LT;positionName>Baltimore Ecosystem Study Information Manager&LT;&lt;/positionName>" & "<br />"
	response.write indent4 & "&lt;address>" & "<br />"
	response.write indent5 & "&lt;deliveryPoint>Cary Institute Of Ecosystem Studies&lt;/deliveryPoint>" & "<br />"
	response.write indent5 & "&lt;deliveryPoint>2801 Sharon Turnpike&lt;/deliveryPoint>" & "<br />"
	response.write indent5 & "&lt;city>Millbrook&lt;/city>" & "<br />"
	response.write indent5 & "&lt;administrativeArea>New York&lt;/administrativeArea>" & "<br />"
	response.write indent5 & "&lt;postalCode>12545&lt;/postalCode>" & "<br />"
	response.write indent5 & "&lt;country>USA&lt;/country>" & "<br />"
	response.write indent4 & "&lt;/address>" & "<br />"
	response.write indent4 & "&lt;electronicMailAddress>walshj@caryinstitute.org&lt;/electronicMailAddress>" & "<br />"
	response.write indent4 & "&lt;/contact>" & "<br />"

	

	response.write indent4 & "&lt;intellectualRights>" & "<br />"
	response.write indent4 & "&lt;para>Publisher: " & trim(rs("publisher")) & " " & rs("datacred") & "&lt;/para>" & "<br />"
	response.write indent4 & "&lt;/intellectualRights>" & "<br />"
	response.write indent4 & "&lt;distribution>" & "<br />"
	response.write indent4 & "&lt;/distribution>" & "<br />"



	' Now write some methods
	'OPEN METHODS AND METHOD LINK DATASETS
	mmdatasetid=trim(rs("dataset_id"))

	strSQLmethodlink="SELECT methodlink.methodid, methodlink.datasetid, methods.methodname, methods.methoddescription FROM Methodlink LEFT JOIN methods ON methodlink.methodid = methods.methodid WHERE trim(methodlink.datasetid)=""" & mmdatasetid & """"
	'response.write "<br />" & strSQLmethodlink & "<br />"
	set rsmethodlink =  Server.CreateObject("ADODB.recordset")
	rsmethodlink.Open strSQLmethodlink, conn

	response.write indent4 & "&lt;methods>"  & "<br />"
	emergencystop=0
	do while not rsmethodlink.EOF and emergencystop<2000
	emergencystop=emergencystop+1
	response.write indent5 & "&lt;methodStep>"  & "<br />"
	response.write indent6 & "&lt;description>"  & "<br />"
	response.write indent7 & "&lt;section>"  & "<br />"
	response.write indent8 & "&lt;title>" & rsmethodlink("methodname") & "&lt;/title>" & "<br />"
	response.write indent8 & "&lt;para>" & rsmethodlink("methoddescription") & "&lt;/para>" & "<br />"
	response.write indent7 & "&lt;/section>"  & "<br />"
	response.write indent6 & "&lt;/description>"  & "<br />"
	response.write indent5 & "&lt;/methodStep>"  & "<br />"
	'response.write rsmethodlink("datasetid") & ", " & rsmethodlink("methodname") & ", " '& rsmethodlink("methoddescription")
	'response.write " that was a method id " & "<br />"
	rsmethodlink.movenext
	loop

	response.write indent4 & "&lt;/methods>"  & "<br />"


	'SECTION BREAK #######################################################################
	'#####################################################################################
	'NOW GET INTO THE SPECIFICS ABOUT THE DATA FILE
	response.write indent4 & "&lt;dataTable>" & "<br />"
	response.write indent5 & "&lt;entityName>"& rs("dataset_id") & "&lt;/entityName>" & "<br />"
	response.write indent5 & "&lt;entityDescription>Main data table for data set " & trim(rs("dataset_id")) &  "&lt;/entityDescription>" & "<br />"
	response.write indent5 & "&lt;physical>" & "<br />"
	response.write indent6 & "&lt;objectName>" & rs("filename") & "&lt;/objectName>" & "<br />"
	response.write indent6 & "&lt;size unit=""" & rs("filesizeunit") & """>" & rs("filesize") & "&lt;/size>" & "<br />"
	response.write indent6 & "&lt;characterEncoding>" & rs("characterencoding") & "&lt;/characterEncoding>" & "<br />"

	'Well now here I'm stuck as to what to do for data that are not text format.  So I will write eml for dataformat that is text, or else I will use some "otherentity" dataformat for othjer than text.  So here goes...
	if trim(rs("dataformat"))="text" then
		response.write indent6 & "&lt;dataFormat>" & "<br />"
		response.write indent7 & "&lt;textFormat>" & "<br />"
		response.write indent7 & "&lt;numHeaderLines>" & rs("numheaderlines") & "&lt;/numHeaderLines>" & "<br />"
		response.write indent7 & "&lt;numFooterLines>" & rs("numfooterlines") & "&lt;/numFooterLines>" & "<br />"
		response.write indent7 & "&lt;recordDelimiter>" & rs("recorddelimiter") & "&lt;/recordDelimiter>" & "<br />"
		response.write indent7 & "&lt;numPhysicalLinesPerRecord>" & rs("linesperrecord") & "&lt;/numPhysicalLinesPerRecord>" & "<br />"
		response.write indent7 & "&lt;attributeOrientation>" & rs("orientation") & "&lt;/attributeOrientation>" & "<br />"
		response.write indent7 & "&lt;simpleDelimited>" & "<br />"
		response.write indent7 & "&lt;fieldDelimiter>" & rs("fielddelimiter") & "&lt;/fieldDelimiter>" & "<br />"
		response.write indent7 & "&lt;quoteCharacter>" & rs("quotecharacter") & "&lt;/quoteCharacter>" & "<br />"
		response.write indent7 & "&lt;/simpleDelimited>" & "<br />"
		response.write indent7 & "&lt;/textFormat>" & "<br />"
		response.write indent6 & "&lt;/dataFormat>" & "<br />"
	else
		response.write indent6 & "&lt;dataFormat>" & "<br />"
		response.write indent6 & "&lt;/dataformat>" & "<br />" 
	end if





	response.write indent6 & "&lt;distribution>" & "<br />"
	response.write indent7 & "&lt;online>" & "<br />"
	response.write indent8 & "&lt;onlineDescription>" & rs("onlinedescription") & "&lt;/onlineDescription>" & "<br />"
	'Wade's looks like this.  Do I need this?
	'NO...  The answer is, Wade is using a data repository at LNO.
	'<url 'function="download">http://metacat.lternet.edu/das/dataAccessServlet?docid=knb-lter-gce.89.17&amp;urlTail=accession=CHM-GC'ED-0303b&amp;filename=CHM-GCED-0303b_1_3.CSV</url>
	'FOR now, I'll just put the line below:
	response.write indent8 & "&lt;url function=""download"">" & rs("onlinelinkage") & "&lt;/url>" & "<br />"





	'REPEAT ACCESS INFORMATION, ONLY THIS TIME FOR THE DOWNLOAD, NOT JUST INFORMATION ABOUT THE FILE
	%>
	<!--#include file="write.eml.accesspermissionsdownload.inc.asp"-->	
	<%


'<distribution>
'<online>
'<onlineDescription>Spreadsheet comma-separated value (CSV) text file with a five line header containing the data set 'title, column titles, units and column types</onlineDescription>
'<url 'function="download">http://metacat.lternet.edu/das/dataAccessServlet?docid=knb-lter-gce.89.17&amp;urlTail=accession=CHM-GC'ED-0303b&amp;filename=CHM-GCED-0303b_1_3.CSV</url>
'</online>
'<access authSystem="knb" order="allowFirst" scope="document">
'<allow>
'<principal>uid=GCE,o=lter,dc=ecoinformatics,dc=org</principal>
'<permission>all</permission>
'</allow>
'<allow>
'<principal>public</principal>
'<permission>read</permission>
'</allow>
'</access>
'</distribution>
'</physical>




'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXX                                                                                  XXXX
'XXXX                                                                                  XXXX
'XXXX                                                                                  XXXX
'XXXX        A L L     C O D E   W O R K S   T O   H E R E  2013/07/08                 XXXX
'XXXX                                                                                  XXXX
'XXXX                                                                                  XXXX
'XXXX                                                                                  XXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX



	
	
	'Close <dataTable> tag --- but a lot of stuff still to write above 
	response.write indent4 & "&LT;/dataTable>" & "<br />"

	'SECTION BREAK #######################################################################
	'#####################################################################################









	'Close <dataset> tag
	response.write indent2 & "&LT;/dataset>" & "<br />"






	'################ SKIPPING THIS FOR NOW, IS HARD CODED, DONT LIKE TH WAY IT WORKS
	'################ BUT DO NOT DISCARD!!!!!
	'Write additional metadata
	'% >
	'<!--#include file="write.eml.addlmetadata.inc.asp"-->	
	'<%
	'#################################################################################

	'closing tag for entire document
	'Write EML bottom of page - footer
	%>
	<!--#include file="write.eml.footer.inc.asp"-->	
	<%
	

	
	rs.movenext

	if rs.EOF then
		response.write "EOF ########################################"
	end if




loop

%>

