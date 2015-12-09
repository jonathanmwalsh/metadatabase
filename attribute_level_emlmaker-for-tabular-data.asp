<% 

'FO WRITING TO SCREN
'indent1="&nbsp;&nbsp;"
'indent2=indent1 & "&nbsp;&nbsp;"
'indent3=indent2 & "&nbsp;&nbsp;"
'indent4=indent3 & "&nbsp;&nbsp;"
'indent5=indent4 & "&nbsp;&nbsp;"
'indent6=indent5 & "&nbsp;&nbsp;"
'indent7=indent6 & "&nbsp;&nbsp;"
'indent8=indent7 & "&nbsp;&nbsp;"
'indent9=indent8 & "&nbsp;&nbsp;"
'indent10=indent9 & "&nbsp;&nbsp;"
'indent11=indent10 & "&nbsp;&nbsp;"
'indent12=indent11 & "&nbsp;&nbsp;"

'FOR WRITING TO FILE
indent1="  "
indent2=indent1 & "  "
indent3=indent2 & "  "
indent4=indent3 & "  "
indent5=indent4 & "  "
indent6=indent5 & "  "
indent7=indent6 & "  "
indent8=indent7 & "  "
indent9=indent8 & "  "
indent10=indent9 & "  "
indent11=indent10 & "  "
indent12=indent11 & "  "

'AND DONT FORGET, WHEN WRITING TO SCREEN, "<" is "<"

'Make database connections and
'Open main datasets
	%>
	<!--#include file="emlmaker.open.datasets.tabular.inc.asp"-->	
	<%


'rs.movelast 'LOL No such command
'find last dataset for file positioning purposes
do while not rs.eof
	mdatasetid=rs("dataset_id")
	response.write rs("dataset_id") & " " & rs("part_of_multi_id_Dataset") & " <<<<< "
	'response.write rs("part_of_multi_id_dataset")="1"
	rs.movenext
loop
response.write "Last dataset ID in recordset is: " & mdatasetid
rs.movefirst

emergencystop=0

'UPDATE COMBINED SET [COMBINED].[publisher] = [COMBINED].[surname]
'UPDATE COMBINED SET publisher = surname 


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



fname=0 ' fname is filename serial number so all files are named with a sequential number in them

do while not rs.EOF and emergencystop<2000

	emergencystop=emergencystop+1
	mRecordID=rs("RecordID")
	fname=MrecordID
	revision_no_for_this_dataset=rs("edition")*10 ' I multiply by 10 so I can usee a decimal pooint in my revision numbers


	cname=ltrim(trim(cstr(fname))) 'cname is the character, rendition of fname, for adding numbers to filenames
	tfname="c:\inetpub\wwwroot\metadata_harvest_attribute_level_eml\knb-lter-bes-" & cname & ".xml"  ' Be careful:  tfname is the ACTUAL filename we are writing to.  So do not change it or you might spew destruction all over the server's disk
	' harvestlistname is the name we report to metacat in the harvestlist.xml file of eml filenames.
	'harvestlistname= "http://belter.org/metacat_harvest/" & cname & ".xml"
	harvestlistname= "knb-lter-bes-" & cname & ".xml"

	'response.write "<br />&nbsp;<br />&nbsp;"



Response.write "<br />OK SO FAR"



	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	'In this part we write to the stream to make harvestlist.xml file

	objTextFile2.WriteLine "<document>"
	objTextFile2.WriteLine "<docid>"
	objTextFile2.WriteLine "<scope>knb-lter-bes</scope>"
	objTextFile2.WriteLine "<identifier>" & cname & "</identifier>"
	'objTextFile2.WriteLine "<revision>" & lter_rev_no & "</revision>"
	objTextFile2.WriteLine "<revision>" & revision_no_for_this_dataset & "</revision>"
	objTextFile2.WriteLine "</docid>"
	objTextFile2.WriteLine "<documentType>eml://eml://ecoinformatics.org/eml-2.1.0</documentType>"
	objTextFile2.WriteLine "<documentURL>http://beslter.org/metadata_harvest_attribute_level_eml/" & harvestlistname & "</documentURL>"
	objTextFile2.WriteLine "</document>"
	objTextFile2.WriteLine " "
	if rs("pastaview")=1 then
		objTextFile5.WriteLine "http://beslter.org/metadata_harvest_attribute_level_eml/" & harvestlistname
	end if


'SAMPLE OF WHAT WE'RE TRYING TO BUILD
'<?xml version="1.0" encoding="UTF-8"?>
'<eml:eml packageId="knb-lter-gce.89.17" system="knb" xmlns:ds="eml://ecoinformatics.org/dataset-2.1.0" xmlns:eml="eml://ecoinformatics.org/eml-2.1.0" xmlns:stmml="http://www.xml-cml.org/schema/stmml-1.1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:schemaLocation="eml://ecoinformatics.org/eml-2.1.0 http://gce-lter.marsci.uga.edu/public/files/schemas/eml-210/eml.xsd">

	
	
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	'In this part we write to the stream to make individual harvest file file
	'tfname="c:\inetpub\wwwroot\metadata_harvest_attribute_level_eml\jonny.txt"
	response.write "<br />  Filename tfname value is " & tfname & " *** "
	Set objTextFile = objFSO.CreateTextFile(tfname, true)




	'Write EML top of page - Header
	%>
	<!--#include file="write.eml.header.inc.asp"-->	
	<%
	
	
	
	'HELPS DEBUGGING - FIGURING OUT WHERE YOU ARE IN THE RECORD SET
	response.write " *** Dataset_id= " & rs("dataset_id") & " * " 
	'response.write  " <br/>" 


	response.write " * Record id= " & rs("RecordID")  & " * " 
	response.write ", File Type= " & rs("file_type") & " * "   
	'response.write "EOF 1 " & rs.EOF & " EOF 2 " & rs2.EOF

	' response.write  " <br/>" 
		response.write ", FA_version= " & rs("fa_update") & " * " 
		'response.write  " <br/>" 




	'Now step through accesspermissions and stuff
	%>
	<!--#include file="write.eml.accesspermissions.inc.asp"-->	
	<%
	
	

	
	'Open <dataset> tag
	objTextFile.WriteLine indent2 & "<dataset scope=""document"">"   'I don't know why using a scope identifier (scope = document) but I see it on examples I am using

	'Title of dataset
	objTextFile.WriteLine indent3 & "<title>" & rs("title") & "</title>" 






	'creator nodes
	
	'OPEN NAMES DATASET, do personal name, do organisations name, 
	%>
	<!--#include file="write.eml.personorgnames.inc.asp"-->	
	<%








	'pubDate  
	'#########################################
	'###### Be aware pubDate for me is yyyy/mm/dd and eml calls for yyyy-mm-dd 
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

	objTextFile.WriteLine indent4 & "<pubDate>" & publicationdate & "</pubDate>" 


	'Write ABSTRACT 
	%>
	<!--#include file="write.eml.abstract.inc.asp"-->	
	<%

	'DO KEYWORDS - parse keyword fields, themekeywords, placekeywords, sift out hte comma separated words   '### Swapped Abstract and keywordset nodes 7/18/2013 per MOB
	%>
	<!--#include file="write.eml.keywords.inc.asp"-->	
	<%


	'Write intellectual rights information
	objTextFile.WriteLine indent4 & "<intellectualRights>" 
	objTextFile.WriteLine indent4 & "<para>Publisher: " & trim(rs("publisher")) & " " & rs("datacred") & "</para>" 
	objTextFile.WriteLine indent4 & "</intellectualRights>" 


	'Write the url of the dataset package
	objTextFile.WriteLine indent4 & "<distribution>" 
	objTextFile.WriteLine indent5 & "<online>" 
	objTextFile.WriteLine indent6 & "<url function=""information"">" & rs("onlineinfolink") & "</url>" 
	objTextFile.WriteLine indent5 & "</online>" 
	objTextFile.WriteLine indent4 & "</distribution>" 

	
	'Write geographic coverage and temporal coverage
	%>
	<!--#include file="write.eml.coverage.inc.asp"-->	
	<%

	
	'Write contact information
	objTextFile.WriteLine indent4 & "<contact><positionName>Baltimore Ecosystem Study Information Manager</positionName>" 
	objTextFile.WriteLine indent4 & "<address>" 
	objTextFile.WriteLine indent5 & "<deliveryPoint>Cary Institute Of Ecosystem Studies</deliveryPoint>" 
	objTextFile.WriteLine indent5 & "<deliveryPoint>2801 Sharon Turnpike</deliveryPoint>" 
	objTextFile.WriteLine indent5 & "<city>Millbrook</city>" 
	objTextFile.WriteLine indent5 & "<administrativeArea>New York</administrativeArea>" 
	objTextFile.WriteLine indent5 & "<postalCode>12545</postalCode>" 
	objTextFile.WriteLine indent5 & "<country>USA</country>" 
	objTextFile.WriteLine indent4 & "</address>" 
	objTextFile.WriteLine indent4 & "<electronicMailAddress>walshj@caryinstitute.org</electronicMailAddress>" 
	objTextFile.WriteLine indent4 & "</contact>" 

	

'	objTextFile.WriteLine indent4 & "<intellectualRights>"    'Commented out since was duplicate three stanzas above.  IntellectualRights is supposed to come just under keywords
'	objTextFile.WriteLine indent4 & "<para>Publisher: " & trim(rs("publisher")) & " " & rs("datacred") & "</para>" 
'	objTextFile.WriteLine indent4 & "</intellectualRights>" 
'	objTextFile.WriteLine indent4 & "<distribution>" 
'	objTextFile.WriteLine indent4 & "</distribution>" 



	' Now write some methods
	'OPEN METHODS AND METHOD LINK DATASETS
	mmdatasetid=trim(rs("dataset_id"))

	strSQLmethodlink="SELECT methodlink.methodid, methodlink.datasetid, methods.methodname, methods.methoddescription FROM Methodlink LEFT JOIN methods ON methodlink.methodid = methods.methodid WHERE trim(methodlink.datasetid)=""" & mmdatasetid & """"
	'response.write "<br />" & strSQLmethodlink 
	set rsmethodlink =  Server.CreateObject("ADODB.recordset")
	rsmethodlink.Open strSQLmethodlink, conn

	objTextFile.WriteLine indent4 & "<methods>"  
	emergencystop=0
	do while not rsmethodlink.EOF and emergencystop<2000
	emergencystop=emergencystop+1
	objTextFile.WriteLine indent5 & "<methodStep>"  
	objTextFile.WriteLine indent6 & "<description>"  
	objTextFile.WriteLine indent7 & "<section>"  
	objTextFile.WriteLine indent8 & "<title>" & rsmethodlink("methodname") & "</title>" 
	m_methoddescription=Replace(rsmethodlink("methoddescription"), vbcrlf, "</para> <para> " & vbcrlf & indent8)

	objTextFile.WriteLine indent8 & "<para>" & m_methoddescription & "</para>" 
	objTextFile.WriteLine indent7 & "</section>"  
	objTextFile.WriteLine indent6 & "</description>"  
	objTextFile.WriteLine indent5 & "</methodStep>"  
	'response.write rsmethodlink("datasetid") & ", " & rsmethodlink("methodname") & ", " '& rsmethodlink("methoddescription")
	'response.write " that was a method id " 
	rsmethodlink.movenext
	loop

	objTextFile.WriteLine indent4 & "</methods>"  


	'SECTION BREAK #######################################################################
	'#####################################################################################
	'NOW GET INTO THE SPECIFICS ABOUT THE DATA FILE
	'response.write " *ggggggggggggggggggg* "
	
	mDataSetMultiID=trim(rs("dataset_id")) ' This is a loop for datasets with multiple entities.  We will repeat the <dataset>...</dataset> elements.  IMPORTANT TO KNOW THIS!
	'If a dataset is just a normal, single, dataset_id, then it will be transparent and just treat this loop 
	'like linear code - right on through, just one pass.
	strsqlMultiID="SELECT COMBINED.RecordID, COMBINED.dataset_id, COMBINED.part_of_multi_id_dataset, COMBINED.child_of_multi_id_dataset, COMBINED.rank_of_multi_id_dataset, COMBINED.edition, COMBINED.id, COMBINED.ExportName, COMBINED.Title, COMBINED.File_Attribute_Type, COMBINED.creatorid, COMBINED.orgname, COMBINED.surname, COMBINED.givenname, COMBINED.themekeywords, COMBINED.placekeywords, COMBINED.abstract, COMBINED.publicationdate, COMBINED.datacred, COMBINED.onlinelinkage, COMBINED.onlineinfolink, COMBINED.geographicdescription, COMBINED.west, COMBINED.east, COMBINED.north, COMBINED.south, COMBINED.temporalbegin, COMBINED.temporalend, COMBINED.publisher, COMBINED.datacred, COMBINED.filename, COMBINED.filesizeunit, COMBINED.filesize, COMBINED.characterencoding, COMBINED.dataformat, COMBINED.numheaderlines, COMBINED.numfooterlines, COMBINED.recorddelimiter, COMBINED.linesperrecord, COMBINED.fielddelimiter, COMBINED.quotecharacter, COMBINED.orientation, COMBINED.onlinedescription, COMBINED.pastaview, File_Attribute_table.File_Attribute_Type, File_Attribute_table.file_type, File_Attribute_status.fa_update FROM ((COMBINED LEFT JOIN File_Attribute_table ON COMBINED.dataset_id = File_Attribute_table.dataset_id) LEFT JOIN File_Attribute_status ON COMBINED.dataset_id = File_Attribute_status.dataset_id) WHERE DataEntityType=1 and trim(COMBINED.dataset_id) = """ & mDataSetMultiID & """ ORDER BY COMBINED.dataset_id, COMBINED.part_of_multi_id_dataset, COMBINED.child_of_multi_id_dataset"
	'Also had pastaview=1

	set rsMultiID = Server.CreateObject("ADODB.recordset")
	rsMultiID.Open strsqlMultiID, conn
	'response.write " *ggggggggggggggggggg* "

	emergencystop=0
	MultiIDCounter=0
	'if rsMultiID.count>1 then
	'	This_is_a_multi=1
	'else
	'	This_is_a_multi=0
	'end if

	do while not rsMultiID.EOF and emergencystop<100
		emergencystop=emergencystop+1
		response.write emergencystop & " ** "
		MultiIDCounter=MultiIDCounter+1
		mrank_of_multi_id_dataset=0


		
		objTextFile.WriteLine indent4 & "<dataTable>" 
		'' Need to change name of entity if it's a repeating entity
		'if MultiIDCounter>1 then
		'	objTextFile.WriteLine indent5 & "<entityName>"& rsMultiID("dataset_id") &  "-" & MultiIDCounter & "</entityName>" 
		'else 
		'	objTextFile.WriteLine indent5 & "<entityName>"& rsMultiID("dataset_id") & "</entityName>" 
		'end if
		' Need to change name of entity if it's a repeating entity
		
		if rsMultiID("rank_of_multi_id_dataset")>0 then
			mrank_of_multi_id_dataset=rsMultiID("rank_of_multi_id_dataset")
		end if
		if rsMultiID("rank_of_multi_id_dataset")>0 then
			objTextFile.WriteLine indent5 & "<entityName>"& rsMultiID("dataset_id") &  "-" & rsMultiID("rank_of_multi_id_dataset") & "</entityName>" 
		else 
			objTextFile.WriteLine indent5 & "<entityName>"& rsMultiID("dataset_id") & "</entityName>" 
		end if

		objTextFile.WriteLine indent5 & "<entityDescription>Data table for data set " & trim(rsMultiID("dataset_id")) &  "</entityDescription>" 
		objTextFile.WriteLine indent5 & "<physical>" 
		objTextFile.WriteLine indent6 & "<objectName>" & rsMultiID("filename") & "</objectName>" 
		objTextFile.WriteLine indent6 & "<size unit=""" & rsMultiID("filesizeunit") & """>" & rsMultiID("filesize") & "</size>" 
		objTextFile.WriteLine indent6 & "<characterEncoding>" & rsMultiID("characterencoding") & "</characterEncoding>" 

		'Well now here I'm stuck as to what to do for data that are not text format.  So I will write eml for dataformat that is text, or else I will use some "otherentity" dataformat for othjer than text.  So here goes...
		if trim(rsMultiID("dataformat"))="text" then
			objTextFile.WriteLine indent6 & "<dataFormat>" 
			objTextFile.WriteLine indent7 & "<textFormat>" 
			objTextFile.WriteLine indent7 & "<numHeaderLines>" & rsMultiID("numheaderlines") & "</numHeaderLines>" 
			objTextFile.WriteLine indent7 & "<numFooterLines>" & rsMultiID("numfooterlines") & "</numFooterLines>" 
			objTextFile.WriteLine indent7 & "<recordDelimiter>" & rsMultiID("recorddelimiter") & "</recordDelimiter>" 
			objTextFile.WriteLine indent7 & "<numPhysicalLinesPerRecord>" & rsMultiID("linesperrecord") & "</numPhysicalLinesPerRecord>" 
			objTextFile.WriteLine indent7 & "<attributeOrientation>" & rsMultiID("orientation") & "</attributeOrientation>" 
			objTextFile.WriteLine indent7 & "<simpleDelimited>" 
			objTextFile.WriteLine indent7 & "<fieldDelimiter>" & rsMultiID("fielddelimiter") & "</fieldDelimiter>" 
			objTextFile.WriteLine indent7 & "<quoteCharacter>" & rsMultiID("quotecharacter") & "</quoteCharacter>" 
			objTextFile.WriteLine indent7 & "</simpleDelimited>" 
			objTextFile.WriteLine indent7 & "</textFormat>" 
			objTextFile.WriteLine indent6 & "</dataFormat>" 
		else
			objTextFile.WriteLine indent6 & "<dataFormat>" 
			objTextFile.WriteLine indent6 & "</dataformat>"  
		end if





		objTextFile.WriteLine indent6 & "<distribution>" 
		objTextFile.WriteLine indent7 & "<online>" 
		objTextFile.WriteLine indent8 & "<onlineDescription>" & rsMultiID("onlinedescription") & "</onlineDescription>" 
		'Wade's looks like this.  Do I need this?
		'NO...  The answer is, Wade is using a data repository at LNO.
		'<url 'function="download">http://metacat.lternet.edu/das/dataAccessServlet?docid=knb-lter-gce.89.17&amp;urlTail=accession=CHM-GC'ED-0303b&amp;filename=CHM-GCED-0303b_1_3.CSV</url>
		'FOR now, I'll just put the line below:
		objTextFile.WriteLine indent8 & "<url function=""download"">" & rsMultiID("onlinelinkage") & "</url>" 
		objTextFile.WriteLine indent7 & "</online>" 





		'REPEAT ACCESS INFORMATION, ONLY THIS TIME FOR THE DOWNLOAD, NOT JUST INFORMATION ABOUT THE FILE
		%>
		<!--#include file="write.eml.accesspermissionsdownload.inc.asp"-->	
		<%


		objTextFile.WriteLine indent7 & "</distribution>" 
		objTextFile.WriteLine indent6 & "</physical>" 










		
		'Open attributes table and filter to current dataset_id
		'response.write rsMultiID("dataset_id")


		'ORIGINAL UNTOUCHED
		'		strSQLattribs="SELECT att.attributeid, att.attributename, att.attributedefinition, att.storagetype, att.measurementscale, att.missingvaluecode, att.missingvaluecodeexplanation FROM attributes att LEFT JOIN attribslink  ON trim(att.attributeid)=trim(attribslink.attributeid) WHERE trim(attribslink.datasetid) = """ & trim(rs("dataset_id")) & """" & " ORDER BY DataSetColumnNo"


		mdataset_multi_id=MultiIDCounter

		
		'strSQLattribs="SELECT att.attributeid, att.attributename, att.attributedefinition, att.storagetype, att.measurementscale, att.missingvaluecode, att.missingvaluecodeexplanation, attribslink.Dataset_multi_id FROM attributes att LEFT JOIN attribslink  ON trim(att.attributeid)=trim(attribslink.attributeid) WHERE trim(attribslink.datasetid) = """ & trim(rs("dataset_id")) & """ and attribslink.dataset_multi_id = """ & mdataset_multi_id &  """ ORDER BY DataSetColumnNo"

		strSQLattribs="SELECT att.attributeid, att.attributename, att.attributedefinition, att.storagetype, att.measurementscale, att.missingvaluecode, att.missingvaluecodeexplanation, attribslink.Dataset_multi_id FROM attributes att LEFT JOIN attribslink  ON trim(att.attributeid)=trim(attribslink.attributeid) WHERE trim(attribslink.datasetid) = """ & trim(rs("dataset_id")) & """ and attribslink.dataset_multi_id = """ & mrank_of_multi_id_dataset &  """ ORDER BY DataSetColumnNo"

		'response.write "<br>" & strSQLattribs & "<br>"
			


		
		'response.write "<br />" & strSQLattribs & "<br />"

		set rsattribs = Server.CreateObject("ADODB.recordset")
		rsattribs.Open strSQLattribs, conn



		' Module to write attributes to file for this dataset_id
		
		
		objTextFile.WriteLine indent6 & "<attributeList>" 
		
		emergencystopatt=0
		
		do while not rsattribs.EOF and emergencystopatt<2000
'response.write "HERE IS val RSATTRIBS MULTI ID" & rsattribs("dataset_multi_id") & "<br>"
'response.write "HERE IS Multi Counter      " & mdataset_multi_id & "<br>"
			emergencystopatt=emergencystopatt+1

			if rs("part_of_multi_id_dataset")=1 then
				objTextFile.WriteLine indent7 & "<attribute id=""" & rsattribs("attributeid") & ".table-" & mrank_of_multi_id_dataset & """>" 
			else
				objTextFile.WriteLine indent7 & "<attribute id=""" & rsattribs("attributeid") & """>" 
			end if
			objTextFile.WriteLine indent8 & "<attributeName>" & rsattribs("attributename") & "</attributeName>" 
			objTextFile.WriteLine indent8 & "<attributeDefinition>" & rsattribs("attributedefinition") & "</attributeDefinition>" 
			objTextFile.WriteLine indent8 & "<storageType>" & rsattribs("storagetype") & "</storageType>" 
			objTextFile.WriteLine indent8 & "<measurementScale>" & rsattribs("measurementscale") & "</measurementScale>" 
					'if errorcode is not blank...  objTextFile.WriteLine indent3 & "<" & rsatt("attributeid") & ">" 
					'objTextFile.WriteLine indent3 & "<" & rsatt("attributeid") & ">" 
			If len(trim(rsattribs("missingvaluecode")))>0 then
				objTextFile.WriteLine indent8 & "<missingValueCode>"
				objTextFile.WriteLine indent9 & "<code>" & rsattribs("missingValueCode") & "</code>"
				objTextFile.WriteLine indent9 & "<codeExplanation>" & rsattribs("MissingValueCodeExplanation") & "</codeExplanation>" 
				objTextFile.WriteLine indent8 & "</missingValueCode>" 
			end if
			objTextFile.WriteLine indent7 & "</attribute>" 

			rsattribs.movenext
		loop


		' close tag
		objTextFile.WriteLine indent6 & "</attributeList>" 



		
		'Close <dataTable> tag --- but a lot of stuff still to write above 
		objTextFile.WriteLine indent4 & "</dataTable>" 

		rsMultiID.movenext


	loop ' End loop for Multipe <Datatable>...</datatable> elements
	'#####################################################################################
	'#####################################################################################
	'#####################################################################################
	'#####################################################################################
	'#####################################################################################
	'#####################################################################################
	'#####################################################################################
	'#####################################################################################
	
	
	
	
	'SECTION BREAK #######################################################################
	'#####################################################################################









	'Close <dataset> tag
	objTextFile.WriteLine indent2 & "</dataset>" 






	'Write additional metadata - It will be for custom units
	%>
	<!--#include file="write.eml.addlmetadata.inc.asp"-->	
	<%
	

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

