<% 


'Open connection to database
accessdb="esri2bes_int_osrs-search_jwalsh.mdb"                 ' Just setting a variable to thej name
db=Server.MapPath(accessdb)                    ' JW enhanced cutesy variables
dbGenericPath = "/emlmaker_attribute_level_eml/"
dbConn = "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & db  'my cutesy way




'Open main spatial data recordset
dbRs = "COMBINED" 'Name of table.
strConn = dbConn
strTable = dbRs

' Open Connection to the database
set xConn = Server.CreateObject("ADODB.Connection")
xConn.Open strConn




'Open Second non-spatial data recordset
dbRs2 = "NonSpatial" 'name of table
'strConn2 = dbConn
strTable2 = dbRs2
' Open Connection to the database
'set xConn2 = Server.CreateObject("ADODB.Connection")
'xConn2.Open strConn2




'Do the selection and create the recordset for the spatial records
strsql = "SELECT * FROM [" & strTable & "] "
'strsql = strsql & " where IsPublic = 1 " 
'strsql = strsql & "'" & wherelike & "'"
strsql=strsql & " order by RecordID ASC, ID ASC"  ' Sorting by DESC in ID field gets us our "last" record within the RecordID.  (THIS WAS FOR ORS - Records could have the same ID.  This is no longer true.  IDs are unique.  What we want is, step thru RecordID field and select greates ID within that subset of a given RecordID.
response.write strsql & "<br>"

set xrs = Server.CreateObject("ADODB.Recordset")
xrs.Open strsql, xConn, 1, 2  'The 1 and 2 refer to index elements like in line 3 lines above order by xxx + xxx + xxx - WATCH FOR TYPE MISMATCHES!

'Do the selection and create the recordset for the non-spatial records
strsql2 = "SELECT * FROM [" & strTable2 & "] "
'strsql2 = strsql2 & " where IsPublic = 1 " 
'strsql = strsql & "'" & wherelike & "'"
strsql2=strsql2 & " order by RecordID ASC"  ' Sorting by DESC in ID field gets us our "last" record within the RecordID.  What we want is, step thru RecordID field and select greates ID within that subset of a given RecordID.
response.write "strsql2=" & strsql2 & "<br>"

set xrs2 = Server.CreateObject("ADODB.Recordset")
xrs2.Open strsql2, xConn, 1, 2  'The 1 and 2 refer to index elements like in line 3 lines above order by xxx + xxx + xxx - WATCH FOR TYPE MISMATCHES!

xrs2.movefirst




'<!-- INSERT PAGE TOP BEGIN PAGE TOP -->


response.write "****************************** BEGIN EML OUTPUT" & "<br>" & Vbcrlf


'<!-- INSERT CONTENT-->



testtext="hello"
'response.write testtext
'response.write testtex

dim objFSO
Set objFSO = server.CreateObject("Scripting.FileSystemObject")
dim objTextFile 'There are many of these, these will be the xml files we write
dim objTextFile2 'This will be the file harvestlist.xml
dim objTextFile3 'This will be the html page the public will see.
dim objTextFile4 'There are many of these, these will be the html files we write for the full metadata records
Set objTextFile = objFSO.CreateTextFile(server.mappath("emloutput.txt"), true)
Set objTextFile2 = objFSO.CreateTextFile("c:\inetpub\wwwroot\metacat_harvest_attribute_level_eml\harvestlist.xml", true)
response.write "****************************** Harvestlist file open for writing " & "<br>" & Vbcrlf
Set objTextFile3 = objFSO.CreateTextFile("c:\inetpub\wwwroot\metacat_harvest_attribute_level_eml\frame7-page_1_auto.asp", true)
objTextFile2.WriteLine "<?xml version=""1.0"" encoding=""UTF-8"" ?>"

objTextFile2.WriteLine "<hrv:harvestList xmlns:hrv=""eml://ecoinformatics.org/harvestList"" >"



Const Filename = "/metacat_harvest_attribute_level_eml/lter_revision_no.txt"	' file to read
Const ForReading = 1, ForWriting = 2, ForAppending = 3
Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0



' Map the logical path to the physical system path
Dim Filepath
dim lter_rev_no
lter_rev_no=1

Filepath = Server.MapPath(Filename)

if objFSO.FileExists(Filepath) Then

	Set TextStream = objFSO.OpenTextFile(Filepath, ForReading, False, _
											  TristateUseDefault)
	' Read file in one hit
		
	Dim Contents
	Contents = TextStream.ReadAll
	Response.write "<pre>Contents of revision file set at: " & ltrim(trim(Contents)) & ".  Change file lter_revision_no.txt to adjust this.</pre><hr>"
	TextStream.Close
	Set TextStream = nothing
	lter_rev_no=trim(ltrim(contents))
	

Else

	Response.Write "<h3><i><font color=red> File " & Filename &_
                       " does not exist</font></i></h3>"

End If

Response.write "<pre>" & lter_rev_no & "</pre><hr>"



'Counters
dim records_processed
records_processed=0
dim publiccount
publiccount=0
dim lno_view_count
lno_view_count=0

'************************************************
'************************************************
'************************************************
'************************************************ 


'BEGIN LOOP TO WRITE SPATIAL AND NON SPATIAL DATA TO EML FILES  USING TABLE "COMBINED"



xrs.requery
xxx= xrs.eof
response.write " EOF EQUALS:"&xxx&"***"
xrs.movefirst
fname=0
mRecordID = 0
emergencystop=0

do while not xrs.EOF 'and emergencystop<2

	emergencystop=emergencystop+1
	'fname=fname+1
	mRecordID=xrs("RecordID")
	fname=MrecordID
	response.write MrecordID=xrs("RecordID")
	cname=ltrim(trim(cstr(fname)))
	tfname="c:\inetpub\wwwroot\metacat_harvest_attribute_level_eml\bes_" & cname & ".xml"  ' Be careful:  tfname is the ACTUAL filename we are writing to.  So do not change it or you might spew destruction all over the server's disk
	' harvestlistname is the name we report to metacat in the harvestlist.xml file of eml filenames.
	'harvestlistname= "http://belter.org/metacat_harvest/" & cname & ".xml"
	harvestlistname= "bes_" & cname & ".xml"



	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	'In this part we write to the stream to make harvestlist.xml file

	objTextFile2.WriteLine "<document>"
	objTextFile2.WriteLine "<docid>"
	objTextFile2.WriteLine "<scope>knb-lter-bes</scope>"
	objTextFile2.WriteLine "<identifier>" & cname & "</identifier>"
	objTextFile2.WriteLine "<revision>" & lter_rev_no & "</revision>"
	objTextFile2.WriteLine "</docid>"
	objTextFile2.WriteLine "<documentType>eml://ecoinformatics.org/eml-2.0.1</documentType>"
	objTextFile2.WriteLine "<documentURL>http://beslter.org/metacat_harvest_attribute_level_eml/" & harvestlistname & "</documentURL>"
	objTextFile2.WriteLine "</document>"
	objTextFile2.WriteLine " "


	
	
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	'In this part we write to the stream to make individual harvest file file
	Set objTextFile = objFSO.CreateTextFile(tfname, true)


	'response.write "********** mmmmm NEW RECORD" & "<br>" & Vbcrlf & "<br>" & tfname & Vbcrlf


	
	response.write "********** NEW RECORD - spatial - " & xrs("title") & Vbcrlf 


	
	'Write file top
	filetop= "<?xml version=""1.0"" encoding=""UTF-8""?> <eml:eml xmlns:eml=""eml://ecoinformatics.org/eml-2.0.1"" xmlns:stmml=""http://www.xml-cml.org/schema/stmml"" xmlns:sw=""eml://ecoinformatics.org/software-2.0.1"" xmlns:cit=""eml://ecoinformatics.org/literature-2.0.1"" xmlns:ds=""eml://ecoinformatics.org/dataset-2.0.1"" xmlns:prot=""eml://ecoinformatics.org/protocol-2.0.1"" xmlns:doc=""eml://ecoinformatics.org/documentation-2.0.1"" xmlns:res=""eml://ecoinformatics.org/resource-2.0.1"" xmlns:xs=""http://www.w3.org/2001/XMLSchema"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xsi:schemaLocation=""eml://ecoinformatics.org/eml-2.0.1 ./eml-2.0.1/eml.xsd"" packageId="""

	filetop=filetop & "knb-lter-bes." & cname & "." & lter_rev_no 'lter_rev_no is the revision number.  It is contolled in a textfile called lter_rev_no.txt
	
	filetop=filetop & """ system=""knb""> "

	objTextFile.WriteLine filetop

	
	
	objTextFile.WriteLine "<dataset>"

	
	
	



	'Start writing elements
	objTextFile.WriteLine "<title>"  & xrs("title") &  " BES ID " & xrs("RecordID") & "-" & xrs("ID") & "</title>"
		objTextFile.WriteLine t& "<creator>"
			objTextFile.WriteLine t& "<individualName>"
				objTextFile.WriteLine t&t& "<surName>" & xrs("publisher") & "</surName>"
			objTextFile.WriteLine t& "</individualName>"
		objTextFile.WriteLine t& "</creator>"

		mystrdate=""
		'objTextFile.WriteLine t&t& Datepart("yyyy", xrs("Publicationdate")) & "-0" & Datepart("m", xrs("Publicationdate")) & "-" & Datepart("d", xrs("Publicationdate"))
		mystrdate= Datepart("yyyy", xrs("Publicationdate"))
		if (datepart("m",xrs("Publicationdate"))<10) then
			mystrdate=mystrdate & "-0" & Datepart("m", xrs("Publicationdate")) 
		else
			mystrdate=mystrdate & "-" & Datepart("m", xrs("Publicationdate")) 
		end if
		if (datepart("d",xrs("Publicationdate"))<10) then
			mystrdate=mystrdate & "-0" & Datepart("d", xrs("Publicationdate"))
		else
			mystrdate=mystrdate & "-" & Datepart("d", xrs("Publicationdate"))
		end if
		objTextFile.WriteLine t& "<pubDate>" & mystrdate & "</pubDate>"
		objTextFile.WriteLine t& "<abstract><para>"
		if len(trim(xrs("Abstract")))>0 then
			objTextFile.WriteLine t&t& replace(xrs("Abstract"), "&", "&amp;")
		else 
			objTextFile.WriteLine t&t& "Not Available"
		end if

		objTextFile.WriteLine t& "</para></abstract>"
		objTextFile.WriteLine t& "<keywordSet>"
				if len(trim(xrs("ThemeKeywords")))>0 then
					objTextFile.WriteLine "<keyword keywordType=""theme"">"& replace(xrs("ThemeKeywords"), ",", ", ")&"</keyword>"
				else
					objTextFile.WriteLine t&t&t& "<keyword keywordType=""theme"">Not Available</keyword>"
				end if
		objTextFile.WriteLine t& "</keywordSet>"
		objTextFile.WriteLine t& "<keywordSet>"
				objTextFile.WriteLine "<keyword keywordType=""place"">"& xrs("PlaceKeywords")&"</keyword>"
		objTextFile.WriteLine t& "</keywordSet>"
		objTextFile.WriteLine t& "<distribution>"
			objTextFile.WriteLine t&t& "<online>"

				if len(trim(xrs("OnlineLinkage") & "*")) < 2 then   ' have to add this goddam asterisk because asp is whacky.  If onlinelinkage field is null, then it  simply doesnt exist.  no comparing it, no displaying it.  had to add star
					objTextFile.WriteLine "<url>" & "http://www.beslter.org" & "</url>"
					response.write "standard url written"
				else 
					objTextFile.WriteLine "<url>" & xrs("OnlineLinkage") & "</url>"
				end if

			objTextFile.WriteLine t&t& "</online>"
		objTextFile.WriteLine t& "</distribution>"


		objTextFile.WriteLine t& "<coverage>"
		objTextFile.WriteLine t& t&    "<geographicCoverage>"
		objTextFile.WriteLine "<geographicDescription>" & "The Baltimore Ecosystem Study ultimately will conduct research and educational activities throughout the Baltimore metropolitan area. This large area includes Baltimore City, Baltimore County, and the counties of Ann Arundel, Carrol, Harford, Howard, and Montgomery. Gwynns Falls includes agricultural lands, recently suburbanized areas, established suburbs, and dense urban areas having residential, commercial and open spaces. In addition, a reference area has been established in a forested catchment of the Gunpowder drainage in Oregon Ridge County Park. " & "</geographicDescription>"
		objTextFile.WriteLine t& t& t&      "<boundingCoordinates>"
		objTextFile.WriteLine t& t& t& t&		   "<westBoundingCoordinate>-77.314183</westBoundingCoordinate>"
		objTextFile.WriteLine t& t& t& t&          "<eastBoundingCoordinate>-76.012008</eastBoundingCoordinate>"
		objTextFile.WriteLine t& t& t& t&          "<northBoundingCoordinate>39.724847</northBoundingCoordinate>"
		objTextFile.WriteLine t& t& t& t&          "<southBoundingCoordinate>38.708367</southBoundingCoordinate>"
		objTextFile.WriteLine t& t& t& t&          "<boundingAltitudes>"
		objTextFile.WriteLine t& t& t& t& t&             "<altitudeMinimum>50</altitudeMinimum>"
		objTextFile.WriteLine t& t& t& t& t&             "<altitudeMaximum>700</altitudeMaximum>"
		objTextFile.WriteLine t& t& t& t& t&             "<altitudeUnits>feet</altitudeUnits>"
		objTextFile.WriteLine t& t& t& t&          "</boundingAltitudes>"
		objTextFile.WriteLine t& t& t&       "</boundingCoordinates>"
		objTextFile.WriteLine t& t&    "</geographicCoverage>"
		objTextFile.WriteLine t& "</coverage>"


		objTextFile.WriteLine t& "<contact>"
			objTextFile.WriteLine t&t& "<individualName>"
				objTextFile.WriteLine t&t&t& "<givenName>Jonathan</givenName>"
				objTextFile.WriteLine t&t&t& "<surName>Walsh</surName>"
			objTextFile.WriteLine t&t& "</individualName>"
			objTextFile.WriteLine t&t& "<organizationName>Cary Institute of Ecosystem Studies</organizationName>"
			objTextFile.WriteLine t&t& "<positionName>Information Manager</positionName>"
			objTextFile.WriteLine t&t& "<address>"
			objTextFile.WriteLine t&t& "<deliveryPoint>IES</deliveryPoint>"
			objTextFile.WriteLine t&t& "<deliveryPoint>Box AB, 65 Sharon Tpke</deliveryPoint>"
			objTextFile.WriteLine t&t& "<city>Millbrook</city>"
			objTextFile.WriteLine t&t& "<administrativeArea>NY</administrativeArea>"
			objTextFile.WriteLine t&t& "<postalCode>12545</postalCode>"
			objTextFile.WriteLine t&t& "<country>USA</country>"
			objTextFile.WriteLine t&t& "</address>"
			objTextFile.WriteLine t&t& "<phone phonetype=""voice"">845-677-7600</phone>"
			objTextFile.WriteLine t&t& "<phone phonetype=""fax"">  </phone>"
			objTextFile.WriteLine t&t& "<electronicMailAddress>walshj@ecostudies.org</electronicMailAddress>"
		objTextFile.WriteLine t& "</contact>"
		objTextFile.WriteLine t& "<publisher>"
			objTextFile.WriteLine t&t& "<organizationName>"&xrs("Themeref")&" "&xrs("Datacred")&" "&xrs("field23")&" "&xrs("Field22")&" Baltimore Ecosystem Study</organizationName>"
			objTextFile.WriteLine t&t& "<address>"
				objTextFile.WriteLine t&t&t& "<deliveryPoint>Room 134 TRC Building</deliveryPoint>"
				objTextFile.WriteLine t&t&t& "<deliveryPoint> University of Maryland, Baltimore County</deliveryPoint>"
				objTextFile.WriteLine t&t&t& "<deliveryPoint> 5200 Westland Blvd</deliveryPoint>"
				objTextFile.WriteLine t&t&t& "<city>Baltimore</city>"
				objTextFile.WriteLine t&t&t& "<administrativeArea>MD</administrativeArea>"
				objTextFile.WriteLine t&t&t& "<postalCode>21227</postalCode>"
			objTextFile.WriteLine t&t& "</address>"
		objTextFile.WriteLine t& "</publisher>"
		



		if xrs("IsPublic")=1 and xrs("lno-view")=1 then

			objTextFile.WriteLine t& "<access authSystem=""knb"" order=""allowFirst"" scope=""document"">"
				objTextFile.WriteLine t&t& "<allow>"
					objTextFile.WriteLine t&t&t& "<principal>uid=""BES"",o=lter,dc=ecoinformatics,dc=org</principal>"
					objTextFile.WriteLine t&t&t& "<permission>all</permission>"
				objTextFile.WriteLine t&t& "</allow>"
				objTextFile.WriteLine t&t& "<allow>"
					objTextFile.WriteLine t&t&t& "<principal>public</principal>"
					objTextFile.WriteLine t&t&t& "<permission>read</permission>"
				objTextFile.WriteLine t&t& "</allow>"
			objTextFile.WriteLine t& "</access>"
		else
			objTextFile.WriteLine t& "<access authSystem=""knb"" order=""allowFirst"" scope=""document"">"
				objTextFile.WriteLine t&t& "<allow>"
					objTextFile.WriteLine t&t&t& "<principal>uid=""BES"",o=lter,dc=ecoinformatics,dc=org</principal>"
					objTextFile.WriteLine t&t&t& "<permission>all</permission>"
				objTextFile.WriteLine t&t& "</allow>"
				objTextFile.WriteLine t&t& "<deny>" ' We're writing deny access right here if we need to
					objTextFile.WriteLine t&t&t& "<principal>public</principal>"
					objTextFile.WriteLine t&t&t& "<permission>read</permission>"
				objTextFile.WriteLine t&t& "</deny>"
			objTextFile.WriteLine t& "</access>"

		end if


		'Just keeping track
		if xrs("IsPublic")=1  then
			publiccount=publiccount+1
		end if
		if xrs("lno-view")=1  and xrs("IsPublic")=1 then
			lno_view_count=lno_view_count+1
		end if

		records_processed=records_processed+1

		response.write "Processed: " 
		response.write publiccount
		response.write "  Public: "
		response.write publiccount
		response.write "  LNO View: "
		response.write lno_view_count
		response.write "   "





	objTextFile.WriteLine "</dataset>"
objTextFile.WriteLine "</eml:eml>"


	
	
	
	
	
	
'	<BR>&nbsp;<BR>&nbsp;<BR>
'	<%
	
	response.write "********** END OF RECORD" & "<br>" & Vbcrlf & "<br>" & Vbcrlf
	
	do while not xrs.EOF  ' Here we skip over the remaining, older versions of the same recordID.  This recordset should now be indexed on RecordID ASC, ID DESC.  BUT NOTE:  **********  You can't compare recordID if past eof.  It borks.  So you can get a condition where youre checking recordid but you're past eof.  So you need an if block inside the loop FIRST checking for eof.  If not eof can compare the record id to mrecordid.  

		'objTextFile.WriteLine " BEFORE MOVENEXT ************************ "
		'objTextFile.WriteLine xrs.eof
		'objTextFile.WriteLine " - "
		'objTextFile.WriteLine fname 
		'objTextFile.WriteLine " - "
		'objTextFile.WriteLine fname<10 
		'objTextFile.WriteLine " - "
		'objTextFile.WriteLine xrs("RecordID") 
		'objTextFile.WriteLine " - "
		'objTextFile.WriteLine MRecordID 
		'objTextFile.WriteLine " - "

		xrs.movenext


		If not xrs.EOF then
			if MRecordID<>xrs("RecordID") then 'we exit loop
				exit do
			end if
		end if
	loop
loop


'************************************************
'************************************************
'************************************************
'************************************************ END LOOP TO CREATE XML FILES FOR COMBINED DATA






	
'************************************************
'************************************************
'************************************************
'************************************************ BEGIN LOOP TO WRITE SPATIAL AND NON SPATIAL DATA TO HTML FILES USING TABLE COMBINED
xrs.requery
xrs.movefirst
'mRecordID = 0
'fname=0 'yes, reinitialize this because we will add a different prefix to the filenames ' but no, do not reinitialize because it messes up scope identifier on other end


	



objTextFile3.WriteLine "<html>" 'make the top of the html document
objTextFile3.WriteLine "<head></head>" 

objTextFile3.WriteLine "<body>"
'Write HTML File top

objTextFile3.WriteLine "<!--#include file=""frame7-page_1_auto_head.html""-->"

objTextFile3.WriteLine Chr(9) &  "<table border='0' bordercolor='blue'>" ' new table - chr(9) is a tab



dim emllinkname
dim emlfnameno
emlfnameno=0
dim rowcounter
rowcounter=0
emergencystop=0


do while not xrs.EOF 'and emergencystop<1

	mRecordID=xrs("RecordID")
	emergencystop=emergencystop+1
'	fname=MrecordID
	emlfnameno=mRecordID
	cname=ltrim(trim(cstr(emlfnameno)))
	emllinkname="http://beslter.org/metacat_harvest_attribute_level_eml/bes_" & cname & ".xml"


	'**********************************************************
	'**********************************************************
	'******  HERE WE WRITE THE FULL RECORD HTML FILE **********
	'**********************************************************
	'**********************************************************

	'Create  name and file handle for individual html file of full metadata record
	fullmetarecordfilename="c:\inetpub\wwwroot\metacat_harvest_attribute_level_eml\html_metadata\bes_" & cname & ".asp"
	fullmetarecordlinkname="/metacat_harvest_attribute_level_eml/html_metadata/bes_" & cname & ".asp"
	response.write fullmetarecordlinkname

	'Just try writing anything to our new objTextFile4 which is html page of full metadata record
	Set objTextFile4 = objFSO.CreateTextFile(fullmetarecordfilename, true)
	objTextfile4.WriteLine "<html>" 'make the top of the html document
	objTextfile4.WriteLine "<head></head>" 

	objTextfile4.WriteLine "<body>"
	'Write HTML File top

	objTextfile4.WriteLine "<!--#include virtual  =""/metacat_harvest_attribute_level_eml/frame7-page_1_auto_bes_metadata_full_record_head.html""-->"

	objTextfile4.WriteLine "<!-- Begin main table in metadata record -->"
	objTextfile4.WriteLine "<table border='1' width='90%'>"


	objTextfile4.WriteLine "<tr><td class='opentext' valign='top'>RecordID</td><td class='opentext' valign='top'>BES_" & xrs("RecordID") & "</td></tr>"
	objTextfile4.WriteLine "<tr><td class='opentext' valign='top'>ID</td><td class='opentext' valign='top'>" & xrs("ID") & "</td></tr>"
	objTextfile4.WriteLine "<tr><td class='opentext' valign='top'>PublicationDate</td><td class='opentext' valign='top'>" & xrs("PublicationDate") & "</td></tr>"
	objTextfile4.WriteLine "<tr><td class='opentext' valign='top'>Title</td><td class='opentext' valign='top'>" & xrs("Title") & "</td></tr>"
	objTextfile4.WriteLine "<tr><td class='opentext' valign='top'>Edition</td><td class='opentext' valign='top'>" & xrs("Edition") & "</td></tr>"
	objTextfile4.WriteLine "<tr><td class='opentext' valign='top'>Publication Place</td><td class='opentext' valign='top'>" & xrs("PublicationPlace") & "</td></tr>"
	objTextfile4.WriteLine "<tr><td class='opentext' valign='top'>Publisher</td><td class='opentext' valign='top'>" & xrs("Publisher") & "</td></tr>"
	objTextfile4.WriteLine "<tr><td class='opentext' valign='top'>Online Linkage</td><td class='opentext' valign='top'>" & "<a href=http://www.beslter.org/preclick/pre-click.asp?url=" &  xrs("OnlineLinkage") & ">" & xrs("OnlineLinkage") & "</a>" & "</td></tr>"
	objTextfile4.WriteLine "<tr><td class='opentext' valign='top'>Abstract</td><td class='opentext' valign='top'>" & Replace(xrs("Abstract"), vbCrLf, "<BR>") & "</td></tr>"
	objTextfile4.WriteLine "<tr><td class='opentext' valign='top'>Purpose</td><td class='opentext' valign='top'>" & xrs("Purpose") & "</td></tr>"
	objTextfile4.WriteLine "<tr><td class='opentext' valign='top'>Supplemental Info</td><td class='opentext' valign='top'>" & xrs("SupplementalInfo") & "</td></tr>"




	objTextfile4.WriteLine "<tr><td class='opentext' valign='top'>Is GIS (1 yes 0 no)?</td><td class='opentext' valign='top'>" & xrs("GIS") & "</td></tr>"
	objTextfile4.WriteLine "<tr><td class='opentext' valign='top'>West</td><td class='opentext' valign='top'>-77.314183" & xrs("West") & "</td></tr>"
	objTextfile4.WriteLine "<tr><td class='opentext' valign='top'>East</td><td class='opentext' valign='top'>-76.012008" & xrs("East") & "</td></tr>"
	objTextfile4.WriteLine "<tr><td class='opentext' valign='top'>North</td><td class='opentext' valign='top'>39.724847" & xrs("North") & "</td></tr>"
	objTextfile4.WriteLine "<tr><td class='opentext' valign='top'>South</td><td class='opentext' valign='top'>38.708367" & xrs("South") & "</td></tr>"
	objTextfile4.WriteLine "<tr><td class='opentext' valign='top'>Presentation Form</td><td class='opentext' valign='top'>" & xrs("PresentationForm") & "</td></tr>"
	objTextfile4.WriteLine "<tr><td class='opentext' valign='top'>Data Credit</td><td class='opentext' valign='top'>" & xrs("DataCred") & "</td></tr>"
	objTextfile4.WriteLine "<tr><td class='opentext' valign='top'>Reference Name</td><td class='opentext' valign='top'>" & xrs("ThemeRef") & "</td></tr>"
	objTextfile4.WriteLine "<tr><td class='opentext' valign='top'>Reference Place</td><td class='opentext' valign='top'>" & xrs("Field23") & "</td></tr>"
	objTextfile4.WriteLine "<tr><td class='opentext' valign='top'>Reference EMail</td><td class='opentext' valign='top'>" & xrs("Field22") & "</td></tr>"
	objTextfile4.WriteLine "<tr><td class='opentext' valign='top'>Theme Keywords</td><td class='opentext' valign='top'>" & xrs("ThemeKeywords") & "</td></tr>"
	objTextfile4.WriteLine "<tr><td class='opentext' valign='top'>Place Keywords</td><td class='opentext' valign='top'>" & xrs("PlaceKeywords") & "</td></tr>"








	objTextfile4.WriteLine "</table>"
	objTextfile4.WriteLine "<!-- End main table in metadata record -->"




	objTextfile4.WriteLine "<!--#include virtual=""/metacat_harvest_attribute_level_eml/frame7-page_1_auto_bes_metadata_full_record_foot.html""-->"

	'**********************************************************
	'**********************************************************
	'****  FINISHED WRITING THE FULL RECORD HTML FILE *********
	'**********************************************************
	'**********************************************************


	






	rowcounter=rowcounter+1

	if int(rowcounter/5)=rowcounter/5 or rowcounter=1 then
		objTextFile3.WriteLine Chr(9) & "<tr bgcolor='#eeeeee'><th align='left'>Title</th><th align='left'>Creator</th><th align='left'>Pub Date</th><th align='left'>Theme Keywords</th><th align='left'>Place Keywords</th><th align='left'>Metadata</th><th align='left'>Data Link</th><th align='left'>EML</th></tr>"
	end if

	objTextFile3.WriteLine Chr(9) & "<tr>" 'New row

	
	
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	'In this part we write to the stream to add to the html file
'	Set objTextFile = objFSO.CreateTextFile(tfname, true)


	'response.write "********** NEW RECORD" & "<br>" & Vbcrlf & "<br>" & tfname & Vbcrlf


	
	response.write "emerg=" & emergencystop & "********** NEW HTML ENTRY non spatial " & xrs("title") & Vbcrlf 

	
	'filetop= "<p>BES Dataset: <br>&nbsp;<br>&nbsp;</p>"
	'objTextFile3.WriteLine filetop

	
	'Start writing elements
'	objTextFile3.WriteLine "<td colspan='7' valign='top' bgcolor='#8080ff'>Title: <b>" 
'		objTextFile3.WriteLine xrs("title") &  "</b> BES Dataset ID: " & xrs("RecordID") & "</td></tr><tr>"
		objTextFile3.WriteLine Chr(9) & "<td class='opentext' valign='top'>"

					objTextFile3.WriteLine "<b>" & xrs("title") & "</b> BES Dataset ID: " & xrs("RecordID") & "</td>"

		objTextFile3.WriteLine Chr(9) & "<td valign='top' class='opentext'>"

					objTextFile3.WriteLine xrs("publisher") & "</td>"

		objTextFile3.WriteLine Chr(9) & "<td valign='top' class='opentext'>"

			objTextFile3.WriteLine Chr(9) & Datepart("yyyy", xrs("Publicationdate")) & "-" & Datepart("m", xrs("Publicationdate")) & "-" & Datepart("d", xrs("Publicationdate")) & "</td>"




			objTextFile3.WriteLine Chr(9) & "<td valign='top' class='opentext'>"

				if len(trim(xrs("themekeywords")))>0 then
					objTextFile3.WriteLine Chr(9) & replace(xrs("ThemeKeywords"), ",", ", ") & "</td>"  
				else
					objTextFile3.WriteLine Chr(9) & "&nbsp;" & "</td>"  
				end if

			objTextFile3.WriteLine Chr(9) & "<td valign='top' class='opentext'>" 

				if len(trim(xrs("PlaceKeywords")))>0 then
					objTextFile3.WriteLine Chr(9) & replace(xrs("PlaceKeywords"), ",", ", ") & "</td>"
				else
					objTextFile3.WriteLine Chr(9) & "Not Available"
				end if

				'Write link to fullrecord file name
				objTextFile3.WriteLine Chr(9) & "<td valign='top' class='opentext'><a href='" &fullmetarecordlinkname & "'>Full Record</a></td>"



				objTextFile3.WriteLine Chr(9) & "<td valign='top' class='opentext'>"
					if len(trim(xrs("OnlineLinkage") & "*")) < 2 then   ' have to add this goddam asterisk because asp is whacky.  If onlinelinkage field is null, then it  simply doesnt exist.  no comparing it, no displaying it.  had to add star
						objTextFile3.WriteLine Chr(9) & "Not Available"
						response.write "standard url written"
					else 
						if instr(xrs("OnlineLinkage"),"\\")>0 or instr(xrs("OnlineLinkage"),"NA in x")>0 then
							objTextFile3.WriteLine "Not Available"
						else
							objTextFile3.WriteLine "<a href='" & xrs("OnlineLinkage") & "'>Dataset</a>"
						end if
						'objTextFile3.WriteLine xrs("OnlineLinkage")

					end if

				objTextFile3.WriteLine Chr(9) & "</td>"

		objTextFile3.WriteLine Chr(9) & "<td valign='top' class='opentext'> <a href='" & emllinkname & "'>" & "EML" & "</a>"
				'objTextFile3.WriteLine emllinkname
			objTextFile3.WriteLine Chr(9) & "</td>"

		objTextFile3.WriteLine Chr(9) & "</tr><tr>"
		objTextFile3.WriteLine Chr(9) & "<td bgcolor='#eeeeee' colspan='6' valign=top>"
'		objTextFile3.WriteLine "Abstract: "
		'if len(trim(xrs("Abstract")))>0 then
		'		objTextFile3.WriteLine replace(xrs("Abstract"), "&", "&amp;") & "</td>"
		'else 
		'		objTextFile3.WriteLine "Not Available" & "</td>"
		'end if



		' Print a blank space to make the web page iis easier to read
		objTextFile3.WriteLine Chr(9) & "</tr><tr>"
		objTextFile3.WriteLine Chr(9) & "<td colspan='7' valign=top>&nbsp; "
		objTextFile3.WriteLine Chr(9) & "</td> "



	
	
	
	
	
	
'	<BR>&nbsp;<BR>&nbsp;<BR>
'	<%
	
	response.write "********** END OF RECORD" & "<br>" & Vbcrlf & "<br>" & Vbcrlf
	
	'Just keeping track
	if xrs("IsPublic")=1  then
		publiccount=publiccount+1
	end if
	if xrs("lno-view")=1  and xrs("IsPublic")=1 then
		lno_view_count=lno_view_count+1
	end if

	records_processed=records_processed+1

	response.write "Processed: " 
	response.write publiccount
	response.write "  Public: "
	response.write publiccount
	response.write "  LNO View: "
	response.write lno_view_count
	response.write "   "


	
	
	
	
	xrs.movenext


	objTextFile3.WriteLine Chr(9) & "</tr>" 'Close New row








loop



'**********************************************
'Now close out html page
objTextFile3.WriteLine Chr(9) &  "</table>" ' new table - chr(9) is a tab
objTextFile3.WriteLine "<!--#include file=""frame7-page_1_auto_foot.html""-->"




'************************************************
'************************************************
'************************************************
'************************************************ END LOOP TO CREATE HTML FILES FOR COMBINED DATA










'************************************************
'************************************************
'************************************************
'************************************************
'************************************************
'************************************************  END PROGRAM, CLEAN UP
'<!---  PAGE BOTTOM PAGE BOTTOM -->	

'Write the final line to the harvest list file that has all the filenames
objTextFile2.WriteLine "</hrv:harvestList>"

Response.write "******************** PAGE BOTTOM" & "<br>" & Vbcrlf
'objTextFile3.WriteLine  "******************** PAGE BOTTOM" & "<br>" & Vbcrlf

'<!-- end embedded table of content-->

Response.Write "****************************** END OF CONTENT" & "<br>" & Vbcrlf
'objTextFile.WriteLine "****************************** END OF CONTENT" & "<br>" & Vbcrlf

set xConn = Nothing
set xrs = Nothing
objTextFile.Close
Set objTextFile = Nothing
Set objFSO = Nothing


objTextFile2.Close
Set objTextFile2 = Nothing



Set fname = Nothing
Set mRecordID = Nothing
Set cname = Nothing
Set tfname = Nothing
Set harvestlistname = Nothing
Set emergencystop = Nothing



Set accessdb = Nothing
Set db = Nothing
Set dbGenericPath = Nothing
Set dbConn = Nothing
Set dbRs = Nothing
Set strConn = Nothing
Set strTable = Nothing
Set xConn = Nothing




%>



