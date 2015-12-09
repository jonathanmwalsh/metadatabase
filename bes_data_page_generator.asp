<% 


'Open connection to database
accessdb="metadatabase.mdb"
db=Server.MapPath(accessdb)
dbGenericPath = "/attribute_level_emlmaker/"
dbConn = "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & db  'my cutesy way




'Open main spatial data recordset
dbRs = "COMBINED" 'Name of table.
strConn = dbConn
strTable = dbRs

' Open Connection to the database
set xConn = Server.CreateObject("ADODB.Connection")
xConn.Open strConn






'Do the selection and create the recordset for the spatial records
strsql = "SELECT * FROM [" & strTable & "] "
'strsql = strsql & " where IsPublic = 1 " 
'strsql = strsql & "'" & wherelike & "'"
' superceded this line because now I can have more tan one copy with the same dataset_id
'strsql=strsql & " order by RecordID ASC, ID ASC"  ' Sorting by DESC in ID field gets us our "last" record within the RecordID.  (THIS WAS FOR ORS - Records could have the same ID.  This is no longer true.  IDs are unique.  What we want is, step thru RecordID field and select greates ID within that subset of a given RecordID.
' So... sort by dataset_id + part_of_multi_id_dataset+ child_of_multi_id_dataset
strsql=strsql & " order by dataset_id ASC, child_of_multi_id_dataset ASC"  ' Sorting by DESC in ID field gets us our "last" record 
'response.write "<br>&nbsp;#####<br>" & strsql & "<br>"

set xrs = Server.CreateObject("ADODB.Recordset")
xrs.Open strsql, xConn, 1, 2  'The 1 and 2 refer to index elements like in line 3 lines above order by xxx + xxx + xxx - WATCH FOR TYPE MISMATCHES!





'<!-- INSERT PAGE TOP BEGIN PAGE TOP -->


'response.write "****************************** BEGIN EML OUTPUT" & "<br>" & Vbcrlf


'<!-- INSERT CONTENT-->



testtext="hello"
'response.write testtext
'response.write testtex

dim obj2FSO
Set obj2FSO = server.CreateObject("Scripting.FileSystemObject")
dim obj2TextFile 'There are many of these, these will be the xml files we write
dim obj2TextFile2 'This will be the file harvestlist.xml
dim obj2TextFile3 'This will be the html page the public will see.
dim obj2TextFile4 'There are many of these, these will be the html files we write for the full metadata records
Set obj2TextFile = obj2FSO.CreateTextFile(server.mappath("emloutput2.txt"), true)
Set obj2TextFile2 = obj2FSO.CreateTextFile("c:\inetpub\wwwroot\metadata_harvest_attribute_level_eml\harvestlist_no_need.xml", true)
Set obj2TextFile3 = obj2FSO.CreateTextFile("c:\inetpub\wwwroot\metadata_harvest_attribute_level_eml\frame7-page_1_auto_no_need.asp", true)
obj2TextFile2.WriteLine "<?xml version=""1.0"" encoding=""UTF-8"" ?>"

obj2TextFile2.WriteLine "<hrv:harvestList xmlns:hrv=""eml://ecoinformatics.org/harvestList"" >"



Const Filename2 = "/metadata_harvest_attribute_level_eml/lter_revision_no.txt"	' file to read
Const ForReading2 = 1, ForWriting2 = 2, ForAppending2 = 3
Const TristateUseDefault2 = -2, TristateTrue2 = -1, TristateFalse2 = 0



' Map the logical path to the physical system path
Dim Filepath2
dim lter_rev_no2
lter_rev_no2=1

Filepath2 = Server.MapPath(Filename)

if obj2FSO.FileExists(Filepath2) Then

	Set TextStream2 = obj2FSO.OpenTextFile(Filepath2, ForReading, False, _
											  TristateUseDefault)
	' Read file in one hit
		
	Dim Contents2
	Contents2 = Textstream2.ReadAll
	'Response.write "<pre>Contents of revision file set at: " & ltrim(trim(Contents2)) & ".  Change file lter_revision_no.txt to adjust this.</pre><hr>"
	Textstream2.Close
	Set Textstream2 = nothing
	lter_rev_no2=trim(ltrim(Contents2))
	

Else

	Response.Write "<h3><i><font color=red> File " & Filename &_
                       " does not exist</font></i></h3>"

End If

Response.write "<pre>" & lter_rev_no2 & "</pre><hr>"



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
'response.write " XXXXXXXXXXXXXXXXXXX EQUALS:"&xxx&"***"
xrs.movefirst
fname=0
mRecordID = 0
emergencystop=0







	
'************************************************
'************************************************
'************************************************
'************************************************ BEGIN LOOP TO WRITE SPATIAL AND NON SPATIAL DATA TO HTML FILES USING TABLE COMBINED
xrs.requery
xrs.movefirst
'mRecordID = 0
'fname=0 'yes, reinitialize this because we will add a different prefix to the filenames ' but no, do not reinitialize because it messes up scope identifier on other end


	



obj2TextFile3.WriteLine "<html>" 'make the top of the html document
obj2TextFile3.WriteLine "<head></head>" 

obj2TextFile3.WriteLine "<body>"
'Write HTML File top

obj2TextFile3.WriteLine "<!--#include file=""frame7-page_1_auto_head.html""-->"

obj2TextFile3.WriteLine Chr(9) &  "<table border='0' bordercolor='blue'>" ' new table - chr(9) is a tab



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
	emllinkname="http://beslter.org/metadata_harvest_attribute_level_eml/zzzoldbes_" & cname & ".xml"


	'**********************************************************
	'**********************************************************
	'******  HERE WE WRITE THE FULL RECORD HTML FILE **********
	'**********************************************************
	'**********************************************************

	'Create  name and file handle for individual html file of full metadata record
	fullmetarecordfilename="c:\inetpub\wwwroot\metadata_harvest_attribute_level_eml\html_metadata\bes_" & cname & ".asp"
	fullmetarecordlinkname="/metadata_harvest_attribute_level_eml/html_metadata/bes_" & cname & ".asp"
	response.write fullmetarecordlinkname

	'Just try writing anything to our new obj2TextFile4 which is html page of full metadata record
	Set obj2TextFile4 = obj2FSO.CreateTextFile(fullmetarecordfilename, true)
	obj2Textfile4.WriteLine "<html>" 'make the top of the html document
	obj2Textfile4.WriteLine "<head></head>" 

	obj2Textfile4.WriteLine "<body>"
	'Write HTML File top

	obj2Textfile4.WriteLine "<!--#include virtual  =""/metadata_harvest_attribute_level_eml/frame7-page_1_auto_bes_metadata_full_record_head.html""-->"

	obj2Textfile4.WriteLine "<!-- Begin main table in metadata record -->"
	obj2Textfile4.WriteLine "<table border='1' width='90%'>"


	obj2Textfile4.WriteLine "<tr><td class='opentext' valign='top'>RecordID</td><td class='opentext' valign='top'>BES_" & xrs("RecordID") & "</td></tr>"
	obj2Textfile4.WriteLine "<tr><td class='opentext' valign='top'>ID</td><td class='opentext' valign='top'>" & xrs("ID") & "</td></tr>"
	obj2Textfile4.WriteLine "<tr><td class='opentext' valign='top'>PublicationDate</td><td class='opentext' valign='top'>" & xrs("PublicationDate") & "</td></tr>"
	obj2Textfile4.WriteLine "<tr><td class='opentext' valign='top'>Title</td><td class='opentext' valign='top'>" & xrs("Title") & "</td></tr>"
	obj2Textfile4.WriteLine "<tr><td class='opentext' valign='top'>Edition</td><td class='opentext' valign='top'>" & xrs("Edition") & "</td></tr>"
	obj2Textfile4.WriteLine "<tr><td class='opentext' valign='top'>Publication Place</td><td class='opentext' valign='top'>" & xrs("PublicationPlace") & "</td></tr>"
	obj2Textfile4.WriteLine "<tr><td class='opentext' valign='top'>Publisher</td><td class='opentext' valign='top'>" & xrs("Publisher") & "</td></tr>"
	obj2Textfile4.WriteLine "<tr><td class='opentext' valign='top'>Online Linkage</td><td class='opentext' valign='top'>" & "<a href=http://www.beslter.org/preclick/pre-click.asp?url=" &  xrs("OnlineLinkage") & ">" & xrs("OnlineLinkage") & "</a>" & "</td></tr>"
	obj2Textfile4.WriteLine "<tr><td class='opentext' valign='top'>Abstract</td><td class='opentext' valign='top'>" & Replace(xrs("Abstract"), vbCrLf, "<BR>") & "</td></tr>"
	obj2Textfile4.WriteLine "<tr><td class='opentext' valign='top'>Purpose</td><td class='opentext' valign='top'>" & xrs("Purpose") & "</td></tr>"
	obj2Textfile4.WriteLine "<tr><td class='opentext' valign='top'>Supplemental Info</td><td class='opentext' valign='top'>" & xrs("SupplementalInfo") & "</td></tr>"




	obj2Textfile4.WriteLine "<tr><td class='opentext' valign='top'>Is GIS (1 yes 0 no)?</td><td class='opentext' valign='top'>" & xrs("GIS") & "</td></tr>"
	obj2Textfile4.WriteLine "<tr><td class='opentext' valign='top'>West</td><td class='opentext' valign='top'>" & xrs("West") & "</td></tr>"
	obj2Textfile4.WriteLine "<tr><td class='opentext' valign='top'>East</td><td class='opentext' valign='top'>" & xrs("East") & "</td></tr>"
	obj2Textfile4.WriteLine "<tr><td class='opentext' valign='top'>North</td><td class='opentext' valign='top'>" & xrs("North") & "</td></tr>"
	obj2Textfile4.WriteLine "<tr><td class='opentext' valign='top'>South</td><td class='opentext' valign='top'>" & xrs("South") & "</td></tr>"
	obj2Textfile4.WriteLine "<tr><td class='opentext' valign='top'>Presentation Form</td><td class='opentext' valign='top'>" & xrs("PresentationForm") & "</td></tr>"
	obj2Textfile4.WriteLine "<tr><td class='opentext' valign='top'>Data Credit</td><td class='opentext' valign='top'>" & xrs("DataCred") & "</td></tr>"
	obj2Textfile4.WriteLine "<tr><td class='opentext' valign='top'>Reference Name</td><td class='opentext' valign='top'>" & xrs("ThemeRef") & "</td></tr>"
	obj2Textfile4.WriteLine "<tr><td class='opentext' valign='top'>Reference Place</td><td class='opentext' valign='top'>" & xrs("Field23") & "</td></tr>"
	obj2Textfile4.WriteLine "<tr><td class='opentext' valign='top'>Reference EMail</td><td class='opentext' valign='top'>" & xrs("Field22") & "</td></tr>"
	obj2Textfile4.WriteLine "<tr><td class='opentext' valign='top'>Theme Keywords</td><td class='opentext' valign='top'>" & xrs("ThemeKeywords") & "</td></tr>"
	obj2Textfile4.WriteLine "<tr><td class='opentext' valign='top'>Place Keywords</td><td class='opentext' valign='top'>" & xrs("PlaceKeywords") & "</td></tr>"

	'Write some attribute information if xrs("pastaview")=1
	if xrs("pastaview")=1 then


		'Sub table
		obj2Textfile4.WriteLine "<tr><td colspan='2'><table border='0' cellpadding='2' bordercolor='#dddddd' width='100%'><tr>" 'OPEN SUB TABLE
		obj2Textfile4.WriteLine "<tr><td class='opentext' valign='top'>&nbsp;<BR>&nbsp;<BR>&nbsp;<BR>LTER Network Information System attribute information...</td></tr>"
		obj2Textfile4.WriteLine "</tr></table>" 'END SUB TABLE

		
		
		obj2Textfile4.WriteLine "<tr><td colspan='2'><table border='1' cellpadding='1' bordercolor='#dddddd' width='100%'><tr>" 'OPEN SUB TABLE



		'headers
		obj2Textfile4.WriteLine "<td class='opentext' valign='top'>Filetype</td>"
		obj2Textfile4.WriteLine "<td class='opentext' valign='top'>onlineinfolink</td>"
		obj2Textfile4.WriteLine "<td class='opentext' valign='top'>filesizeunit</td>"
		obj2Textfile4.WriteLine "<td class='opentext' valign='top'>filesize</td>"
		obj2Textfile4.WriteLine "<td class='opentext' valign='top'>characterencoding</td>"
		obj2Textfile4.WriteLine "<td class='opentext' valign='top'>dataformat</td>"
		obj2Textfile4.WriteLine "<td class='opentext' valign='top'>numheaderlines</td>"
		obj2Textfile4.WriteLine "<td class='opentext' valign='top'>numfooterlines</td>"
		obj2Textfile4.WriteLine "<td class='opentext' valign='top'>recorddelimiter</td>"
		obj2Textfile4.WriteLine "</tr><tr>" 'END SUB TABLE ROW BEGIN NEW ROW

		'contents

		obj2Textfile4.WriteLine "<td class='opentext' valign='top'>&nbsp;" & xrs("filetype") & "</td>"
		obj2Textfile4.WriteLine "<td class='opentext' valign='top'>" & xrs("onlineinfolink") & "</td>"
		obj2Textfile4.WriteLine "<td class='opentext' valign='top'>" & xrs("filesizeunit") & "</td>"
		obj2Textfile4.WriteLine "<td class='opentext' valign='top'>" & xrs("filesize") & "</td>"
		obj2Textfile4.WriteLine "<td class='opentext' valign='top'>" & xrs("characterencoding") & "</td>"
		obj2Textfile4.WriteLine "<td class='opentext' valign='top'>" & xrs("dataformat") & "</td>"
		obj2Textfile4.WriteLine "<td class='opentext' valign='top'>" & xrs("numheaderlines") & "</td>"
		obj2Textfile4.WriteLine "<td class='opentext' valign='top'>" & xrs("numfooterlines") & "</td>"
		obj2Textfile4.WriteLine "<td class='opentext' valign='top'>" & xrs("recorddelimiter") & "</td>"



		obj2Textfile4.WriteLine "</tr></table>" 'END SUB TABLE
		

		
		
		obj2Textfile4.WriteLine "<tr><td colspan='2'><table border='1' cellpadding='1' bordercolor='#dddddd' width='100%'><tr>" 'OPEN SUB TABLE



		'headers
		obj2Textfile4.WriteLine "<td class='opentext' valign='top'>linesperrecord</td>"
		obj2Textfile4.WriteLine "<td class='opentext' valign='top'>fielddelimiter</td>"
		obj2Textfile4.WriteLine "<td class='opentext' valign='top'>quotecharacter</td>"
		obj2Textfile4.WriteLine "<td class='opentext' valign='top'>orientation</td>"
		obj2Textfile4.WriteLine "<td class='opentext' valign='top'>dataentitytype (1=tabular 2='otherentity')</td>"
		obj2Textfile4.WriteLine "<td class='opentext' valign='top'>creatorid</td>"
		obj2Textfile4.WriteLine "<td class='opentext' valign='top'>temporalbegin</td>"
		obj2Textfile4.WriteLine "<td class='opentext' valign='top'>temporalend</td>"
		obj2Textfile4.WriteLine "<td class='opentext' valign='top'>PASTA view</td>"
		obj2Textfile4.WriteLine "<td class='opentext' valign='top'>link to EML record</td>"

		obj2Textfile4.WriteLine "</tr><tr>" 'END SUB TABLE ROW BEGIN NEW ROW

		'contents

		obj2Textfile4.WriteLine "<td class='opentext' valign='top'>" & xrs("linesperrecord") & "</td>"
		obj2Textfile4.WriteLine "<td class='opentext' valign='top'>" & xrs("fielddelimiter") & "</td>"
		obj2Textfile4.WriteLine "<td class='opentext' valign='top'>" & xrs("quotecharacter") & "</td>"
		obj2Textfile4.WriteLine "<td class='opentext' valign='top'>" & xrs("orientation") & "</td>"
		obj2Textfile4.WriteLine "<td class='opentext' valign='top'>" & xrs("dataentitytype") & "</td>"
		obj2Textfile4.WriteLine "<td class='opentext' valign='top'>" & xrs("creatorid") & "</td>"
		obj2Textfile4.WriteLine "<td class='opentext' valign='top'>" & xrs("temporalbegin") & "</td>"
		obj2Textfile4.WriteLine "<td class='opentext' valign='top'>" & xrs("temporalend") & "</td>"
		obj2Textfile4.WriteLine "<td class='opentext' valign='top'>" & xrs("pastaview") & "</td>"
		'if it's a multi-attribute dataset must link to the EML record belonging to the parent because it's the only one.
		If xrs("part_of_multi_id_dataset")=1 then
			'response.write "<br>processing multi_id_dataset<br>"
			if xrs("child_of_multi_id_dataset")=0 then 'there will be just one of these for evey multi id set.
				'response.write "<br> SETTTTING AN EML LINK NAME  ---"

				emllinkingname="http://beslter.org/metadata_harvest_attribute_level_eml/" & "knb-lter-bes-" & xrs("recordid") & ".xml"
				response.write emllinkingname & "<br>"
			end if
			response.write "<br> EMLlinkname is " & emllinkingname & "<br>"
			'Otherwise, don't set it. 
			'Only set the emllinkingname for the parent.  There is only one eml file for the partent and all children
			'Since dataset is sorted by dataset_id PLUS child_of_multi_id_dataset, this will keep the firrst emllinkingname for all children as the parent
			'multi_id_emllinkname
			 'and xrs("child_of_multi_id_dataset")=1 then

		else 
			emllinkingname="http://beslter.org/metadata_harvest_attribute_level_eml/" & "knb-lter-bes-" & xrs("recordid") & ".xml"
		end if

		if xrs("pastaview")=1 then
			obj2Textfile4.WriteLine "<td class='opentext' valign='top'><a href='" & emllinkingname & "'>HERE</a></td>"
		else
			obj2Textfile4.WriteLine "<td class='opentext' valign='top'>" & "N/A" & "</td>"
		end if




		obj2Textfile4.WriteLine "</tr></table>" 'END SUB TABLE
		
	end if 'pastview=1 and we write out attribute level data





	obj2Textfile4.WriteLine "</table>"
	obj2Textfile4.WriteLine "<!-- End main table in metadata record -->"




	obj2Textfile4.WriteLine "<!--#include virtual=""/metadata_harvest_attribute_level_eml/frame7-page_1_auto_bes_metadata_full_record_foot.html""-->"

	'**********************************************************
	'**********************************************************
	'****  FINISHED WRITING THE FULL RECORD HTML FILE *********
	'**********************************************************
	'**********************************************************


	






	rowcounter=rowcounter+1
	if int(rowcounter/5)=rowcounter/5 or rowcounter=1 then

		obj2TextFile3.WriteLine Chr(9) & "<tr bgcolor='#eeeeee'><th align='left'>Title</th><th align='left'>Creator</th><th align='left'>Pub Date</th><th align='left'>Theme Keywords</th><th align='left'>Place Keywords</th><th align='left'>Metadata</th><th align='left'>Data Link</th><th align='left'>EML</th></tr>"
	end if

	obj2TextFile3.WriteLine Chr(9) & "<tr>" 'New row

	
	
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	'In this part we write to the stream to add to the html file
'	Set obj2TextFile = obj2FSO.CreateTextFile(tfname, true)


	'response.write "********** NEW RECORD" & "<br>" & Vbcrlf & "<br>" & tfname & Vbcrlf


	
	response.write " Counter=" & emergencystop & " *** Title: " & xrs("title") 

	
	'filetop= "<p>BES Dataset: <br>&nbsp;<br>&nbsp;</p>"
	'obj2TextFile3.WriteLine filetop

	
	'Start writing elements
'	obj2TextFile3.WriteLine "<td colspan='7' valign='top' bgcolor='#8080ff'>Title: <b>" 
'		obj2TextFile3.WriteLine xrs("title") &  "</b> BES Dataset ID: " & xrs("RecordID") & "</td></tr><tr>"
		obj2TextFile3.WriteLine Chr(9) & "<td class='opentext' valign='top'>"

					obj2TextFile3.WriteLine "<b>" & xrs("title") & "</b> BES Dataset ID: " & xrs("RecordID") & "</td>"

		obj2TextFile3.WriteLine Chr(9) & "<td valign='top' class='opentext'>"

					obj2TextFile3.WriteLine xrs("publisher") & "</td>"

		obj2TextFile3.WriteLine Chr(9) & "<td valign='top' class='opentext'>"

			obj2TextFile3.WriteLine Chr(9) & Datepart("yyyy", xrs("Publicationdate")) & "-" & Datepart("m", xrs("Publicationdate")) & "-" & Datepart("d", xrs("Publicationdate")) & "</td>"




			obj2TextFile3.WriteLine Chr(9) & "<td valign='top' class='opentext'>"

				if len(trim(xrs("themekeywords")))>0 then
					obj2TextFile3.WriteLine Chr(9) & replace(xrs("ThemeKeywords"), ",", ", ") & "</td>"  
				else
					obj2TextFile3.WriteLine Chr(9) & "&nbsp;" & "</td>"  
				end if

			obj2TextFile3.WriteLine Chr(9) & "<td valign='top' class='opentext'>" 

				if len(trim(xrs("PlaceKeywords")))>0 then
					obj2TextFile3.WriteLine Chr(9) & replace(xrs("PlaceKeywords"), ",", ", ") & "</td>"
				else
					obj2TextFile3.WriteLine Chr(9) & "Not Available"
				end if

				'Write link to fullrecord file name
				obj2TextFile3.WriteLine Chr(9) & "<td valign='top' class='opentext'><a href='" &fullmetarecordlinkname & "'>Full Record</a></td>"



				obj2TextFile3.WriteLine Chr(9) & "<td valign='top' class='opentext'>"
					if len(trim(xrs("OnlineLinkage") & "*")) < 2 then   ' have to add this goddam asterisk because asp is whacky.  If onlinelinkage field is null, then it  simply doesnt exist.  no comparing it, no displaying it.  had to add star
						obj2TextFile3.WriteLine Chr(9) & "Not Available"
						response.write "standard url written"
					else 
						if instr(xrs("OnlineLinkage"),"\\")>0 or instr(xrs("OnlineLinkage"),"NA in x")>0 then
							obj2TextFile3.WriteLine "Not Available"
						else
							obj2TextFile3.WriteLine "<a href='" & xrs("OnlineLinkage") & "'>Dataset</a>"
						end if
						'obj2TextFile3.WriteLine xrs("OnlineLinkage")

					end if

				obj2TextFile3.WriteLine Chr(9) & "</td>"

		obj2TextFile3.WriteLine Chr(9) & "<td valign='top' class='opentext'> <a href='" & emllinkname & "'>" & "EML" & "</a>"
				'obj2TextFile3.WriteLine emllinkname
			obj2TextFile3.WriteLine Chr(9) & "</td>"

		obj2TextFile3.WriteLine Chr(9) & "</tr><tr>"
		obj2TextFile3.WriteLine Chr(9) & "<td bgcolor='#eeeeee' colspan='6' valign=top>"
'		obj2TextFile3.WriteLine "Abstract: "
		'if len(trim(xrs("Abstract")))>0 then
		'		obj2TextFile3.WriteLine replace(xrs("Abstract"), "&", "&amp;") & "</td>"
		'else 
		'		obj2TextFile3.WriteLine "Not Available" & "</td>"
		'end if



		' Print a blank space to make the web page iis easier to read
		obj2TextFile3.WriteLine Chr(9) & "</tr><tr>"
		obj2TextFile3.WriteLine Chr(9) & "<td colspan='7' valign=top>&nbsp; "
		obj2TextFile3.WriteLine Chr(9) & "</td> "



	
	
	
	
	
	
'	<BR>&nbsp;<BR>&nbsp;<BR>
'	<%
	
	response.write "**** END OF RECORD" & "<br>"
	
	'Just keeping track
	if xrs("IsPublic")=1  then
		publiccount=publiccount+1
	end if
	if xrs("lnoview")=1  and xrs("IsPublic")=1 then
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


	obj2TextFile3.WriteLine Chr(9) & "</tr>" 'Close New row








loop



'**********************************************
'Now close out html page
obj2TextFile3.WriteLine Chr(9) &  "</table>" ' new table - chr(9) is a tab
obj2TextFile3.WriteLine "<!--#include file=""frame7-page_1_auto_foot.html""-->"




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
obj2TextFile2.WriteLine "</hrv:harvestList>"

'Response.write "******************** PAGE BOTTOM" & "<br>" & Vbcrlf
'obj2TextFile3.WriteLine  "******************** PAGE BOTTOM" & "<br>" & Vbcrlf

'<!-- end embedded table of content-->

Response.Write "****************************** END OF CONTENT" & "<br>" & Vbcrlf
'obj2TextFile.WriteLine "****************************** END OF CONTENT" & "<br>" & Vbcrlf

set xConn = Nothing
set xrs = Nothing
obj2TextFile.Close
Set obj2TextFile = Nothing
Set obj2FSO = Nothing


obj2TextFile2.Close
Set obj2TextFile2 = Nothing



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
