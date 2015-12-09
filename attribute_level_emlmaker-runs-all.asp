
<%
response.write "Do not run this script until the schema situation is handled.  In other words, there is a new EML schema.  So I modified the script.  But now if I run ANY of the old recodsa in table named COMBINED, they will get written up with the new schema.  They are in PASTA with the old one so that's not Kosher."

Response.write "@@@@@@@@@@@@@   Suggest adding another field In the database for new schema/old schema and only processing ever aghain new schema.  Note:  I did post off old schema to another table BUT, then the records dont show up on beslter.org."

'MAKE SCRIPT CRASH LOL

2+2=5
%>





<% 'Open files

dim objFSO
Set objFSO = server.CreateObject("Scripting.FileSystemObject")
dim objTextFile 'There are many of these, these will be the xml files we write
dim objTextFile2 'This will be the file harvestlist.xml
dim objTextFile3 'This will be the html page the public will see.
dim objTextFile4 'There are many of these, these will be the html files we write for the full metadata records
dim objTextfile5 ' This will be a straight text file of the urls of the files we create
Set objTextFile = objFSO.CreateTextFile(server.mappath("emloutput.txt"), true)
Set objTextFile2 = objFSO.CreateTextFile("c:\inetpub\wwwroot\metadata_harvest_attribute_level_eml\harvestlist.xml", true)
Set objTextFile5 = objFSO.CreateTextFile("c:\inetpub\wwwroot\metadata_harvest_attribute_level_eml\pastacheckurls.txt", true)
'response.write "<br>------------------------------------------hi-----------------------------------------------<br>"
Set objTextFile3 = objFSO.CreateTextFile("c:\inetpub\wwwroot\metadata_harvest_attribute_level_eml\frame7-page_1_auto.asp", true)
objTextFile2.WriteLine "<?xml version=""1.0"" encoding=""UTF-8"" ?>"

objTextFile2.WriteLine "<hrv:harvestList xmlns:hrv=""eml://ecoinformatics.org/harvestList"" >"



Const Filename = "/metadata_harvest_attribute_level_eml/lter_revision_no.txt"	' file to read
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

%>
















<%

'Run the scripts for 
'Tabular,
'OLther-Entity
	%>

	<!--#include file="attribute_level_emlmaker-for-tabular-data.asp"-->	
	<!--#include file="attribute_level_emlmaker-for-other-entity.asp"-->


	<%
	' now regenerate html metadata page for the nice humans to read.


	response.write "<br> Okay, now writing the html data catalog for the nice humans to read."
	%>
	<!--#include file="bes_data_page_generator.asp"-->   <!-- Regenerates the HTML page for the nice humans to read-->

	<% 




	set xConn = Nothing
	set xrs = Nothing
	objTextFile.Close
	Set objTextFile = Nothing
	Set objFSO = Nothing

	objTextFile2.WriteLine "</hrv:harvestList>"

	objTextFile2.Close
	Set objTextFile2 = Nothing

'	objTextFile5.Close
'	Set objTextFile5 = Nothing



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




