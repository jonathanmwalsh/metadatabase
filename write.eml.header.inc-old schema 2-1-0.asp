<%

	'Module to write EML header

	lterrevisionno=rs("edition")*10
	dataset_id_num=trim(MID(rs("dataset_id"),5,6)) 'thje number 6 here is arbitrary.  It's jsut that I know my numbers are four digits long.  using "6" lets me modify may scheme to use dataset id numbers two digits longer withjout things crashing/
	if mid(dataset_id_num,1,1)="0" then
		emergstop=0
		do while mid(dataset_id_num,1,1)="0" and emergstop < 100
			dataset_id_num=mid(dataset_id_num,2,len(trim(dataset_id_num))-1)
			emergstop=emergstop+1
		loop
	end if


	'response.write lterrevisionno

	objTextFile.WriteLine "<?xml version=""1.0"" encoding=""UTF-8""?>" 

	objTextFile.WriteLine "<eml:eml packageId=""knb-lter-bes." & dataset_id_num & "." & lterrevisionno & """ system=""BES""" 
	objTextFile.WriteLine indent1 & "  xmlns:ds=""eml://ecoinformatics.org/dataset-2.1.0"""  
	objTextFile.WriteLine indent1 & "  xmlns:eml=""eml://ecoinformatics.org/eml-2.1.0"""  
	objTextFile.WriteLine indent1 & " xmlns:stmml=""http://www.xml-cml.org/schema/stmml-1.1"""
	objTextFile.WriteLine indent1 & "  xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xsi:schemaLocation=""eml://ecoinformatics.org/eml-2.1.0   http://nis.lternet.edu/schemas/EML/eml-2.1.0/eml.xsd"">"  

%>
