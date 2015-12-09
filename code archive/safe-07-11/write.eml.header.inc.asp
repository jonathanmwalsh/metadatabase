<%

	'Module to write EML header

	'Note, we need a lot of XML literal equivelents...  We are about to write:
	'<?xml version="1.0" encoding="UTF-8"?>
	'<eml:eml packageId="knb-lter-sbx.1.5" system="https://pasta.lternet.edu"
	'xmlns:eml="eml://ecoinformatics.org/eml-2.1.0"
	'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:schemaLocation="eml://ecoinformatics.org/eml-2.1.0 http://nis.lternet.edu/schemas/EML/eml-2.1.0/eml.xsd">
	'But we need &quot; , &LT;, &GT; 

	response.write "&LT;?xml version=&quot;1.0&quot; encoding=&quot;UTF-8&quot;?&GT;" & "<br />"
	response.write "&LT;eml:eml packageId=&quot;knb-lter-bes.1.5&quot; system=&quot;https://pasta.lternet.edu&quot;" & "<br />"
	response.write indent1 & "  xmlns:eml=&quot;eml://ecoinformatics.org/eml-2.1.0"  & "<br />"
	response.write indent1 & "  xmlns:xsi=&quot;http://www.w3.org/2001/XMLSchema-instance&quot; xsi:schemaLocation=&quot;eml://ecoinformatics.org/eml-2.1.0   http://nis.lternet.edu/schemas/EML/eml-2.1.0/eml.xsd&quot;&GT;"  & "<br />"

%>
