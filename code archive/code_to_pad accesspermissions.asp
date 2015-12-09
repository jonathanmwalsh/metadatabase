<% 
response.write "hi"


strsql="SELECT * FROM accesspermissions"

set conn=Server.CreateObject("ADODB.Connection")
conn.Provider="Microsoft.Jet.OLEDB.4.0"
conn.Open "c:/inetpub/wwwroot/attribute_level_emlmaker/esri2bes_int_osrs-search_jwalsh.mdb"

'OPEN MAIN DATASET
set rs = Server.CreateObject("ADODB.recordset")
rs.Open strSQL, conn, 3, 3


rs.movefirst

emergencystop=0

mrepldataset_ID="ddd"
mreplnum=0

'odd even flag
flag=1


do while not rs.EOF and emergencystop<2000

	emergencystop=emergencystop+1
	response.write rs("dataset_id") & " * " 
	response.write emergencystop

	if flag=1 then
		puttext="uid=BES,o=lter,dc=ecoinformatics,dc=org"
		putpermission="all"
		flag=0
	else
		puttext="public"
		putpermission="read"
		flag=1
	end if
	
	if mreplnum<10 then
		mreplnumchar=" 000" & mreplnum
	elseif mreplnum<100 then
		mreplnumchar=" 00" & mreplnum
	elseif mreplnum<1000 then
		mreplnumchar=" 0" & mreplnum
	else
		mreplnumchar="  " & mreplnum
	end if 
		
	mreplnumchar="BES_" & trim(mreplnumchar)
	response.write " GGGGG "& mreplnumchar & " GGGGG "

	rs("allowdeny")="allow"
	rs("principal")=puttext
	rs("permission")=putpermission
		
	response.write  " <br/>" 
	
	
	rs.movenext
	if rs.EOF then
		response.write "EOF ########################################"
	end if




loop

%> 

