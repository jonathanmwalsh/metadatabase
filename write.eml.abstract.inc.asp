<% 
	'strip out line feeds and replace them with </para><para> tag pair
	'test_string=rs("abstract")
	m_abstract=Replace(rs("abstract"), vbcrlf, "</para><para>" & vbcrlf & indent5)
	'response.write m_abstract
	objTextFile.WriteLine indent4 & "<abstract>" 
	objTextFile.WriteLine indent5 & "<para>" 
	objTextFile.WriteLine indent5 & m_abstract 
	objTextFile.WriteLine indent5 & "</para>" 
	objTextFile.WriteLine indent4 & "</abstract>" 


%>



