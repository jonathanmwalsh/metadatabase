<% 


	'###################################################################################################
	'########################################## KEYWORDS
	'###################################################################################################
	response.write indent4 & "&LT;keywordSet>" & "<br />"
	if len(trim(rs("ThemeKeywords")))>0 then
		newword="anything"
		stoppy=0 'Stoppy keeps the loop from going indefinitely
		testword=trim(rs("ThemeKeywords"))
		'testword="peter, paul, mary, george, ralphie"
		do while stoppy<100 'Stoppy keeps the loop from going indefinitely
			stoppy=stoppy+1
			'response.write "<br /> ####################### <br />"
			'response.write vbcrlf
			'response.write "stoppy= " & stoppy & "***"
			'response.write " testword is *" & testword  & "***"
			testword=replace(testword, ",  ", ", ")
			origlen=len(testword)
			'response.write " length original = " & origlen & "***"
			'response.write " " & testword & "<br />"
			templen=instr(testword, ",")
			'response.write " templen= " & templen & " *** "
			nextfulllen=origlen-templen
			'response.write " Nextfullen= " & nextfulllen & " *** "
			'response.write " position of commma " & instr(testword, ",") & "<br />"
			if instr(testword,",")>0 then
				thekeyword= trim(left(testword,templen-1)) '-1 because to lop off comma/space
			else 
				thekeyword= trim(testword)  ' since no comma, estlen drops to zero, so just set it to testword
			end if

			'response.write " new keyword is *" & thekeyword & "*<br />"
			newword= mid(testword, templen +1, nextfulllen) 
			'response.write " substring (newword) = *" & newword & "*<br />"
			'response.write "length of newword=" & len(newword)
			'WRITE IT TO FILE
			response.write indent5 & "&LT;keyword keywordType=""theme"">"& thekeyword &"&LT;/keyword>" & "<br />"
			if instr(testword,",")=0 then 'no more commas, must be on the last keyword.  Out we go.
				'response.write "    EXITING DO <br />"
				exit do
			end if
			testword=newword
			testwordlen=len(testword)
			'response.write "<br /> ####################### <br />"
		loop
	else
		response.write indent4 & "&LT;keyword keywordType=""theme"">Not Available</keyword>" & "<br />"
	end if
	response.write indent4 & "&LT;/keywordSet>" & "<br />"

	
	response.write indent4 & "&LT;keywordSet>" & "<br />"
	if len(trim(rs("PlaceKeywords")))>0 then
		newword="anything"
		stoppy=0 'Stoppy keeps the loop from going indefinitely
		testword=trim(rs("PlaceKeywords"))
		'testword="peter, paul, mary, george, ralphie"
		do while stoppy<100 'Stoppy keeps the loop from going indefinitely
			stoppy=stoppy+1
			'response.write "<br /> ####################### <br />"
			'response.write vbcrlf
			'response.write "stoppy= " & stoppy & "***"
			'response.write " testword is *" & testword  & "***"
			testword=replace(testword, ",  ", ", ")
			origlen=len(testword)
			'response.write " length original = " & origlen & "***"
			'response.write " " & testword & "<br />"
			templen=instr(testword, ",")
			'response.write " templen= " & templen & " *** "
			nextfulllen=origlen-templen
			'response.write " Nextfullen= " & nextfulllen & " *** "
			'response.write " position of commma " & instr(testword, ",") & "<br />"
			if instr(testword,",")>0 then
				thekeyword= trim(left(testword,templen-1)) '-1 because to lop off comma/space
			else 
				thekeyword= trim(testword)  ' since no comma, estlen drops to zero, so just set it to testword
			end if

			'response.write " new keyword is *" & thekeyword & "*<br />"
			newword= mid(testword, templen +1, nextfulllen) 
			'response.write " substring (newword) = *" & newword & "*<br />"
			'response.write "length of newword=" & len(newword)
			'WRITE IT TO FILE
			response.write indent5 & "&LT;keyword keywordType=""theme"">"& thekeyword &"&LT;/keyword>" & "<br />"
			if instr(testword,",")=0 then 'no more commas, must be on the last keyword.  Out we go.
				'response.write "    EXITING DO <br />"
				exit do
			end if
			testword=newword
			testwordlen=len(testword)
			'response.write "<br /> ####################### <br />"
		loop
	else
		response.write indent4 & "&LT;keyword keywordType=""theme"">Not Available</keyword>" & "<br />"
	end if
	response.write indent4 & "&LT;/keywordSet>" & "<br />"

	
	
	
	

	'response.write <keywordThesaurus>LTER Controlled Vocabulary</keywordThesaurus>

	'###################################################################################################
	'########################################## KEYWORDS
	'###################################################################################################



%>



