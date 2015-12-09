<% 

	objTextFile.WriteLine indent4 & "<coverage>" 
	objTextFile.WriteLine indent5 & "<geographicCoverage>" 
	objTextFile.WriteLine indent6 & "<geographicDescription>" & rs("geographicdescription") & "</geographicDescription>" 
	objTextFile.WriteLine indent6 & "<boundingCoordinates>" 
	objTextFile.WriteLine indent6 & "<westBoundingCoordinate>" & rs("west") & "</westBoundingCoordinate>" 
	objTextFile.WriteLine indent6 & "<eastBoundingCoordinate>" & rs("east") & "</eastBoundingCoordinate>" 
	objTextFile.WriteLine indent6 & "<northBoundingCoordinate>" & rs("north") & "</northBoundingCoordinate>" 
	objTextFile.WriteLine indent6 & "<southBoundingCoordinate>" & rs("south") & "</southBoundingCoordinate>" 
	objTextFile.WriteLine indent6 & "</boundingCoordinates>" 
	objTextFile.WriteLine indent5 & "</geographicCoverage>" 

	
	'convert temporalbegin date format
	mdate=rs("temporalbegin")
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
		datepartmonthformat=datepartmonth
	end if

	 temporalbegin= datepartyear &  "-" & _
	 datepartmonthformat & _
	"-" & datepartdayformat  

	
	'convert temporalend date format
	mdate=rs("temporalend")
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
		datepartmonthformat=datepartmonth
	end if

	 temporalend= datepartyear &  "-" & _
	 datepartmonthformat & _
	"-" & datepartdayformat  

	objTextFile.WriteLine indent5 & "<temporalCoverage>" 
	objTextFile.WriteLine indent6 & "<rangeOfDates>" 
	objTextFile.WriteLine indent7 & "<beginDate>" 
	objTextFile.WriteLine indent8 & "<calendarDate>" & temporalbegin & "</calendarDate>" 
	objTextFile.WriteLine indent7 & "</beginDate>" 
	objTextFile.WriteLine indent7 & "<endDate>" 
	objTextFile.WriteLine indent8 & "<calendarDate>" & temporalend & "</calendarDate>" 
	objTextFile.WriteLine indent7 & "</endDate>" 
	objTextFile.WriteLine indent6 & "</rangeOfDates>" 
	objTextFile.WriteLine indent5 & "</temporalCoverage>" 

	objTextFile.WriteLine indent4 & "</coverage>" 
	
	
	
%>



