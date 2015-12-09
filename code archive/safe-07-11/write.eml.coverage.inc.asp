<% 

	response.write indent4 & "&lt;coverage>" & "<br />"
	response.write indent5 & "&lt;geographicCoverage>" & "<br />"
	response.write indent6 & "&lt;geographicDescription>" & rs("geographicdescription") & "</geographicDescription>" & "<br />"
	response.write indent6 & "&lt;boundingCoordinates>" & "<br />"
	response.write indent6 & "&lt;westBoundingCoordinate>" & rs("west") & "</westBoundingCoordinate>" & "<br />"
	response.write indent6 & "&lt;eastBoundingCoordinate>" & rs("east") & "</eastBoundingCoordinate>" & "<br />"
	response.write indent6 & "&lt;northBoundingCoordinate>" & rs("north") & "</northBoundingCoordinate>" & "<br />"
	response.write indent6 & "&lt;southBoundingCoordinate>" & rs("south") & "</southBoundingCoordinate>" & "<br />"
	response.write indent6 & "&lt;/boundingCoordinates>" & "<br />"
	response.write indent5 & "&lt;/geographicCoverage>" & "<br />"

	
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
		datepartmonthformat=dateparmonth
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
		datepartmonthformat=dateparmonth
	end if

	 temporalend= datepartyear &  "-" & _
	 datepartmonthformat & _
	"-" & datepartdayformat  

	response.write indent5 & "&lt;temporalCoverage>" & "<br />"
	response.write indent5 & "&lt;rangeOfDates>" & "<br />"
	response.write indent5 & "&lt;beginDate>" & "<br />"
	response.write indent6 & "&lt;calendarDate>" & temporalbegin & "&lt;/calendarDate>" & "<br />"
	response.write indent5 & "&lt;/beginDate>" & "<br />"
	response.write indent5 & "&lt;endDate>" & "<br />"
	response.write indent6 & "&lt;calendarDate>" & temporalend & "&lt;/calendarDate>" & "<br />"
	response.write indent5 & "&lt;/endDate>" & "<br />"

	response.write indent4 & "&lt;/coverage>" & "<br />"
	
	
	
%>



