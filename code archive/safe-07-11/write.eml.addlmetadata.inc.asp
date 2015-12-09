<% 
	response.write indent2 & "&LT;additionalMetadata>"  & "<br />"'optional
    response.write indent3 & "&LT;metadata>" & "<br />"
    response.write indent4 & "&LT;unitList>" & "<br />"
    response.write indent5 & "&LT;unit abbreviation=""m-1"" id=""reciprocalMeter"" multiplerToSI=""1""" & "<br />"
    response.write indent5 & "name=""reciprocalMeter"" parentSI=""meter"" unitType=""lengthReciprocal"">" & "<br />"
    response.write indent6 & "&LT;description>per meter, describes optical properties&LT;/description>" & "<br />"
    response.write indent5 & "&LT;/unit>" & "<br />"
    response.write indent5 & "&LT;unit abbreviation=""ft"" id=""foot"" multiplerToSI=""2.54""" & "<br />"
    response.write indent5 & "name=""ft"" parentSI=""meter"" unitType=""length"">" & "<br />"
    response.write indent6 & "&LT;description>how long is something, describes physical properties&LT;/description>" & "<br />"
    response.write indent5 & "&LT;/unit>" & "<br />"
    response.write indent4 & "&LT;/unitList>" & "<br />"
    response.write indent3 & "&LT;/metadata>" & "<br />"
    response.write indent2 & "&LT;/additionalMetadata>" & "<br />"

%>



