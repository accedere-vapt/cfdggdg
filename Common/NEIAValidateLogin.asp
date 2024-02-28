<%
'For debugging
'Response.Write Session("sEmployeeNo")


if Session("sEmployeeNo") = "" then
	
	strDescription = "Your session has expired. Please login again."
	Response.Write  "<FONT class=clsError>" & strDescription & "</FONT>"
	Response.Write "<BR><BR><HR>"
	Response.Write "<FONT class=clsTextLabel>Technical Explanation</fONT>"
	Response.Write "<BR>"
	Response.Write "Error Number: 11"
	Response.Write "<BR>"
	Response.Write "Error Description: Login expired" 
	Response.Write "<BR>"
	Response.Write "<BR>"
	Response.Write "<BR>"

%>
<CENTER>
<INPUT TYPE="button" name="btnSubmit"  value="Login Again" class ="clsButton"  onClick="document.location.href='/ecgc/login.asp'">
</CENTER>
	
<%	
end if
Response.end
%>