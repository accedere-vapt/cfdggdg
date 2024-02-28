<%
   Set objRCountry = Server.CreateObject ("ADODB.Recordset")
   Set objConn = Server.CreateObject ("ADODB.Connection")
	objConn.Open astrConn 'strcon	Open the Connection to the Database
	objRCountry.Open "select ctry_cd,ctry_name from ctry_mst", objConn
	 'To get next value from sequence using the connection.
%>    
    <!--<select id="selBuyerCountry" name="selBuyerCountry"> -->
<!-- <option value=India selected>India</option> -->
<%
if not objRCountry.EOF then
while not objRCountry.EOF
%>
      
   <!--  <option value="<%=objRCountry("ctry_cd")%>"><%=objRCountry("ctry_name")%></option> -->
    <option value="<%=objRCountry("ctry_name")%>"><%=objRCountry("ctry_name")%></option>
<%
objRCountry.MoveNext
wend
end if
objRCountry.close
Set objRCountry = Nothing
objConn.close
Set objConn = Nothing
%>

