<HTML>

<HEAD>

</HEAD>
<BODY>
<OBJECT id=CrystalReport1 style="LEFT: 0px; TOP: 0px" 
classid=clsid:00025601-0000-0000-C000-000000000046 
data=data:application/x-oleobject;base64,AVYCAAAAAADAAAAAAAAARgBQBQAAAAAA3frdut363brlAgAA5QIAAAEAAQAAAAAAAAAAAAAAAAAsAWQAZADqAQIAAAACAAAABwAAAAAAAAACAAAAAAAAAAAAAAAAAAAAAQABAAEAAQABPAAAAAEBAQEBAAEAAABwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA== 
VIEWASTEXT runat="server"></OBJECT>

<% 

'on error resume next
Function GenRpt(intRptID,intDest,intFileType,intProp_no,strParam)
'**********This Function is used to generate the Crystal report using Seagate Crystal Report Control
'**********Input 1. intRptID = This is a Unique id given to Each report and will be used to make query to RPTLIB table
'**********Input 2. intDest = This is used to define the Destination of generated Crystal Report
'**********Input 3. intFileType = This is used to define the generated report file type
'**********Input 4. First Parameter passed as formula to report and also used to create the generated file name
'**********Input 5. Another Parameter
'**********Output   1 if Report generated succesfully and 0 for No report  					 	 

Dim intcreated 
Dim objCn
Dim objRs 
Dim strSQL
Dim strRptFileName, strRptSPName
Dim strGenRptName
Dim strRptParams
Dim strFormulas	
	Set objCn = Server.CreateObject("ADODB.Connection")
	Set objRs = Server.CreateObject("ADODB.Recordset")

	objCn.Open astrconn

	strSQL = "Select * from RPTLIB where RL_RPTID ="& intRptID

	objRs.CursorLocation=3

	objRs.Open strSQL,objCn

	strRptFileName = objRs.Fields("RL_RPTFILENAME")	
	strRptFileName = strTemplateCrystalReportsPath & Trim(strRptFileName)
	
	strRptParams = split(strParam,"|")
	
	
	strGenRptName = strGenCrystalReportsPath & trim(objRs("RL_RPTFILECD"))& "_" & intProp_no & ".rpt"		
	
	if objRs.Fields("RL_SPFLAG")= "Y" then				
		
		strRptSPName = objRs.Fields("RL_SPNAME")
		
		strSQL = "Begin " & trim(strRptSPName) & "(" & intProp_no & ", " & trim(session.SessionID) & ");End;"			
		
		objCn.Execute strSQL
		
	end if	
	
	objRs.Close
	set objRs = Nothing
	objCn.Close
	Set objCn = Nothing
	
	
	CrystalReport1.ReportFileName = strRptFileName	
	CrystalReport1.Destination = intDest	
	for intCouter = 0 to ubound(strRptParams)
		
		CrystalReport1.Formulas(cint(intCouter))= TRIM(strRptParams(intCouter))
				
	Next 
		
	CrystalReport1.Connect = astrReportConn	
	if intdest = 2 then		
		
		CrystalReport1.PrintFileType= intFileType	
		CrystalReport1.PrintFileName = strGenRptName		
	end if	
	Response.Write "Before calling"		
	CrystalReport1.Action = 1
	Response.Write "After Printing"
	Response.End
	intcreated =CrystalReport1.RecordsSelected
	
	if intcreated > 0 then			
		GenRpt=  1
	else			
		GenRpt = 0
	end if

end function

%></BODY></HTML>
