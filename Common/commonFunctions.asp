<%
'**********************************************************
' This File Contains all common  Functions
'**********************************************************
'Function name : FnDate
'Purpose       : This will convert the date from mm/dd/yyyy to dd/mm/yyyy format
'Input         : 
'Output        : 
'Author        : TataInfotech
'Date          : 4-11-2001
'**********************************************************
function FnDate(Eff_date)
     if Eff_date <> ""  then
         dim dt,x,y,z,temp
         dt = cdate(Eff_date)
         x = datepart("d" ,dt)
         y = datepart("m" ,dt)
         z = datepart("yyyy" ,dt)
         if len(x)<2 then
         x = 0 & x
         end if
         if len(y)<2 then
         y = 0 & y
         end if
         temp = x & "/" & y & "/" & z
         Fndate=temp
    else
          Fndate=""
   end if
end function

'**********************************************************
'Function name : fnDisplayDatePicker
'Purpose       : This will include the date picker with date field
'Input         : 
'Output        : 
'Author        : TataInfotech
'Date          : 4-11-2001
'**********************************************************

function fnDisplayDatePicker(strFieldName)
	
	Dim strTemp
	
	strTemp = "<a href=""javascript:show_calendar('" & strFieldName & "');"" " & vbcrlf
	strTemp = strTemp & " onmouseover=""window.status='Date Picker';return true;"" "  & vbcrlf
	strTemp = strTemp & " onmouseout=""window.status='';return true;"">" & vbcrlf
	strTemp = strTemp & " <img src=""../../../images/show-calendar.gif"" border=0></a>" & vbcrlf
	'strTemp = strTemp & " <img src=""" & Application("strImagesPath") & "show-calendar.gif""  border=0></a>" & vbcrlf
	fnDisplayDatePicker = strTemp
end function

'**********************************************************
'Function name : fnDisplayNewDatePicker
'Purpose       : This will include the date picker with date field
'Input         : Field Name, relative directory structure
'Output        : 
'Author        : TataInfotech
'Date          : 4-11-2001
'**********************************************************

function fnDisplayNewDatePicker(strFieldName, intDirectoryLevel)
	Dim  strTemp
	strTemp = "<a href=""javascript:show_calendar('" & strFieldName & "');"" " & vbcrlf
	strTemp = strTemp & " onmouseover=""window.status='Date Picker';return true;"" "  & vbcrlf
	strTemp = strTemp & " onmouseout=""window.status='';return true;"">"
	Select Case  intDirectoryLevel
		Case 1
			strTemp = strTemp & " <img src=""../images/show-calendar.gif"" border=0></a>" & vbcrlf
		Case 2
			strTemp = strTemp & " <img src=""../../images/show-calendar.gif"" border=0></a>" & vbcrlf
		Case 3
			strTemp = strTemp & " <img src=""../../../images/show-calendar.gif""  border=0></a>" & vbcrlf
		Case 4
			strTemp = strTemp & " <img src=""../../../../images/show-calendar.gif""  border=0></a>" & vbcrlf
		Case 5
			strTemp = strTemp & " <img src=""../../../../../images/show-calendar.gif""  border=0></a>" & vbcrlf		
	End Select
	'strTemp = strTemp & " <img src=""" & Application("strImagesPath") & "show-calendar.gif"" width=24 height=22 border=0></a>" & vbcrlf
	fnDisplayNewDatePicker = strTemp
end function



'**********************************************************
'Function name : fnGetStatus
'Purpose       : This checks whether the status of the entity is active or Inactive
'                
'Input         : 
'Output        : It returns a Active or Inactive in place of 'A' or 'I' from database
'Author        : Tata Infotech
'Date          : 7-10-1999
'Remark(Yogesh): If Status is 'S' then return 'Seperated'. Changed on 10/12/2001.
'				 Request from Amit Jain.
'**********************************************************
function fnGetStatus(strStatus)
     if strStatus = "A" then
	        fnGetStatus = "Active"
    else
	        fnGetStatus = "In-Active"
     end if
	if strStatus = "S" then
	        fnGetStatus = "Seperated"
	end if
end function

'**********************************************************
'Function name : fnDecodeStatus
'Purpose       : This function returns the Description of the Status Field
'Input         : The function should be called as 
'				 fnDecodeStatus("<Param1>")
'				 Param1 = "A" or "I"
'Output        : The Function returns the expanded form of the Status ("Active" or "In-Active")
'Author        : Yogesh Joshi
'Date          : 22-11-2001
'Reference	   : For implementation, please see Country Master (cudViewCtryMst.asp)
'Remark(Yogesh): If Status is 'S' then return 'Seperated'. Changed on 10/12/2001.
'				 Request from Amit Jain
'**********************************************************
function fnDecodeStatus(strStatus)
	if strStatus = "A" then
		fnDecodeStatus = "Active"
	else
		fnDecodeStatus = "In-Active"
	end if
	if strStatus = "S" then
	    fnDecodeStatus = "Seperated"
	end if
end function

'**********************************************************
'Function name : fnGetEmpDtls
'Purpose       : This will get the logical location code, logical location name, employee name, employee designation
'Input         : employee no
'Output        : 
'Author        : Geeta Negi
'Date          : 21-11-2001
'**********************************************************
Function fnGetEmpDtls(intEmpNo)
		
		Set conCommon = Server.CreateObject ("ADODB.Connection")
		Set rsCommon = Server.CreateObject ("ADODB.Recordset")
		conCommon.Open astrConn
		strCommon = "Select a.emp_alias, a.desig_cd, a.loc_cd, c.description from HRD_EMP_MST a, DOP_USER_MST b, " &_
					" LOGICAL_LOC_MST c where a.emp_no = b.emp_no  and a.loc_cd = c.logicalloc_cd and " &_
					" b.emp_no = '" & intEmpNo & "'"
		'Response.Write strCommon
		 
		rsCommon.Open strCommon, conCommon
		
		strEmpName = rsCommon.Fields(0)
		strEmpDesig = rsCommon.Fields(1)
		strBranchCd = rsCommon.Fields(2)
		strBranchName = rsCommon.Fields(3)
		
		rsCommon.Close
		conCommon.Close 
		
		set rsCommon = nothing
		set conCommon = nothing

	End Function


'**********************************************************
'Function name : fnGetLOV
'Purpose       : This function returns the values from the LOV_MST table.
'Input         : The function should be called as 
'				 fnGetLOV("<Param1>", "<Param2>", "<Param3>", "<Param4>")
'				 Param1 = Lov_Cd, Param2 = Lov_Sub_Cd, 
'				 Param3 = Lov_value, Param4 = Lov_Value
'				 Param3 should be the Lov_Value to be Selected in the Dropdown
'				 Param4 are the comma-seperated Lov_Value/s to be Excluded from the Dropdown
'				 If Param3 and Param4 are not required, please pass NULL or SPACES.
'Output        : The Function returns the values to be populated in the Dropdown
'				 in between the <Option> tags.
'Author        : Yogesh Joshi
'Date          : 19-11-2001
'Reference	   : For implementation, please see Country Master (cudAddCtryExp.asp)
'Remark(Amit C): Coded for Multiple Exclusion on 17-Dec-2001.
'**********************************************************
function fnGetLOV(strLovCd, strLovSubCd, strSelected, strExclude)


	Dim objGetLOVConn, rs
	Dim SQL, strReturn	
		
	if strLovSubCd = "" or strLovSubCd = "" then
		'return null if incorrect parameters
		strReturn = "<OPTION></OPTION>"
	else
	'else for if strLovSubCd = "" or strLovSubCd = "" then
	    set objGetLOVConn = Server.CreateObject( "ADODB.Connection" )
		set rs = Server.CreateObject( "ADODB.Recordset" )
		objGetLOVConn.Open astrConn
		if Err.number <> 0 then
			 'Check if connection is opened properly
			 Response.Redirect "Common/Error.asp?aintCode=4&aintPage=-1&astrErrDescription=" & Err.description
		     Response.End  
		end if	 
		SQL = "Select lov_value, lov_desc from LOV_MST where lov_cd = '" & strLovCd & "' and lov_sub_cd = '" & strLovSubCd & "' "
		rs.Open SQL,objGetLOVConn
		
		if (rs.EOF and rs.BOF) then
			'return null if no records found
			strReturn = "<OPTION></OPTION>"
		else
		'else for if (rs.EOF and rs.BOF) then
			strReturn = ""
			while not rs.EOF
			'return Option code if records are present	
			
			
				if rs("lov_value") = strSelected then
					'Select the record in dropdown
					strReturn = strReturn & "<OPTION VALUE='"& rs("lov_value")& "' selected>" &rs("lov_desc")& "</OPTION>"
				else
					if strExclude <> "" and instr(strExclude,rs("lov_value")) > 0 then
						'Skip the record if value needs to be excluded else Include as simple option in dropdown
					else
						strReturn = strReturn & "<OPTION VALUE='"& rs("lov_value")& "'>" &rs("lov_desc")& "</OPTION>"						
					'end if for if rs("lov_value") <> strExclude and instr(strExclude,rs("lov_value")) > 0 then
					end if
				'end if for if rs("lov_value") = strSelected then
				end if
				'Goto next record
				rs.MoveNext
			wend

			rs.Close
			set rs = nothing
		end if
		'end if for if (rs.EOF and rs.BOF) then
	end if
	'end if for if strLovSubCd = "" or strLovSubCd = "" then
	
	fnGetLOV = strReturn
end function

'**********************************************************
'Function name : fnDecodeLOV
'Purpose       : This function returns the Description from the LOV_MST table.
'Input         : The function should be called as 
'				 fnDecodeLOV("<Param1>", "<Param2>", "<Param3>")
'				 Param1 = Lov_Cd, Param2 = Lov_Sub_Cd, 
'				 Param3 = Lov_value
'				 Param3 should be the Lov_value to be Decoded
'Output        : The Function returns the Decoded Lov values' description
'Author        : Yogesh Joshi
'Date          : 22-11-2001
'Reference	   : For implementation, please see Country Master (cudViewCtryMst.asp)
'**********************************************************
function fnDecodeLOV(strLovCd, strLovSubCd, strLovValue)
	Dim objDecodeLOVConn, rs
	Dim SQL
	if strLovSubCd = "" or strLovSubCd = "" then
		'return null if incorrect parameters
		fnDecodeLOV = ""
	else		'else for if strLovSubCd = "" or strLovSubCd = "" then
	    set objDecodeLOVConn = Server.CreateObject( "ADODB.Connection" )
		set rs = Server.CreateObject( "ADODB.Recordset" )
		objDecodeLOVConn.Open astrConn
		if Err.number <> 0 then
			 'Check if connection is opened properly
			 Response.Redirect "Common/Error.asp?aintCode=4&aintPage=-1&astrErrDescription=" & Err.description
		end if	 

		SQL = "Select lov_desc from LOV_MST where lov_cd = '" & strLovCd & "' and lov_sub_cd = '" & strLovSubCd & "' and lov_value = '" & strLovValue & "'"
		rs.Open SQL,objDecodeLOVConn
		if (rs.EOF and rs.BOF) then
			fnDecodeLOV = ""	'return null if no records found
		else	'else for if (rs.EOF and rs.BOF) then
			fnDecodeLOV = rs("lov_desc")
		end if	'end if for if (rs.EOF and rs.BOF) then

		rs.Close
		set rs = nothing
		objDecodeLOVConn.Close
		set objDecodeLOVConn = nothing

	end if		'end if for if strLovSubCd = "" or strLovSubCd = "" then
	
end function


'**********************************************************
'Function name : FnDOPCheck
'Purpose       : This function returns the values from the DOP_FIN_ROLE_PARAM_VALUE table.
'Input         : The function should be called as 
'				 FnDOPCheck(Param_Cd)
'				 Param_Cd = Parameter Code from Parameter Master 
'Output        : The Function returns the value to be used in the java validate function
'Author        : Manish Nagar
'Date          : 23-11-2001
'Reference	   : 
'**********************************************************




function FnDOPCheck(Param_Cd)
		Set conCommon = Server.CreateObject ("ADODB.Connection")
		Set rsCommon = Server.CreateObject ("ADODB.Recordset")
		
		
		strUserID = Session("sUserID")
		strCommon = "Select max(value) from DOP_FIN_ROLE_PARAM_VALUE where upper(parameter_cd) = upper('" & trim(Param_Cd) & "') and fin_role_cd in (select fin_role_cd from dop_user_fin_role where user_id = '" & strUserID & "' and status = 'A')"
		
		dim strDOPCheck
		
		conCommon.Open astrConn
		
		rsCommon.Open strCommon, conCommon
		
		
		'Response.Write rsCommon.Fields(0)
		'Response.End 
		if not rsCommon.EOF then
		strDOPCheck = rsCommon.Fields(0)
		FnDOPCheck = strDOPCheck
 		else
 		FnDOPCheck = ""
 		end if		
 		
 		
		rsCommon.Close
		conCommon.Close 
		
		set rsCommon = nothing
		set conCommon = nothing
		
		
		
 	

End Function

'**********************************************************
'Function name : FnDOPCheckDec
'Purpose       : This function returns the values from the DOP_FIN_ROLE_PARAM_VALUE table.
'Input         : The function should be called as 
'				 FnDOPCheck(Param_Cd)
'				 Param_Cd = Parameter Code from Parameter Master 
'Output        : The Function returns the value to be used in the java validate function
'Author        : Manish Nagar
'Date          : 23-11-2001
'Reference	   : 
'**********************************************************




function FnDOPCheckDec(Param_Cd)
		Set conCommon = Server.CreateObject ("ADODB.Connection")
		Set rsCommon = Server.CreateObject ("ADODB.Recordset")
		strUserID = Session("sUserID")
		strCommon = "Select nvl(max(dec_value),0) from DOP_FIN_ROLE_PARAM_VALUE where upper(parameter_cd) = upper('" & trim(Param_Cd) & "') and fin_role_cd in (select fin_role_cd from dop_user_fin_role where user_id = '" & strUserID & "' and status = 'A')"
		
		conCommon.Open astrConn
		rsCommon.Open strCommon, conCommon
		'Response.Write rsCommon.Fields(0)
		'Response.End 
		if not rsCommon.EOF then
		strDOPCheck = rsCommon.Fields(0)
		FnDOPCheckDec = strDOPCheck
		else
		FnDOPCheckDec=0
 		end if		
		rsCommon.Close
		conCommon.Close 
		
		set rsCommon = nothing
		set conCommon = nothing

End Function


'**********************************************************
'Function name : EncryptString
'Purpose       : This function returns the Encrypted Value and Decrypted value for a string.
'Input         : The function should be called as 
'				 EncryptString (sRaw) 
'				 sRaw = Any String
'Output        : The Function returns the Encrypted or Decrypted value.
'Author        : Manish Nagar
'Date          : 28-11-2001
'Reference	   : 
'**********************************************************


Function EncryptString (sRaw) 

   Const SECRET = "20fjsdlfanrtg[34trjsngf[tp9ertyh" 'any random characters
    Dim sTmp 
    Dim iCount 
    sTmp = ""


    For iCount = 1 To len(sRaw)
        if asc(mid(sRaw, iCount, 1)) = asc(mid(SECRET, iCount, 1)) then
        sTmp = sTmp & mid(sRaw, iCount, 1)
        else
        sTmp = sTmp & chr(asc(mid(sRaw, iCount, 1)) XOr asc(mid(SECRET, iCount, 1)))
		end if
    Next 
	
    EncryptString = sTmp
Exit Function

End Function




'**********************************************************
'Function name : Encrypt_Password
'Purpose       : This function returns the Encrypted Value for a string.
'Input         : The function should be called as 
'				 Encrypt_Password(password)
'				 password = Password String
'Output        : The Function returns the Encrypted value.
'Author        : Manish Nagar
'Date          : 28-11-2001
'Reference	   : 
'**********************************************************





'Public pass(255)
'Public alt(255) 
'Public txt 


Public Function Encrypt_Password(password)

    txt = ""


    For n = 1 To Len(Trim(password))
        pass(n) = Asc(Mid(Trim(password), n, 1))


        If pass(n) = 32 Then
            alt(n) = Chr(pass(n))
        Else
            alt(n) = Chr(pass(n) - n)
        End If

        txt = txt + alt(n)
    Next 

    '
    Encrypt_Password = txt
End Function



'**********************************************************
'Function name : deEncrypt_Password
'Purpose       : This function returns the Decrypted Value for a for function Encrypt_Password(password) string.
'Input         : The function should be called as 
'				 deEncrypt_Password(password)
'				 password = Password String
'Output        : The Function returns the Dencrypted value.
'Author        : Manish Nagar
'Date          : 28-11-2001
'Reference	   : 
'**********************************************************


Public Function deEncrypt_Password(password)

    txt = ""


    For n = 1 To Len(Trim(password))
        pass(n) = Asc(Mid(Trim(password), n, 1))


        If pass(n) = 32 Then
            alt(n) = Chr(pass(n))
        Else
            alt(n) = Chr(pass(n) + n )
        End If

        txt = txt + alt(n)
    Next 

    '
    deEncrypt_Password = txt
End Function


'**********************************************************
'Function name : fnGetGLDetails
'Purpose       : This will get General Ledger Details
'Input         : GL Code fields
'Output        : 
'Author        : Amit C
'Date          : 03-12-2001
'**********************************************************
Function fnGetGLDtls(astrMainGlCode, astrSubGlCd1, astrSubGlCd2, astrSubGlCd3, astrSubGlCd4)

	fnGetGLDtls = ""
	
	Dim conCommon
	Dim rsCommon
	Dim strSQL
	
	Set conCommon = Server.CreateObject ("ADODB.Connection")
	Set rsCommon = Server.CreateObject ("ADODB.Recordset")
	conCommon.Open astrConn
	if Err.number <> 0 then
		'Response.Write "<FONT class=clsError>fnGetGLDtls: Connection Error</FONT><BR>"
		'Response.Write "Error No. " & Err.number & " "
		'Response.Write "Error Desc. " & Err.Description & " "
		'Response.End 
		Response.Redirect "../../../Common/Error.asp?aintCode=4&aintPage=-1&astrErrNumber="&Err.Number&"&astrErrDescription="&Err.description
	end if	
	
	strSQL = "Select active, gl_is_group, gl_type, personal_ledger_level, bal_ind, zero_bal_flg "
	strSQL = strSQL & "from PF_ENTITY_GL_MST "
	strSQL = strSQL & "where entity_cd='PF' "
	strSQL = strSQL & "and maingl_cd='" & astrMainGlCode & "' "
	If astrSubGLCd1 <> "" Then
		strSQL = strSQL & "and subgl_cd1='" & astrSubGLCd1 & "' "
	End If
	If astrSubGLCd2 <> "" Then
		strSQL = strSQL & "and subgl_cd2='" & astrSubGLCd2 & "' "
	End If
	If astrSubGLCd3 <> "" Then
		strSQL = strSQL & "and subgl_cd3='" & astrSubGLCd3 & "' "
	End If
	If astrSubGLCd4 <> "" Then
		strSQL = strSQL & "and subgl_cd4='" & astrSubGLCd4 & "' "		
	End If
	
	
	rsCommon.Open strSQL, conCommon
	
	if Err.number <> 0 then
		'Response.Write "<FONT class=clsError>fnGetGLDtls: Connection Error</FONT><BR>"
		'Response.Write "Error No. " & Err.number & " "
		'Response.Write "Error Desc. " & Err.Description & " "
		'Response.End 
		Response.Redirect "../../../Common/Error.asp?aintCode=4&aintPage=-1&astrErrNumber="&Err.Number&"&astrErrDescription="&Err.description
	end if	
	
	if not rsCommon.EOF then
		fnGetGLDtls = rsCommon.GetRows()
	else
		fnGetGLDtls = ""
	end if
	
	rsCommon.Close
	conCommon.Close 
		
	set rsCommon = nothing
	set conCommon = nothing

End Function



'**********************************************************
'Function name : fnCheckCode
'Purpose       : This will check if record for the code 
'				 entered exists in the table or not 
'Input         : strColumnName, strEnteredValue, strTableName
'Output        : will return "0" if records not found, else "1"
'Author        : Yogesh Joshi
'Date          : 03-Jan-2002
'**********************************************************
	function fnCheckCode(strColumnName, strEnteredValue, strTableName)
		set objConnCheckCode = Server.CreateObject( "ADODB.Connection" )
		set objRsCheckCode = Server.CreateObject( "ADODB.Recordset" )
		objConnCheckCode.Open astrConn	
		if Err.number <> 0 then
			 Response.Redirect "../../../Common/Error.asp?aintCode=4&aintPage=-1&astrErrDescription=" & Err.description
		end if
		
		CheckCodeSQL = "select * from " & strTableName & " where UPPER(" & strColumnName & ") = UPPER('" & strEnteredValue & "')"

	   	objRsCheckCode.Open CheckCodeSQL,objConnCheckCode
		if objRsCheckCode.EOF then
			objRsCheckCode.Close
			objConnCheckCode.close
			set objRsCheckCode = nothing
			set objConnCheckCode = nothing
			fnCheckCode = "0"
		else
			objRsCheckCode.Close
			objConnCheckCode.close
			set objRsCheckCode = nothing
			set objConnCheckCode = nothing
			fnCheckCode = "1"
		end if
	end function

'**********************************************************
'Function name : fnstrConvertDt
'Purpose       : Convert date from "DD/MM/YYYY" format
'				 to "MM/DD/YYYY" format
'Input         : Date ("DD/MM/YYYY" format)
'Output        : Date ("MM/DD/YYYY" format)
'Author        : Milind Khedaskar
'Date          : 10-11-2001
'**********************************************************
function fnstrConvertDt(astrDt)

	fnstrConvertDt = mid(astrDt,4,2) & "/" & left(astrDt,2) & "/" & right(astrDt,4)

end function



'**********************************************************
'Function name : fnGetNameFromCode
'Purpose       : This will retreive the name from the code 
'Input         : strColumnName, strEnteredValue, strTableName
'Output        : will return "0" if records not found, else will return the name
'Author        : Yogesh Joshi
'Date          : 14-Jan-2002
'**********************************************************
function fnGetNameFromCode(strTableName, strCodeField, strNameField, strEnteredCode)
	set objConnNameFromCode = Server.CreateObject( "ADODB.Connection" )
	set objRsNameFromCode = Server.CreateObject( "ADODB.Recordset" )
	objConnNameFromCode.Open astrConn	
	if Err.number <> 0 then
		 Response.Redirect "../../../Common/Error.asp?aintCode=4&aintPage=-1&astrErrDescription=" & Err.description
	end if
		
	strNameFromCodeSQL = "select " & strNameField & " from " & strTableName & " where UPPER(" & strCodeField & ") = UPPER('" & strEnteredCode & "')"

   	objRsNameFromCode.Open strNameFromCodeSQL,objConnNameFromCode
	if objRsNameFromCode.EOF then
		objRsNameFromCode.Close
		objConnNameFromCode.close
		set objRsNameFromCode = nothing
		set objConnNameFromCode = nothing
		fnGetNameFromCode = "0"
	else
		strReturnName = objRsNameFromCode(strNameField)
		objRsNameFromCode.Close
		objConnNameFromCode.close
		set objRsNameFromCode = nothing
		set objConnNameFromCode = nothing
		fnGetNameFromCode = strReturnName
	end if
end function


%>
