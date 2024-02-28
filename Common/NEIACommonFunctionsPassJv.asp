<%' on error resume next %>
<% 'Response.Buffer = false %>
<%
'**********************************************************
' This File Contains all common Accounts Functions
'**********************************************************

' Module Level Variables
Dim mstrMonth, mstrFiscalYear
Dim mconEcgcDb 
Dim mstrSQL

'**********************************************************
'Function name : fngetCurrFiscalYearStartDt
'Purpose       : Retrieve the Current Fiscal Year Start Date
'Input         : 
'Output        : return Current Fiscal Year Start Date
'				 also store it in Session Variable
'				 sstrCurrFiscalYearStartDt
'Author        : Milind Khedaskar
'Date          : 28-11-2001
'**********************************************************
Function fngetCurrFiscalYearStartDt()
	
	on error resume next
	
	Dim lconEcgcDb 
	Dim lstrSQL
	
	fngetCurrFiscalYearStartDt = ""
	
	Dim lrsFiscalYear
	
	Set lconEcgcDb = Server.CreateObject ("ADODB.Connection")
	Set lrsFiscalYear = Server.CreateObject ("ADODB.Recordset")

	lrsFiscalYear.LockType = 2 ' adLockPessimistic

	lstrSQL = "select to_char(curr_fisc_yr_start_dt,'DD/MM/YYYY') "
	lstrSQL = lstrSQL & "from NEIA_FISCAL_YR "
	lstrSQL = lstrSQL & "where entity_cd = 'NEIA' "
	'Response.Write lstrSQL
		
	lconEcgcDb.Open astrConn
	
	if Err.number <> 0 then
		'Response.Write "<FONT class=clsError>fngetCurrFiscalYearStartDt: Connection Error</FONT><BR>"
		'Response.Write "Error No. " & Err.number & " "
		'Response.Write "Error Desc. " & Err.Description & " "
		'Response.End 
		Response.Redirect "../../../Common/Error.asp?aintCode=4&aintPage=-1&astrErrNumber="&Err.Number&"&astrErrDescription="&Err.description
	end if	
		
	lrsFiscalYear.Open lstrSQL, lconEcgcDb
	if Err.number <> 0 then
		'Response.Write "<FONT class=clsError>fngetCurrFiscalYearStartDt: Record not found</FONT><BR>"
		'Response.Write "Error No. " & Err.number & " "
		'Response.Write "Error Desc. " & Err.Description & " "
		'Response.End 
		Response.Redirect "../../../Common/Error.asp?aintCode=4&aintPage=-1&astrErrNumber="&Err.Number&"&astrErrDescription="&Err.description
	end if	
				
	if not lrsFiscalYear.EOF then
		Session("sstrCurrFiscalYearStartDt") = lrsFiscalYear.Fields(0)
		fngetCurrFiscalYearStartDt = Session("sstrCurrFiscalYearStartDt")
	else
		Response.Redirect "../../../Common/Error.asp?aintCode=4502&aintPage=-1&astrErrNumber="&Err.Number&"&astrErrDescription="&Err.description
	end if	
	
	lrsFiscalYear.Close
	lconEcgcDb.Close 
	
	set lrsFiscalYear = nothing
	set lconEcgcDb = nothing
	
End Function

'**********************************************************
'Function name : fnintIsCalendarOpen
'Purpose       : Check whether calendar month is open or not
'Input         : Logical Location Code
'				 Transaction Type
'				 Transaction Date				
'Output        : True is the Calendar Month is open else return false
'Author        : Milind Khedaskar
'Date          : 10-11-2001
'Mod	ification History
'Date		   : 11-02-2002. Changes made for defect # 153.
'                Made modifications for returning integer values 
'                (instead of boolean) and not do any redirection
'                on error. Changed name of function (from 
'                fnintIsCalendarOpen to fnintIsCalendarOpen). Also 
'                removed old and commented code.
'**********************************************************
Function fnintIsCalendarOpen(astrLogicalLocCd, astrTxnType, astrTxnDt )
	
	on error resume next
	
	Dim lrsCalendar
		
	Dim ldtmTmpDt
	Dim lstrTxnMonth,lstrTxnYear

	Set mconEcgcDb = Server.CreateObject ("ADODB.Connection")
	Set lrsCalendar = Server.CreateObject ("ADODB.Recordset")

	' Get Month and Fiscal Year
	ldtmTmpDt = cdate(fnstrConvertDt(astrTxnDt))
	lstrTxnMonth = monthname(month(ldtmTmpDt),true)
	lstrTxnYear = datepart("yyyy" ,ldtmTmpDt)
	
	mstrSQL = "select * "
	mstrSQL = mstrSQL & "from NEIA_CALENDAR_MST "
	mstrSQL = mstrSQL & "where logicalloc_cd = '" & astrLogicalLocCd & "' "
	mstrSQL = mstrSQL & "and month = '" & lstrTxnMonth & "' "
	mstrSQL = mstrSQL & "and fiscal_yr = '" & lstrTxnYear & "' "
	mstrSQL = mstrSQL & "and gl_txn_type = '" & astrTxnType & "' "
	mstrSQL = mstrSQL & "and closed = 'N' "
	mstrSQL = mstrSQL & "and entity_cd='NEIA' "
	'Response.Write mstrSQL
	'Response.end
		
	mconEcgcDb.Open astrConn
	
	if Err.number <> 0 then
		'Response.Write "<FONT class=clsError>fnintIsCalendarOpen: Connection Error</FONT><BR>"
		'Response.Write "Error No. " & Err.number & " " & "Error Desc. " & Err.Description & " "
		'Response.End 
		'Response.Redirect "../../../Common/Error.asp?aintCode=4&aintPage=-1&astrErrNumber="&Err.Number&"&astrErrDescription="&Err.description
		fnintIsCalendarOpen = -4
		exit function		
	end if	
		
	lrsCalendar.Open mstrSQL, mconEcgcDb

	if Err.number <> 0 then
		'Response.Write "<FONT class=clsError>fnintIsCalendarOpen:Connection error.</FONT><BR>"
		'Response.Write "Error No. " & Err.number & " " & "Error Desc. " & Err.Description & " "
		'Response.End 
		'Response.Redirect "../../../Common/Error.asp?aintCode=4&aintPage=-1&astrErrNumber="&Err.Number&"&astrErrDescription="&Err.description
		fnintIsCalendarOpen = -4
		exit function		
	end if	
				
	if lrsCalendar.EOF then
		fnintIsCalendarOpen = 1
	else
		fnintIsCalendarOpen = 0
	end if	
		
	lrsCalendar.Close
	mconEcgcDb.Close 
		
	set lrsCalendar = nothing
	set mconEcgcDb = nothing

End Function

'**********************************************************
'Function name : fnGetMonthAndFiscalYear
'Purpose       : Get Month and Fiscal Year
'Input         : Transaction Date (DD/MM/YYYY format)
'Output        : Month in mstrMonth 
'				 Fiscal Year in mstrFiscalYear
'Author        : Milind Khedaskar
'Date          : 10-11-2001
'**********************************************************
sub fnGetMonthAndFiscalYear(astrTxnDt)

	dim ldtmTmpDt
	ldtmTmpDt = cdate(fnstrConvertDt(astrTxnDt))
	
	' Call this function
	fnGetMonthAndFiscalYearByDate ldtmTmpDt
	 
end sub

'**********************************************************
'Function name : fnGetMonthAndFiscalYearByDate
'Purpose       : Get Month and Fiscal Year
'Input         : Transaction Date (Date format)
'Output        : Month in mstrMonth 
'				 Fiscal Year in mstrFiscalYear
'				 Returns number of days between Transaction Date and Current Fiscal Year
'Author        : Milind Khedaskar
'Date          : 12-11-2001
'Modification History 
'Date          : 05-01-2002 - For Defect # 76 (Amit C)
'**********************************************************
Function fnGetMonthAndFiscalYearByDate(adtmTxnDt)

	dim lstrTxnMonth,lstrTxnYear
	dim ldtmCurrFiscalYearStartDt
	
	'Changes for defect # 76 made by Amit C on 05-Jan-02.
	'Added new variable.
	dim lintYears
	dim lintDays
	
	fnGetMonthAndFiscalYearByDate = 0
	
	'Response.Write "Fiscal Year : " & Session("sstrCurrFiscalYearStartDt") & "<BR>"
	
	if Session("sstrCurrFiscalYearStartDt") = "" then
		'Current Fiscal Year start date not found
		'Get the current fiscal year start date
		fngetCurrFiscalYearStartDt
	end if

	lstrTxnMonth = monthname(month(adtmTxnDt),true)
	'Transaction year is changed from adtmTxnDt to Fiscal year start date.
	lstrTxnYear = datepart("yyyy" ,Session("sstrCurrFiscalYearStartDt"))
	
	ldtmCurrFiscalYearStartDt = cdate(fnstrConvertDt(Session("sstrCurrFiscalYearStartDt")))

	'Changes for defect # 76 made by Amit C on 05-Jan-02.
	'Determine year difference between current fiscal year and transaction date.
	lintYears = 1	
	if adtmTxnDt < ldtmCurrFiscalYearStartDt then
		do while dateadd("yyyy",lintYears,adtmTxnDt) < ldtmCurrFiscalYearStartDt
			lintYears = lintYears + 1
		loop
		lstrTxnYear = (cint(lstrTxnYear) - lintYears) & "-" & (cint(lstrTxnYear) - lintYears + 1)
	else
		do while dateadd("yyyy",lintYears,ldtmCurrFiscalYearStartDt) <= adtmTxnDt
			lintYears = lintYears + 1
		loop	
		lstrTxnYear = (cint(lstrTxnYear) + lintYears - 1) & "-" & (cint(lstrTxnYear) + lintYears)
	end if
	
	mstrFiscalYear = lstrTxnYear
	mstrMonth = lstrTxnMonth
	
	'Returns number of days between Current Fiscal Year Start Date and the Transaction Date
	'Moved from above
	lintDays = datediff("d",ldtmCurrFiscalYearStartDt,adtmTxnDt)
	fnGetMonthAndFiscalYearByDate = lintDays
	
End Function


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


'*************************************************************************
'Function name : fnintPassGLTxn
'Purpose       : To Pass GL Transactions
'Input         : Connection Object
'                Transaction Header Array
'				 Transaction Details Array
'Output        : Returns Transaction Number if the transaction 
'				 is successfully passed
'				 else returns error number (-ve value)
'Author        : Milind Khedaskar
'Date          : 10-11-2001
'
'Modification History
'Date          : 11-02-2002 Changes made for defect # 153
'                Logic was modified to store only transaction totals
'                in NEIA_GL_BR_BAL and NEIA_GL_BAL_SUMM tables. Year
'                opening balances would be stored in NEIA_GL_BR_BAL_YR
'                and NEIA_GL_BAL_SUMM_YR tables. Old commented code was
'                also removed.
'Date		   : 08-05-2002 Changes in the sequence no genration of gl txn no (fiscal year)
'Date		   : 03-08-2002 LTC Validations added
'*************************************************************************
function fnintPassGLTxn(aconEcgcDb, aarrTxnHdr,aarrTxnDtl)
	
	on error resume next
	
	' General Variables
	Dim lstrEntityCd
	Dim lstrLogicalLocCd
	
	' Recordset Variables
	Dim lrsFiscalYr
	Dim lrsGLTxn
	Dim lrsEntityGLMst
	Dim lrsGLBrBal
	Dim lrsGLBalSumm
	Dim lrsGLErr

	' GL Codes
	Dim lintMainGLCd
	Dim lintSubGLCd1
	Dim lintSubGLCd2
	Dim lintSubGLCd3
	Dim lintSubGLCd4
	
	' GL Master Variables
	Dim lstrBalInd
	Dim lstrZeroBalFlg
	Dim lstrGLType
	Dim lstrPersonalLedgerLevel
	
	' GL Transaction Detail Variables
	Dim lstrTxnType
	Dim llngTxnNo
	Dim lstrTxnDt
	Dim lstrTxnRef

	' Processing Variables	
	Dim lintRecords
	Dim lintCalendarOpen

	Dim ldblOpenCrBal
	Dim ldblOpenDrBal

	Dim	ldblCurrCrBal
	Dim ldblCurrDrBal

	' Other Information	
	Dim lstrPersonalLedgerCd
	Dim ldblTxnAmt
	Dim ldblTotTxnAmt
	Dim lintDays
	Dim lblnContinueProcessing

	' Branch Balance Variables
	Dim larrGLBrBal
	Dim ldblGLBrBal
	
	' Balance Summary Variables
	Dim ldblGLSummBal
	
	' Error Information Variables
	Dim lintErrNo
	Dim ldblErrBalAmt
	
	' Error Flags
	Dim lblnErr
	Dim lblnOldErr
	
	' Date Variables
	Dim ldtmTxnDt
	Dim lstrFiscalYearForTxnDt
	Dim lstrMonthForTxnDt
	
	Dim ldtmTodaysDt
	Dim lstrFiscalYearForTodaysDt
	
	Dim lstrFiscalYearForProcessing
	Dim lstrMonthForProcessing
	
	Dim lstrFinalBalanceDt
	Dim ldtmFinalBalanceDt
		
	' Default return value
	fnintPassGLTxn = 0
	ldblTotTxnAmt = 0
	
	' Set Variables which are required frequently
	lstrEntityCd = aarrTxnHdr(0)
	lstrLogicalLocCd = aarrTxnHdr(1)
	lstrTxnType = aarrTxnHdr(2)
	lstrTxnDt = aarrTxnHdr(3)
	lstrTxnRef = aarrTxnHdr(4)
	lstrUserID = aarrTxnHdr(5)
	'Response.Write "here"
	'Response.End 
	'Determine is calendar month is open or closed for the input transaction type	
	lintCalendarOpen = fnintIsCalendarOpen(lstrLogicalLocCd,lstrTxnType,lstrTxnDt)
	
	if lintCalendarOpen <> 0 then
		'response.write lintCalendarOpen	
		'response.end
		if lintCalendarOpen = 1 then 
			'Calendar Month is not open
			'Response.Write "<FONT class=clsError>Calendar Month is Closed, Hence can not pass Transaction</FONT><BR>"
			'Response.End 
			'Response.Redirect "../../../Common/Error.asp?aintCode=4501&aintPage=-1&astrErrNumber="&Err.Number&"&astrErrDescription="&Err.description
			fnintPassGLTxn = -4501
		else
			'Some other error occured
			fnintPassGLTxn = lintCalendarOpen
		end if
		exit function
	end if

	'Check if transaction date is older than a year or if it is greater than current date.
	'If yes, give an error.
	ldtmTxnDt = cdate(fnstrConvertDt(lstrTxnDt))
	ldtmTodaysDt = date
			
	lintDays = fnGetMonthAndFiscalYearByDate(ldtmTxnDt)
	lstrFiscalYearForTxnDt = mstrFiscalYear
	lstrMonthForTxnDt = mstrMonth
	
	If lintDays < -366 then
		'Transaction is too old 
		'Response.Write "<FONT class=clsError>fnintPassGLTxn: Transaction is older than 365 days</FONT><BR>"
		'Response.End
		'Response.Redirect "../../../Common/Error.asp?aintCode=4512&aintPage=-1&astrErrNumber="&Err.Number&"&astrErrDescription="&Err.description
		fnintPassGLTxn = -4512
		exit function
	End if
			
	If ldtmTxnDt > ldtmTodaysDt then
		'Response.Write "<FONT class=clsError>fnintPassGLTxn: Can not pass Post-dated Transactions</FONT><BR>"
		'Response.End
		'Response.Redirect "../../../Common/Error.asp?aintCode=4513&aintPage=-1&astrErrNumber="&Err.Number&"&astrErrDescription="&Err.description
		fnintPassGLTxn = -4513
		exit function
	End if
	
	'If the fiscal year in which the transaction has been passed is already closed, then
	'give an error
	mstrSQL = "Select last_yr_closed "
	mstrSQL = mstrSQL & "from NEIA_FISCAL_YR "
	mstrSQL = mstrSQL & "where entity_cd='NEIA' "
	'Response.Write mstrSQL & "<BR>"
	'Response.End 

	set lrsFiscalYr = Server.CreateObject("ADODB.Recordset")

	lrsFiscalYr.Open mstrSQL, aconEcgcDb
	if Err.number <> 0 then
		'Response.Write mstrSQL
		'Response.Write "Error number : " & err.number & " Error description : " & err.description
		'Response.End 
		'Response.Redirect "../../../Common/Error.asp?aintCode=4&aintPage=-1&astrErrNumber="&Err.Number&"&astrErrDescription="&Err.description
		fnintPassGLTxn = -4
		exit function
	end if	
	
	if not lrsFiscalYr.EOF and lrsFiscalYr.Fields("last_yr_closed") = lstrFiscalYearForTxnDt then
		'Response.Write "<FONT class=clsError>fnintPassGLTxn: Fiscal Year Closed</FONT><BR>"
		'Response.End 
		'Response.Redirect "../../../Common/Error.asp?aintCode=4534&aintPage=-1&astrErrNumber="&Err.Number&"&astrErrDescription="&Err.description
		fnintPassGLTxn = -4534
		exit function
	end if	
	
	Set lrsGLTxn = Server.CreateObject ("ADODB.Recordset")

	lrsGLTxn.LockType = 2 ' adLockPessimistic

	mstrSQL = "Select nvl(max(gl_txn_no),0) "
	mstrSQL = mstrSQL & "from NEIA_GL_TXN_HDR "
	mstrSQL = mstrSQL & "where logicalloc_cd='" & lstrLogicalLocCd &"' "
	mstrSQL = mstrSQL & "and gl_txn_type='" & lstrTxnType & "' "
	mstrSQL = mstrSQL & "and entity_cd='NEIA' "
	mstrSQL = mstrSQL & "and fiscal_yr='" & lstrFiscalYearForTxnDt & "' "
	
	'Response.Write mstrSQL & "<BR>"
	'Response.End 

	lrsGLTxn.Open mstrSQL, aconEcgcDb
	if Err.number <> 0 then
		'Response.Write mstrSQL
		'Response.End 
		'Response.Redirect "../../../Common/Error.asp?aintCode=4&aintPage=-1&astrErrNumber="&Err.Number&"&astrErrDescription="&Err.description
		fnintPassGLTxn = -4
		exit function
	end if	
	
	'Running sequence Number for the particular ECGC branch
	llngTxnNo = clng(lrsGLTxn.Fields(0).value)
	
	if llngTxnNo = 0 then 
		' Initialising the gl txn no on the fiscal year
		llngTxnNo = clng(left(lstrFiscalYearForTxnDt,4)+"000001")
	else
		' Get the next gl txn no
		llngTxnNo = llngTxnNo + 1
	end if
	
	'Response.Write llngTxnNo
	'Response.End 

	lrsGLTxn.Close()
	set lrsGLTxn = nothing
	
	' GL Transaction Header
	mstrSQL ="Insert into NEIA_GL_TXN_HDR(entity_cd,logicalloc_cd,gl_txn_type, gl_txn_no, txn_dt,reference,user_id,last_trans_date,fiscal_yr) "
	mstrSQL = mstrSQL & "values ("
	mstrSQL = mstrSQL & "'" & lstrEntityCd & "',"
	mstrSQL = mstrSQL & "'" & lstrLogicalLocCd & "',"
	mstrSQL = mstrSQL & "'" & lstrTxnType & "',"
	mstrSQL = mstrSQL & llngTxnNo & ","
	mstrSQL = mstrSQL & "to_date('" & lstrTxnDt & "', 'dd/mm/yyyy') ,"
	mstrSQL = mstrSQL & "'" & lstrTxnRef & "',"
	mstrSQL = mstrSQL & "'" & lstrUserID & "',"
	mstrSQL = mstrSQL & "sysdate, "
	mstrSQL = mstrSQL & "'" & lstrFiscalYearForTxnDt & "'"
	mstrSQL = mstrSQL & ")"			

	'Response.Write mstrSQL & "<BR>"
	'Response.End 
	

	aconEcgcDb.Execute mstrSQL

	If Err.number <> 0 then
		'Response.Write mstrSQL
		'Response.End 
		'Response.Redirect "../../../Common/Error.asp?aintCode=4&aintPage=-1&astrErrNumber="&Err.Number&"&astrErrDescription="&Err.description
		fnintPassGLTxn = -4
		exit function
	end if
	'for i = 0 to ubound(aarrTxnDtl,2) -1
	'	response.write "<br>"
	'	if aarrTxnDtl(0,i) = "" then
	'	Response.Write "here"
	'	end if
	'	Response.Write aarrTxnDtl(0,i) 
	'	Response.Write aarrTxnDtl(1,i) 
	'	Response.Write aarrTxnDtl(2,i)
	'	Response.Write aarrTxnDtl(3,i) 
	'	Response.Write aarrTxnDtl(4,i) 
	'	Response.Write aarrTxnDtl(5,i) 
	'	Response.Write aarrTxnDtl(6,i) 
	'	Response.Write aarrTxnDtl(7,i) 
	'	Response.Write aarrTxnDtl(8,i) 
		
	'next
	
	'Response.End
	
	for i = 0 to ubound(aarrTxnDtl,2) -1
		
		lintMainGLCd = aarrTxnDtl(0,i)
		
		' If the GL Sub codes are NULL, make them 0	
		lintSubGLCd1 = 0
		if (aarrTxnDtl(1,i) <> "") then
			lintSubGLCd1 = aarrTxnDtl(1,i)
		end if
		lintSubGLCd2 = 0
		if (aarrTxnDtl(2,i) <> "") then
			lintSubGLCd2 = aarrTxnDtl(2,i)
		end if
		lintSubGLCd3 = 0
		if (aarrTxnDtl(3,i) <> "") then
			lintSubGLCd3 = aarrTxnDtl(3,i)
		end if
		lintSubGLCd4 = 0
		if (aarrTxnDtl(4,i) <> "") then
			lintSubGLCd4 = aarrTxnDtl(4,i)
		end if
		
		lstrPersonalLedgerCd = " "
		if aarrTxnDtl(5,i) <> "" then
			lstrPersonalLedgerCd = aarrTxnDtl(5,i)
		end if

		lstrDrCrFlg = ucase(aarrTxnDtl(6,i))
		ldblTxnAmt = cdbl(aarrTxnDtl(7,i))
	
		' Select Entity GL Master record
		Set lrsEntityGLMst = Server.CreateObject ("ADODB.Recordset")
		
		lrsEntityGLMst.LockType = 2 ' adLockPessimistic
		
		mstrSQL = "Select bal_ind, zero_bal_flg, gl_type, personal_ledger_level, active, gl_is_group "
		mstrSQL = mstrSQL & "from NEIA_ENTITY_GL_MST "
		mstrSQL = mstrSQL & "where entity_cd='NEIA' "
		mstrSQL = mstrSQL & "and maingl_cd='" & lintMainGLCd & "' "
		mstrSQL = mstrSQL & "and subgl_cd1='" & lintSubGLCd1 & "' "
		mstrSQL = mstrSQL & "and subgl_cd2='" & lintSubGLCd2 & "' "
		mstrSQL = mstrSQL & "and subgl_cd3='" & lintSubGLCd3 & "' "
		mstrSQL = mstrSQL & "and subgl_cd4='" & lintSubGLCd4 & "' "
		'Response.Write mstrSQL & "<BR>"
		'Response.End

		lrsEntityGLMst.Open mstrSQL, aconEcgcDb
		if Err.number <> 0 then
			'Response.Write mstrSQL
			'Response.End
			'Response.Redirect "../../../Common/Error.asp?aintCode=4&aintPage=-1&astrErrNumber="&Err.Number&"&astrErrDescription="&Err.description
			fnintPassGLTxn = -4
			exit function
		end if
		
		if lrsEntityGLMst.BOF or lrsEntityGLMst.EOF then
			'Record not found in Entity GL Master, can not proceed
			'Response.Write "<FONT class=clsError>fnintPassGLTxn: Entity GL Master Select - record not found</FONT><BR>"
			'Response.End
			'Response.Redirect "../../../Common/Error.asp?aintCode=4503&aintPage=-1&astrErrNumber="&Err.Number&"&astrErrDescription="&Err.description
			fnintPassGLTxn = -4503
			exit function
		end if

		if lrsEntityGLMst.Fields("active") = "N" then
			'GL Record is not active
			'Response.Redirect "../../../Common/Error.asp?aintCode=4514&aintPage=-1&astrErrNumber="&Err.Number&"&astrErrDescription="&Err.description
			fnintPassGLTxn = -4514
			exit function
		end if

		if lrsEntityGLMst.Fields("gl_is_group") = "Y" then
			'GL is a group
			'Response.Redirect "../../../Common/Error.asp?aintCode=4515&aintPage=-1&astrErrNumber="&Err.Number&"&astrErrDescription="&Err.description
			fnintPassGLTxn = -4515
			exit function
		end if
		
		lstrBalInd = ucase(lrsEntityGLMst.Fields(0))
		lstrZeroBalFlg = ucase(lrsEntityGLMst.Fields(1))
		lstrGLType = ucase(lrsEntityGLMst.Fields(2))
		lstrPersonalLedgerLevel = ucase(lrsEntityGLMst.Fields(3))
		'Response.Write lstrPersonalLedgerLevel
		'Response.End
		
		lrsEntityGLMst.Close()

		set lrsEntityGLMst = nothing

		' VALIDATE PERSONAL LEDGER CODE
		
		Dim arrTemp
		Dim arrPersonalLedgerCd
		
		'Response.Write lstrPersonalLedgerLevel & " : " & lstrPersonalLedgerCd 
		
		if lstrPersonalLedgerLevel <> "" then
			
			Select Case ucase(lstrPersonalLedgerLevel)
		
				case "EXPORTER"
					arrTemp = fnGetExporterDtls(lstrPersonalLedgerCd)
				
				case "POLICY"
					arrTemp = fnNEIAGetPolicyDtls(aconEcgcDb, lstrPersonalLedgerCd)
					
				'POLICY CHANGES
				'				case "POLICY SCR"
				'					arrTemp = fnNEIAGetPolicyDtls(aconEcgcDb, lstrPersonalLedgerCd, "SCR")
				'
				'				case "POLICY SSP"
				'					arrTemp = fnNEIAGetPolicyDtls(aconEcgcDb, lstrPersonalLedgerCd, "SSP")

				case "EMPLOYEE"
				
					arrTemp = fnGetEmployeeDtls(lstrPersonalLedgerCd)

				case "BANK BRANCH"
					
					arrPersonalLedgerCd = split(lstrPersonalLedgerCd," ")
					if ubound(arrPersonalLedgerCd) = 1 then
					' Both Bank Code and Branch Code are entered by user
						arrTemp = fnGetBankBrDtls(arrPersonalLedgerCd(0), arrPersonalLedgerCd(1))
					else
						'Response.Redirect "../../../Common/Error.asp?aintCode=2&aintPage=-1&astrErrNumber="&Err.Number&"&astrErrDescription=Bank Branch record does not exists"
						arrTemp = ""
					end if
					
				Case "BANK BRANCH EXPORTER"

					arrPersonalLedgerCd = split(lstrPersonalLedgerCd," ")
					arrTemp = ""
					if ubound(arrPersonalLedgerCd) = 2 then
						' Bank Code, Branch Code and Exporter Code are entered by user
						
						' Validate Bank and Branch Code
						arrTemp = fnGetBankBrDtls(arrPersonalLedgerCd(0), arrPersonalLedgerCd(1))
						
						if IsArray(arrTemp) then
							' Validate Exporter Code
							arrTemp = ""
							arrTemp = fnGetExporterDtls(arrPersonalLedgerCd(2))
							
						end if
						
					else
						arrTemp = ""
					end if

				Case "GUARANTEE - INPSG"
		
					arrTemp = fnNEIAGetINGDtls(aconEcgcDb, "INPSG", lstrPersonalLedgerCd)
				
				Case "GUARANTEE - INPCG"
		
					arrTemp = fnNEIAGetINGDtls(aconEcgcDb, "INPCG", lstrPersonalLedgerCd)
		
				Case "GUARANTEE - INEFG"
		
					arrTemp = fnNEIAGetINGDtls(aconEcgcDb, "INEFG", lstrPersonalLedgerCd)
				
				Case "GUARANTEE - INENEIAG"
		
					arrTemp = fnNEIAGetINGDtls(aconEcgcDb, "INENEIAG", lstrPersonalLedgerCd)
						
				Case "GUARANTEE - INEPG"
		
					arrTemp = fnNEIAGetINGDtls(aconEcgcDb, "INEPG", lstrPersonalLedgerCd)
						
				Case "GUARANTEE - WTEPG"
		
					arrTemp = fnNEIAGetINGDtls(aconEcgcDb, "WTEPG", lstrPersonalLedgerCd)
	
				Case "GUARANTEE - INTG"
		
					arrTemp = fnNEIAGetINGDtls(aconEcgcDb, "INTG", lstrPersonalLedgerCd)
				
				Case "GUARANTEE - BIPCG"
					
					arrTemp = fnNEIAGetINGDtls(aconEcgcDb, "BIPCG", lstrPersonalLedgerCd)
					
				Case "LTC POLICY"
						'	Response.Write lstrPersonalLedgerCd&"<br>"
						'	Response.End
					arrTemp = fnNEIAGetLTCPolicyDtls(aconEcgcDb, lstrPersonalLedgerCd)
				
				
				Case "LTC GUARANTEE"
		
					arrTemp = fnNEIAGetLTCGteeDtls(aconEcgcDb, lstrPersonalLedgerCd)
					
				Case "LTC OIIS POLICY"
		
					arrTemp = fnNEIAGetLTCOiisDtls(aconEcgcDb, lstrPersonalLedgerCd)
					
				Case "LTC BC POLICY"
		
					arrTemp = fnNEIAGetLTCBclocDtls(aconEcgcDb, "BC", lstrPersonalLedgerCd)
					
				Case "LTC LOC POLICY"
		
					arrTemp = fnNEIAGetLTCBclocDtls(aconEcgcDb, "LC", lstrPersonalLedgerCd)
					
				Case "CONTRACTPARTY+INVESTMENTNUMBER"

					arrTemp = fnGetContractorDtls(lstrPersonalLedgerCd,lstrLogicalLocCd)
				case else
		
			End Select
			
			if not IsArray(arrTemp) then
				'Response.Write "<FONT class=clsError>fnintPassGLTxn: Invalid Personal Ledger Code</FONT><BR>"
				'Response.End
				'Response.Redirect "../../../Common/Error.asp?aintCode=4504&aintPage=-1&astrErrNumber="&Err.Number&"&astrErrDescription="&Err.description
					
				fnintPassGLTxn = -4504 'Temporary Comment For passing JV For Contract Party
				exit function
			end if
		
		else
			' Personal ledger level is not defined, invalid personal ledger code entered
			if trim(lstrPersonalLedgerCd) <> "" then
		'	Response.Write ucase(lstrPersonalLedgerCd)
				fnintPassGLTxn = -4504 'Temporary Comment For passing JV For Contract Party
				exit function 'Temporary Comment For passing JV For Contract Party
			end if
		end if
		
		' VALIDATIONS OF PERSONAL LEDGER CODE COMPLETE
		
		lstrPersonalLedgerCd = ucase(lstrPersonalLedgerCd)
		
		mstrSQL ="Insert into NEIA_GL_TXN_DTL(entity_cd,logicalloc_cd,gl_txn_type, gl_txn_no, sr_no, maingl_cd,subgl_cd1,subgl_cd2,subgl_cd3,subgl_cd4,personal_ledger_cd,dr_cr_flg,txn_amt,txn_rmk,user_id,last_trans_date) "
		mstrSQL = mstrSQL & "values ("
		mstrSQL = mstrSQL & "'" & lstrEntityCd & "',"
		mstrSQL = mstrSQL & "'" & lstrLogicalLocCd & "',"
		mstrSQL = mstrSQL & "'" & lstrTxnType & "',"
		mstrSQL = mstrSQL & llngTxnNo & ","
		mstrSQL = mstrSQL & i+1 & ","
		mstrSQL = mstrSQL & "'" & lintMainGLCd & "',"
		mstrSQL = mstrSQL & "'" & lintSubGLCd1 & "',"
		mstrSQL = mstrSQL & "'" & lintSubGLCd2 & "',"
		mstrSQL = mstrSQL & "'" & lintSubGLCd3 & "',"
		mstrSQL = mstrSQL & "'" & lintSubGLCd4 & "',"
		mstrSQL = mstrSQL & "'" & lstrPersonalLedgerCd & "',"
		mstrSQL = mstrSQL & "'" & lstrDrCrFlg & "',"
		mstrSQL = mstrSQL & ldblTxnAmt & ","
		mstrSQL = mstrSQL & "'" & aarrTxnDtl(8,i) & "',"
		mstrSQL = mstrSQL & "'" & lstrUserID & "',"
		mstrSQL = mstrSQL & "sysdate "
		mstrSQL = mstrSQL & ")"			

		'Response.Write mstrSQL & "<BR>"		
		'Response.End
		err.Clear
		
		aconEcgcDb.Execute mstrSQL
		
		If Err.number <> 0 then
			'Response.Write "<FONT class=clsError>fnintPassGLTxn: TXN Dtl Insert</FONT><BR>"
			'Response.End
			'Response.Redirect "../../../Common/Error.asp?aintCode=4&aintPage=-1&astrErrNumber="&Err.Number&"&astrErrDescription="&Err.description
			fnintPassGLTxn = -4
			exit function
		end if
		
		' added the validation on 12/03/2003

		if ucase(lstrDrCrFlg) = "CR" then
			ldblTotTxnAmt = CDBL(ldblTotTxnAmt) + ldblTxnAmt
			'Response.Write ldblTotTxnAmt
		else
			'Response.Write ldblTotTxnAmt
			ldblTotTxnAmt = CDBL(ldblTotTxnAmt) - ldblTxnAmt
		end if


		'Include file to update monthly and yearly totals for fiscal year 
		'of the transaction
		lstrFiscalYearForProcessing = lstrFiscalYearForTxnDt
		lstrMonthForProcessing = lstrMonthForTxnDt

		%>
		
		<!-- #include file="NEIAIncUpdMonthTotalspassjv.asp"-->	
		<!-- #include file="NEIAIncUpdYearTotalsPASSJV.asp"-->	
		
		<%

		'Check if transaction was passed for previous fiscal year. If yes and if gl type
		'is 'ASST' or 'LIAB' then update opening balance and current balance of the 
		'present (ongoing) fiscal year.
		
		fnGetMonthAndFiscalYearByDate ldtmTodaysDt
		lstrFiscalYearForTodaysDt = mstrFiscalYear
		if lstrFiscalYearForTxnDt <> lstrFiscalYearForTodaysDt then
			if lstrGLType = "ASST" or lstrGLType = "LIAB" then

				'Include file to update yearly totals for present (ongoing) fiscal year
				lstrFiscalYearForProcessing = lstrFiscalYearForTodaysDt
				
				%>
				
				<!-- #include file="NEIAIncUpdYearTotalsPASSJV.asp"-->	
		
				<%
			
			end if
		end if
		
		'Determine the final balance available in the GL and check if there is any error.
		'If the GL Type is 'ASST' or 'LIAB' then determine balance available on today's 
		'date. If GL Type is 'INCM' or 'LIAB' then check if the fiscal year of transaction
		'is same as current fiscal year. If yes, determine balance as on today's date else
		'determine final balance available for the last fiscal year.
	
		if lstrGLType = "ASST" or lstrGLType = "LIAB" then
			ldtmFinalBalanceDt = Date
		else
			if lstrFiscalYearForTxnDt = lstrFiscalYearForTodaysDt then
				ldtmFinalBalanceDt = Date
			else
				lstrFinalBalanceDt = fngetCurrFiscalYearStartDt
				ldtmFinalBalanceDt = Cdate(fnstrConvertDt(lstrFinalBalanceDt)) - 1
			end if
		end if
		
		'Response.Write "Date .." & ldtmFinalBalanceDt
		'Changes made for defect # 82 on 09-Jan-02 by Amit C.
		larrGLBrBal = fngetGLBranchBalance (aconEcgcDb, lstrLogicalLocCd, lintMainGLCd, lintSubGLCd1, lintSubGLCd2, lintSubGLCd3, lintSubGLCd4, lstrPersonalLedgerCd, FnDate(ldtmFinalBalanceDt))
		
		'Response.write "<BR>Success .." & larrGLBrBal(0) & " Balance returned .. " & larrGLBrBal(1)
		'Response.Write "error .." & err.number & " " & err.description
		'Response.End
		err.Clear	
		if Err.number <> 0 then
			'Response.Write "<FONT class=clsError>fnintPassGLTxn: Branch Balance Select</FONT><BR>"
			'Response.End
			'Response.Redirect "../../../Common/Error.asp?aintCode=4&aintPage=-1&astrErrNumber="&Err.Number&"&astrErrDescription="&Err.description
			fnintPassGLTxn = -4
			exit function
		end if
		
		'Changes made for defect # 82 on 09-Jan-02 by Amit C.
		'If an error has occured in fngetGLBranchBalance then return error passed by that function.
		if larrGLBrBal(0) = 1 then
			fnintPassGLTxn = - cint(larrGLBrBal(1))
			exit function
		end if
		
		ldblGLBrBal =  cdbl(larrGLBrBal(1))
		
		'The GL Balance returned by the procedure will not include the current transaction
		'amount. Add/subtract the transaction amount from the balance returned by the 
		'procedure.
		
		if ucase(lstrDrCrFlg) = "DR" then
			ldblGLBrBal = CDBL(ldblGLBrBal) - ldblTxnAmt
		else
			ldblGLBrBal = CDBL(ldblGLBrBal) + ldblTxnAmt
		end if
		'Response.Write "<BR> Balance Ind .." & lstrDrCrFlg & " Balance .." & ldblGLBrBal
		'Response.Write lstrPersonalLedgerCd &"<BR>" 
		' Select Error record
		Set lrsGLErr = Server.CreateObject ("ADODB.Recordset")
		lrsGLErr.LockType = 2 ' adLockPessimistic

		mstrSQL = "Select error_no,  bal_amt "
		mstrSQL = mstrSQL & "from NEIA_GL_ERROR_DTL "
		mstrSQL = mstrSQL & "where logicalloc_cd='" & lstrLogicalLocCd &"' "
		mstrSQL = mstrSQL & "and error_stat='OPEN' "
		mstrSQL = mstrSQL & "and maingl_cd='" & lintMainGLCd & "' "
		mstrSQL = mstrSQL & "and subgl_cd1='" & lintSubGLCd1 & "' "
		mstrSQL = mstrSQL & "and subgl_cd2='" & lintSubGLCd2 & "' "
		mstrSQL = mstrSQL & "and subgl_cd3='" & lintSubGLCd3 & "' "
		mstrSQL = mstrSQL & "and subgl_cd4='" & lintSubGLCd4 & "' "
		mstrSQL = mstrSQL & "and personal_ledger_cd='" & lstrPersonalLedgerCd & "' "
		mstrSQL = mstrSQL & "and fiscal_yr='" & lstrFiscalYearForProcessing & "' "
		mstrSQL = mstrSQL & "and entity_cd='NEIA' "

		'Response.Write mstrSQL & "<BR>"
		'Response.End
		

		lrsGLErr.Open mstrSQL, aconEcgcDb

		if Err.number <> 0 then
			'Response.Write "<FONT class=clsError>fnintPassGLTxn: Error record</FONT><BR>"
			'Response.End
			'Response.Redirect "../../../Common/Error.asp?aintCode=4&aintPage=-1&astrErrNumber="&Err.Number&"&astrErrDescription="&Err.description
			fnintPassGLTxn = -4
			exit function
		end if

		lintErrNo = 0
		ldblErrBalAmt = 0
		lblnOldErr = false  ' Error record found
			
		if not lrsGLErr.EOF then
			lintErrNo = cint(lrsGLErr.Fields(0))
			ldblErrBalAmt = cdbl(lrsGLErr.Fields(1))
			lblnOldErr = true
		end if
		lrsGLErr.Close()
		set lrsGLErr = nothing			

		lblnErr = true ' Flag indicating whether transaction resulted in error or not

		if lstrBalInd = "CR" then
			if ldblGLBrBal > 0 then
				lblnErr = false
			else if lstrZeroBalFlg = "Y" and ldblGLBrBal = 0 then
				lblnErr = false
				end if
			end if
		end if

		if lstrBalInd = "DR" then
			if ldblGLBrBal < 0 then
				lblnErr = false
			else if lstrZeroBalFlg = "Y" and ldblGLBrBal = 0 then
				lblnErr = false
				end if
			end if
		end if

		if lstrBalInd = "BOTH" then
			if lstrZeroBalFlg = "Y" and ldblGLBrBal = 0 then
				lblnErr = false
			end if
		end if
			
		if lblnErr then
		' Error found
			if lblnOldErr then
			' Error record present hence update error record
				mstrSQL = "UPDATE NEIA_GL_ERROR_DTL "
				mstrSQL = mstrSQL & "set bal_amt = " & ldblGLBrBal & ", "
				mstrSQL = mstrSQL & "user_id='" & lstrUserID & "',"
				mstrSQL = mstrSQL & "last_trans_date=sysdate "
				mstrSQL = mstrSQL & "where logicalloc_cd='" & lstrLogicalLocCd &"' "
				mstrSQL = mstrSQL & "and error_no=" & lintErrNo & " "
				mstrSQL = mstrSQL & "and error_stat='OPEN' "
				mstrSQL = mstrSQL & "and maingl_cd='" & lintMainGLCd & "' "
				mstrSQL = mstrSQL & "and subgl_cd1='" & lintSubGLCd1 & "' "
				mstrSQL = mstrSQL & "and subgl_cd2='" & lintSubGLCd2 & "' "
				mstrSQL = mstrSQL & "and subgl_cd3='" & lintSubGLCd3 & "' "
				mstrSQL = mstrSQL & "and subgl_cd4='" & lintSubGLCd4 & "' "
				mstrSQL = mstrSQL & "and personal_ledger_cd='" & lstrPersonalLedgerCd & "' "
				mstrSQL = mstrSQL & "and fiscal_yr='" & lstrFiscalYearForProcessing & "' "
				mstrSQL = mstrSQL & "and entity_cd='NEIA' "

				'Response.Write mstrSQL & "<BR>"
				'Response.end
				err.Clear
				aconEcgcDb.Execute mstrSQL, lintRecords
                
				'Response.Write  lintRecords
				'Response.End
				
				If lintRecords <> 1 or Err.number <> 0 then
					'Response.Write "<FONT class=clsError>fnintPassGLTxn: Error Update</FONT><BR>"
					'Response.End
					'Response.Redirect "../../../Common/Error.asp?aintCode=4&aintPage=-1&astrErrNumber="&Err.Number&"&astrErrDescription="&Err.description
					fnintPassGLTxn = -4
					exit function
				End If
			else
			
			' Create new error record	
				' Get next error number
				Set lrsGLErr = Server.CreateObject ("ADODB.Recordset")
				lrsGLErr.LockType = 2 ' adLockPessimistic

				mstrSQL = "Select nvl(max(error_no),0) "
				mstrSQL = mstrSQL & "from NEIA_GL_ERROR_DTL "
				mstrSQL = mstrSQL & "where logicalloc_cd='" & lstrLogicalLocCd &"' "
				mstrSQL = mstrSQL & "and maingl_cd='" & lintMainGLCd & "' "
				mstrSQL = mstrSQL & "and subgl_cd1='" & lintSubGLCd1 & "' "
				mstrSQL = mstrSQL & "and subgl_cd2='" & lintSubGLCd2 & "' "
				mstrSQL = mstrSQL & "and subgl_cd3='" & lintSubGLCd3 & "' "
				mstrSQL = mstrSQL & "and subgl_cd4='" & lintSubGLCd4 & "' "
				mstrSQL = mstrSQL & "and personal_ledger_cd='" & lstrPersonalLedgerCd & "' "
				mstrSQL = mstrSQL & "and fiscal_yr='" & lstrFiscalYearForProcessing & "' "
				mstrSQL = mstrSQL & "and entity_cd='NEIA' "

				'Response.Write mstrSQL & "<BR>"
				'Response.End
				
				err.Clear
				lrsGLErr.Open mstrSQL, aconEcgcDb

				if Err.number <> 0 then
					'Response.Write "<FONT class=clsError>fnintPassGLTxn: Balance Summary Select</FONT><BR>"
					'Response.End
					'Response.Redirect "../../../Common/Error.asp?aintCode=4&aintPage=-1&astrErrNumber="&Err.Number&"&astrErrDescription="&Err.description
					fnintPassGLTxn = -4
					exit function
				end if

				lintErrNo = cint(lrsGLErr.Fields(0)) + 1
				lrsGLErr.Close()
				set lrsGLErr = nothing

				mstrSQL ="Insert into NEIA_GL_ERROR_DTL(entity_cd,logicalloc_cd, maingl_cd,subgl_cd1,subgl_cd2,subgl_cd3,subgl_cd4,personal_ledger_cd,fiscal_yr,error_no,bal_amt,initial_txn_dt,initial_txn_type,initial_txn_no,error_stat,user_id,last_trans_date) "
				mstrSQL = mstrSQL & "values ("
				mstrSQL = mstrSQL & "'" & lstrEntityCd & "',"
				mstrSQL = mstrSQL & "'" & lstrLogicalLocCd & "',"
				mstrSQL = mstrSQL & "'" & lintMainGLCd & "',"
				mstrSQL = mstrSQL & "'" & lintSubGLCd1 & "',"
				mstrSQL = mstrSQL & "'" & lintSubGLCd2 & "',"
				mstrSQL = mstrSQL & "'" & lintSubGLCd3 & "',"
				mstrSQL = mstrSQL & "'" & lintSubGLCd4 & "',"
				mstrSQL = mstrSQL & "'" & lstrPersonalLedgerCd & "', "
				mstrSQL = mstrSQL & "'" & lstrFiscalYearForProcessing & "', "
				mstrSQL = mstrSQL & lintErrNo & ","
				mstrSQL = mstrSQL & ldblGLBrBal & ","
				mstrSQL = mstrSQL & "to_date('" & lstrTxnDt & "', 'dd/mm/yyyy') ,"
				mstrSQL = mstrSQL & "'" & lstrTxnType & "',"
				mstrSQL = mstrSQL & llngTxnNo & ","
				mstrSQL = mstrSQL & "'OPEN',"
				mstrSQL = mstrSQL & "'" & lstrUserID & "',"
				mstrSQL = mstrSQL & "sysdate "
				mstrSQL = mstrSQL & ")"			

				'Response.Write mstrSQL & "<BR>"		
				'Response.End
				
                err.Clear 
				aconEcgcDb.Execute mstrSQL
					
				If Err.number <> 0 then
					'Response.Write "<FONT class=clsError>fnintPassGLTxn: Error Insert</FONT><BR>"
					'Response.End
					'Response.Redirect "../../../Common/Error.asp?aintCode=4&aintPage=-1&astrErrNumber="&Err.Number&"&astrErrDescription="&Err.description
					fnintPassGLTxn = -4
					exit function
				end if
			end if
		else
		
		' Error not found
			if lblnOldErr then
			' Since Error record is present, close error record
				mstrSQL = "UPDATE NEIA_GL_ERROR_DTL "
				mstrSQL = mstrSQL & "set bal_amt = " & ldblGLBrBal & ", "
				mstrSQL = mstrSQL & "error_stat = 'CLS', "
				mstrSQL = mstrSQL & "final_txn_no=" & llngTxnNo & ","
				mstrSQL = mstrSQL & "final_txn_type='" & lstrTxnType & "',"
				mstrSQL = mstrSQL & "final_txn_dt=to_date('" & lstrTxnDt & "', 'dd/mm/yyyy') ,"
				mstrSQL = mstrSQL & "user_id='" & lstrUserID & "',"
				mstrSQL = mstrSQL & "last_trans_date=sysdate "
				mstrSQL = mstrSQL & "where logicalloc_cd='" & lstrLogicalLocCd &"' "
				mstrSQL = mstrSQL & "and error_no=" & lintErrNo & " "
				mstrSQL = mstrSQL & "and error_stat='OPEN' "
				mstrSQL = mstrSQL & "and maingl_cd='" & lintMainGLCd & "' "
				mstrSQL = mstrSQL & "and subgl_cd1='" & lintSubGLCd1 & "' "
				mstrSQL = mstrSQL & "and subgl_cd2='" & lintSubGLCd2 & "' "
				mstrSQL = mstrSQL & "and subgl_cd3='" & lintSubGLCd3 & "' "
				mstrSQL = mstrSQL & "and subgl_cd4='" & lintSubGLCd4 & "' "
				mstrSQL = mstrSQL & "and personal_ledger_cd='" & lstrPersonalLedgerCd & "' "
				mstrSQL = mstrSQL & "and fiscal_yr='" & lstrFiscalYearForProcessing & "' "
				mstrSQL = mstrSQL & "and entity_cd='NEIA' "

				'Response.Write mstrSQL & "<BR>"
				err.Clear
				aconEcgcDb.Execute mstrSQL, lintRecords

				If lintRecords <> 1 or Err.number <> 0 then
					'Response.Write "<FONT class=clsError>fnintPassGLTxn: Error Close</FONT><BR>"
					'Response.End
					'Response.Redirect "../../../Common/Error.asp?aintCode=4&aintPage=-1&astrErrNumber="&Err.Number&"&astrErrDescription="&Err.description
					fnintPassGLTxn = -4
					exit function
				End If
			end if
		end if

		' Processing Ends
		
		' ****************************************************************************
	'response.write ldblTotTxnAmt & "<br>"
	next ' for i = 1 to ubound(aarrTxnDtl,2)
	'validation added on 12/03/2003 
	'Response.Write ldblTotTxnAmt
	'Response.End
	
	if abs(ldblTotTxnAmt) > 0.001 then
	
		fnintPassGLTxn = -4547
		exit function
	
	end if

	' Transaction successful
	fnintPassGLTxn = llngTxnNo

end function


'**********************************************************
'Function name : fnintReverseGLTxn
'Purpose       : To Reverse the JV Transaction
'Input         : Connection Object
'                Logical Location Code
'				 Transaction Type
'				 Transaction No.
'				 Reversal Reason
'				 USER ID
'Output        : Returns Transaction Number if the transaction 
'				 is successfully passed
'				 else returns negative error number
'Author        : Milind Khedaskar
'Date          : 19-11-2001
'Modification History
'			     Reversal Date added
'			     Transaction update
'**********************************************************
function fnintReverseGLTxn(aconEcgcDb, astrLogicalLocCd, astrTxnType, alngTxnNo, astrReversalDt, astrReversalReason, astrUserID)
	
	on error resume next
	
	Dim lrsGLTxn
	dim larrTxnHdr(8)
	dim larrTxnDtl()
	Dim lintTxnRows
	Dim lstrTodaysDt
	Dim i
	Dim lintRecords
	Dim lngTxnNo
	Dim lstrRefNo
	
	Set lrsGLTxn = Server.CreateObject ("ADODB.Recordset")
	
	lrsGLTxn.LockType = 2 ' adLockPessimistic
	lrsGLTxn.CursorLocation = 3
	
	mstrSQL = "Select * "
	mstrSQL = mstrSQL & "from NEIA_GL_TXN_HDR "
	mstrSQL = mstrSQL & "where logicalloc_cd='" & astrLogicalLocCd &"' "
	mstrSQL = mstrSQL & "and gl_txn_type='" & astrTxnType & "' "
	mstrSQL = mstrSQL & "and gl_txn_no=" & alngTxnNo & " "
	mstrSQL = mstrSQL & "and entity_cd='NEIA' "

	'Response.Write mstrSQL & "<BR>"
'Response.end
	lrsGLTxn.Open mstrSQL, aconEcgcDb
	if Err.number <> 0 then
		'Response.Write "<FONT class=clsError>fnintReverseGLTxn: lrsGLTxn Open</FONT><BR>"
		'Response.End 
		'Response.Redirect "../../../Common/Error.asp?aintCode=4&aintPage=-1&astrErrNumber="&Err.Number&"&astrErrDescription="&Err.description
		fnintReverseGLTxn = -4
		exit function
	end if	

	if lrsGLTxn.EOF then
		'Response.Write "<FONT class=clsError>fnintReverseGLTxn: Transaction Record not found</FONT><BR>"
		'Response.End 
		'Response.Redirect "../../../Common/Error.asp?aintCode=4506&aintPage=-1&astrErrNumber="&Err.Number&"&astrErrDescription="&Err.description
		fnintReverseGLTxn = -4506
		exit function
	end if

	if IsNumeric(trim(lrsGLTxn.Fields("reversal_txn_no"))) Then
		' Transaction already reversed
		fnintReverseGLTxn = -4516
		exit function
	End if
	
	' Fill up larrTxnHdr
	'lstrTodaysDt = FnDate(date)
	lstrRefNo = lrsGLTxn.Fields("REFERENCE")
	
	larrTxnHdr(0) = "NEIA"
	larrTxnHdr(1) = astrLogicalLocCd
	larrTxnHdr(2) = "RV"
	larrTxnHdr(3) = astrReversalDt 
	larrTxnHdr(4) = lstrRefNo
	larrTxnHdr(5) = astrUserID
	larrTxnHdr(6) = alngTxnNo
	larrTxnHdr(7) = astrTxnType
	larrTxnHdr(8) = astrReversalReason
	
	lrsGLTxn.Close

	' Get Transaction Details
	mstrSQL = "Select * "
	mstrSQL = mstrSQL & "from NEIA_GL_TXN_DTL "
	mstrSQL = mstrSQL & "where logicalloc_cd='" & astrLogicalLocCd &"' "
	mstrSQL = mstrSQL & "and gl_txn_type='" & astrTxnType & "' "
	mstrSQL = mstrSQL & "and gl_txn_no=" & alngTxnNo & " "
	mstrSQL = mstrSQL & "and entity_cd='NEIA' "
	'Response.Write mstrSQL & "<BR>"
	
	lrsGLTxn.Open mstrSQL, aconEcgcDb
	
	if Err.number <> 0 then
		'Response.Write "<FONT class=clsError>fnintReverseGLTxn: lrsGLTxn Open</FONT><BR>"
		'Response.End 
		'Response.Redirect "../../../Common/Error.asp?aintCode=4&aintPage=-1&astrErrNumber="&Err.Number&"&astrErrDescription="&Err.description
		fnintReverseGLTxn = -4
		exit function
	end if
	
	if lrsGLTxn.EOF then
		'Response.Write "<FONT class=clsError>fnintReverseGLTxn: lrsGLTxn Open</FONT><BR>"
		'Response.End 
		'Response.Redirect "../../../Common/Error.asp?aintCode=4506&aintPage=-1&astrErrNumber="&Err.Number&"&astrErrDescription="&Err.description
		fnintReverseGLTxn = -4506
		exit function
	end if	

	' Fill us larrTxnDtl
	
	lintTxnRows = lrsGLTxn.RecordCount
	redim larrTxnDtl(8,lintTxnRows-1)
	i = 0
	while NOT lrsGLTxn.EOF

		strDrCrFlg = lrsGLTxn.Fields("DR_CR_FLG")
		if strDrCrFlg = "DR" then
			strDrCrFlg = "CR"
		else
			strDrCrFlg = "DR"
		end if
		
		larrTxnDtl(0,i) = lrsGLTxn.Fields("MAINGL_CD")
		larrTxnDtl(1,i) = lrsGLTxn.Fields("SUBGL_CD1")
		larrTxnDtl(2,i) = lrsGLTxn.Fields("SUBGL_CD2")
		larrTxnDtl(3,i) = lrsGLTxn.Fields("SUBGL_CD3")
		larrTxnDtl(4,i) = lrsGLTxn.Fields("SUBGL_CD4")
		larrTxnDtl(5,i) = lrsGLTxn.Fields("PERSONAL_LEDGER_CD")
		larrTxnDtl(6,i) = strDrCrFlg
		larrTxnDtl(7,i) = lrsGLTxn.Fields("TXN_AMT")
		larrTxnDtl(8,i) = lrsGLTxn.Fields("TXN_RMK")
		
		lrsGLTxn.MoveNext
		
		i = i + 1
		
	wend

	lrsGLTxn.Close
	
	' Pass the reverse transaction
	lngTxnNo = fnintPassGLTxn(aconEcgcdb, larrTxnHdr, larrTxnDtl)
	
	' Assign the Result
	fnintReverseGLTxn = lngTxnNo
	
	if clng(lngTxnNo) > 0 then
		' Transaction successful
		
		' Update Transaction
		mstrSQL ="Update NEIA_GL_TXN_HDR "
		mstrSQL = mstrSQL & "set "
		mstrSQL = mstrSQL & "reversal_txn_type='RV', "
		mstrSQL = mstrSQL & "reversal_txn_no=" & lngTxnNo & ", "
		mstrSQL = mstrSQL & "reversal_reason='" & astrReversalReason & "', "
		mstrSQL = mstrSQL & "user_id='" & Session("sUserID") & "', "
		mstrSQL = mstrSQL & "last_trans_date=sysdate "
		mstrSQL = mstrSQL & "where gl_txn_no=" & alngTxnNo & " "
		mstrSQL = mstrSQL & "AND gl_txn_type='" & astrTxnType & "' "
		mstrSQL = mstrSQL & "AND logicalloc_cd='" & astrLogicalLocCd & "'"
		mstrSQL = mstrSQL & "and entity_cd='NEIA' "

		'Response.Write mstrSQL & "<BR>"

		aconEcgcDb.Execute mstrSQL, lintRecords

		If Err.number <> 0 then
			fnintReverseGLTxn = -4
			exit function
		End If
		If lintRecords <> 1 then
			fnintReverseGLTxn = -4506
			exit function
		End If
		
		' Update Petty cash records
		if astrTxnType = "PC" then
		
			mstrSQL ="Update NEIA_petty_cash set "
			mstrSQL = mstrSQL & "gl_txn_no=null, "
			mstrSQL = mstrSQL & "user_id='" & Session("sUserID") & "', "
			mstrSQL = mstrSQL & "last_trans_date=sysdate "
			mstrSQL = mstrSQL & "where entity_cd='NEIA' "
			mstrSQL = mstrSQL & "and logicalloc_cd='" & strLogicalLocCd &"' "
			mstrSQL = mstrSQL & "and gl_txn_type='" & astrTxnType &"' "
			mstrSQL = mstrSQL & "and gl_txn_no=" & alngTxnNo
			'Response.Write mstrSQL & "<BR>"
			aconEcgcDb.Execute mstrSQL, lintRecords
			If Err.number <> 0 then
				fnintReverseGLTxn = -4
				exit function
			End If
			If lintRecords = 0 then
				fnintReverseGLTxn = -2
				exit function
			End If
		
		end if
		
		' Update Receipt records
		if astrTxnType = "RE" then
		
			mstrSQL ="Update NEIA_rcpts set "
			mstrSQL = mstrSQL & "gl_flg='R', "
			mstrSQL = mstrSQL & "user_id='" & Session("sUserID") & "', "
			mstrSQL = mstrSQL & "last_trans_date=sysdate "
			mstrSQL = mstrSQL & "where entity_cd='NEIA' "
			mstrSQL = mstrSQL & "and logicalloc_cd='" & strLogicalLocCd &"' "
			mstrSQL = mstrSQL & "and rcpt_no=" & lstrRefNo
			'Response.Write mstrSQL & "<BR>"
			aconEcgcDb.Execute mstrSQL, lintRecords
			If Err.number <> 0 then
				fnintReverseGLTxn = -4
				exit function
			End If
			If lintRecords = 0 then
				fnintReverseGLTxn = -2
				exit function
			End If
		
		end if

	End if ' if cint(lngTxnNo) > 0 then
		
end function


'**********************************************************
'Function name : fngetGLBranchBalance
'Purpose       : Retrieve the GL Branch Balance
'Input         : Connection Object
'                Logical Location Code 
'				 Main GL Code 
'                Sub GL Code 1
'                Sub GL Code 2
'                Sub GL Code 3
'                Sub GL Code 4
'                Personal Ledger Code
'                Transaction Date in DD/MM/YYYY format
'
'Output        : return a 2 element array
'				 1st element will indicate success (0) or error (1) in function
'                2nd element will contain gl balance or error number
'
'Author        : Milind Khedaskar
'Date          : 28-11-2001
'
'Modification history	
'Date		   : 8-Jan-02. Changes for defect # 82.
'                Made changes for returning result in an array - larrResult
'                larrResult(0) will be set to 1 if any error occurs else will be 0
'                larrResult(1) will be set to error number if an error occurs else 
'                will contain the gl balance. Also removed old commented code from 
'                the function. - Amit C
'Date		   : 12-Feb-02. Changes for defect # 153.
'                Made changes for returning balance from the modified balance tables.
'*********************************************************************************************
Function fngetGLBranchBalance(aconEcgcDb, astrLogicalLocCd, aintMainGLCd, aintSubGLCd1, aintSubGLCd2, aintSubGLCd3, aintSubGLCd4, astrPersonalLedgerCd, astrTxnDt)
'response.write astrLogicalLocCd & "<br>"
'response.write aintMainGLCd & "<br>"
'response.write aintSubGLCd1 & "<br>"
'response.write  aintSubGLCd2 & "<br>"
'response.write astrPersonalLedgerCd & "<br>"
'response.end	
	'on error resume next
	
	Dim larrTempDt
	Dim lstrTempDt
	Dim larrResult(1)
	Dim lobjCmd
	Dim param0, param1, param2, param3, param4
	Dim param5, param6, param7, param8, param9, param10
	
	'Convert date to dd-mon-yy format - default format used in Oracle
	larrTempDt = split(astrTxnDt,"/")
	lstrTempDt = larrTempDt(0) + "-" + MonthName(larrTempDt(1),True) + "-" + mid(larrTempDt(2),3,2)
	
	Set lobjCmd=Server.CreateObject("ADODB.Command")

	Set param0 = Server.CreateObject("ADODB.Parameter")
	Set param1 = Server.CreateObject("ADODB.Parameter")
	Set param2 = Server.CreateObject("ADODB.Parameter")
	Set param3 = Server.CreateObject("ADODB.Parameter")
	Set param4 = Server.CreateObject("ADODB.Parameter")
	Set param5 = Server.CreateObject("ADODB.Parameter")
	Set param6 = Server.CreateObject("ADODB.Parameter")
	Set param7 = Server.CreateObject("ADODB.Parameter")
	Set param8 = Server.CreateObject("ADODB.Parameter")
	Set param9 = Server.CreateObject("ADODB.Parameter")
	Set param10 = Server.CreateObject("ADODB.Parameter")

	lobjCmd.ActiveConnection=astrConn	    
	'lobjCmd.ActiveConnection=aconEcgcDb

	With lobjCmd

		.CommandType=adCmdStoredProc
		.CommandText="NEIA_Get_Branch_Bal"

		Set param0 = .CreateParameter("ipEntity_Cd", adVarChar, adParamInput, 20, "NEIA")
		.Parameters.Append param0

		Set param1 = .CreateParameter("ipLogicalLoc_cd", adVarChar, adParamInput, 20, astrLogicalLocCd)
		.Parameters.Append param1
		
		Set param2 = .CreateParameter("ipMainGl_Cd", adInteger, adParamInput, 4, aintMainGLCd)
		.Parameters.Append param2

		Set param3 = .CreateParameter("ipSubGl_Cd1", adInteger, adParamInput, 3, aintSubGLCd1)
		.Parameters.Append param3

		Set param4 = .CreateParameter("ipSubGl_Cd2", adInteger, adParamInput, 3, aintSubGLCd2)
		.Parameters.Append param4

		Set param5 = .CreateParameter("ipSubGl_Cd3", adInteger, adParamInput, 3, aintSubGLCd3)
		.Parameters.Append param5

		Set param6 = .CreateParameter("ipSubGl_Cd4", adInteger, adParamInput, 3, aintSubGLCd4)
		.Parameters.Append param6

		Set param7 = .CreateParameter("ipPersonal_Ledger_Cd", adVarChar, adParamInput, 150, ucase(astrPersonalLedgerCd))
		.Parameters.Append param7

		Set param8 = .CreateParameter("ipBalance_Dt", adVarChar, adParamInput, 10, lstrTempDt)
		.Parameters.Append param8

		Set param9 = .CreateParameter("opCr_Balance", adDecimal, adParamOutput, 14, 0)
		.Parameters.Append param9

		Set param10 = .CreateParameter("opDr_Balance", adDecimal, adParamOutput, 14, 0)
		.Parameters.Append param10

			
		.Execute
		
		if Err.number <> 0 then	
			'Clear the error so that the calling routine is not affected
			Err.Clear
			larrResult(0) = 1						'Error
			larrResult(1) = 4507					'Error in determining branch general ledger balance
			
			fngetGLBranchBalance = larrResult
			exit function
		end if
		
	End With

	larrResult(0) = 0						'Success
	larrResult(1) = param9 - param10		'Balance
 	
	
'----------------------------------------------------------------------------------------------------------------------
'		if astrPersonalLedgerCd = "98546" then
'			Response.Write "start"&"<br>"
'			Response.Write "param0  : " & param0 &"<br>"
'			Response.Write "param1  : " & param1 &"<br>"
'			Response.Write "param2  : " & param2 &"<br>"
'			Response.Write "param3  : " & param3 &"<br>"
'			Response.Write "param4  : " & param4 &"<br>"
'			Response.Write "param5  : " & param5 &"<br>"
'			Response.Write "param6  : " & param6 &"<br>"
'			Response.Write "param7  : " & param7 &"<br>"
'			Response.Write "param8  : " & param8 &"<br>"
'			Response.Write "param9  : " & param9 &"<br>"
'			Response.Write "param10  : " & param10 &"<br>"
'			Response.Write "end"&"<br>"
'			Response.End
'		end if
'----------------------------------------------------------------------------------------------------------------------
	fngetGLBranchBalance = larrResult
	
End Function


'**********************************************************
'Function name : fngetGLBalanceSummary
'Purpose       : Retrieve the GL Balance Summary
'Input         : Main GL Code 
'                Sub GL Code 1
'                Sub GL Code 2
'                Sub GL Code 3
'                Sub GL Code 4
'                Transaction Date
'
'Output        : return balance
'				 
'Author        : Milind Khedaskar
'Date          : 28-11-2001
'**********************************************************
Function fngetGLBalanceSummary(aintMainGLCd, aintSubGLCd1, aintSubGLCd2, aintSubGLCd3, aintSubGLCd4, astrTxnDt)
	
	on error resume next
	
	Dim larrTempDt
	Dim lstrTempDt

	Dim objCmd
	Dim param0, param1, param2, param3, param4
	Dim param5, param6, param7, param8
	
	'Convert date to dd-mon-yy format - default format used in Oracle
	larrTempDt = split(astrTxnDt,"/")
	lstrTempDt = larrTempDt(0) + "-" + MonthName(larrTempDt(1),True) + "-" + mid(larrTempDt(2),3,2)
	'Response.Write lstrTempDt
		
	Set objCmd=Server.CreateObject("ADODB.Command")

	Set param0 = Server.CreateObject("ADODB.Parameter")
	Set param1 = Server.CreateObject("ADODB.Parameter")
	Set param2 = Server.CreateObject("ADODB.Parameter")
	Set param3 = Server.CreateObject("ADODB.Parameter")
	Set param4 = Server.CreateObject("ADODB.Parameter")
	Set param5 = Server.CreateObject("ADODB.Parameter")
	Set param6 = Server.CreateObject("ADODB.Parameter")
	Set param7 = Server.CreateObject("ADODB.Parameter")
	Set param8 = Server.CreateObject("ADODB.Parameter")

	objCmd.ActiveConnection=astrConn	    

	With objCmd

		.CommandType=adCmdStoredProc
		.CommandText="NEIA_Get_Bal_Summ"

		Set param0 = .CreateParameter("ipEntity_Cd", adVarChar, adParamInput, 20, "NEIA")
		.Parameters.Append param0

		Set param1 = .CreateParameter("ipMainGl_Cd", adInteger, adParamInput, 4, aintMainGLCd)
		.Parameters.Append param1

		Set param2 = .CreateParameter("ipSubGl_Cd1", adInteger, adParamInput, 3, aintSubGLCd1)
		.Parameters.Append param2

		Set param3 = .CreateParameter("ipSubGl_Cd2", adInteger, adParamInput, 3, aintSubGLCd2)
		.Parameters.Append param3

		Set param4 = .CreateParameter("ipSubGl_Cd3", adInteger, adParamInput, 3, aintSubGLCd3)
		.Parameters.Append param4

		Set param5 = objCmd.CreateParameter("ipSubGl_Cd4", adInteger, adParamInput, 3, aintSubGLCd4)
		.Parameters.Append param5

		Set param6 = objCmd.CreateParameter("ipBalance_Dt", adVarChar, adParamInput, 10, lstrTempDt)
		.Parameters.Append param6

		Set param7 = objCmd.CreateParameter("opCr_Balance", adDecimal, adParamOutput, 14, 0)
		.Parameters.Append param7

		Set param8 = objCmd.CreateParameter("opDr_Balance", adDecimal, adParamOutput, 14, 0)
		.Parameters.Append param8

		.Execute
		
	End With

	fngetGLBalanceSummary = param7 - param8
	
End Function


'**********************************************************
'Function name : fnGetExporterDtls
'Purpose       : This will get Exporter Details
'Input         : IE Code
'Output        : Exporter Details in array
'Author        : Milind Khedaskar
'Date          : 28-11-2001
'**********************************************************
Function fnGetExporterDtls(astrIECd)
	
	fnGetExporterDtls = ""
	
	Dim conCommon
	Dim rsCommon
	Dim strSQL
	
	Set conCommon = Server.CreateObject ("ADODB.Connection")
	Set rsCommon = Server.CreateObject ("ADODB.Recordset")
	conCommon.Open astrConn
	if Err.number <> 0 then
		'Response.Write "<FONT class=clsError>fnGetExporterDtls: Connection Error</FONT><BR>"
		'Response.Write "Error No. " & Err.number & " "
		'Response.Write "Error Desc. " & Err.Description & " "
		'Response.End 
		Response.Redirect "../../../Common/Error.asp?aintCode=4&aintPage=-1&astrErrNumber="&Err.Number&"&astrErrDescription="&Err.description
	end if	
	
	strSQL = "Select ie_cd, expo_name " 
	strSQL = strSQL & "from expo_mst "
	strSQL = strSQL & "where ie_cd = UPPER('" & astrIECd & "')"

	'Response.Write strSQL		
	rsCommon.Open strSQL, conCommon
	
	if Err.number <> 0 then
		'Response.Write "<FONT class=clsError>fnGetExporterDtls: Connection Error</FONT><BR>"
		'Response.Write "Error No. " & Err.number & " "
		'Response.Write "Error Desc. " & Err.Description & " "
		'Response.End 
		Response.Redirect "../../../Common/Error.asp?aintCode=4&aintPage=-1&astrErrNumber="&Err.Number&"&astrErrDescription="&Err.description
	end if	
	
	if not rsCommon.EOF then
	
		fnGetExporterDtls = rsCommon.GetRows()
	
	else
		
		'Response.Write "<FONT class=clsError>fnGetExporterDtls: Record not found</FONT><BR>"
		'Response.Write "Error No. " & Err.number & " "
		'Response.Write "Error Desc. " & Err.Description & " "
		'Response.End 
		Response.Redirect "../../../Common/Error.asp?aintCode=2&aintPage=-1&astrErrNumber="&Err.Number&"&astrErrDescription=Exporter record does not exist."
	
	
	end if
	
	rsCommon.Close
	conCommon.Close 
		
	set rsCommon = nothing
	set conCommon = nothing

End Function

'**********************************************************
'Function name : fnGetEmployeeDtls
'Purpose       : This will get Employee Details
'Input         : IE Code
'Output        : Employee Details in array
'Author        : Milind Khedaskar
'Date          : 28-11-2001
'**********************************************************
Function fnGetEmployeeDtls(aintEmpNo)
	
	fnGetEmployeeDtls = ""
	
	Dim conCommon
	Dim rsCommon
	Dim strSQL
	Dim intEmpNo
	
	Set conCommon = Server.CreateObject ("ADODB.Connection")
	Set rsCommon = Server.CreateObject ("ADODB.Recordset")
	conCommon.Open astrConn
	if Err.number <> 0 then
		'Response.Write "<FONT class=clsError>fnGetEmployeeDtls: Connection Error</FONT><BR>"
		'Response.Write "Error No. " & Err.number & " "
		'Response.Write "Error Desc. " & Err.Description & " "
		'Response.End 
		Response.Redirect "../../../Common/Error.asp?aintCode=4&aintPage=-1&astrErrNumber="&Err.Number&"&astrErrDescription="&Err.description
	end if	
	
	intEmpNo = aintEmpNo
	if not IsNumeric(intEmpNo) then
		intEmpNo = 0
	end if
	
	strSQL = "Select emp_no, emp_alias " 
	strSQL = strSQL & "from hrd_emp_mst "
	strSQL = strSQL & "where emp_no = " & intEmpNo
	
	'Response.Write strSQL & "<BR>"
	rsCommon.Open strSQL, conCommon
	
	if Err.number <> 0 then
		'Response.Write "<FONT class=clsError>fnGetEmployeeDtls: Connection Error</FONT><BR>"
		'Response.Write "Error No. " & Err.number & " "
		'Response.Write "Error Desc. " & Err.Description & " "
		'Response.End 
		Response.Redirect "../../../Common/Error.asp?aintCode=4&aintPage=-1&astrErrNumber="&Err.Number&"&astrErrDescription="&Err.description
	end if	
	
	if not rsCommon.EOF then
	
		fnGetEmployeeDtls = rsCommon.GetRows()
	
	else
		
		'Response.Write "<FONT class=clsError>fnGetEmployeeDtls: Record not found</FONT><BR>"
		'Response.Write "Error No. " & Err.number & " "
		'Response.Write "Error Desc. " & Err.Description & " "
		'Response.End 
		Response.Redirect "../../../Common/Error.asp?aintCode=2&aintPage=-1&astrErrNumber="&Err.Number&"&astrErrDescription=Employee record does not exist."
	
	
	end if
	
	rsCommon.Close
	conCommon.Close 
		
	set rsCommon = nothing
	set conCommon = nothing

End Function

'**********************************************************
'Function name : fnGetBankBrDtls
'Purpose       : This will get Bank Branch Details
'Input         : Bank Code
'                Bank Branch Code
'Output        : Bank Branch Details in array
'Author        : Milind Khedaskar
'Date          : 28-11-2001
'**********************************************************
Function fnGetBankBrDtls(astrBankCd,astrBankBrCd)
	
	fnGetBankBrDtls = ""
	
	Dim conCommon
	Dim rsCommon
	Dim strSQL
	
	Set conCommon = Server.CreateObject ("ADODB.Connection")
	Set rsCommon = Server.CreateObject ("ADODB.Recordset")
	conCommon.Open astrConn
	if Err.number <> 0 then
		'Response.Write "<FONT class=clsError>fnGetBankBrDtls: Connection Error</FONT><BR>"
		'Response.Write "Error No. " & Err.number & " "
		'Response.Write "Error Desc. " & Err.Description & " "
		'Response.End 
		Response.Redirect "../../../Common/Error.asp?aintCode=4&aintPage=-1&astrErrNumber="&Err.Number&"&astrErrDescription="&Err.description
	end if	
	
	strSQL = "Select a.bank_cd, a.bank_name, b.bank_br_cd, b.bank_br_name " 
	strSQL = strSQL & "from bank_fi_mst a, bank_branch_mst b "
	strSQL = strSQL & "where a.bank_cd = b.bank_cd "
	strSQL = strSQL & "and a.bank_cd = UPPER('" & astrBankCd & "') "
	strSQL = strSQL & "and b.bank_br_cd = UPPER('" & astrBankBrCd & "') "
		
	rsCommon.Open strSQL, conCommon
	
	if Err.number <> 0 then
		'Response.Write "<FONT class=clsError>fnGetBankBrDtls: Connection Error</FONT><BR>"
		'Response.Write "Error No. " & Err.number & " "
		'Response.Write "Error Desc. " & Err.Description & " "
		'Response.End 
		Response.Redirect "../../../Common/Error.asp?aintCode=4&aintPage=-1&astrErrNumber="&Err.Number&"&astrErrDescription="&Err.description
	end if	
	
	if not rsCommon.EOF then
	
		fnGetBankBrDtls = rsCommon.GetRows()
	
	else
		
		'Response.Write "<FONT class=clsError>fnGetBankBrDtls: Record not found</FONT><BR>"
		'Response.Write "Error No. " & Err.number & " "
		'Response.Write "Error Desc. " & Err.Description & " "
		'Response.End 
		Response.Redirect "../../../Common/Error.asp?aintCode=2&aintPage=-1&astrErrNumber="&Err.Number&"&astrErrDescription=Bank Branch record does not exist."
	
	
	end if
	
	rsCommon.Close
	conCommon.Close 
		
	set rsCommon = nothing
	set conCommon = nothing

End Function

'**********************************************************
'Function name : fnGetPolicyDtls
'Purpose       : This will get Policy Details
'Input         : Policy No
'Output        : Policy Details in array
'Author        : Milind Khedaskar
'Date          : 21-11-2001
'Modification History :
'28-Dec-2001 - Changes in view of policy no (defect 73)
'**********************************************************
Function fnGetPolicyDtls(astrPolNo)
	
	fnGetPolicyDtls = ""
	
	Dim conCommon
	Dim rsCommon
	Dim strSQL
	
	Set conCommon = Server.CreateObject ("ADODB.Connection")
	Set rsCommon = Server.CreateObject ("ADODB.Recordset")
	conCommon.Open astrConn
	if Err.number <> 0 then
		'Response.Write "<FONT class=clsError>fnGetPolicyDtls: Connection Error</FONT><BR>"
		'Response.Write "Error No. " & Err.number & " "
		'Response.Write "Error Desc. " & Err.Description & " "
		'Response.End 
		Response.Redirect "../../../Common/Error.asp?aintCode=4&aintPage=-1&astrErrNumber="&Err.Number&"&astrErrDescription="&Err.description
	end if	
	
	strSQL = "Select policy_no " 
	strSQL = strSQL & "from VIEW_POLICY_NO "
	strSQL = strSQL & "where policy_no = '" & astrPolNo & "'"
		
	'Response.Write strSQL
	rsCommon.Open strSQL, conCommon
	
	if Err.number <> 0 then
		'Response.Write "<FONT class=clsError>fnGetPolicyDtls: Connection Error</FONT><BR>"
		'Response.Write "Error No. " & Err.number & " "
		'Response.Write "Error Desc. " & Err.Description & " "
		'Response.End 
		Response.Redirect "../../../Common/Error.asp?aintCode=4&aintPage=-1&astrErrNumber="&Err.Number&"&astrErrDescription="&Err.description
	end if	
	
	if not rsCommon.EOF then
	
		fnGetPolicyDtls = rsCommon.GetRows()
	
	else
		
		'Response.Write "<FONT class=clsError>fnGetPolicyDtls: Record not found</FONT><BR>"
		'Response.Write "Error No. " & Err.number & " "
		'Response.Write "Error Desc. " & Err.Description & " "
		'Response.End 
		Response.Redirect "../../../Common/Error.asp?aintCode=2&aintPage=-1&astrErrNumber="&Err.Number&"&astrErrDescription=Policy record does not exist."
	
	
	end if
	
	rsCommon.Close
	conCommon.Close 
		
	set rsCommon = nothing
	set conCommon = nothing

End Function

'**********************************************************
'Function name : fnGetGteeDtls
'Purpose       : This will get Gtee Details
'Input         : gtee no and gtee type
'Output        : 
'Author        : Milind Khedaskar
'Date          : 28-11-2001
'**********************************************************
Function fnGetGteeDtls(astrIECd)
	
	fnGetGteeDtls = ""
	
	' Gtee lookup not ready, so does the table structure
	exit function
	
	Dim conCommon
	Dim rsCommon
	Dim strSQL
	
	Set conCommon = Server.CreateObject ("ADODB.Connection")
	Set rsCommon = Server.CreateObject ("ADODB.Recordset")
	conCommon.Open astrConn
	if Err.number <> 0 then
		'Response.Write "<FONT class=clsError>fnGetGteeDtls: Connection Error</FONT><BR>"
		'Response.Write "Error No. " & Err.number & " "
		'Response.Write "Error Desc. " & Err.Description & " "
		'Response.End 
		Response.Redirect "../../../Common/Error.asp?aintCode=4&aintPage=-1&astrErrNumber="&Err.Number&"&astrErrDescription="&Err.description
	end if	
	
	' strSQL to be prepared
		
	rsCommon.Open strSQL, conCommon
	
	if Err.number <> 0 then
		'Response.Write "<FONT class=clsError>fnGetGteeDtls: Connection Error</FONT><BR>"
		'Response.Write "Error No. " & Err.number & " "
		'Response.Write "Error Desc. " & Err.Description & " "
		'Response.End 
		Response.Redirect "../../../Common/Error.asp?aintCode=4&aintPage=-1&astrErrNumber="&Err.Number&"&astrErrDescription="&Err.description
	end if	
	
	if not rsCommon.EOF then
	
		fnGetGteeDtls = rsCommon.GetRows()
	
	else
		
		'Response.Write "<FONT class=clsError>fnGetGteeDtls: Record not found</FONT><BR>"
		'Response.Write "Error No. " & Err.number & " "
		'Response.Write "Error Desc. " & Err.Description & " "
		'Response.End 
		Response.Redirect "../../../Common/Error.asp?aintCode=2&aintPage=-1&astrErrNumber="&Err.Number&"&astrErrDescription=Guarantee record does not exist."
	
	
	end if
	
	rsCommon.Close
	conCommon.Close 
		
	set rsCommon = nothing
	set conCommon = nothing

End Function

'**********************************************************
'Function name : fnGetStatusAgencyDtls
'Purpose       : This will get Status Agency Details
'Input         : Agency Code
'Output        : Status Agency Details in array
'Author        : Milind Khedaskar
'Date          : 07-12-2001
'**********************************************************
Function fnGetStatusAgencyDtls(astrAgencyCd)
	
	fnGetStatusAgencyDtls = ""
	
	Dim conCommon
	Dim rsCommon
	Dim strSQL
	
	Set conCommon = Server.CreateObject ("ADODB.Connection")
	Set rsCommon = Server.CreateObject ("ADODB.Recordset")
	conCommon.Open astrConn
	if Err.number <> 0 then
		'Response.Write "<FONT class=clsError>fnGetStatusAgencyDtls: Connection Error</FONT><BR>"
		'Response.Write "Error No. " & Err.number & " "
		'Response.Write "Error Desc. " & Err.Description & " "
		'Response.End 
		Response.Redirect "../../../Common/Error.asp?aintCode=4&aintPage=-1&astrErrNumber="&Err.Number&"&astrErrDescription="&Err.description
	end if	
	
	strSQL = "Select AGENCY_CD, AGENCY_NAME " 
	strSQL = strSQL & "from BUD_STATUS_AGENGY_MST "
	strSQL = strSQL & "where AGENCY_CD = '" & astrAgencyCd & "'"
		
	rsCommon.Open strSQL, conCommon
	
	if Err.number <> 0 then
		'Response.Write "<FONT class=clsError>fnGetStatusAgencyDtls: Connection Error</FONT><BR>"
		'Response.Write "Error No. " & Err.number & " "
		'Response.Write "Error Desc. " & Err.Description & " "
		'Response.End 
		Response.Redirect "../../../Common/Error.asp?aintCode=4&aintPage=-1&astrErrNumber="&Err.Number&"&astrErrDescription="&Err.description
	end if	
	
	if not rsCommon.EOF then
	
		fnGetStatusAgencyDtls = rsCommon.GetRows()
	
	else
		
		'Response.Write "<FONT class=clsError>fnGetStatusAgencyDtls: Record not found</FONT><BR>"
		'Response.Write "Error No. " & Err.number & " "
		'Response.Write "Error Desc. " & Err.Description & " "
		'Response.End 
		Response.Redirect "../../../Common/Error.asp?aintCode=2&aintPage=-1&astrErrNumber="&Err.Number&"&astrErrDescription=Status Agency record does not exist."
	
	
	end if
	
	rsCommon.Close
	conCommon.Close 
		
	set rsCommon = nothing
	set conCommon = nothing

End Function

'**********************************************************
'Function name : fnGetContractPartyDtls
'Purpose       : This will get Policy Details
'Input         : Contract Party Code
'Output        : Contract Party Details in array
'Author        : Milind Khedaskar
'Date          : 07-12-2001
'**********************************************************
Function fnGetContractPartyDtls(astrContractPartyCd)
	
	fnGetContractPartyDtls = ""
	
	Dim conCommon
	Dim rsCommon
	Dim strSQL
	
	Set conCommon = Server.CreateObject ("ADODB.Connection")
	Set rsCommon = Server.CreateObject ("ADODB.Recordset")
	conCommon.Open astrConn
	if Err.number <> 0 then
		'Response.Write "<FONT class=clsError>fnGetContractPartyDtls: Connection Error</FONT><BR>"
		'Response.Write "Error No. " & Err.number & " "
		'Response.Write "Error Desc. " & Err.Description & " "
		'Response.End 
		Response.Redirect "../../../Common/Error.asp?aintCode=4&aintPage=-1&astrErrNumber="&Err.Number&"&astrErrDescription="&Err.description
	end if	
	
	strSQL = "Select CONTRACT_PARTY_CD, CONTRACT_PARTY_NAME " 
	strSQL = strSQL & "from NEIA_CONTRACT_PARTY_MST "
	strSQL = strSQL & "where UPPER(CONTRACT_PARTY_CD) = UPPER('" & astrContractPartyCd & "')"
		
	rsCommon.Open strSQL, conCommon
	
	if Err.number <> 0 then
		'Response.Write "<FONT class=clsError>fnGetContractPartyDtls: Connection Error</FONT><BR>"
		'Response.Write "Error No. " & Err.number & " "
		'Response.Write "Error Desc. " & Err.Description & " "
		'Response.End 
		Response.Redirect "../../../Common/Error.asp?aintCode=4&aintPage=-1&astrErrNumber="&Err.Number&"&astrErrDescription="&Err.description
	end if	
	
	if not rsCommon.EOF then
	
		fnGetContractPartyDtls = rsCommon.GetRows()
	
	else
		
		'Response.Write "<FONT class=clsError>fnGetContractPartyDtls: Record not found</FONT><BR>"
		'Response.Write "Error No. " & Err.number & " "
		'Response.Write "Error Desc. " & Err.Description & " "
		'Response.End 
		Response.Redirect "../../../Common/Error.asp?aintCode=2&aintPage=-1&astrErrNumber="&Err.Number&"&astrErrDescription=Contract Party record does not exists"
	
	
	end if
	
	rsCommon.Close
	conCommon.Close 
		
	set rsCommon = nothing
	set conCommon = nothing

End Function


'**********************************************************
'Function name : fnValidateInward
'Purpose       : This will get Inward Details
'Input         : Inward No and Inward Item No
'Output        : 
'Author        : Geeta Negi
'Date          : 28-11-2001
'**********************************************************
Function fnValidateInward(intIwdNo, intIwdItemNo)

	strCommon = "Select to_char(iwd_dt, 'dd/mm/yyyy'), iwd_status, ecgc_branch_cd from iwd_header a, iwd_dtl b where " &_
				" a.iwd_no = b.iwd_no and a.iwd_no = '" & intIwdNo & "'" & " and b.iwd_item_no = '" & intIwdItemNo & "'"
		
	Set conCommon = Server.CreateObject ("ADODB.Connection")
	Set rsCommon = Server.CreateObject ("ADODB.Recordset")

	conCommon.Open astrConn
	rsCommon.Open strCommon, conCommon
	
	If rsCommon.EOF then
		fnDisplayError 101,0
		fnValidateInward = 0
	Else
		If rsCommon.Fields(1) <> "PEN" Then
			fnDisplayError 122,0
			fnValidateInward = 0
		Else
			strBranchCd = rsCommon.Fields(2)
			fnValidateInward = rsCommon.Fields(0)
		End If
	End If
		
	rsCommon.Close 
	conCommon.Close 

	set rsCommon = nothing
	set conCommon = nothing

End Function


'**********************************************************
'Function name : fnGetGLDetails
'Purpose       : This will get General Ledger Details
'Input         : GL Code fields
'Output        : 
'Author        : Amit C
'Date          : 03-12-2001
'**********************************************************
Function fnGetGLDetails(astrMainGlCode, astrSubGlCd1, astrSubGlCd2, astrSubGlCd3, astrSubGlCd4)

	fnGetGLDetails = ""
	
	Dim conCommon
	Dim rsCommon
	Dim strSQL
	
	Set conCommon = Server.CreateObject ("ADODB.Connection")
	Set rsCommon = Server.CreateObject ("ADODB.Recordset")
	conCommon.Open astrConn
	if Err.number <> 0 then
		'Response.Write "<FONT class=clsError>fnGetGLDetails: Connection Error</FONT><BR>"
		'Response.Write "Error No. " & Err.number & " "
		'Response.Write "Error Desc. " & Err.Description & " "
		'Response.End 
		Response.Redirect "../../../Common/Error.asp?aintCode=4&aintPage=-1&astrErrNumber="&Err.Number&"&astrErrDescription="&Err.description
	end if	
	
	strSQL = "Select active, gl_is_group, gl_type, personal_ledger_level, bal_ind, zero_bal_flg "
	strSQL = strSQL & "from NEIA_ENTITY_GL_MST "
	strSQL = strSQL & "where entity_cd='NEIA' "
	strSQL = strSQL & "and maingl_cd='" & astrMainGLCd & "' "
	strSQL = strSQL & "and subgl_cd1='" & astrSubGLCd1 & "' "
	strSQL = strSQL & "and subgl_cd2='" & astrSubGLCd2 & "' "
	strSQL = strSQL & "and subgl_cd3='" & astrSubGLCd3 & "' "
	strSQL = strSQL & "and subgl_cd4='" & astrSubGLCd4 & "' "		

	rsCommon.Open strSQL, conCommon
	
	if Err.number <> 0 then
		'Response.Write "<FONT class=clsError>fnGetGLDetails: Connection Error</FONT><BR>"
		'Response.Write "Error No. " & Err.number & " "
		'Response.Write "Error Desc. " & Err.Description & " "
		'Response.End 
		Response.Redirect "../../../Common/Error.asp?aintCode=4&aintPage=-1&astrErrNumber="&Err.Number&"&astrErrDescription="&Err.description
	end if	
	
	if not rsCommon.EOF then
		fnGetGLDetails = rsCommon.GetRows()
	else
		'Response.Write "<FONT class=clsError>fnGetPolicyDtls: Record not found</FONT><BR>"
		'Response.Write "Error No. " & Err.number & " "
		'Response.Write "Error Desc. " & Err.Description & " "
		'Response.End 
		Response.Redirect "../../../Common/Error.asp?aintCode=2&aintPage=-1&astrErrNumber="&Err.Number&"&astrErrDescription="&Err.description
	end if
	
	rsCommon.Close
	conCommon.Close 
		
	set rsCommon = nothing
	set conCommon = nothing

End Function


'*****************************************************************************************
'Function name : fnGetGLBalanceUtil
'Purpose       : This will get reserved and utilised amount for given general ledger.
'                personal ledger and a specific use-for code or the total amount reserved
'                against all use-for codes
'Input         : Array containing logical location and general ledger details
'Output        : Three element array. 
'				 1st element will indicate success (0) or error (1) in function
'				 2nd element will contain reserved amount information or error number
'				 3rd element will contain utilized amount information or zero
'Author        : Amit C
'Date          : 31-01-2002
'******************************************************************************************
Function fnGetGLBalanceUtil(aarrGlBalUtil)

	on error resume next
	
	Dim lconEcgcdb, lrsGLBalUtil
	Dim lstrSQL

	Dim luseForCd, lcurrFiscalYr	
	Dim larrResult(2)

	fnGetMonthAndFiscalYearByDate(date)
	lcurrFiscalYr = mstrFiscalYear
	'Response.Write lcurrFiscalYr
	'Response.End
	
	Set lconEcgcDb = Server.CreateObject ("ADODB.Connection")
	set lrsGLBalUtil = Server.CreateObject( "ADODB.Recordset" )

	' Create database connection
	lconEcgcDb.Open astrConn
	if Err.number <> 0 then
		'Response.Redirect "../../../Common/Error.asp?aintCode=4&aintPage=-1&astrErrNumber="&Err.Number&"&astrErrDescription="&Err.description
		larrResult(0) = 1					'Error
		larrResult(1) = 4					'Error number
		larrResult(2) = 0					'Set last element of result array to zero
		fnGetBalanceUtil = larrResult
		exit function
	end if	
	
	'Determine if amount reserved information is to be fetched for a single
	'use-for-code or if the total amount reserved in the gl is to be fetched
	if aarrGlBalUtil(8) = "AllUseForCodes" then
		luseForCd = "%"
	else
		luseForCd = aarrGlBalUtil(8)
	end if
	
	'Determine total amount reserved
	lstrSQL = "Select nvl(sum(amount),0) from NEIA_gl_bal_util "
	lstrSQL = lstrSQL & "where entity_cd = '" & aarrGlBalUtil(0) & "' and "
	lstrSQL = lstrSQL & "logicalloc_cd = '" & aarrGlBalUtil(1) & "' and "
	lstrSQL = lstrSQL & "maingl_cd = " & aarrGlBalUtil(2) & " and "
	lstrSQL = lstrSQL & "subgl_cd1 = " & aarrGlBalUtil(3) & " and "
	lstrSQL = lstrSQL & "subgl_cd2 = " & aarrGlBalUtil(4) & " and "
	lstrSQL = lstrSQL & "subgl_cd3 = " & aarrGlBalUtil(5) & " and "
	lstrSQL = lstrSQL & "subgl_cd4 = " & aarrGlBalUtil(6) & " and "
	lstrSQL = lstrSQL & "personal_ledger_cd = '" & aarrGlBalUtil(7) & "' and "
    lstrSQL = lstrSQL & "fiscal_yr = '" & lcurrFiscalYr  & "' and "
    lstrSQL = lstrSQL & "use_for_cd like '" & luseForCd  & "' and "
    lstrSQL = lstrSQL & "status = 'R'" 

	'Response.Write lstrSQL
	'Response.End
    
	lrsGLBalUtil.Open lstrSQL, lconEcgcDb
	if Err.number <> 0 then
		'Response.Redirect "../../../Common/Error.asp?aintCode=4&aintPage=-1&astrErrNumber="&Err.Number&"&astrErrDescription="&Err.description
		larrResult(0) = 1					'Error
		larrResult(1) = 4					'Error number
		larrResult(2) = 0					'Set last element of result array to zero
		fnGetBalanceUtil = larrResult
		exit function
	else
		'Response.Write lrsGLBalUtil.Fields(0)
		'Response.end
		larrResult(1) = lrsGLBalUtil.Fields(0)	'Total reserved amount
	end if	
	
	'Close the recordset
	lrsGLBalUtil.Close

	'Determine total amount utilised 
	lstrSQL = "Select nvl(sum(amount),0) from NEIA_gl_bal_util "
	lstrSQL = lstrSQL & "where entity_cd = '" & aarrGlBalUtil(0) & "' and "
	lstrSQL = lstrSQL & "logicalloc_cd = '" & aarrGlBalUtil(1) & "' and "
	lstrSQL = lstrSQL & "maingl_cd = " & aarrGlBalUtil(2) & " and "
	lstrSQL = lstrSQL & "subgl_cd1 = " & aarrGlBalUtil(3) & " and "
	lstrSQL = lstrSQL & "subgl_cd2 = " & aarrGlBalUtil(4) & " and "
	lstrSQL = lstrSQL & "subgl_cd3 = " & aarrGlBalUtil(5) & " and "
	lstrSQL = lstrSQL & "subgl_cd4 = " & aarrGlBalUtil(6) & " and "
	lstrSQL = lstrSQL & "personal_ledger_cd = '" & aarrGlBalUtil(7) & "' and "
    lstrSQL = lstrSQL & "fiscal_yr = '" & lcurrFiscalYr  & "' and "
    lstrSQL = lstrSQL & "use_for_cd like '" & luseForCd  & "' and "
    lstrSQL = lstrSQL & "status = 'U'" 

	'Response.Write lstrSQL
	'Response.End
    
	lrsGLBalUtil.Open lstrSQL, lconEcgcDb
	if Err.number <> 0 then
		'Response.Redirect "../../../Common/Error.asp?aintCode=4&aintPage=-1&astrErrNumber="&Err.Number&"&astrErrDescription="&Err.description
		larrResult(0) = 1					'Error
		larrResult(1) = 4					'Error number
		larrResult(2) = 0					'Set last element of result array to zero
		fnGetBalanceUtil = larrResult
		exit function
	else
		larrResult(2) = lrsGLBalUtil.Fields(0)	'Total Utilised Amount
	end if	
	
	larrResult(0) = 0						'Success
	fnGetGLBalanceUtil = larrResult
	
End Function		'Function fnGetGLBalanceUtil


'**************************************************************************************
'Function name : fnintStoreGLBalUtil
'Purpose       : This function will either create a record in NEIA_gl_bal_util to 
'                reserve a certain amount from available gl balance or will update the 
'                status of an already reserved record to utilized
'Input         : Connection object (so that the calling program can control rollback
'                of database updates happening in the function)
'                An array containing logical location and general ledger details
'Output        : Will return 0 on successful completion else an error number
'Author        : Amit C
'Date          : 31-01-2002
'**************************************************************************************
Function fnintStoreGLBalUtil(aconEcgcdb, aarrGlBalUtil)

	on error resume next

	Dim lrsGLBalUtil
	Dim lstrSQL

	Dim lcurrFiscalYr, lintRecords	

	fnGetMonthAndFiscalYearByDate(date)
	lcurrFiscalYr = mstrFiscalYear
	'Response.Write lcurrFiscalYr
	'Response.End
	
	set lrsGLBalUtil = Server.CreateObject( "ADODB.Recordset" )
	

        'Check if record already exists in gl_bal_util. If yes, then update amount else insert a new
        'record.
		lstrSQL = "Select status from NEIA_gl_bal_util where "
		lstrSQL = lstrSQL & "entity_cd = '" & aarrGlBalUtil(0) & "' and "
		lstrSQL = lstrSQL & "logicalloc_cd = '" & aarrGlBalUtil(1) & "' and "
		lstrSQL = lstrSQL & "maingl_cd = " & aarrGlBalUtil(2) & " and "
		lstrSQL = lstrSQL & "subgl_cd1 = " & aarrGlBalUtil(3) & " and "
		lstrSQL = lstrSQL & "subgl_cd2 = " & aarrGlBalUtil(4) & " and "
		lstrSQL = lstrSQL & "subgl_cd3 = " & aarrGlBalUtil(5) & " and "
		lstrSQL = lstrSQL & "subgl_cd4 = " & aarrGlBalUtil(6) & " and "
		lstrSQL = lstrSQL & "personal_ledger_cd = '" & aarrGlBalUtil(7) & "' and "
		lstrSQL = lstrSQL & "fiscal_yr = '" & lcurrFiscalYr & "' and "
		lstrSQL = lstrSQL & "module_cd = '" & aarrGlBalUtil(8) & "' and "
		lstrSQL = lstrSQL & "use_for_cd = '" & aarrGlBalUtil(9) & "'"
'		Response.write lstrSQL
'		Response.end
		
		lrsGLBalUtil.Open lstrSQL, aconEcgcDb
		
		if lrsGLBalUtil.EOF then 

			'Create new record in NEIA_gl_bal_util
			lstrSQL = "Insert into NEIA_gl_bal_util (entity_cd, logicalloc_cd, maingl_cd, subgl_cd1, subgl_cd2, subgl_cd3, subgl_cd4, personal_ledger_cd, "
			lstrSQL = lstrSQL & "fiscal_yr, module_cd, use_for_cd, amount, status, remarks, user_id, last_trans_date) "
			lstrSQL = lstrSQL & "values('" & aarrGlBalUtil(0) & "', "
			lstrSQL = lstrSQL & "'" & aarrGlBalUtil(1) & "', "
			lstrSQL = lstrSQL & aarrGlBalUtil(2) & ", "
			lstrSQL = lstrSQL & aarrGlBalUtil(3) & ", "
			lstrSQL = lstrSQL & aarrGlBalUtil(4) & ", "
			lstrSQL = lstrSQL & aarrGlBalUtil(5) & ", "
			lstrSQL = lstrSQL & aarrGlBalUtil(6) & ", "
			lstrSQL = lstrSQL & "'" & aarrGlBalUtil(7) & "', "
			lstrSQL = lstrSQL & "'" & lcurrFiscalYr & "', "
			lstrSQL = lstrSQL & "'" & aarrGlBalUtil(8) & "', "
			lstrSQL = lstrSQL & "'" & aarrGlBalUtil(9) & "', "
			if isnull(aarrGlBalUtil(10)) or aarrGlBalUtil(10)="" then 
				lstrSQL = lstrSQL & "'0', "
			else
				lstrSQL = lstrSQL & "'" & aarrGlBalUtil(10) & "', "
			end if			
			lstrSQL = lstrSQL & "'R', " 
			lstrSQL = lstrSQL & "'" & aarrGlBalUtil(12) & "', "
			lstrSQL = lstrSQL & "'" & aarrGlBalUtil(13) & "', "
			lstrSQL = lstrSQL & "sysdate"
			lstrSQL = lstrSQL & ")"			

			'Response.Write lstrSQL
			'Response.End
		
			aconEcgcDb.Execute lstrSQL

			if Err.number <> 0 then
				'Response.Redirect "../../../Common/Error.asp?aintCode=4&aintPage=-1&astrErrNumber="&Err.Number&"&astrErrDescription="&Err.description
				fnintStoreGLBalUtil = 4				'Return error
				exit function
			end if	
			
		else

			'Check whether GL Balance needs to be reserved or whether already reserved balance 
			'needs to be marked as utilized
			if aarrGlBalUtil(11) = "R" then
					'If status of record is already utilized then return an error.
					if ucase(lrsGLBalUtil.Fields("status")) = "U" then
						fnintStoreGLBalUtil = 4543
						exit function
					end if
			
					'Update amount reserved information
					lstrSQL = "Update NEIA_gl_bal_util "
					lstrSQL = lstrSQL & "set amount = " & aarrGlBalUtil(10) & ", "
					lstrSQL = lstrSQL & "user_id = '" & aarrGlBalUtil(13) & "', "
					lstrSQL = lstrSQL & "last_trans_date = sysdate "
					lstrSQL = lstrSQL & "where entity_cd = '" & aarrGlBalUtil(0) & "' and "
					lstrSQL = lstrSQL & "logicalloc_cd = '" & aarrGlBalUtil(1) & "' and "
					lstrSQL = lstrSQL & "maingl_cd = " & aarrGlBalUtil(2) & " and "
					lstrSQL = lstrSQL & "subgl_cd1 = " & aarrGlBalUtil(3) & " and "
					lstrSQL = lstrSQL & "subgl_cd2 = " & aarrGlBalUtil(4) & " and "
					lstrSQL = lstrSQL & "subgl_cd3 = " & aarrGlBalUtil(5) & " and "
					lstrSQL = lstrSQL & "subgl_cd4 = " & aarrGlBalUtil(6) & " and "
					lstrSQL = lstrSQL & "personal_ledger_cd = '" & aarrGlBalUtil(7) & "' and "
					lstrSQL = lstrSQL & "fiscal_yr = '" & lcurrFiscalYr  & "' and "
					lstrSQL = lstrSQL & "module_cd = '" & aarrGlBalUtil(8) & "' and "
					lstrSQL = lstrSQL & "use_for_cd = '" & aarrGlBalUtil(9) & "' and "
					lstrSQL = lstrSQL & "status = 'R'" 

					'Response.Write lstrSQL
					'Response.End
    
					aconEcgcDb.Execute lstrSQL, lintRecords

					if Err.number <> 0 then
						'Response.Redirect "../../../Common/Error.asp?aintCode=4&aintPage=-1&astrErrNumber="&Err.Number&"&astrErrDescription="&Err.description
						fnintStoreGLBalUtil = 4
						exit function
					end if	
		
					if lintRecords <> 1 then
						'Response.Redirect "../../../Common/Error.asp?aintCode=4&aintPage=-1&astrErrNumber="&Err.Number&"&astrErrDescription="&Err.description
						fnintStoreGLBalUtil = 4539
						exit function
					end if

			else
	
					'Update reserved amount status to utilized
					lstrSQL = "Update NEIA_gl_bal_util "
					lstrSQL = lstrSQL & "set amount = " & aarrGlBalUtil(10) & ", "
					lstrSQL = lstrSQL & "status = 'U', " 
					lstrSQL = lstrSQL & "remarks = '" & aarrGlBalUtil(12) & "', "
					lstrSQL = lstrSQL & "user_id = '" & aarrGlBalUtil(13) & "', "
					lstrSQL = lstrSQL & "last_trans_date = sysdate "
					lstrSQL = lstrSQL & "where entity_cd = '" & aarrGlBalUtil(0) & "' and "
					lstrSQL = lstrSQL & "logicalloc_cd = '" & aarrGlBalUtil(1) & "' and "
					lstrSQL = lstrSQL & "maingl_cd = " & aarrGlBalUtil(2) & " and "
					lstrSQL = lstrSQL & "subgl_cd1 = " & aarrGlBalUtil(3) & " and "
					lstrSQL = lstrSQL & "subgl_cd2 = " & aarrGlBalUtil(4) & " and "
					lstrSQL = lstrSQL & "subgl_cd3 = " & aarrGlBalUtil(5) & " and "
					lstrSQL = lstrSQL & "subgl_cd4 = " & aarrGlBalUtil(6) & " and "
					lstrSQL = lstrSQL & "personal_ledger_cd = '" & aarrGlBalUtil(7) & "' and "
					lstrSQL = lstrSQL & "fiscal_yr = '" & lcurrFiscalYr  & "' and "
					lstrSQL = lstrSQL & "module_cd = '" & aarrGlBalUtil(8) & "' and "
					lstrSQL = lstrSQL & "use_for_cd = '" & aarrGlBalUtil(9) & "' and "
					lstrSQL = lstrSQL & "status = 'R'" 

					'Response.Write lstrSQL
					'Response.End
    
					aconEcgcDb.Execute lstrSQL, lintRecords
					if Err.number <> 0 then
						'Response.Redirect "../../../Common/Error.asp?aintCode=4&aintPage=-1&astrErrNumber="&Err.Number&"&astrErrDescription="&Err.description
						fnintStoreGLBalUtil = 4
						exit function
					end if	
					
					if lintRecords <> 1 then
						'Response.Redirect "../../../Common/Error.asp?aintCode=4&aintPage=-1&astrErrNumber="&Err.Number&"&astrErrDescription="&Err.description
						fnintStoreGLBalUtil = 4539
						exit function
					end if

			end if

		end if
		
	fnintStoreGLBalUtil = 0				'Success
	
End Function	'Function fnintStoreGLBalUtil


'**********************************************************
'Function name : fnNEIAGetPolicyDtls
'Purpose       : This will get Policy Details
'Input         : Connection Object and Policy No
'Output        : Policy Details in array
'Author        : Milind Khedaskar
'Date          : 22-03-2002
'**********************************************************
Function fnNEIAGetPolicyDtls(aconEcgcDb, astrPolNo)
'POLICY CHANGES
'Function fnNEIAGetPolicyDtls(aconEcgcDb, astrPolNo, astrPolType)
	
	on error resume next
	
	fnNEIAGetPolicyDtls = ""
	
	Dim conCommon
	Dim rsCommon
	Dim strSQL
	
	Set rsCommon = Server.CreateObject ("ADODB.Recordset")
	
	strSQL = "Select policy_no " 
	strSQL = strSQL & "from VIEW_POLICY_NO "
	strSQL = strSQL & "where policy_no = '" & astrPolNo & "' "
	'POLICY CHANGES
	'strSQL = strSQL & "and pol_type = '" & astrPolType & "' "
		
	'Response.Write strSQL
	rsCommon.Open strSQL, aconEcgcDb
	if Err.number <> 0 then
		Response.Redirect "../../../Common/Error.asp?aintCode=4&aintPage=-1&astrErrNumber="&Err.Number&"&astrErrDescription="&Err.description
	end if	
	
	if not rsCommon.EOF then
	
		fnNEIAGetPolicyDtls = rsCommon.GetRows()
	
	else
		
		Response.Redirect "../../../Common/Error.asp?aintCode=2&aintPage=-1&astrErrNumber="&Err.Number&"&astrErrDescription=Policy record does not exist."
	
	end if
	
	rsCommon.Close
	set rsCommon = nothing

End Function


'**********************************************************
'Function name : fnNEIAGetINGDtls
'Purpose       : This will get Guarantee Details
'Input         : Connection Object, Gtee type and Gtee No.
'Output        : Gtee no. in array
'Author        : Milind Khedaskar
'Date          : 22-03-2002
'**********************************************************
Function fnNEIAGetINGDtls(aconEcgcDb, astrGteeType, aintGteeNo)

	on error resume next

	fnNEIAGetINGDtls = ""
	
	strINGGtee = "Select GTEE_NO "
	
	Select Case ucase(astrGteeType)
		case "INPCG", "INPSG"
			strINGGtee = strINGGtee & "from INPCG_INPSG_REC "
		case "INEFG", "INENEIAG"
			strINGGtee = strINGGtee & "from INEFG_INENEIAG_REC "
		case "INEPG"
			strINGGtee = strINGGtee & "from INEPG_REC "
		case  "WTEPG"
			strINGGtee = strINGGtee & "from WTEPG_REC "
		case "INTG"
			strINGGtee = strINGGtee & "from INTG_REC "
		case "BIPCG"
			strINGGtee = strINGGtee & "from BIPCG_REC_HDR "	
	End Select
	
	strINGGtee = strINGGtee & "where GTEE_NO='" & aintGteeNo & "' "
	
	Set rsINGGtee = Server.CreateObject ("ADODB.Recordset")

	rsINGGtee.Open strINGGtee, aconEcgcDb
	if Err.number <> 0 then
		Response.Redirect "../../../Common/Error.asp?aintCode=4&aintPage=-1&astrErrNumber="&Err.Number&"&astrErrDescription="&Err.description
	end if
	
	If not rsINGGtee.EOF then
		fnNEIAGetINGDtls = rsINGGtee.GetRows()
	Else
		Response.Redirect "../../../Common/Error.asp?aintCode=2&aintPage=-1&astrErrNumber="&Err.Number&"&astrErrDescription=Guarantee record does not exist."
	End If
		
	rsINGGtee.Close 
	set rsINGGtee = nothing

End Function

'**********************************************************
'Function name : fnNEIAGetLTCPolicyDtls
'Purpose       : This will get LTC Policy Details
'Input         : Connection Object and Policy No
'Output        : LTC Policy Details in array
'Author        : Milind Khedaskar
'Date          : 03-08-2002
'**********************************************************
Function fnNEIAGetLTCPolicyDtls(aconEcgcDb, astrPolNo)
	
	fnNEIAGetLTCPolicyDtls = ""
	
	Dim conCommon
	Dim rsCommon
	Dim strSQL
	
	Set rsCommon = Server.CreateObject ("ADODB.Recordset")
	
	strSQL = "Select pol_no " 
	strSQL = strSQL & "from LTC_POL_DTL "
	strSQL = strSQL & "where UPPER(pol_no) = UPPER('" & astrPolNo & "')"
		
	
	rsCommon.Open strSQL, aconEcgcDb
	if Err.number <> 0 then
		Response.Redirect "../../../Common/Error.asp?aintCode=4&aintPage=-1&astrErrNumber="&Err.Number&"&astrErrDescription="&Err.description
	end if	
	
	if not rsCommon.EOF then
	
		fnNEIAGetLTCPolicyDtls = rsCommon.GetRows()
	
	else
	Response.Write "fssaff"
	Response.End
		Response.Redirect "../../../Common/Error.asp?aintCode=2&aintPage=-1&astrErrNumber="&Err.Number&"&astrErrDescription=LTC Policy record does not exist."
	
	end if
	
	rsCommon.Close
	set rsCommon = nothing

End Function

'**********************************************************
'Function name : fnNEIAGetLTCGteeDtls
'Purpose       : This will get LTC Gtee Details
'Input         : Connection Object and LTC Gtee No
'Output        : LTC Gtee Details in array
'Author        : Milind Khedaskar
'Date          : 03-08-2002
'**********************************************************
Function fnNEIAGetLTCGteeDtls(aconEcgcDb, astrGteeNo)
	
	on error resume next
	
	fnNEIAGetLTCGteeDtls = ""
	
	Dim conCommon
	Dim rsCommon
	Dim strSQL
	
	Set rsCommon = Server.CreateObject ("ADODB.Recordset")
	
	strSQL = "Select gtee_no " 
	strSQL = strSQL & "from LTC_GTEE_DTL "
	strSQL = strSQL & "where UPPER(gtee_no) = UPPER('" & astrGteeNo & "')"
		
	'Response.Write strSQL
	rsCommon.Open strSQL, aconEcgcDb
	if Err.number <> 0 then
		Response.Redirect "../../../Common/Error.asp?aintCode=4&aintPage=-1&astrErrNumber="&Err.Number&"&astrErrDescription="&Err.description
	end if	
	
	if not rsCommon.EOF then
	
		fnNEIAGetLTCGteeDtls = rsCommon.GetRows()
	
	else
		
		Response.Redirect "../../../Common/Error.asp?aintCode=2&aintPage=-1&astrErrNumber="&Err.Number&"&astrErrDescription=LTC Gtee record does not exist."
	
	end if
	
	rsCommon.Close
	set rsCommon = nothing

End Function

'**********************************************************
'Function name : fnNEIAGetLTCOiisDtls
'Purpose       : This will get LTC OIIS Policy Details
'Input         : Connection Object and LTC OIIS Policy No
'Output        : LTC OIIS Policy Details in array
'Author        : Milind Khedaskar
'Date          : 03-08-2002
'**********************************************************
Function fnNEIAGetLTCOiisDtls(aconEcgcDb, astrPolNo)
	
	on error resume next
	
	fnNEIAGetLTCOiisDtls = ""
	
	Dim conCommon
	Dim rsCommon
	Dim strSQL
	
	Set rsCommon = Server.CreateObject ("ADODB.Recordset")
	
	strSQL = "Select pol_no " 
	strSQL = strSQL & "from LTC_OIIS_POL_DTL "
	strSQL = strSQL & "where UPPER(pol_no) = UPPER('" & astrPolNo & "')"
		
	'Response.Write strSQL
	rsCommon.Open strSQL, aconEcgcDb
	if Err.number <> 0 then
		Response.Redirect "../../../Common/Error.asp?aintCode=4&aintPage=-1&astrErrNumber="&Err.Number&"&astrErrDescription="&Err.description
	end if	
	
	if not rsCommon.EOF then
	
		fnNEIAGetLTCOiisDtls = rsCommon.GetRows()
	
	else
		
		Response.Redirect "../../../Common/Error.asp?aintCode=2&aintPage=-1&astrErrNumber="&Err.Number&"&astrErrDescription=LTC OIIS Policy record does not exist."
	
	end if
	
	rsCommon.Close
	set rsCommon = nothing

End Function

'**********************************************************
'Function name : fnNEIAGetLTCBclocDtls
'Purpose       : This will get BCLOC Policy Details
'Input         : Connection Object, BCLOC Type and BCLOC Policy No
'Output        : BCLOC Policy Details in array
'Author        : Milind Khedaskar
'Date          : 03-08-2002
'**********************************************************
Function fnNEIAGetLTCBclocDtls(aconEcgcDb, astrPolType, astrPolNo)
	
	on error resume next
	
	fnNEIAGetLTCBclocDtls = ""
	
	Dim conCommon
	Dim rsCommon
	Dim strSQL
	
	Set rsCommon = Server.CreateObject ("ADODB.Recordset")
	
	strSQL = "Select pol_no, pol_type " 
	strSQL = strSQL & "from LTC_BCLOC_POL_DTL "
	strSQL = strSQL & "where UPPER(pol_no) = UPPER('" & astrPolNo & "') "
	strSQL = strSQL & "and UPPER(pol_type) = UPPER('" & astrPolType & "') "
		
	'Response.Write strSQL
	rsCommon.Open strSQL, aconEcgcDb
	if Err.number <> 0 then
		Response.Redirect "../../../Common/Error.asp?aintCode=4&aintPage=-1&astrErrNumber="&Err.Number&"&astrErrDescription="&Err.description
	end if	
	
	if not rsCommon.EOF then
	
		fnNEIAGetLTCBclocDtls = rsCommon.GetRows()
	
	else
		
		Response.Redirect "../../../Common/Error.asp?aintCode=2&aintPage=-1&astrErrNumber="&Err.Number&"&astrErrDescription=LTC " & astrPolType & " Policy record does not exist."
	
	end if
	
	rsCommon.Close
	set rsCommon = nothing

End Function

'**********************************************************
'Function name : fnGetContractorDtls
'Purpose       : This will get the CONTRACT_PARTY_CD, CONTRACT_PARTY_NAME, STATUS
'Input         : contract code, location code
'Output        :
'Author        : Kaustubh Deshpande
'Date          : 05-04-2005
'**********************************************************
Function fnGetContractorDtls(intContractCd,strLocCd)

		Set conCommon = Server.CreateObject ("ADODB.Connection")
		Set rsCommon = Server.CreateObject ("ADODB.Recordset")

		conCommon.Open astrConn

		'strCommon = " Select CONTRACT_PARTY_CD, CONTRACT_PARTY_NAME, STATUS " &_
		'			" from   NEIA_CONTRACT_PARTY_MST where CONTRACT_PARTY_CD = '" & intContractCd & "'" &_
		'			" and    LOGICALLOC_CD = '" & strLocCd & "' and status = 'A'"

		strCommon= "SELECT DISTINCT A.PERSONAL_LEDGER_CD contract_party_cd,B.CONTRACT_PARTY_NAME contract_party_name ,DECODE(B.status,'I','In-active','Active')   FROM NEIA_GL_TXN_DTL A ,NEIA_CONTRACT_PARTY_MST B WHERE "
		strCommon=strCommon & " (MAINGL_cD,SUBGL_CD1,SUBGL_CD2,SUBGL_CD3,SUBGL_CD4) IN "
		strCommon= strCommon & " (SELECT MAINGL_cD,SUBGL_CD1,SUBGL_CD2,SUBGL_CD3,SUBGL_CD4 FROM NEIA_ENTITY_GL_MST WHERE UPPER(PERSONAL_LEDGER_LEVEL) LIKE '%CONTRACT%')"
		strCommon= strCommon & " AND SUBSTR(PERSONAL_LEDGER_cD,1,INSTR(A.PERSONAL_LEDGER_cD,' ')-1)=B.CONTRACT_PARTY_CD  "
		if iRequestQuery <> "" then
			strCommon = strCommon & "and A.PERSONAL_LEDGER_cd = '" & intContractCd & "'"
		end if
		strCommon = strCommon & "ORDER BY A.PERSONAL_LEDGER_CD"

		rsCommon.Open strCommon, conCommon

		if not rsCommon.EOF then
			fnGetContractorDtls = rsCommon.GetRows()
		end if

		'strContractCd = rsCommon.Fields(0)
		'strContractName = rsCommon.Fields(1)
		'strStatus = rsCommon.Fields(2)

		rsCommon.Close
		conCommon.Close

		set rsCommon = nothing
		set conCommon = nothing

End Function



%>