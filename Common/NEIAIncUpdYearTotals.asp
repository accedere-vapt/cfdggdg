<%

'***********************
'START update of debit/credit total for the transaction year in branch balance 
'year table

mstrSQL = "UPDATE NEIA_GL_BR_BAL_YR set "
if lstrDrCrFlg = "CR" then
	if lstrFiscalYearForTxnDt <> lstrFiscalYearForProcessing then
		mstrSQL = mstrSQL & "open_cr_bal = open_cr_bal + " & ldblTxnAmt & ", "
	end if
	mstrSQL = mstrSQL & "last_cr_bal = last_cr_bal + " & ldblTxnAmt & ", "
else
	if lstrFiscalYearForTxnDt <> lstrFiscalYearForProcessing then
		mstrSQL = mstrSQL & "open_dr_bal = open_dr_bal + " & ldblTxnAmt & ", "
	end if
	mstrSQL = mstrSQL & "last_dr_bal = last_dr_bal + " & ldblTxnAmt & ", "				
end if
mstrSQL = mstrSQL & "user_id='" & lstrUserID & "',"
mstrSQL = mstrSQL & "last_trans_date=sysdate "
mstrSQL = mstrSQL & "where logicalloc_cd='" & lstrLogicalLocCd &"' "
mstrSQL = mstrSQL & "and maingl_cd='" & lintMainGLCd & "' "
mstrSQL = mstrSQL & "and subgl_cd1='" & lintSubGLCd1 & "' "
mstrSQL = mstrSQL & "and subgl_cd2='" & lintSubGLCd2 & "' "
mstrSQL = mstrSQL & "and subgl_cd3='" & lintSubGLCd3 & "' "
mstrSQL = mstrSQL & "and subgl_cd4='" & lintSubGLCd4 & "' "
mstrSQL = mstrSQL & "and personal_ledger_cd='" & lstrPersonalLedgerCd & "' "
mstrSQL = mstrSQL & "and fiscal_yr='" & lstrFiscalYearForProcessing & "' "
mstrSQL = mstrSQL & "and entity_cd='NEIA' "

'Response.Write mstrsql
'Response.End

		
aconEcgcDb.Execute mstrSQL, lintRecords

If Err.number <> 0 then
	'Response.Write "<FONT class=clsError>fnintPassGLTxn: Branch Balance Update</FONT><BR>"
	'Response.End
	'Response.Redirect "../../../Common/Error.asp?aintCode=4&aintPage=-1&astrErrNumber="&Err.Number&"&astrErrDescription="&Err.description
	fnintPassGLTxn = -4
	exit function
end if


							
'If a record is not found, then insert a new record
If lintRecords = 0 and Err.number = 0 then

	ldblOpenCrBal = 0
	ldblOpenDrBal = 0
	ldblLastCrBal = 0
	ldblLastDrBal = 0

	if lstrDrCrFlg = "CR" then
		ldblLastCrBal = ldblTxnAmt
	else
		ldblLastDrBal = ldblTxnAmt
	end if
							
	mstrSQL = "INSERT into NEIA_GL_BR_BAL_YR (entity_cd,logicalloc_cd,maingl_cd,subgl_cd1,subgl_cd2,subgl_cd3,subgl_cd4,personal_ledger_cd,fiscal_yr,open_cr_bal,open_dr_bal,last_cr_bal,last_dr_bal,user_id,last_trans_date) "
	mstrSQL = mstrSQL & "values( "
	mstrSQL = mstrSQL & "'NEIA', "
	mstrSQL = mstrSQL & "'" & lstrLogicalLocCd &"', "
	mstrSQL = mstrSQL & "'" & lintMainGLCd & "', "
	mstrSQL = mstrSQL & "'" & lintSubGLCd1 & "', "
	mstrSQL = mstrSQL & "'" & lintSubGLCd2 & "', "
	mstrSQL = mstrSQL & "'" & lintSubGLCd3 & "', "
	mstrSQL = mstrSQL & "'" & lintSubGLCd4 & "', "
	mstrSQL = mstrSQL & "'" & lstrPersonalLedgerCd & "', "
	mstrSQL = mstrSQL & "'" & lstrFiscalYearForProcessing & "', "
	mstrSQL = mstrSQL & "'" & ldblOpenCrBal & "', "
	mstrSQL = mstrSQL & "'" & ldblOpenDrBal & "', "
	mstrSQL = mstrSQL & "'" & ldblLastCrBal & "', "
	mstrSQL = mstrSQL & "'" & ldblLastDrBal & "', "
	mstrSQL = mstrSQL & "'" & lstrUserID & "',"
	mstrSQL = mstrSQL & "sysdate "
	mstrSQL = mstrSQL & ")"


	aconEcgcDb.Execute mstrSQL

	If Err.number <> 0 then
		'Response.Write "<FONT class=clsError>fnintPassGLTxn: Branch Balance Insert</FONT><BR>"
		'Response.End
		'Response.Redirect "../../../Common/Error.asp?aintCode=4&aintPage=-1&astrErrNumber="&Err.Number&"&astrErrDescription="&Err.description
		fnintPassGLTxn = -4
		exit function
	end if
								
end if ' If lintRecords = 0 and Err.number = 0 then

'END of debit/credit total update for the transaction year in branch balance 
'year table
'***********************


'***********************
'START update of debit/credit total for the transaction year in balance summary 
'year table	
		
mstrSQL = "UPDATE NEIA_GL_BAL_SUMM_YR set "
if lstrDrCrFlg = "CR" then
	if lstrFiscalYearForTxnDt <> lstrFiscalYearForProcessing then
		mstrSQL = mstrSQL & "open_cr_bal = open_cr_bal + " & ldblTxnAmt & ", "
	end if
	mstrSQL = mstrSQL & "last_cr_bal = last_cr_bal + " & ldblTxnAmt & ", "
else
	if lstrFiscalYearForTxnDt <> lstrFiscalYearForProcessing then
		mstrSQL = mstrSQL & "open_dr_bal = open_dr_bal + " & ldblTxnAmt & ", "
	end if
	mstrSQL = mstrSQL & "last_dr_bal = last_dr_bal + " & ldblTxnAmt & ", "				
end if
mstrSQL = mstrSQL & "user_id='" & lstrUserID & "',"
mstrSQL = mstrSQL & "last_trans_date=sysdate "
mstrSQL = mstrSQL & "where maingl_cd='" & lintMainGLCd & "' "
mstrSQL = mstrSQL & "and subgl_cd1='" & lintSubGLCd1 & "' "
mstrSQL = mstrSQL & "and subgl_cd2='" & lintSubGLCd2 & "' "
mstrSQL = mstrSQL & "and subgl_cd3='" & lintSubGLCd3 & "' "
mstrSQL = mstrSQL & "and subgl_cd4='" & lintSubGLCd4 & "' "
mstrSQL = mstrSQL & "and fiscal_yr='" & lstrFiscalYearForProcessing & "' "
mstrSQL = mstrSQL & "and entity_cd='NEIA' "
'Response.Write mstrSQL & "<BR>"		
		
aconEcgcDb.Execute mstrSQL, lintRecords
		
If Err.number <> 0 then
	'Response.Write "<FONT class=clsError>fnintPassGLTxn: Branch Balance Update</FONT><BR>"
	'Response.End
	'Response.Redirect "../../../Common/Error.asp?aintCode=4&aintPage=-1&astrErrNumber="&Err.Number&"&astrErrDescription="&Err.description
	fnintPassGLTxn = -4
	exit function
end if
							
'If a record is not found, then insert a new record
If lintRecords = 0 and Err.number = 0 then

	ldblOpenCrBal = 0
	ldblOpenDrBal = 0
	ldblLastCrBal = 0
	ldblLastDrBal = 0

	if lstrDrCrFlg = "CR" then
		ldblLastCrBal = ldblTxnAmt
	else
		ldblLastDrBal = ldblTxnAmt
	end if
							
	mstrSQL = "INSERT into NEIA_GL_BAL_SUMM_YR (entity_cd,maingl_cd,subgl_cd1,subgl_cd2,subgl_cd3,subgl_cd4,fiscal_yr,open_cr_bal,open_dr_bal,last_cr_bal,last_dr_bal,user_id,last_trans_date) "
	mstrSQL = mstrSQL & "values( "
	mstrSQL = mstrSQL & "'NEIA', "
	mstrSQL = mstrSQL & "'" & lintMainGLCd & "', "
	mstrSQL = mstrSQL & "'" & lintSubGLCd1 & "', "
	mstrSQL = mstrSQL & "'" & lintSubGLCd2 & "', "
	mstrSQL = mstrSQL & "'" & lintSubGLCd3 & "', "
	mstrSQL = mstrSQL & "'" & lintSubGLCd4 & "', "
	mstrSQL = mstrSQL & "'" & lstrFiscalYearForProcessing & "', "
	mstrSQL = mstrSQL & "'" & ldblOpenCrBal & "', "
	mstrSQL = mstrSQL & "'" & ldblOpenDrBal & "', "
	mstrSQL = mstrSQL & "'" & ldblLastCrBal & "', "
	mstrSQL = mstrSQL & "'" & ldblLastDrBal & "', "
	mstrSQL = mstrSQL & "'" & lstrUserID & "',"
	mstrSQL = mstrSQL & "sysdate "
	mstrSQL = mstrSQL & ")"
	'Response.Write mstrSQL & "<BR>"		

	aconEcgcDb.Execute mstrSQL

	If Err.number <> 0 then
		'Response.Write "<FONT class=clsError>fnintPassGLTxn: Branch Balance Insert</FONT><BR>"
		'Response.End
		'Response.Redirect "../../../Common/Error.asp?aintCode=4&aintPage=-1&astrErrNumber="&Err.Number&"&astrErrDescription="&Err.description
		fnintPassGLTxn = -4
		exit function
	end if
								
end if  'If lintRecords = 0 and Err.number = 0 then


'END of debit/credit total update for the transaction year in balance summary 
'year table
'***********************

%>