function f_bValidateFiscalYear(oTextBox,bFieldType,sTextBoxName)
{
	// Validate Fiscal Year in yyyy-yyyy format

	//If the field is mandatory, check for null
	if(bFieldType == 1)
	{
		if(gf_bCheckNull(oTextBox))
		{
			oTextBox.focus();
			alert(sTextBoxName + " cannot be null");
			return(false);
		}
	}
	else
	{
		if (oTextBox.value.length==0)
		{
			return(true);
		}
	}
	
	if(gf_bCheckAnySpaces(oTextBox))
	{
		oTextBox.focus();
		alert(sTextBoxName + " should not contain spaces");
		return(false);
	}

	var intTextBoxLength = 0;   //Length of the string
	var strTextBox = "";   //String to be checked
	var arrYear; // Array to get individual numbers
	var i; 
   
	intTextBoxLength = oTextBox.value.length;
	strTextBox = oTextBox.value;
	
	// The First character should be number
	if (isNaN(strTextBox.charAt(0)))
	{
		oTextBox.focus();
		alert(sTextBoxName + " should be in yyyy-yyyy format");
		return(false);      
	}
   
	// The Last character should be number
	if (isNaN(strTextBox.charAt(intTextBoxLength-1)))
	{
		oTextBox.focus();
		alert(sTextBoxName + " should be in yyyy-yyyy format");
		return(false);      
	}

	// check for consecutive hyphens
	i=0;
	while (i<intTextBoxLength-1)
	{
		if ((strTextBox.charAt(i) == '-') && (strTextBox.charAt(i+1) == '-'))
		{
			oTextBox.focus();
			alert(sTextBoxName + " should be in yyyy-yyyy format");
			return(false);
		}
		i++;
	}

	// Validate Fiscal Year. It should have only number & hyphen
	arrYear = strTextBox.split('-');
	
	if (arrYear.length != 2)
	// Check for only one hyphen
	{
		oTextBox.focus();
		alert(sTextBoxName + " should be in yyyy-yyyy format");
		return(false);
	}
	
	oTextBox.value = arrYear[0];
	if (! gf_bCheckInteger(oTextBox))
	{
		oTextBox.focus();
		oTextBox.value = strTextBox;
		alert(sTextBoxName + " should be in yyyy-yyyy format");
		return(false);
	}
	else if (oTextBox.value.length != 4)
	{
		oTextBox.focus();
		oTextBox.value = strTextBox;
		alert(sTextBoxName + " should be in yyyy-yyyy format");
		return(false);
	}
   
    oTextBox.value = arrYear[1];
	if (! gf_bCheckInteger(oTextBox))
	{
		oTextBox.focus();
		oTextBox.value = strTextBox;
		alert(sTextBoxName + " should be in yyyy-yyyy format");
		return(false);
	}
	else if (oTextBox.value.length != 4)
	{
		oTextBox.focus();
		oTextBox.value = strTextBox;
		alert(sTextBoxName + " should be in yyyy-yyyy format");
		return(false);
	}
	
	oTextBox.value = strTextBox;
	
	if (parseFloat(arrYear[1]) != parseFloat(arrYear[0]) + 1 )
	// Check if the next year is current year plus one
	{
		oTextBox.focus();
		alert(sTextBoxName + " should have next year as current year plus one.");
		return(false);
	}
	
	return true;
}

function fn_bValidateGLCodesAndCheckAmt()
{
	// Check the debit and credit amount
	var intRows, totAmt, intFinalRows, intMainGLCd;
	
	var objTxtBoxMainGLCd;
	var objTxtBoxSubGLCd1;
	var objTxtBoxSubGLCd2;
	var objTxtBoxSubGLCd3;
	var objTxtBoxSubGLCd4;
	
	// Total number of rows in table
	intRows = tblGLTxn.tBodies(0).rows.length;
	
	// Number of Transaction detail rows 
	// (Total number of rows in table - column header row)
	document.frmMain.hdnTxnRows.value = intRows - 1;
	
	totAmt = 0;
	
	for (i=1;i<intRows;i++)
	{
		objTxtBoxMainGLCd = eval("document.frmMain.txtMainGLCd$"+i);
		objTxtBoxSubGLCd1 = eval("document.frmMain.txtSubGLCd1$"+i);
		objTxtBoxSubGLCd2 = eval("document.frmMain.txtSubGLCd2$"+i);
		objTxtBoxSubGLCd3 = eval("document.frmMain.txtSubGLCd3$"+i);
		objTxtBoxSubGLCd4 = eval("document.frmMain.txtSubGLCd4$"+i);
			
		if (!f_bValidateGLCode(objTxtBoxMainGLCd,objTxtBoxSubGLCd1,objTxtBoxSubGLCd2,objTxtBoxSubGLCd3,objTxtBoxSubGLCd4,1,"General Ledger"))
		{
			return false;
		}
		if (!gf_bValidateAmount(eval("document.frmMain.txtAmt$" + i),1,1,"Amount"))
		{
			return false;
		}
		fltAmt = parseFloat(eval("document.frmMain.txtAmt$" + i + ".value"));
				
		if (eval("document.frmMain.selDrCrFlg$" + i + ".value") == "DR")
			totAmt = totAmt - fltAmt;
		else
			totAmt = totAmt + fltAmt;
	}
	
	
	if ( !(totAmt < 0.001 && totAmt > -0.001) )
	{
		if (totAmt < 0 )
			totAmt = -1 * totAmt;
		alert("Please check your entries.\nThere is a difference of " + totAmt + " between Debit and Credit entries");
		return (false);
	}
	return (true);
	
	//Previous commented
	//if (totAmt != 0)
	//{
	//	if (totAmt < 0 )
	//		totAmt = -1 * totAmt;
	//	alert("Please check your entries.\nThere is a difference of " + totAmt + " between Debit and Credit entries");
	//	return (false);
	//}
	//return (true);
}

// Author: Amit C 
function gfnDeleteLastRow(atblTable)
{
	// Delete last row in the table	
	
	var tblTable = atblTable;
	var intRow = tblTable.tBodies(0).rows.length;
	
	if (intRow == 1) 
		return;
	
	var strTable = tblTable.outerHTML;
	strTable = strTable.substr(1, (strTable.lastIndexOf("<TR>")) - 1);
	
	tblTable.parentElement.innerHTML =  strTable + 
									"</TBODY></TABLE>";									

}