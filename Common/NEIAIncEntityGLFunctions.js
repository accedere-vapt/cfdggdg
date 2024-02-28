//**********************************************************
// This File Contains all Entity GL Functions
//**********************************************************

function f_bValidateGLCode (oTextBox0, oTextBox1, oTextBox2, oTextBox3, oTextBox4, bFieldType, sTextBoxName)
{
    //If the field is non-mandatory check if all values are null. If yes, return true
    if (bFieldType == 0)
    {
   
       if (gf_bCheckNull(oTextBox0) &&
           gf_bCheckNull(oTextBox1) &&
           gf_bCheckNull(oTextBox2) &&
           gf_bCheckNull(oTextBox3) &&
           gf_bCheckNull(oTextBox4)
          ) 
       {
         return true;
       }
    }
    
	if (!gf_bValidateWholeNumber(oTextBox0,1,1,sTextBoxName + " main gl code"))
	{
		return false;
	}
	
	if (!gf_bValidateWholeNumber(oTextBox1,0,1,sTextBoxName + " sub gl code 1"))
	{
		return false;
	}	
	
	if (!gf_bValidateWholeNumber(oTextBox2,0,1,sTextBoxName + " sub gl code 2"))
	{
		return false;
	}
	
	if (!gf_bValidateWholeNumber(oTextBox3,0,1,sTextBoxName + " sub gl code 3"))
	{
		return false;
	}
	
	if (!gf_bValidateWholeNumber(oTextBox4,0,1,sTextBoxName + " sub gl code 4"))
	{
		return false;
	}
	
	
	// Intermediate subledgers cannot be blank
	if (oTextBox1.value == "" &&
	    (oTextBox2.value != "" || oTextBox3.value != "" || oTextBox4.value != ""))
	{
		oTextBox1.focus();
		alert(sTextBoxName + " cannot have blank subledger codes in between.");
		return false;
	}
	
	if (oTextBox2.value == "" &&
	    (oTextBox3.value != "" || oTextBox4.value != ""))
	{
		oTextBox2.focus();
		alert(sTextBoxName + " cannot have blank subledger codes in between.");
		return false;
	}
	
	if (oTextBox3.value == "" &&
	    oTextBox4.value != "")
	{
		oTextBox3.focus();
		alert(sTextBoxName + " cannot have blank subledger codes in between.");
		return false;
	}
	
	return true;
}

function gf_bValidateWholeNumber(oTextBox,bFieldType,bNumberZero,sTextBoxName)
{

   //If the field is mandatory check for null
   if(bFieldType == 1)
   {
   
      if(gf_bCheckNull(oTextBox))
      {
   
         alert(sTextBoxName + " cannot be null");
         return(false);
      
      }
      
   }
   
   if(gf_bCheckAnySpaces(oTextBox))
   {
   
      alert(sTextBoxName + " should not contain spaces");
      return(false);
      
   }
   
   //Check if it is numeric
   if(isNaN(oTextBox.value))
   {
   
      oTextBox.focus();
      alert(sTextBoxName + " should be numeric");
      return(false);
      
   }
  
   if(bNumberZero == 0)  //Check if Number can be equal to zero but non-negative
   {
   
      if(gf_bCheckLessthanZero(oTextBox))
      {
   
         alert(sTextBoxName + " cannot be negative");
         return(false);      
      
      }
      
   }   
   else
   {
      if(gf_bCheckZeroOrNegative(oTextBox))
      {
   
         alert(sTextBoxName + " cannot be less than or equal to zero");
         return(false);      
      
      }
   
   }
   if(gf_bCheckTwoDecimalPlaces(oTextBox))
   {
   
      alert(sTextBoxName + " is in incorrect format");
      return(false);      
      
   }
   return(true);
   
}
