<%
Dim msDgtInWords1(10)
Dim msDgtInWords2(10)
Dim msDgtInWords3(10)
Dim msPlacesInWords(6)
Dim msNumInWords

Sub Initialize()

	' Procedure to initialize all arrays

	msDgtInWords1(1) = "ONE"
	msDgtInWords1(2) = "TWO"
	msDgtInWords1(3) = "THREE"
	msDgtInWords1(4) = "FOUR"
	msDgtInWords1(5) = "FIVE"
	msDgtInWords1(6) = "SIX"
	msDgtInWords1(7) = "SEVEN"
	msDgtInWords1(8) = "EIGHT"
	msDgtInWords1(9) = "NINE"

	msDgtInWords2(1) = "ELEVEN"
	msDgtInWords2(2) = "TWELVE"
	msDgtInWords2(3) = "THIRTEEN"
	msDgtInWords2(4) = "FOURTEEN"
	msDgtInWords2(5) = "FIFTEEN"
	msDgtInWords2(6) = "SIXTEEN"
	msDgtInWords2(7) = "SEVENTEEN"
	msDgtInWords2(8) = "EIGHTEEN"
	msDgtInWords2(9) = "NINETEEN"

	msDgtInWords3(1) = "TEN"
	msDgtInWords3(2) = "TWENTY"
	msDgtInWords3(3) = "THIRTY"
	msDgtInWords3(4) = "FORTY"
	msDgtInWords3(5) = "FIFTY"
	msDgtInWords3(6) = "SIXTY"
	msDgtInWords3(7) = "SEVENTY"
	msDgtInWords3(8) = "EIGHTY"
	msDgtInWords3(9) = "NINETY"

	msPlacesInWords(1) = "TEN"
	msPlacesInWords(2) = "HUNDRED"
	msPlacesInWords(3) = "THOUSAND"
	msPlacesInWords(4) = "LAKH"
	msPlacesInWords(5) = "CRORE"

End Sub

Function ConvertToWords(number)

	' Function to convert given amount in words
	' Maximum length of Number : 11 and 2 decimal places
	' Input is number to be converted of Type Double
	' Output is the amount in words of Type String
	' If Input is a Number which cannot be handled then returns NULL

	Dim lnNumLen
	Dim lnNumPos
	Dim lnDgt
	Dim lnNumDecPlace
	Dim lnplace
	Dim lnNextDgt
	Dim lnEffNumPos
	
	Initialize

	lnNumLen = Len(CStr(number))
	
	lnNumDecPlace = InStr(1, number, ".")

	If lnNumDecPlace <> 0 Then
	    lnNumDecPlace = lnNumLen - lnNumDecPlace 'Indicates the Number of decimal places
	End If

	'If Number is of Type which cannot be handled then return NULL
	If (lnNumLen > 11 And lnNumDecPlace = 0) Or (lnNumLen > 14 And lnNumDecPlace <> 0) Or (lnNumDecPlace > 2) Then
	    msNumInWords = "NULL"
	    ConvertToWords = msNumInWords
	    Exit Function
	End If

	'If Number of decimal places is 0
	If lnNumDecPlace = 0 Then
	    msNumInWords = "AND PAISE ZERO ONLY"
	Else
	    msNumInWords = ""
	End If

	'if Number is 0 then append ZERO
	If number = 0 Then
	    msNumInWords = "RUPEES ZERO " & msNumInWords
	    ConvertToWords = msNumInWords
	    Exit Function
	End If

	'lnNumPos indicates the position of the digit which is being processed
	For lnNumPos = 1 To lnNumLen
	    
	    lnDgt = Mid(number, lnNumLen - lnNumPos + 1, 1)
	    If lnDgt <> "." Then
	        'If the number has no decimal place or 1 decimal place then add offset to position
	        'lnEffNumPos = IIf(lnNumDecPlace = 0, lnNumPos + 3, (IIf(lnNumDecPlace = 1, lnNumPos + 1, lnNumPos)))
	        if lnNumDecPlace = 0 then
				lnEffNumPos = lnNumPos + 3
	        else 
				if lnNumDecPlace = 1 then
					lnEffNumPos = lnNumPos + 1
				else
					lnEffNumPos = lnNumPos
				end if
			end if
	        
	        Select Case lnEffNumPos
	        
	        Case 1, 4, 6, 7, 9, 11, 13, 14:
	                If lnNumDecPlace = 2 And lnEffNumPos < 3 Then msNumInWords = msNumInWords & "ONLY"
	                If lnDgt <> 0 Then
	                    lnplace = 0
	                    'Append places
	                    Select Case lnEffNumPos
	                    
	                    Case 6, 13: lnplace = 2
	                    
	                    Case 7, 9, 11, 14:
	                            If lnEffNumPos = 7 Or lnEffNumPos = 14 Then lnplace = 3
	                            If lnEffNumPos = 9 Then lnplace = 4
	                            If lnEffNumPos = 11 Then lnplace = 5
	                            
	                    Case Else: lnplace = 0
	                    
	                    End Select
	                    
	                    If lnplace <> 0 Then msNumInWords = msPlacesInWords(lnplace) & " " & msNumInWords
	                    
	                    If lnNumPos < lnNumLen Then
	                        lnNextDgt = Mid(number, lnNumLen - lnNumPos, 1)
	                    Else
	                        lnNextDgt = 0
	                    End If
	                                        
	                    'If Next Digit is 1 and Effective Number positon is other than 6 or 13 then don't append
	                    If lnNextDgt <> 1 Or lnEffNumPos = 6 Or lnEffNumPos = 13 Then
	                        msNumInWords = msDgtInWords1(CInt(lnDgt)) & " " & msNumInWords
	                    End If
	                End If
	                
	                
	        Case 2, 5, 8, 10, 12:
	                If lnNumDecPlace = 1 And lnEffNumPos < 3 Then
	                    msNumInWords = msDgtInWords3(CInt(lnDgt)) & " " & "ONLY"
	                Else
	                    If lnDgt <> 0 Then
	                        lnplace = 0
	                        'Append places
	                        Select Case lnEffNumPos
	                        
	                        Case 8: If (Mid(number, lnNumLen - lnNumPos + 2, 1)) = 0 Then lnplace = 3
	                        
	                        Case 10: If (Mid(number, lnNumLen - lnNumPos + 2, 1)) = 0 Then lnplace = 4
	                                
	                        Case 12: If (Mid(number, lnNumLen - lnNumPos + 2, 1)) = 0 Then lnplace = 5
	                        
	                        Case Else: lnplace = 0
	                        
	                        End Select
	                        
	                        If lnplace <> 0 Then msNumInWords = msPlacesInWords(lnplace) & " " & msNumInWords
	                        
	                        'If Digit is 1 and Previous Digit is not 0
	                        If lnDgt = 1 And (Mid(number, lnNumLen - lnNumPos + 2, 1) <> 0) Then
	                            msNumInWords = msDgtInWords2(Mid(number, lnNumLen - lnNumPos + 2, 1)) & " " & msNumInWords
	                        Else
	                            msNumInWords = msDgtInWords3(lnDgt) & " " & msNumInWords
	                        End If
	                    End If
	                End If
	                
	                If (lnNumDecPlace = 1 Or lnNumDecPlace = 2) And lnEffNumPos < 3 Then
	                    msNumInWords = "AND PAISE " & msNumInWords
	                End If
	    
	        End Select
	    End If
	Next

	'If amount is in paise only i.e. only decimal places
	If InStr(1, number, ".") <> 0 Then
	    If CDbl(Left(number, (InStr(1, number, ".") - 1))) = 0 Then
	        msNumInWords = "ZERO " & msNumInWords
	        ConvertToWords = msNumInWords
	        Exit Function
	    End If
	End If

	msNumInWords = msNumInWords	
	ConvertToWords = msNumInWords
End Function

%>