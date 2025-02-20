' ===========================================================================
' Module       : modWarntLR
' Description  : This module contains functions and procedures for handling warranty claims.
'                It includes utilities for processing warranty transactions, generating reports, and interacting with the warranty database.
'
' Key Features :
'               - Warranty claim validation and processing.
'               - Automated report generation for warranty cases.
'               - Database interactions for warranty record management.
'
' External Dependencies :
'               - GlobalFn (global utility functions).
'               - SumProp (summary calculations for warranty claims).
'
' Functions    :
'               #1: - fnDealerCode     : Determines the dealer based on its identifier.
'               #2: - fnstrWstart      : Creates a string from month and year in "mmyy" format.
'               #3: - fnDevidePartNo   : Splits a part number into prefix, core, and suffix, returning the specified part as a string.
'
' Developer    : Oleh Bondarenko
' Created      : yyyy-mm-dd
' Last Updated : [òåêóùàÿ äàòà yyyy-mm-dd] by Oleh Bondarenko - Added comments for GitHub upload
' ===========================================================================
Option Compare Database
' ---------------------------------------------------------------------------
' Public Enum PartOfNmb
' Enumeration defining different parts of a part number.
' Used in fnDevidePartNo function to determine which part to return.
' ---------------------------------------------------------------------------
Public Enum PartOfNmb
    dpnPref
    dpnMain
    dpnSuf
End Enum
Option Explicit

' ---------------------------------------------------------------------------
' Function #1   : fnDealerCode
' Purpose       : Determines the dealer code based on the given identifier.
' Parameters    : DealerID (Byte) : The identifier of the dealer.
' Returns       : Byte - The corresponding dealer code.
' Behavior      :
'    - Maps the given DealerID to a predefined dealer code.
'    - Returns 0 if DealerID is out of range.
' ---------------------------------------------------------------------------
Public Function fnDealerCode(DealerID As Byte) As Byte

    Select Case DealerID
        Case 0                      'Dealer in Kiev
            fnDealerCode = 2
        Case 1                      'Dealer in Simferopol
            fnDealerCode = 9
        Case 2                      'Dealer in Lviv
            fnDealerCode = 7
        Case 3                      'Dealer in Kharkov
            fnDealerCode = 8
        Case 4                      'Dealer in Odessa
            fnDealerCode = 6
        Case 5                      'Dealer in Dnepr
            fnDealerCode = 5
        Case 6                      'Dealer in Donetsk
            fnDealerCode = 4
        Case 7                      'Dealer in Kremenchug
            fnDealerCode = 10
    End Select
    
'   No default case - if DealerID is out of range, function returns 0 (default)

End Function

' ---------------------------------------------------------------------------
' Function #2   : fnstrWstart
' Purpose       : Creates a formatted string "mmyy" from a given date.
' Parameters    : Wstart (Date) : The input date.
' Returns       : String - The formatted "mmyy" representation of the date.
' Behavior      :
'    - Extracts the month and year from the date.
'    - Formats them as a two-digit month followed by a two-digit year.
' ---------------------------------------------------------------------------
Public Function fnstrWstart(Wstart As Date) As String

    Dim strMonth, strYear As String
    
'       Extract month and format as two-digit string:
        strMonth = LTrim(Str(Month([Wstart])))
        strMonth = IIf(Len(strMonth) = 1, "0" & strMonth, strMonth)

        strYear = Right(Str(Year(Wstart)), 2)   ' - Extract last two digits of the year:
    
        fnstrWstart = strMonth & strYear        ' - Combine month and year to return "mmyy" format

End Function

' ---------------------------------------------------------------------------
' Function #3   : fnDevidePartNo
' Purpose       : Splits a part number into prefix, core (main part), and suffix.
' Parameters    :
'    - strPartNo (String)       : The full part number to be processed.
'    - btPartOfNm (PartOfNmb)   : Specifies which part of the number to return.
' Returns       : String - The requested part of the number.
' Behavior      :
'    - Iterates through each character in the part number.
'    - Classifies characters as either part of the prefix, main number, or suffix.
' ---------------------------------------------------------------------------
Public Function fnDevidePartNo(strPartNo As String, btPartOfNm As PartOfNmb) As String

    Dim i As Byte               ' Loop counter
    Dim strChar As String       ' Current character being processed
    Dim strPref As String       ' Prefix (letters at the start)
    Dim strMain As String       ' Main numeric part
    Dim strSuf As String        ' Suffix (letters at the end)
    
'   Initialize variables
    strPref = ""
    strMain = ""
    strSuf = ""
    
'   Loop through each character in the part number:
    For i = 1 To Len(strPartNo)
        strChar = Mid(strPartNo, i, 1)              ' - Extract current character
        
        If Asc(strChar) > 57 Then                   ' If it's a letter (not a digit):
            If i < 5 Then                           '   If position is before 5th character
                strPref = strPref & strChar         '   - Add to prefix
            Else
                strSuf = strSuf & strChar           '   - Otherwise, add to suffix
            End If
        Else                                        ' If it's a digit:
            strMain = strMain & strChar             '   - Add to main numeric part
        End If
    Next i
    
'   Return the requested part of the part number:
    Select Case btPartOfNm
        Case 0                          ' Return prefix
            fnDevidePartNo = strPref
        Case 1                          ' Return main numeric part
            fnDevidePartNo = strMain
        Case 2                          ' Return suffix
            fnDevidePartNo = strSuf
    End Select

End Function
