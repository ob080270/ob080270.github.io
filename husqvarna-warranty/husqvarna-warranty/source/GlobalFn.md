' ===========================================================================
' Module        : GlobalFn
' Description   : This module contains globally used functions and procedures
'                 for various operations across different forms and reports.
'                 It includes utility functions, data validation routines,
'                 and calculations that are frequently used within the project.
'
' Key Features  :
'                 - Data processing and validation utilities
'                 - Mathematical and statistical calculations
'                 - Date and time formatting functions
'                 - String manipulation and conversion functions
'
' Functions :
'                 - #1 fnPartNoSpc   : Processes part numbers without spaces.
'                 - #2 SndMail       : Sends an email using predefined parameters.
'                 - #3 fnSrchVIN     : Searches for a vehicle identification number (VIN).
'                 - #4 fnDfSel       : Selects a default function parameter.
'                 - #5 fnDfDescr     : Retrieves a description for a given code.
'                 - #6 fnDescrLkUp   : Looks up a description from a lookup table.
'
' Procedures :
'                 - #7 CheckClmTag   : Validates claim tag data.
'
' Developer     : Oleh Bondarenko
' Created       : 2003-04-17
' Last Updated  : 2025-02-19 by Oleh Bondarenko - Added comments for GitHub upload
' ===========================================================================
Option Compare Database
Option Explicit
Private objMsg As clMessage

' --------------------------------------------------------------------------
' Function #1   : fnPartNoSpc
' Purpose       : Removes spaces from a given part number string.
' Parameters    : strPartNo (String) : Input part number that may contain spaces.
' Returns       : String - The part number with all spaces removed.
' Behavior      :
'                   - Uses the VBA Replace function to remove all spaces.
'                   - Ensures the returned value does not contain leading, trailing, or in-between spaces.
'
' Example Usage : Debug.Print fnPartNoSpc(" 123 456 789 ")  Output: "123456789"
'
' Notes         :
'                   - Assumes input is a valid string.
'                   - If the input is Null, an empty string is returned.
' --------------------------------------------------------------------------
Public Function fnPartNoSpc(arg As String) As String
    
    If IsNull(strPartNo) Then
        fnPartNoSpc = ""
    Else
        fnPartNoSpc = Replace(strPartNo, " ", "")       ' Remove all spaces from the string
    End If

End Function

' --------------------------------------------------------------------------
' Function #2   : SndMail
' Purpose       : Generates and prepares an email notification based on
'                 the specified mail type and relevant warranty process.
' Parameters    : btMailType (Byte) : Type of email to be sent.
' Returns       : None
' Behavior      :
'                   - Uses Microsoft Outlook to create and display an email.
'                   - Populates email subject and body based on the selected type.
'                   - Sets appropriate recipients, reminders, and follow-up flags.
'                   - Supports different types of notifications:
'                     1. Sales to Warranty Admin about a new warranty start date.
'                     2. Service Consultant to Warranty Admin about a new warranty claim.
'                     3. Service Consultant to Sales about a missing motorcycle in the database.
'                     4. Warranty Admin to Service Consultant about claim registration in ARCO.
' Example Usage :
'                   Call SndMail(1) ' Sends notification about warranty start
'                   Call SndMail(3) ' Notifies sales about a missing motorcycle
' Notes         :
'                   - Requires Microsoft Outlook to be installed and configured.
'                   - Uses the clMessage class (objMsg) to manage message metadata.
'                   - Emails are displayed in Outlook before sending.
'                   - Ensure that the relevant forms (frCust, frAP, frWA) are open
'                     before calling this function to avoid runtime errors.
' --------------------------------------------------------------------------
Public Function SndMail(btMailType As Byte)
'   Declare Outlook objects:
    Dim objOutlook As Outlook.Application
    Dim objMailItem As MailItem
    Dim strBody As String
    Dim frm, sfrm As Form
    Dim CrLf, Tb As String
    
'   Define new line and tab characters:
    CrLf = Chr(10) + Chr(13)
    Tb = Chr(9)
    
'   Initialize Outlook application and message object:
    Set objOutlook = New Outlook.Application
    Set objMailItem = objOutlook.CreateItem(olMailItem)
    Set objMsg = New clMessage
    
    With objMailItem
        Select Case btMailType      ' Select email type
            Case 1
'           Sales to Warranty Admin about new warranty start date:
                Set frm = Forms!frCust!sfVhc.Form
                objMsg.pptSubj = "Warranty Activation in ARCO: " & frm!vhVIN & "   |   " & frm!vhMod & "   |   " & frm!vhWS
                ' Create a message body string:
                strBody = "Please activate the warranty for the specified HUSQVARNA motorcycle." & CrLf
                strBody = strBody & "Customer and motorcycle data have been entered into the HQ database." & CrLf
                strBody = strBody & "-----------------------" & CrLf
                strBody = strBody & "Yours faithfully," & CrLf
                strBody = strBody & objMsg.pptUsr & CrLf
                strBody = strBody & "AWT Bavaria"
                .To = "WarrAdmin@bmw.ua"
            Case 2
'           Service Consultant to Warranty Admin about a new warranty claim:
                Set frm = Forms!frAP
                Set sfrm = Forms!frAP!sfClm.Form
                objMsg.pptSubj = "New Warranty Claim(s) HUSQVARNA: " & frm!apVIN & " | Claim: " & sfrm.Tag & " | Act: " & frm!apAct
                ' Create a message body string:
                strBody = "A new warranty claim(s) has been created for the specified HUSQVARNA motorcycle: #" & sfrm.Tag & CrLf
                strBody = strBody & "Please register it in ARCO." & CrLf
                strBody = strBody & "-----------------------" & CrLf
                strBody = strBody & "Yours faithfully," & CrLf
                strBody = strBody & objMsg.pptUsr & CrLf
                strBody = strBody & "AWT Bavaria"
                .To = "WarrAdmin@bmw.ua"
                .ReminderTime = frm!apRpr - 1
                .FlagRequest = "Send to Husqvarna"
                .FlagDueBy = DateAdd("d", 2, Now())
            Case 3
'           Service Consultant to Sales about missing motorcycle in database
                Set frm = Forms!frAP
                Set sfrm = Forms!frAP!sfClm.Form
                objMsg.pptSubj = "Motorcycle Husqvarna: " & frm!apVIN & " - Not Found in Database"
                ' Create a message body string:
                strBody = "The specified HUSQVARNA motorcycle is missing from the database." & CrLf
                strBody = strBody & "Please add it to the database as soon as possible to proceed with the warranty claim." & CrLf
                strBody = strBody & "-----------------------" & CrLf
                strBody = strBody & "Yours faithfully," & CrLf
                strBody = strBody & objMsg.pptUsr & CrLf
                strBody = strBody & "AWT Bavaria"
                .To = "Sales@bmw.ua; ServAdvisor@bmw.ua"
                .ReminderTime = DateAdd("h", 2, Now())
                .FlagRequest = "Register in the database"
                .FlagDueBy = DateAdd("h", 2, Now())
            Case 4
'           Warranty Admin to Service Consultant about claim registration in ARCO:
                Set frm = Forms!frWA
                Set sfrm = Forms!frWA!sfWA_Itm.Form
                objMsg.pptSubj = "Entered in ARCO: " & frm!apVIN & " | Claim: " & frm!clNr & " | Act: " & frm!apAct
                ' Create a message body string:
                strBody = "The warranty has been registered in ARCO." & CrLf
                strBody = strBody & "-----------------------" & CrLf
                strBody = strBody & "Best regards," & CrLf
                strBody = strBody & objMsg.pptUsr & CrLf
                strBody = strBody & "AWT Bavaria"
                .To = "WarrMain@bmw.ua"
        End Select
        .CC = "ServAdvisor@bmw.ua"
        .Subject = objMsg.pptSubj
        .Body = strBody
        .FlagIcon = olRedFlagIcon
        .ReminderSet = True
        .Display                    ' Displays the email for review before sending
    End With
    
'   Cleanup:
    Set objMailItem = Nothing
    Set objOutlook = Nothing
    Set objMsg = Nothing

End Function

' ------------------------------------------------------------------------------------------------------
' Function #3   : fnSrchVIN
' Purpose       : Searches for a specified VIN (Vehicle Identification Number) in the vehicle database.
' Parameters    : strVIN (String) : The VIN to be searched in the database.
' Returns       : Boolean - True if the VIN is found, False otherwise.
' Behavior      :
'                   - Uses the DLookup function to check if the VIN exists in the "tblVhc" table.
'                   - If a matching VIN is found, the function returns True.
'                   - If no match is found, it returns False.
' Example Usage :
'                   If fnSrchVIN("WB101010XJ1234567") Then
'                       MsgBox "VIN found in database."
'                   Else
'                       MsgBox "VIN not found."
'                   End If
' Notes         :
'                   - Assumes that the "tblVhc" table contains a field named "vhVIN".
'                   - The input VIN must be formatted correctly (e.g., 17-character standard format).
'                   - The function does not handle case sensitivity in VIN comparison.
'                   - Returns Null if an error occurs, so additional error handling may be required.
' ------------------------------------------------------------------------------------------------------
Public Function fnSrchVIN(strVIN As String) As Boolean
    Dim fndVIN As String
    
    fndVIN = DLookup("[vhVIN]", "tblVhc", "[vhVIN] = '" & strVIN & "'")

End Function

' --------------------------------------------------------------------------
' Function #4   : fnDfSel
' Purpose       : Controls the value returned by the defect code selection form.
' Returns       : String - The selected defect code, or an empty string if no valid selection is made.
' Behavior      :
'                   - Checks the concatenation of two list selections (lstSbGr & lstSmp).
'                   - If the concatenated length is less than 7 characters, no selection is considered.
'                   - Ensures that the first character of the selected subgroup list matches
'                     the tree view value (txtTreeValue).
'                   - If the selection is invalid, an empty string is returned.
'                   - Special condition: If txtTreeValue equals "î", the selection is still considered valid.
' Example Usage :
'                   Dim selectedDf As String
'                   selectedDf = fnDfSel()
'                   If selectedDf <> "" Then
'                       MsgBox "Selected Defect Code: " & selectedDf
'                   Else
'                       MsgBox "No valid selection made."
'                   End If
' Notes         :
'                   - Assumes that the defect selection form (frDfSelect) is open when calling the function.
'                   - Uses Nz to handle potential null values when concatenating selections.
'                   - The logic ensures that only relevant defect codes are returned based on tree view filtering.
'                   - The function does not perform additional validation on the final code format.
' --------------------------------------------------------------------------
Public Function fnDfSel() As String
    Dim frm As Form
    
    Set frm = Forms!frDfSelect                                  ' Set reference to the defect selection form
    
    If Nz(Len(frm!lstSbGr & frm!lstSmp)) < 7 Then               ' If the concatenated length of subgroup and sample list < 7,
        fnDfSel = ""                                            ' - no valid selection is made
    Else
        If Left(frm!lstSbGr, 1) = frm!txtTreeValue Then         ' Validate that the first character of the selected subgroup matches the tree value
            fnDfSel = frm!lstSbGr & frm!lstSmp                  ' - Valid selection
        Else
            If frm!txtTreeValue = "î" Then                      ' Special case: If tree value equals "î",
                fnDfSel = frm!lstSbGr & frm!lstSmp              ' - selection is valid
            Else
                fnDfSel = ""                                    ' - Invalid selection
            End If
        End If
    End If

End Function

' --------------------------------------------------------------------------
' Function #5   : fnDfDescr
' Purpose       : Returns the description of a defect code based on the concatenation of subgroup and symptom descriptions.
' Parameters    :
'                   - strSbGr (String) : Subgroup code.
'                   - strSmp (String)  : Symptom code.
' Returns       : String - A formatted description combining the subgroup and symptom descriptions.
' Behavior      :
'                   - Uses DLookup to retrieve the subgroup description from "qryDescrLkUp".
'                   - Uses DLookup to retrieve the symptom description from "tblSymptm".
'                   - Concatenates both descriptions with a " - " separator.
' Example Usage :
'                   Dim defectDescription As String
'                   defectDescription = fnDfDescr("A1", "05")
'                   MsgBox "Defect Description: " & defectDescription
' Notes         :
'                   - Assumes that "qryDescrLkUp" contains a field "sgrNmRu" for subgroup names.
'                   - Assumes that "tblSymptm" contains a field "smpNmRu" for symptom descriptions.
'                   - If either lookup fails, the function may return "Null - Null".
'                   - Ensure that valid subgroup and symptom codes are provided as input.
' --------------------------------------------------------------------------
Public Function fnDfDescr(strSbGr As String, strSmp As String) As String
    Dim strSgpD As String   ' Subgroup description
    Dim strSmpD As String   ' Symptom description

'       Retrieve descriptions from respective tables:
        strSgpD = DLookup("[sgrNmRu]", "qryDescrLkUp", "[Code] = '" & strSbGr & "'")
        strSmpD = DLookup("[smpNmRu]", "tblSymptm", "[smpID] = '" & strSmp & "'")

'   Concatenate the descriptions:
    fnDfDescr = strSgpD & " - " & strSmpD
    
End Function

' --------------------------------------------------------------------------
' Function #6   : fnDescrLkUp
' Purpose       : Populates the description field based on the LookUp field after an identifier field is updated.
' Parameters    : frm (Form) : The form where the update event occurs.
' Returns       : None
' Behavior      :
'                   - If the LookUp description field (luDescr) is not empty, it is copied to the clTechn field.
'                   - If the function is triggered in the "sfItm" subform:
'                     - It assigns the claim number from the parent form.
'                     - Determines whether the item is a part ("P") or a labor item ("W").
'                     - If the item is labor ("W"), it links it to the corresponding claim.
'                   - If the item number starts with "DEBIT", it was originally intended to set warranty-related fields, but this is commented out.
'                   - If the form is not "sfItm", it simply assigns luDescr to clTechn and exits.
' Error Handling:
'                   - Catches runtime errors and displays an error message with the error number.
'                   - Uses "Resume Next" to continue execution after displaying the error.
' Example Usage :
'                   Call fnDescrLkUp(Me)
' Notes         :
'                   - This function is triggered by an AfterUpdate event on an identifier field.
'                   - Assumes that the fields luDescr, clTechn, itClm, Type, and itChrg exist
'                     in the form structure.
'                   - The commented-out section related to "DEBIT" might need further review
'                     if it should be enabled again.
' --------------------------------------------------------------------------
Public Function fnDescrLkUp(frm As Form)
' Called by the AfterUpdate event of an identifier field
' Populates the description field with the LookUp field value
On Error GoTo ErrorHandler                                          ' If a run period error occurs
                                                                    '  - match part-work-type before exiting the procedure
    If IsNull(frm!luDescr) = False Then                             ' If the LookUp description is not null,
        frm!clTechn = frm!luDescr                                   '  - copy it to clTechn
    End If
       
'   If the function is executed in the "sfItm" subform:
    If frm.Name = "sfItm" Then
'       Assign claim number from the parent form to the subform field:
        'frm!sfClNo = frm.Parent("clNo")
        frm!itClm = frm.Parent("sfClm")!clNr
'       Determine whether the item is a part or a labor item:
        Select Case Len(frm!itNr)
            Case 9
                frm!Type = "P"                      ' - Part
            Case Else
                frm!Type = "W"                      ' - Labor item
                frm!itChrg = frm.Parent("clNo")
        End Select
        
        ' Commented-out logic for handling "DEBIT" items
        ' If needed, this section should be uncommented and tested.
        ' If Left(Screen.ActiveControl, 5) = "DEBIT" Then
        '     Forms!frClaims!sf2Parts.Form!ckAWTwrnt = True
        '     Forms!frClaims!sf2Parts.Form!intAW = 1
        '     If IsNull(Forms!frClaims!sf2Parts.Form!wkDescr) Then
        '         Forms!frClaims!sf2Parts.Form!wkDescr = "DEBIT AUDIT 2011"
        '     End If
        ' End If
        
    Else                                            ' - For all other forms, just assign luDescr to clTechn and exit
        frm!clTechn = frm!luDescr
        Exit Function
    End If
    
ErrorHandler:
    ' Display an error message and continue execution:
    MsgBox "Error #" & Err.Number & " : " & Err.Description, vbExclamation, "Application Error"
    Resume Next

End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure #7  : CheckClmTag
' Purpose       : Checks the subform descriptor for claims to verify whether a claim number has been created.
' Parameters    : None
' Returns       : None (Sub procedure)
' Behavior      :
'                   - Accesses the subform "sfClm" within the main form "frAP".
'                   - If the Tag property of the subform is not empty:
'                     - Calls the SndMail function with parameter 2 to notify the Warranty Administrator about new warranty claims.
'                     - Clears the Tag property after sending the notification.
' Example Usage :
'                   Call CheckClmTag
' Notes         :
'                   - The Tag property is used as a temporary marker to track whether a claim number exists.
'                   - Assumes that "frAP" is open and contains the "sfClm" subform.
'                   - Relies on the SndMail function to handle email notification.
'                   - The commented-out Recordset code suggests potential enhancement for database validation.
' -----------------------------------------------------------------------------------------------------------------------
Public Sub CheckClmTag()
    Dim frm As Form
    Set frm = [Forms]![frAP]![sfClm].[Form]

    If frm.Tag <> "" Then                       ' If the descriptor (Tag property) is not empty
        Call SndMail(2)                         ' Send an email notification to the Warranty Administrator about new claims
        frm.Tag = ""                            ' Clear the descriptor after sending the notification
    End If
    
End Sub
