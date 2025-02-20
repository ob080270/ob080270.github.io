' ===========================================================================
' Description   : Form module for managing customer information.
' Child Form    : None
'
' Key Features  :
' - Manages customer data display and interactions.
' - Handles form events such as opening, data updates, and search actions.
' - Implements validation and filtering mechanisms.
'
' Events        :
' 1. cmdCls_Click           - Closes the customer form.
' 2. Form_AfterUpdate       - Executes after updating a record.
' 3. Form_Open              - Executes when the form is opened.
' 4. swEdit_AfterUpdate     - Toggles the edit mode of the form.
' 5. cmdNewCust_Click       - Creates a new customer record.
' 6. cmdFndVIN_Click        - Initiates motorcycle search by VIN.
' 7. swFindMoto_AfterUpdate - Enables/disables motorcycle search by VIN or plate number.
' 8. swUsr_AfterUpdate      - Filters customers by the assigned salesperson.
' 9. txtSearchPlate_Change  - Clears the VIN search field when plate number is entered.
' 10. txtSearchVIN_Change   - Clears the plate number search field when VIN is entered.
'
' Developer    : Oleh Bondarenko
' Created      : 2011-10-17
' Last Updated : 2025-02-20 by Oleh Bondarenko - Added comments for GitHub upload
' ===========================================================================
Option Explicit
Option Compare Database
' -------------------------------------------------------------------
' Event #1        : cmdCls_Click
' Purpose         : Closes the customer form without saving changes.
' Behavior        :
' - Executes the DoCmd.Close method to close the "frCust" form.
' - Uses the acSaveNo argument to prevent automatic saving of changes.
' -------------------------------------------------------------------
Private Sub cmdCls_Click()

    DoCmd.Close acForm, "frCust", acSaveNo

End Sub

' -------------------------------------------------------------------
' Event #2        : Form_AfterUpdate
' Purpose         : Validates form fields after a record update.
' Behavior        :
' - Checks required fields (e.g., Name, Phone, Email, Address).
' - Displays a message box listing missing fields.
' - Ensures data completeness for customer records.
' -------------------------------------------------------------------
Private Sub Form_AfterUpdate()
' After editing a record - check the fields to see if they are filled in

    Dim msg As String                               ' - Message text for the MsgBox
    Dim Style As Integer                            ' - MsgBox style
    Dim Title As String                             ' - MsgBox title
    Dim Responce As Integer                         ' - User response
    Dim CrLf As String * 2                          ' - Newline characters
    
    CrLf = Chr(13) & Chr(10)                        ' - Define newline characters
    
'   Check required fields and append missing ones to the message:
    If IsNull(Me!cstFIO) Then
        msg = "Customer Name " & CrLf
    End If
    If IsNull(Me!cstTel) Then
        msg = msg & "Phone Number " & CrLf
    End If
    If IsNull(Me!cstEM) Then
        msg = msg & "e-mail " & CrLf
    End If
    If IsNull(Me!cstBD) Then
        msg = msg & "Birthday " & CrLf
    End If
    If IsNull(Me!cstAddr) Then
        msg = msg & "Address " & CrLf
    End If
    If IsNull(Me!cstJob) Then
        msg = msg & "Occupation " & CrLf
    End If
    If IsNull(Me!cstCR) Then
        msg = msg & "Purchase Influence Factor " & CrLf
    End If
    If IsNull(Me!cstPurp) Then
        msg = msg & "Purpose " & CrLf
    End If
    If IsNull(Me!cstPR) Then
        msg = msg & "Purchase Decision Factor " & CrLf
    End If
    If IsNull(Me!cstPrvBr) Then
        msg = msg & "Previous Motorcycle Brand " & CrLf
    End If
    If IsNull(Me!cstPrvMod) Then
        msg = msg & "Previous Motorcycle Model " & CrLf
    End If
    If IsNull(Me!cstPrvDS) Then
        msg = msg & "Year of Previous Motorcycle Purchase " & CrLf
    End If
    If IsNull(Me!cstPrvV) Then
        msg = msg & "Previous Motorcycle Engine Capacity " & CrLf
    End If
    If IsNull(Me!cstPrvTp) Then
        msg = msg & "Previous Motorcycle Type" & CrLf
    End If
    
    ' Show a warning message if any required fields are missing:
    Title = "Missing Required Fields:"
    Style = vbOKOnly + vbExclamation
    Responce = MsgBox(msg, Style, Title)

End Sub

' -------------------------------------------------------------------
' Event #3        : Form_Open
' Purpose         : Initializes the customer form upon opening.
' Behavior        :
' - Maximizes the form window.
' - Disables edit mode by default.
' - Disables motorcycle search by default.
' - Calls `swEdit_AfterUpdate` (Event #4) to enforce initial form settings.
' -------------------------------------------------------------------
Private Sub Form_Open(Cancel As Integer)
    
    DoCmd.Maximize              ' - Maximize the form window
    Me!swEdit = False           ' - Set the edit mode toggle to "Disabled"
    Me!swFindMoto = False       ' - Set the motorcycle search button to the “Not pressed” state
    Call swEdit_AfterUpdate     ' - Perform a set of actions related to the prohibition of form editing

End Sub

' -------------------------------------------------------------------
' Event #4        : swEdit_AfterUpdate
' Purpose         : Enables or disables edit mode for the customer form.
' Behavior        :
' - Toggles the editability of key customer fields based on `swEdit` state.
' - Changes background colors to indicate edit mode.
'     - White (16777215) > Editable
'     - Gray (12632256) > Read-only
'     - Light Gray (8421631) > Editable fields in subform
' - Enables or disables adding new records.
' - Controls visibility of the "Editing Allowed" label.
' - Adjusts subform (`sfVhc`) fields accordingly.
' -------------------------------------------------------------------
Private Sub swEdit_AfterUpdate()
'
    Dim frm As Form                             ' - Reference to the subform
    
'   Define color constants for clarity:
    Const COLOR_WHITE As Long = 16777215        ' - Editable fields
    Const COLOR_GRAY As Long = 12632256         ' - Read-only fields
    Const COLOR_LIGHT_GRAY As Long = 8421631    ' - Editable subform fields
    
    Set frm = sfVhc.Form                        ' - Set reference to the subform
    
    If Me!swEdit Then
'       Enable editing mode:
        Me!cstFIO.Locked = False
        Me!cstFIO.BackColor = COLOR_WHITE
        Me!cstTel.Locked = False
        Me!cstTel.BackColor = COLOR_WHITE
        Me!cstEM.Locked = False
        Me!cstEM.BackColor = COLOR_WHITE
        Me!cstBD.Locked = False
        Me!cstBD.BackColor = COLOR_WHITE
        Me!cstAddr.Locked = False
        Me!cstAddr.BackColor = COLOR_WHITE
        Me!cstJob.Locked = False
        Me!cstJob.BackColor = COLOR_WHITE
        Me!cstCR.Locked = False
        Me!cstCR.BackColor = COLOR_WHITE
        Me!cstPurp.Locked = False
        Me!cstPurp.BackColor = COLOR_WHITE
        Me!cstPR.Locked = False
        Me!cstPR.BackColor = COLOR_WHITE
        Me!cstPrvBr.Locked = False
        Me!cstPrvBr.BackColor = COLOR_WHITE
        Me!cstPrvMod.Locked = False
        Me!cstPrvMod.BackColor = COLOR_WHITE
        Me!cstPrvDS.Locked = False
        Me!cstPrvDS.BackColor = COLOR_WHITE
        Me!cstPrvV.Locked = False
        Me!cstPrvV.BackColor = COLOR_WHITE
        Me!cstPrvTp.Locked = False
        Me!cstPrvTp.BackColor = COLOR_WHITE
        
        Me!sfVhc.Locked = False                 ' - Enable subform editing
        frm!vhVIN.BackColor = COLOR_LIGHT_GRAY
        frm!vhPlt.BackColor = COLOR_LIGHT_GRAY
        frm!vhMod.BackColor = COLOR_LIGHT_GRAY
        frm!vhWS.BackColor = COLOR_LIGHT_GRAY
        
        Me.AllowAdditions = True                ' - Enable adding new records
        Me!cmdNewCust.Enabled = True            ' - Enable the "New Customer" button
        Me!lbEdit.Visible = True                ' - Show the "Editing Allowed" label
        Me!cstFIO.TextAlign = 1                 ' - Align customer name text to the left
    Else    ' - Disable editing mode:
        Me!cstFIO.Locked = True
        Me!cstFIO.BackColor = COLOR_GRAY
        Me!cstTel.Locked = True
        Me!cstTel.BackColor = COLOR_GRAY
        Me!cstEM.Locked = True
        Me!cstEM.BackColor = COLOR_GRAY
        Me!cstBD.Locked = True
        Me!cstBD.BackColor = COLOR_GRAY
        Me!cstAddr.Locked = True
        Me!cstAddr.BackColor = COLOR_GRAY
        Me!cstJob.Locked = True
        Me!cstJob.BackColor = COLOR_GRAY
        Me!cstCR.Locked = True
        Me!cstCR.BackColor = COLOR_GRAY
        Me!cstPurp.Locked = True
        Me!cstPurp.BackColor = COLOR_GRAY
        Me!cstPR.Locked = True
        Me!cstPR.BackColor = COLOR_GRAY
        Me!cstPrvBr.Locked = True
        Me!cstPrvBr.BackColor = COLOR_GRAY
        Me!cstPrvMod.Locked = True
        Me!cstPrvMod.BackColor = COLOR_GRAY
        Me!cstPrvDS.Locked = True
        Me!cstPrvDS.BackColor = COLOR_GRAY
        Me!cstPrvV.Locked = True
        Me!cstPrvV.BackColor = COLOR_GRAY
        Me!cstPrvTp.Locked = True
        Me!cstPrvTp.BackColor = COLOR_GRAY
        
        Me!sfVhc.Locked = True                  ' - Disable subform editing
        frm!vhVIN.BackColor = COLOR_GRAY
        frm!vhPlt.BackColor = COLOR_GRAY
        frm!vhMod.BackColor = COLOR_GRAY
        frm!vhWS.BackColor = COLOR_GRAY
        
        Me.AllowAdditions = False               ' - Disable adding new records
        Me!cmdNewCust.Enabled = False           ' - Disable the "New Customer" button
        Me!lbEdit.Visible = False               ' - Hide the "Editing Allowed" label
        Me!cstFIO.TextAlign = 2                 ' - Align customer name text to the center
    End If

End Sub

' -------------------------------------------------------------------
' Event #5        : cmdNewCust_Click
' Purpose         : Creates a new customer record.
' Behavior        :
' - Moves the form to a new record.
' - Sets focus to the "Customer Name" field (`cstFIO`) for immediate input.
' -------------------------------------------------------------------
Private Sub cmdNewCust_Click()

        DoCmd.GoToRecord , , acNewRec   ' - Navigate to a new blank record
        Me!cstFIO.SetFocus              ' - Set focus to the "Customer Name" field for input
        
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Event #6        : cmdFndVIN_Click
' Purpose         : Initiates search for a motorcycle by VIN.
' Behavior        :
' - Changes the form's record source to `qsSearchByMoto`, which executes a predefined query filtering motorcycles by VIN.
' - This action updates the displayed customer records based on the search criteria.
' -----------------------------------------------------------------------------------------------------------------------
Private Sub cmdFndVIN_Click()
' Set the form's record source to the query that filters by VIN

    Me.RecordSource = "qsSearchByMoto"
    
End Sub

' ---------------------------------------------------------------------------------------------
' Event #7        : swFindMoto_AfterUpdate
' Purpose         : Enables or disables the search for motorcycles based on VIN or plate number.
' Behavior        :
' - If `swFindMoto` is enabled:
'   - Checks if a VIN or plate number is entered.
'   - Determines whether a full or partial VIN/plate number was provided.
'   - Updates the form’s `RecordSource` accordingly.
'   - Displays a message if no matching motorcycles are found.
' - If `swFindMoto` is disabled:
'   - Resets the search fields.
'   - Restores the default record source (`tblCust`).
' ---------------------------------------------------------------------------------------------
Private Sub swFindMoto_AfterUpdate()

    If swFindMoto Then                                                          ' - If Enable search mode
        If Len(Me!txtSearchVIN) > 0 Then                                        ' - Check the VIN search field:
            Select Case Len(Me!txtSearchVIN)
                Case Is < 17                                                    '   - Partial VIN entered
                    If DCount("vhCust", "subSearchVIN2") = 0 Then               ' - If the motorcycle is not found:
                        MsgBox "No motorcycles found matching the criteria..."
                        Me!swFindMoto = False
                        Me!txtSearchVIN = ""
                        Me!txtSearchPlate = ""
                    Else                                                        ' - If the motorcycle is found:
                        Me.RecordSource = "qsSearchByMotoPartVIN"               '   - Updates the form’s `RecordSource` accordingly
                    End If
                Case 17                                                         ' - Full VIN entered
                    If DCount("vhCust", "subSearchVIN") = 0 Then                ' - If the motorcycle is not found:
                        MsgBox "No motorcycles found matching the criteria..."
                        Me!swFindMoto = False
                        Me!txtSearchVIN = ""
                        Me!txtSearchPlate = ""
                    Else                                                        ' - If the motorcycle is found:
                        Me.RecordSource = "qsSearchByMoto"                      '   - Updates the form’s `RecordSource` accordingly
                    End If
                Case Else                                                       ' - If VIN too long (error)
                    Beep                                                        '   - Inform about error with sound
                    MsgBox "VIN cannot exceed 17 characters"                    '   - Give a message
                    Me!txtSearchVIN.SetFocus                                    '   - set focus to VIN input field
                    Exit Sub                                                    '   - Exit the procedure and transfer the initiative to the user
            End Select
        End If
                                                                                ' Check the plate number search field:
        If Len(Me!txtSearchPlate) > 0 Then                                      ' - if the txtSearchPlate field is not empty
            Select Case Len(Me!txtSearchPlate)                                  '   then check its length:
                Case Is < 8
                    If DCount("vhCust", "subSearchPP") = 0 Then                 '   - Search by partial match of Plate Number
                        MsgBox "No motorcycles found matching the criteria..."
                        Me!swFindMoto = False
                        Me!txtSearchVIN = ""
                        Me!txtSearchPlate = ""
                    Else                                                        ' - If the motorcycle is found:
                        Me.RecordSource = "qsSearchByMotoPP"                    '   - Updates the form’s `RecordSource` accordingly
                    End If
                Case Else
'                   If 8 or more characters are entered,
'                   we consider it a full license plate:
                    If DCount("vhCust", "subSearchVIN") = 0 Then                ' - If the motorcycle is not found:
                        MsgBox "No motorcycles found matching the criteria..."  '   - Give a message
                        Me!swFindMoto = False
                        Me!txtSearchVIN = ""
                        Me!txtSearchPlate = ""
                    Else                                                        ' - If the motorcycle is found:
                        Me.RecordSource = "qsSearchByMoto"                      '   - Updates the form’s `RecordSource` accordingly
                    End If
            End Select
        End If

    Else    ' Disable search mode and reset fields:
        Me.RecordSource = "tblCust"
        Me!txtSearchVIN = ""
        Me!txtSearchPlate = ""
    End If
    
End Sub

' -------------------------------------------------------------------
' Event #8        : swUsr_AfterUpdate
' Purpose         : Filters customer records based on the assigned salesperson.
' Behavior        :
' - If `swUsr` is enabled:
'   - Applies a filter to show only records where the `cstUsr` field matches the current user.
' - If `swUsr` is disabled:
'   - Removes the filter to show all customers.
' -------------------------------------------------------------------
Private Sub swUsr_AfterUpdate()

    If swUsr Then                                       ' Enable filtering by salesperson
        Me.Filter = "[cstUsr] ='" & Me!cstUsr & "'"
        Me.FilterOn = True
    Else                                                ' Disable filtering and show all records
        Me.FilterOn = False
    End If
    
End Sub

' -------------------------------------------------------------------
' Event #9        : txtSearchPlate_Change
' Purpose         : Ensures mutual exclusivity between VIN and plate number search.
' Behavior        :
' - When the user enters a value in the plate number search field (`txtSearchPlate`),
'   the VIN search field (`txtSearchVIN`) is cleared automatically.
' - This prevents simultaneous filtering by both fields.
' -------------------------------------------------------------------
Private Sub txtSearchPlate_Change()
    
    Me!txtSearchVIN = ""                                ' Clear the VIN search field when a plate number is entered
    
End Sub

' -------------------------------------------------------------------
' Event #10       : txtSearchVIN_Change
' Purpose         : Ensures mutual exclusivity between VIN and plate number search.
' Behavior        :
' - When the user enters a value in the VIN search field (`txtSearchVIN`),
'   the plate number search field (`txtSearchPlate`) is cleared automatically.
' - This prevents simultaneous filtering by both fields.
' -------------------------------------------------------------------
Private Sub txtSearchVIN_Change()

    Me!txtSearchPlate = ""                              ' Clear the plate number search field when a VIN is entered
    
End Sub
