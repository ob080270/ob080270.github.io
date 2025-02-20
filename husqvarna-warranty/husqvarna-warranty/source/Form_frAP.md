' ===========================================================================
' Description  : Form module for managing Accounts Payable (AP) operations.
' Child Form   :
'                - sfClm: Claim - Repair
'                - sfItm: Work and Parts
'
' Key Features :
' - Handles AP data management.
' - Supports filtering and sorting of AP records.
' - Implements data validation and automation for AP transactions.
'
' Events :
' 1. apVIN_AfterUpdate        - Handles VIN updates and searches for motorcycles.
'                                - Checks if the entered VIN exists in the database.
'                                - If VIN is not found, triggers an email notification.
'                                - Updates related fields (vhMod, vhWS) with new data.
' 2. cbClmFnd_AfterUpdate     - Updates the claim record corresponding to the entered claim.
' 3. cmdNewRec_Click          - Creates a new AP record.
' 4. Form_Open                - Sets the form's editing mode.
' 5. sfClm_Enter              - Handles focus entering the sfClm subform.
' 6. sfClm_Exit               - Handles focus exiting the sfClm subform.
' 7. sfItm_Enter              - Handles focus entering the sfItm subform.
' 8. swEdit_AfterUpdate       - Updates the editing mode of the form.
' 9. cmdExitApp_Click         - Closes the application.
'
' Developer    : Oleh Bondarenko
' Created      : 2012-06-12
' Last Updated : 2025-02-19 by Oleh Bondarenko - Added comments for GitHub upload
' ===========================================================================
Option Explicit
Option Compare Database
' ---------------------------------------------------------------------------
' Event #1        : apVIN_AfterUpdate
' Purpose         : Handles the update of the VIN field.
' Behavior        :
'                  - Checks if the entered VIN exists in the database (tblVhc).
'                  - If the VIN is not found, triggers an email notification.
'                  - Updates the related fields (vhMod, vhWS) with the latest data.
' External Calls  :
'                  - - SndMail() (Module: GlobalFn) - Sends an email notification if the VIN is missing.
' ---------------------------------------------------------------------------
Private Sub apVIN_AfterUpdate()

    Dim frm As Form         ' Reference to the current form
    
    Set frm = Forms!frAP
'   Check if the entered VIN exists in the database
    If IsNull(DLookup("[vhVIN]", "tblVhc", "[vhVIN] = '" & frm!apVIN & "'")) Then
'       If VIN is not found, send an email notification to the responsible party:
'       Service Consultant to Sales about a missing motorcycle in the database (see module GlobalFn)
        Call SndMail(3)
    End If

'   Refresh dependent fields to reflect the latest data
    Me!vhMod.Requery
    Me!vhWS.Requery

End Sub

' ---------------------------------------------------------------------------
' Event #2        : cbClmFnd_AfterUpdate
' Purpose         : Searches an invoice that corresponds to the specified claim number.
' Behavior        :
'                  - Retrieves the claim number from the dataset (query "qsCbClmFnd").
'                  - Searches for the corresponding claim record in the form's dataset.
'                  - Moves the form to the found claim record, if it exists.
'                  - Uses DLookup to retrieve the claim number.
'                  - Uses RecordsetClone and FindLast to locate the claim record.
' ---------------------------------------------------------------------------
Private Sub cbClmFnd_AfterUpdate()
    Dim fndAct As Long                              ' - Variable to store the found invoice number
    Dim rst As Recordset                            ' - Recordset for navigating form data
    
    fndAct = Nz(DLookup("[apAct]", "qsCbClmFnd"))   ' - Retrieve the claim number from the dataset
    
    Set rst = Me.RecordsetClone                     ' - Get a clone of the form's recordset
        rst.FindLast "[apAct] = " & fndAct          ' - Locate the record with the corresponding claim number
        Me.Bookmark = rst.Bookmark                  ' - Move the form to the located record
        
'   Clean up the recordset object
    rst.Close
    Set rst = Nothing
        
End Sub

' ---------------------------------------------------------------------------
' Event #3        : cmdNewRec_Click
' Purpose         : Creates a new record in the tblAP.
' Behavior        :
'                  - Calls CheckClmTag to ensure necessary claim data is valid.
'                  - Moves the form to a new blank record.
'                  - Sets focus to the primary input field (apAct) for user entry.
' External Calls  :
'                  - CheckClmTag() - Verifies claim tagging requirements before proceeding (see module GlobalFn).
' ---------------------------------------------------------------------------
Private Sub cmdNewRec_Click()
'   Ensure that all necessary claim tagging requirements are met before proceeding:
    Call CheckClmTag
    
    DoCmd.GoToRecord , , acNewRec   ' - Move to a new blank record in the form
    Me!apAct.SetFocus               ' - Set focus to the main field for entering a new AP record
    
End Sub

' ---------------------------------------------------------------------------
' Event #4        : Form_Open
' Purpose         : Initializes the form when it is opened.
' Behavior        :
'                  - Sets the editing mode based on user permissions.
'                  - Configures UI elements accordingly.
'                  - Ensures the form opens with the correct default settings.
' ---------------------------------------------------------------------------
Private Sub Form_Open(Cancel As Integer)

    swEdit = False                  ' - Disable editing mode by default to prevent unintended modifications
    cmdNewRec.Enabled = False       ' - Disable the "New Record" button (cmdNewRec)
    DoCmd.Maximize                  ' - Maximize the form to fit the screen for better visibility and user experience

End Sub

' ---------------------------------------------------------------------------
' Event #5        : sfClm_Enter
' Purpose         : Handles actions when the focus enters the sfClm subform.
' Behavior        :
'                  - Ensures correct data context when the user interacts with the claims subform.
'                  - Adjusts UI elements by modifying the background style of editable fields.
'                  - Enables visual differentiation of fields when editing mode is active.
' ---------------------------------------------------------------------------
Private Sub sfClm_Enter()
    Dim frClm As Form       ' - Reference to the sfClm subform
    
    Set frClm = sfClm.Form  ' - Assign the subform to the variable
    
    If Me!swEdit Then       ' - Check if edit mode is enabled
'       Make a transparent background for the following fields
'       in the main form:
        Me!apAct.BackStyle = 0
        Me!apOrd.BackStyle = 0
        Me!apRpr.BackStyle = 0
        Me!apMlg.BackStyle = 0
        Me!apTp.BackStyle = 0
        Me!apVIN.BackStyle = 0
        Me!apDlr.BackStyle = 0

'       Make an opaque background for the following fields
'       in the sfClm subform:
        frClm!clNr.BackStyle = 1
        frClm!clCompl.BackStyle = 1
        frClm!clDf.BackStyle = 1
        frClm!clTechn.BackStyle = 1
    End If

End Sub

' ---------------------------------------------------------------------------
' Event #6        : sfClm_Exit
' Purpose         : Handles actions when the focus leaves the sfClm subform.
' Behavior        :
'                  - Restores default UI appearance when exiting the claims subform.
'                  - Adjusts the background style of key fields to indicate they are no longer active.
' ---------------------------------------------------------------------------
Private Sub sfClm_Exit(Cancel As Integer)
    Dim frClm As Form                       ' - Reference to the sfClm subform
    
    Set frClm = sfClm.Form                  ' - Assign the subform to the variable
    
    If Me!swEdit Then                       ' - Check if edit mode is enabled
'       Restore default background style (transparent)
'       for fields in sfClm:
        frClm!clNr.BackStyle = 0
        frClm!clCompl.BackStyle = 0
        frClm!clDf.BackStyle = 0
        frClm!clTechn.BackStyle = 0
    End If

End Sub

' ---------------------------------------------------------------------------
' Event #7        : sfItm_Enter
' Purpose         : Handles actions when the focus enters the sfItm subform.
' Behavior        :
'                  - Ensures correct data context when interacting with the items subform.
'                  - Automatically sets focus to the item number field if it is empty.
'                  - Highlights editable fields if edit mode is enabled.
' ---------------------------------------------------------------------------
Private Sub sfItm_Enter()
    Dim frm As Form                     ' - Reference to the sfItm subform

    Set frm = Forms!frAP!sfItm.Form     ' - Assign the subform to the variable
    
    If IsNull(frm!itNr) Then            ' If the item number field (itNr) is empty
        frm!itNr.SetFocus               ' - move the focus to it
    End If

    If Me!swEdit Then                   ' - Check if edit mode is enabled
'       Highlight editable fields
'       in the main form:
        Me!apAct.BackStyle = 0
        Me!apOrd.BackStyle = 0
        Me!apRpr.BackStyle = 0
        Me!apMlg.BackStyle = 0
        Me!apTp.BackStyle = 0
        Me!apVIN.BackStyle = 0
        Me!apDlr.BackStyle = 0
    End If

End Sub

' ---------------------------------------------------------------------------
' Event #8        : swEdit_AfterUpdate
' Purpose         : Toggles edit mode on and off based on the swEdit checkbox.
' Behavior        :
'                  - Enables or disables editing for key fields in the main form.
'                  - Controls the ability to add and edit records in subforms.
'                  - Updates UI elements to visually indicate the current mode.
'                  - Calls CheckClmTag when switching to non-edit mode to validate claims.
' ---------------------------------------------------------------------------
Private Sub swEdit_AfterUpdate()

    Dim frClm, frIt As Form                                 ' - Declare form variables for subforms
    
'   Assign references to subforms:
    Set frClm = sfClm.Form
    Set frIt = sfItm.Form
    
    If Me!swEdit Then                                       ' - If editing mode is enabled
        Me!apAct.Locked = False                             ' - Allow editing of field
        Me!apAct.BackStyle = 1                              ' - Set background to visible
        Me!apOrd.Locked = False
        Me!apOrd.BackStyle = 1
        Me!apRpr.Locked = False
        Me!apRpr.BackStyle = 1
        Me!apMlg.Locked = False
        Me!apMlg.BackStyle = 1
        Me!apTp.Locked = False
        Me!apTp.BackStyle = 1
        Me!apVIN.Locked = False
        Me!apVIN.BackStyle = 1
        Me!apDlr.Locked = False
        Me!apDlr.BackStyle = 1

'       Enable adding and editing in the claims subform:
        frClm.AllowAdditions = True                         ' - Allow new claims to be added
        frClm.AllowEdits = True                             ' - Allow editing of existing claims

'       Unlock the items subform and enable adding records:
        sfItm.Locked = False
        frIt.AllowAdditions = True
        'frIt.AllowEdits = True
        'frIt.AllowDeletions = True
        
        Me.AllowAdditions = True                            ' - Allow adding new records in the main form
        lbEdit.Visible = True                               ' - Show the "Edit Mode" label
        cmdNewRec.Enabled = True                            ' - enable the "New Record" button
    
    Else ' If editing mode is disabled:
'       Lock key fields in the main form
'       and set their background to transparent:
        Me!apAct.Locked = True
        Me!apAct.BackStyle = 0
        Me!apOrd.Locked = True
        Me!apOrd.BackStyle = 0
        Me!apRpr.Locked = True
        Me!apRpr.BackStyle = 0
        Me!apMlg.Locked = True
        Me!apMlg.BackStyle = 0
        Me!apTp.Locked = True
        Me!apTp.BackStyle = 0
        Me!apVIN.Locked = True
        Me!apVIN.BackStyle = 0
        Me!apDlr.Locked = True
        Me!apDlr.BackStyle = 0
        
        frClm.AllowAdditions = False                        ' - Disable adding in the claims subform
        frClm.AllowEdits = False                            ' - Disable editing in the claims subform
        
        sfItm.Locked = True                                 ' - Lock the items subform
        frIt.AllowAdditions = False                         ' and disable adding records
        'frIt.AllowEdits = False
        'frIt.AllowDeletions = False
        
        Me.AllowAdditions = False                           ' - Disable adding new records in the main form
        lbEdit.Visible = False                              ' - Hide the "Edit Mode" label
        cmdNewRec.Enabled = False                           ' and disable the "New Record" button
        
        Call CheckClmTag                                    ' - Check the descriptor for ready-made claims
        
    End If
    
    apAct.SetFocus
    
End Sub

' ---------------------------------------------------------------------------
' Event #9        : cmdExitApp_Click
' Purpose         : Closes the application when the exit button is clicked.
' Behavior        :
'                  - Attempts to close the application using DoCmd.Quit.
'                  - Implements error handling to display a message if an issue occurs.
'                  - Ensures a graceful exit by handling any runtime errors.
' ---------------------------------------------------------------------------
Private Sub cmdExitApp_Click()
On Error GoTo Err_cmdExitApp_Clic

    DoCmd.Quit                      ' - Close the application

Exit_cmdExitApp_Click:
    Exit Sub                        ' - Exit the procedure safely

Err_cmdExitApp_Click:
    MsgBox Err.Description          ' - Display the error message if an issue occurs
    Resume Exit_cmdExitApp_Click    ' - Return to the exit point to safely exit the procedure
    
End Sub
