Sub borg_empnum_change(form As Object)
If BORG.empnum.text <> "admin832174" Then
    BORG.vis_btn.Visible = False
    BORG.invis_btn.Visible = False
Else
    BORG.vis_btn.Visible = True
    BORG.invis_btn.Visible = True
End If


End Sub


Sub borg_booGhostShow_Click(form As Object)
With form
If form.booGhostShow.Value = True Then
    form.frameGhostMaster.Visible = True
    form.labelUpdater.Top = 312
    form.Height = 396
ElseIf form.booGhostShow.Value = False Then
    form.frameGhostMaster.Visible = False
    form.labelUpdater.Top = 246
    form.Height = 330
End If
End With
'Call SaveOptions
End Sub

Sub borg_booMoreControls_Click(form As Object)
If form.booMoreControls.Value = False Then
    form.btn_CloseCan.Visible = False
    form.btn_OpenCan.Visible = False
    form.btn_UnAssign.Visible = False
    form.frame_closeScrn.Height = 60
    form.CanSelectGUI.Top = 6
    form.clscrn_refresh.Top = 6
    'form.btn_ManifestAll.Top = 30
    form.btn_ManifestOne.Top = 30
    form.btn_AddIce.Top = 30
ElseIf form.booMoreControls.Value = True Then
    'form.btn_CloseCan.Visible = True 'reconcile keeps can from closing..keeping this hidden for now.
    form.btn_OpenCan.Visible = True
    form.btn_UnAssign.Visible = True
    form.frame_closeScrn.Height = 84
    form.CanSelectGUI.Top = 30
    form.clscrn_refresh.Top = 30
    'form.btn_ManifestAll.Top = 54
    form.btn_ManifestOne.Top = 54
    form.btn_AddIce.Top = 54
End If
'Call SaveOptions
End Sub

Sub borg_btn_AddCan_Click(form As Object)
mycannum = trim(form.txt_canNum)
mysplit = form.combo_splitName
myDest = form.txt_Dest
myType = form.combo_hazType

If mycannum = "" Or mysplit = "" Or myDest = "" Or myType = "" Then
    form.labelUpdater.Caption = "ERROR: PLEASE FILL IN ALL INFORMATION BEFORE ADDING A NEW CAN"
    Exit Sub
End If

x = 2
Do Until Sheet4.Cells(x, 1) = ""
    If Sheet4.Cells(x, 1).text = mycannum Then
        If mycannum <> "bulk*" Then
            Exit Do
        Else
            If Sheet4.Cells(x, 2) = mysplit And _
               Sheet4.Cells(x, 4) = myType Then
               Exit Do
            End If
        End If
    End If
    x = x + 1
Loop

Sheet4.Cells(x, 1) = mycannum
Sheet4.Cells(x, 2) = mysplit
Sheet4.Cells(x, 3) = myDest
Sheet4.Cells(x, 4) = myType
Sheet4.Cells(x, 5) = "--"

form.txt_canNum = ""
form.combo_hazType = ""
form.combo_splitName = ""
form.txt_Dest = ""
'
Call FunctionModule.UpdateCanList

'call function saveWorkBook
Application.ActiveWorkbook.Save
form.txt_canNum.SetFocus
Dim Ctl As Control
    For Each Ctl In form.Controls
        'MsgBox TypeName(Ctl)
        Select Case TypeName(Ctl)
            Case "TextBox"
                Ctl.TabKeyBehavior = False
        End Select
    Next Ctl
End Sub

Sub borg_btn_AddIce_Click(form As Object)
If BZmodule.bz_connected = False Then
    BORG.labelUpdater.Caption = "ERROR: Login to BDG and Bluezone to use this feature"
    Exit Sub
End If
If BORG.CanSelectGUI.Value = "" Then
    AddIce ("none")
Else
    AddIce (BORG.CanSelectGUI.Value)
End If
form.clscrn_refresh_Click

End Sub

Sub borg_btn_assignCan_Click(form As Object)
If BZmodule.bz_connected = False Then
    BORG.labelUpdater.Caption = "ERROR: Login to BDG and Bluezone to use this feature"
    Exit Sub
End If

Call FlexAssign.FlexAssignDirectory(form.txt_canNum.text)
GrabCloseScreen
End Sub

Sub borg_BTN_AutoAssign_Click(form As Object)
If BZmodule.bz_connected = False Then
    BORG.labelUpdater.Caption = "ERROR: Login to BDG and Bluezone to use this feature"
    Exit Sub
End If
If Sheet4.Cells(3, 1) = "" Then
    BORG.labelUpdater.Caption = "ERROR: Please add cans to the can menu on top."
    Exit Sub
End If
Call FlexAssign.FlexAssignDirectory
GrabCloseScreen

End Sub

Sub borg_btn_cancheck_Click(form As Object)
If BZmodule.bz_connected = False Then
    BORG.labelUpdater.Caption = "ERROR: Login to BDG and Bluezone to use this feature"
    Exit Sub
End If
form.Hide

famislogingui.empnum = BORG.empnum

famis.famislogin
form.Show
form.CanSelectGUI.Value = ""
Call DGscreenChooser("close")

GrabCloseScreen
End Sub

Sub borg_btn_clearCans_Click(form As Object)
'delete cans
Sheet4.Range("A3:E999").Delete xlUp
End Sub

Sub borg_btn_CloseCan_Click(form As Object)
If BZmodule.bz_connected = False Then
    BORG.labelUpdater.Caption = "ERROR: Login to BDG and Bluezone to use this feature"
    Exit Sub
End If
CloseCan (form.CanSelectGUI.Value)
Call form.clscrn_refresh_Click
End Sub

Sub borg_btn_login_Click(form As Object)
If form.empnum.text = "" Or form.PasswordBox = "" Then
    MsgBox ("Please enter your FedEx ID and IMS password")
    Exit Sub
End If
Application.Visible = True
Application.Visible = False

Call OpenBlueZone.BZOpenSession

End Sub

Sub borg_btn_ManifestAll_Click(form As Object)
If BZmodule.bz_connected = False Then
    BORG.labelUpdater.Caption = "ERROR: Login to BDG and Bluezone to use this feature"
    Exit Sub
End If
If Sheet4.Cells(3, 1) = "" Then
    form.labelUpdater.Caption = "ERROR: No cans set up in BestDG. Set up your cans in the can menu above to print multiple manifests."
    Exit Sub
End If

Dim datarow As Integer
Dim cannum As String
Dim candest As String
Dim haztype As String
Dim excelrow As Integer

datarow = 3
Do While Sheet4.Cells(datarow, 1) <> ""
    cannum = Sheet4.Cells(datarow, 1).text
    candest = Sheet4.Cells(datarow, 3).text
    haztype = Sheet4.Cells(datarow, 4).text
    
    Call GhostAssign.filterClear
    Call Module1.DEL
    Call Module1.SETUP
        
    excelrow = GhostAssign.GrabAssign(cannum)
    
    Call Assign023(cannum)
    Call DGscreenChooser("viewawb")

    'setup format and variables for VAWB section
    Sheet1.Columns("A:A").NumberFormat = "000000000000"
    Sheet1.Columns("C:C").NumberFormat = "0000"
    Sheet1.Columns("J:J").NumberFormat = "0.00000"

    Call OpenBlueZone.GoViewAWB(excelrow)
    
    BORG.labelUpdater.Caption = "Running Fixes"
    Call Module4.APOPfix
    
    BORG.labelUpdater.Caption = "Sorting your data..."
    Call Module4.SORT_MACRO(cannum, candest, haztype)
    
    BORG.labelUpdater.Caption = "Counting Gas"
    Call Module4.gasCount
    
    BORG.labelUpdater.Caption = "Counting Pieces"
    Call Module4.pieceCount
    
    BORG.labelUpdater.Caption = "Printing your data..."
    Call Module1.printFun
    
    Call DGscreenChooser("close")
    datarow = datarow + 1
Loop

End Sub

Sub borg_btn_ManifestOne_Click(form As Object)
If BZmodule.bz_connected = False Then
    BORG.labelUpdater.Caption = "ERROR: Login to BDG and Bluezone to use this feature"
    Exit Sub
End If
Call GhostAssign.filterClear
Call Module1.DEL
Module1.DELmanifestSheet

If BORG.CanSelectGUI.Value = "" Then
    BORG.labelUpdater.Caption = "PLEASE SELECT A CAN TO MANIFEST FROM THE CLOSED SCREEN CAN CHOOSER"
    Exit Sub
End If
'Call TheGrab(Empty, " ", 3)'old
Call Module1.SETUP
'Call GhostAssign.GrabAssigned(3, form.CanSelectGUI.text)
Dim excelrow As Integer
excelrow = GhostAssign.GrabAssign(form.CanSelectGUI.text)
Call Assign023(form.CanSelectGUI.text)
Call DGscreenChooser("viewawb")

'setup format and variables for VAWB section
Sheet1.Columns("A:A").NumberFormat = "000000000000"
Sheet1.Columns("C:C").NumberFormat = "0000"
Sheet1.Columns("J:J").NumberFormat = "0.00000"

Call OpenBlueZone.GoViewAWB(excelrow)

BORG.labelUpdater.Caption = "Running Fixes"
Call Module4.APOPfix

BORG.labelUpdater.Caption = "Sorting your data..."
Dim cannum As String
Dim candest As String
Dim haztype As String
Dim datarow As Integer
datarow = 3
Do Until Sheet3.Cells(datarow, 12) = form.CanSelectGUI.text
    'u = Trim(form.CanSelectGUI.text)
    'uu = Trim(Sheet3.Cells(datarow, 12))
    datarow = datarow + 1
Loop
cannum = Sheet3.Cells(datarow, 12).text
candest = Sheet3.Cells(datarow, 13).text
haztype = ""

Call Module4.SORT_MACRO(cannum, candest, haztype)

BORG.labelUpdater.Caption = "Counting Gas"
Call Module4.gasCount

BORG.labelUpdater.Caption = "Counting Pieces"
Call Module4.pieceCount

If BORG.PrintQ = True Then
    BORG.labelUpdater.Caption = "Printing your data..."
    Call Module1.printFun
End If


Call DGscreenChooser("close")

End Sub

Sub borg_btn_OpenCan_Click(form As Object)

If BZmodule.bz_connected = False Then
    BORG.labelUpdater.Caption = "ERROR: Login to BDG and Bluezone to use this feature"
    Exit Sub
End If

If BORG.CanSelectGUI.Value = "" Then
    MsgBox ("Please select a value from the list")
    GrabCloseScreen
End If

openclosedcan (BORG.CanSelectGUI.Value)
GrabCloseScreen

End Sub

Sub borg_btn_removeCan_Click(form As Object)
For intRow = 0 To form.listCan.ListCount - 1
    If form.listCan.Selected(intRow) = True Then
        '.Delete Shift:=xlUp
        myRange = "A" & Trim(str(intRow + 3)) & ":E" & Trim(str(intRow + 3))
        Sheet4.Range(myRange).Delete ([xlUp])
        Exit Sub
    End If
Next
End Sub

Sub borg_btn_UnAssign_Click(form As Object)
If BZmodule.bz_connected = False Then
    BORG.labelUpdater.Caption = "ERROR: Login to BDG and Bluezone to use this feature"
    Exit Sub
End If
If BORG.CanSelectGUI.Value = "" Then
    BORG.labelUpdater.Caption = "ERROR: Please select a value from the list."
    BORG.CanSelectGUI.Value = ""
    GrabCloseScreen
    Exit Sub
End If

Call UnassignCan(BORG.CanSelectGUI.Value)
BORG.labelUpdater.Caption = "Succesfully Unassigned freight from " & BORG.CanSelectGUI.text & "."
BORG.CanSelectGUI.Value = ""
Call GrabCloseScreen

End Sub

Sub borg_btnClose_Click(form As Object)
Call SaveOptions
Unload form
Call BZmodule.BZcloseSessions
Call ThisWorkbook.CloseandSave
Application.Quit
End Sub


Sub borg_btnDump_Click(form As Object)
If BZmodule.bz_connected = False Then
    BORG.labelUpdater.Caption = "ERROR: Login to BDG and Bluezone to use this feature"
    Exit Sub
End If

Dim eRows As Integer
eRows = GetMaxRow
Dim can As String
can = BORG.CanSelectGUI

Call GhostAssign.filterClear
BORG.labelUpdater.Caption = "Dumping " & can & " into AWB data."
'Call GhostAssign.GrabAssigned(eRows, can)
Call GhostAssign.GrabAssign(form.CanSelectGUI.text, "A", GetMaxRow)
Call GhostAssign.GrabAssign(form.CanSelectGUI.text, "I", GetMaxRow)
Call dupFind
Call DGscreenChooser("Close")
Call OpenBlueZone.UnassignCan(BORG.CanSelectGUI.text)
Call FormsHandler.borg_clscrn_refresh_Click(form)
BORG.labelUpdater.Caption = can & " has been dumped into AWB data and is ready to be ghost assigned."
End Sub

Sub borg_btnGhostManIt_Click(form As Object)
If form.ghostCombo.Value <> "" Then
    Call GhostAssign.filterCanSort(BORG.ghostCombo.text)
Else
    form.labelUpdater.Caption = "Please select an item from the drop down menu."
End If

End Sub

Sub borg_btnGrabUnassigned_Click(form As Object)
If BZmodule.bz_connected = False Then
    BORG.labelUpdater.Caption = "ERROR: Login to BDG and Bluezone to use this feature"
    Exit Sub
End If

BORG.labelUpdater.Caption = "Clearing up old data..."

Call GhostAssign.filterClear
Call Module1.DEL

BORG.labelUpdater.Caption = "Grabbing Assign Screen Items"
Call GhostAssign.GrabUnassigned(3)

Call GhostAssign.dupFind
BORG.labelUpdater.Caption = "Finished grabbing items in assign screen"

BORG.labelUpdater.Caption = "Running Fixes"
Call Module4.APOPfix

'BORG.labelUpdater.Caption = "Sorting your data..."
'Call Module4.SORT_MACRO
'
'BORG.labelUpdater.Caption = "Counting Gas"
'Call Module4.gasCount
'
'BORG.labelUpdater.Caption = "Counting Pieces"
'Call Module4.pieceCount

'If BORG.PrintQ = True Then
'    BORG.labelUpdater.Caption = "Printing your data..."
'    Call Module1.printFun
'End If
Call DGscreenChooser("close")
End Sub

Sub borg_btnMinimize_Click(form As Object)
BORG.Hide
miniform.Show
End Sub

Sub borg_btnSettings_Change(form As Object)
If BORG.btnSettings.Value = True Then
    BORG.settingsFrame.Visible = True
Else
    BORG.settingsFrame.Visible = False
    For Each W In Application.Workbooks
        W.Save
    Next W
End If

End Sub

Sub borg_btnSettings_Click(form As Object)

If BORG.btnSettings.Value = True Then
    BORG.settingsFrame.Visible = True
Else
    BORG.settingsFrame.Visible = False
End If
End Sub


Sub borg_Can_flight_Click(form As Object)
If form.StationSort.Value = True And form.Can_flight.Value = True Then
    MsgBox ("Select only one sort option for manifesting")
    form.Can_flight.Value = False
End If
End Sub
Sub borg_CanSelectGUI_Change(form As Object)

End Sub

Sub borg_clscrn_refresh_Click(form As Object)
Call deletecans
form.CanSelectGUI.Value = ""
If BZmodule.bz_connected = False Then
    BORG.labelUpdater.Caption = "ERROR: Login to BDG and Bluezone to use this feature"
    Exit Sub
End If

'BORG.CanSelectGUI.Clear
Call DGscreenChooser("close")

Call GrabCloseScreen

End Sub

Sub borg_combo_splitName_Change(form As Object)
'when new split is selected load up Dest in txtDest
col = 2
Do Until Sheet6.Cells(2, col) = form.combo_splitName
    col = col + 1
Loop

form.txt_Dest = UCase(Sheet6.Cells(4, col))

End Sub

Sub borg_CommandButton9_Click(form As Object)
If BZmodule.bz_connected = False Then
    BORG.labelUpdater.Caption = "ERROR: Login to BDG and Bluezone to use this feature"
    Exit Sub
End If

'clear up filter
Call GhostAssign.filterClear
'clear up old predit assign stuff
Sheet1.Range("U3:U9999").Clear

Call GhostAssign.GhostSort
Call FormsHandler.borg_ghostrefresh_Click(form)

End Sub

Sub borg_dropdownMenu_Change(form As Object)
Select Case form.dropdownMenu.Value
    Case "Job Aid"
        'code
        x = MsgBox("Job Aid not yet implemented", vbCritical, "Not implemented error")
    Case "Split Manager"
        form.Hide
        sortmen.Show
    Case "Ship Center"
        'form.Hide
        'ShipCntr.Show
        x = MsgBox("Ship Center not yet implemented", vbCritical, "Not implemented error")
End Select

form.dropdownMenu.Value = "Menu"
End Sub

Sub borg_ghostrefresh_Click(form As Object)

BORG.ghostCombo.Clear
BORG.ghostrefresh.Visible = False
Call FunctionModule.GhostList
Call GhostAssign.dupFind
BORG.ghostrefresh.Visible = True

End Sub

Sub borg_invis_btn_Click(form As Object)
Application.Visible = False

End Sub

Sub borg_listCan_Change(form As Object)
For intRow = 0 To form.listCan.ListCount - 1
    If form.listCan.Selected(intRow) = True Then
        If Sheet4.Cells(intRow + 3, 1) <> "" Then
            form.txt_canNum = Sheet4.Cells(intRow + 3, 1).Value
            form.combo_splitName = Sheet4.Cells(intRow + 3, 2).Value
            form.combo_hazType = Sheet4.Cells(intRow + 3, 4).Value
        Else
            form.txt_canNum = ""
            form.combo_splitName = ""
            form.combo_hazType = ""
            form.txt_Dest = ""
        End If
        Exit Sub
    End If
Next
form.listCan.Value = ""
End Sub


Sub borg_StationSort_Click(form As Object)
If form.StationSort.Value = True And form.Can_flight.Value = True Then
    MsgBox ("Select only one sort option for manifesting")
    form.StationSort.Value = False
End If

Call SaveOptions
End Sub

Sub borg_tgl_btnLogin_Change(form As Object)
If form.tgl_btnLogin.Value = True Then
    form.loginFrame.Visible = True
Else
    form.loginFrame.Visible = False
End If
End Sub

Sub borg_UserForm_Initialize(form As Object)

Dim Ctl As Control
    For Each Ctl In form.Controls
        'MsgBox TypeName(Ctl)
        Select Case TypeName(Ctl)
            Case "TextBox"
                Ctl.TabKeyBehavior = False
        End Select
    Next Ctl

BORG.btn_AddCan.Height = 18
BORG.btn_AddCan.Left = 264
BORG.btn_AddCan.Top = 2
BORG.btn_AddCan.Width = 54

BORG.btn_removeCan.Height = 18
BORG.btn_removeCan.Left = 258
BORG.btn_removeCan.Top = 26
BORG.btn_removeCan.Width = 60

BORG.btn_clearCans.Height = 18
BORG.btn_clearCans.Left = 258
BORG.btn_clearCans.Top = 48
BORG.btn_clearCans.Width = 60

BORG.btn_cancheck.Visible = True
BORG.btn_cancheck.Height = 18
BORG.btn_cancheck.Left = 258
BORG.btn_cancheck.Top = 70
BORG.btn_cancheck.Width = 60

BORG.btn_ManifestAll.Height = 18
BORG.btn_ManifestAll.Left = 258
BORG.btn_ManifestAll.Top = 96
BORG.btn_ManifestAll.Width = 60

Call RetrieveOptions
Call clearSVConCans

Call Module1.DEL

If form.booGhostShow.Value = False Then Call FormsHandler.borg_booGhostShow_Click(form)
If form.booMoreControls.Value = False Then Call FormsHandler.borg_booMoreControls_Click(form)

'set up haztype combo box
form.combo_hazType.AddItem "ADG"
form.combo_hazType.AddItem "IDG"
form.combo_hazType.AddItem "ALL"

'set up split menu
Call FunctionModule.UpdateSplitList
Call FunctionModule.UpdateCanList

form.dropdownMenu.Clear
form.dropdownMenu.AddItem "Ship Center"
form.dropdownMenu.AddItem "Split Manager"
form.dropdownMenu.AddItem "Job Aid"

End Sub

Sub borg_vis_btn_Click(form As Object)
Application.Visible = True
End Sub

Sub borg_userform_queryClose(form As Object)
Call FormsHandler.borg_btnClose_Click(form)
End Sub



Sub visibleapp()
Application.Visible = True

End Sub

