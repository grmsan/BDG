VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} BORG 
   Caption         =   "G.O.A.T. DG"
   ClientHeight    =   7470
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7065
   OleObjectBlob   =   "BORG.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "BORG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub booGhostShow_Click()
If Me.booGhostShow.Value = True Then
    Me.frameGhostMaster.Visible = True
    Me.labelUpdater.Top = 312
    Me.Height = 396
ElseIf Me.booGhostShow.Value = False Then
    Me.frameGhostMaster.Visible = False
    Me.labelUpdater.Top = 246
    Me.Height = 330
End If
'Call SaveOptions
End Sub

Sub booMoreControls_Click()
If Me.booMoreControls.Value = False Then
    Me.btn_CloseCan.Visible = False
    Me.btn_OpenCan.Visible = False
    Me.btn_UnAssign.Visible = False
    Me.frame_closeScrn.Height = 60
    Me.CanSelectGUI.Top = 6
    Me.clscrn_refresh.Top = 6
    Me.btn_ManifestAll.Top = 30
    Me.btn_ManifestOne.Top = 30
    Me.btn_AddIce.Top = 30
ElseIf Me.booMoreControls.Value = True Then
    Me.btn_CloseCan.Visible = True
    Me.btn_OpenCan.Visible = True
    Me.btn_UnAssign.Visible = True
    Me.frame_closeScrn.Height = 84
    Me.CanSelectGUI.Top = 30
    Me.clscrn_refresh.Top = 30
    Me.btn_ManifestAll.Top = 54
    Me.btn_ManifestOne.Top = 54
    Me.btn_AddIce.Top = 54
End If
'Call SaveOptions
End Sub

Private Sub btn_AddCan_Click()
myCannum = Me.txt_canNum
mySplit = Me.combo_splitName
myDest = Me.txt_Dest
myType = Me.combo_hazType

If myCannum = "" Or mySplit = "" Or myDest = "" Or myType = "" Then
    Me.labelUpdater.Caption = "ERROR: PLEASE FILL IN ALL INFORMATION BEFORE ADDING A NEW CAN"
    Exit Sub
End If

x = 2
Do Until Sheet4.Cells(x, 1) = ""
    If Sheet4.Cells(x, 1).Text = myCannum Then Exit Do
    x = x + 1
Loop

Sheet4.Cells(x, 1) = myCannum
Sheet4.Cells(x, 2) = mySplit
Sheet4.Cells(x, 3) = myDest
Sheet4.Cells(x, 4) = myType
Sheet4.Cells(x, 5) = "--"

Me.txt_canNum = ""
Me.combo_hazType = ""
Me.combo_splitName = ""
Me.txt_Dest = ""
'
Call FunctionModule.UpdateCanList

'call function saveWorkBook
Application.ActiveWorkbook.Save
Me.txt_canNum.SetFocus
End Sub

Private Sub btn_AddIce_Click()

If BORG.CanSelectGUI.Value = "" Then
    AddIce ("none")
Else
    AddIce (BORG.CanSelectGUI.Value)
End If

Call BORG.clscrn_refresh_Click

End Sub

Private Sub btn_assignCan_Click()
Call FlexAssign.FlexAssignDirectory(Me.txt_canNum.Text)
GrabCloseScreen
End Sub

Private Sub BTN_AutoAssign_Click()
Call FlexAssign.FlexAssignDirectory
GrabCloseScreen
End Sub

Private Sub btn_cancheck_Click()
famislogingui.EmpNum = BORG.EmpNum

famis.famislogin
famis.famisDestCheck

BORG.CanSelectGUI.Clear
Call DGscreenChooser("close", host)
Call DGscreenChooser("close", host)
GrabCloseScreen
End Sub

Sub btn_clearCans_Click()
'delete cans
Sheet4.Range("A3:E999").Delete xlUp
End Sub

Private Sub btn_CloseCan_Click()
CloseCan (BORG.CanSelectGUI.Value)
Call BORG.clscrn_refresh_Click
End Sub

Sub btn_login_Click()
Call OpenBlueZone.BZcloseSessions
Call OpenBlueZone.BZOpenSession
End Sub

Private Sub btn_ManifestOne_Click()
Call GhostAssign.filterClear
Call Module1.DEL

If BORG.CanSelectGUI.Value = "" Then
    BORG.labelUpdater.Caption = "PLEASE SELECT A CAN TO MANIFEST FROM THE CLOSED SCREEN CAN CHOOSER"
    Exit Sub
End If
'Call TheGrab(Empty, " ", 3)'old
Call Module1.SETUP
Call GhostAssign.GrabAssigned(3, Me.CanSelectGUI.Text)

BORG.labelUpdater.Caption = "Running Fixes"
Call Module4.APOPfix

BORG.labelUpdater.Caption = "Sorting your data..."
Call Module4.SORT_MACRO

BORG.labelUpdater.Caption = "Counting Gas"
Call Module4.gasCount

BORG.labelUpdater.Caption = "Counting Pieces"
Call Module4.pieceCount

BORG.labelUpdater.Caption = "Printing your data..."
If BORG.PrintQ = True Then Call Module1.printFun

Call DGscreenChooser("close", host)

End Sub

Private Sub btn_OpenCan_Click()

If BORG.CanSelectGUI.Value = "" Then
    MsgBox ("Please select a value from the list")
    BORG.CanSelectGUI.Clear
    GrabCloseScreen
End If

openclosedcan (BORG.CanSelectGUI.Value)
BORG.CanSelectGUI.Clear
GrabCloseScreen

End Sub

Private Sub btn_removeCan_Click()
For intRow = 0 To Me.listCan.ListCount - 1
    If Me.listCan.Selected(intRow) = True Then
        '.Delete Shift:=xlUp
        myRange = "A" & Trim(str(intRow + 3)) & ":E" & Trim(str(intRow + 3))
        Sheet4.Range(myRange).Delete ([xlUp])
        Exit Sub
    End If
Next
End Sub

Private Sub btn_UnAssign_Click()

If BORG.CanSelectGUI.Value = "" Then
    MsgBox ("Please select a value from the list")
    BORG.CanSelectGUI.Clear
    GrabCloseScreen
End If

Call UnassignCan(BORG.CanSelectGUI.Value)
BORG.CanSelectGUI.Clear
Call GrabCloseScreen

End Sub

Sub btnClose_Click()
Call SaveOptions
Unload Me
Call OpenBlueZone.BZcloseSessions
Call ThisWorkbook.CloseandSave
Application.Quit
End Sub


Private Sub btnDump_Click()
Dim eRows As Integer
eRows = GetMaxRow
Dim can As String
can = BORG.CanSelectGUI

Call GhostAssign.filterClear
Call GhostAssign.GrabAssigned(eRows, can)
Call OpenBlueZone.UnassignCan(BORG.CanSelectGUI.Text)
Call BORG.clscrn_refresh_Click
End Sub

Private Sub btnGhostManIt_Click()
If Me.ghostCombo.Value <> "" Then
    Call GhostAssign.filterCanSort(BORG.ghostCombo.Text)
Else
    Me.labelUpdater.Caption = "Please select an item from the drop down menu."
End If

End Sub

Private Sub btnGrabUnassigned_Click()
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
Call DGscreenChooser("close", "empty")
End Sub

Private Sub btnMinimize_Click()
Me.Hide
miniform.Show
End Sub

Private Sub btnSettings_Change()
If BORG.btnSettings.Value = True Then
    BORG.settingsFrame.Visible = True
Else
    BORG.settingsFrame.Visible = False
    For Each W In Application.Workbooks
        W.Save
    Next W
End If

End Sub

Sub btnSettings_Click()

If BORG.btnSettings.Value = True Then
    BORG.settingsFrame.Visible = True
Else
    BORG.settingsFrame.Visible = False
End If
End Sub


Private Sub Can_flight_Click()
If Me.StationSort.Value = True And Me.Can_flight.Value = True Then
    MsgBox ("Select only one sort option for manifesting")
    Me.Can_flight.Value = False
End If
End Sub

Private Sub CanSelectGUI_Change()

End Sub

Sub clscrn_refresh_Click()
If Me.loginStatusOff.Visible = True Then
    Me.labelUpdater.Caption = "PLEASE LOGIN TO BLUEZONE"
    Me.tgl_btnLogin.Value = True
    Exit Sub
End If

BORG.CanSelectGUI.Clear
Call DGscreenChooser("close", host)
Call GrabCloseScreen
End Sub

Private Sub combo_splitName_Change()
'when new split is selected load up Dest in txtDest
col = 2
Do Until Sheet6.Cells(2, col) = Me.combo_splitName
    col = col + 1
Loop

Me.txt_Dest = UCase(Sheet6.Cells(4, col))

End Sub

Private Sub CommandButton9_Click()
'clear up filter
Call GhostAssign.filterClear
'clear up old predit assign stuff
Sheet1.Range("U3:U9999").Clear

Call GhostAssign.GhostSort
Call BORG.ghostrefresh_Click

End Sub

Private Sub dropdownMenu_Change()
Select Case Me.dropdownMenu.Value
    Case "Job Aid"
        'code
        x = MsgBox("Job Aid not yet implemented", vbCritical, "Not implemented error")
    Case "Split Manager"
        Me.Hide
        sortmen.Show
    Case "Ship Center"
        'Me.Hide
        'ShipCntr.Show
        x = MsgBox("Ship Center not yet implemented", vbCritical, "Not implemented error")
End Select

Me.dropdownMenu.Value = "Menu"
End Sub

Sub ghostrefresh_Click()

BORG.ghostCombo.Clear
BORG.ghostrefresh.Visible = False
Call FunctionModule.GhostList
Call GhostAssign.dupFind
BORG.ghostrefresh.Visible = True

End Sub

Private Sub invis_btn_Click()
Application.Visible = False

End Sub

Private Sub listCan_Change()
For intRow = 0 To Me.listCan.ListCount - 1
    If Me.listCan.Selected(intRow) = True Then
        If Sheet4.Cells(intRow + 3, 1) <> "" Then
            Me.txt_canNum = Sheet4.Cells(intRow + 3, 1).Value
            Me.combo_splitName = Sheet4.Cells(intRow + 3, 2).Value
            Me.combo_hazType = Sheet4.Cells(intRow + 3, 4).Value
        Else
            Me.txt_canNum = ""
            Me.combo_splitName = ""
            Me.combo_hazType = ""
            Me.txt_Dest = ""
        End If
        Exit Sub
    End If
Next
Me.listCan.Value = ""
End Sub


Private Sub StationSort_Click()
If Me.StationSort.Value = True And Me.Can_flight.Value = True Then
    MsgBox ("Select only one sort option for manifesting")
    Me.StationSort.Value = False
End If

Call SaveOptions
End Sub

Sub tgl_btnLogin_Change()
If BORG.tgl_btnLogin.Value = True Then
    BORG.loginFrame.Visible = True
Else
    BORG.loginFrame.Visible = False
End If
End Sub

Sub UserForm_Initialize()
Call RetrieveOptions
Call clearSVConCans
If Me.booGhostShow.Value = False Then Call BORG.booGhostShow_Click
If Me.booMoreControls.Value = False Then Call BORG.booMoreControls_Click

'set up haztype combo box
Me.combo_hazType.AddItem "ADG"
Me.combo_hazType.AddItem "IDG"
Me.combo_hazType.AddItem "ALL"

'set up split menu
Call FunctionModule.UpdateSplitList
Call FunctionModule.UpdateCanList

Me.dropdownMenu.Clear
Me.dropdownMenu.AddItem "Ship Center"
Me.dropdownMenu.AddItem "Split Manager"
Me.dropdownMenu.AddItem "Job Aid"

End Sub

Private Sub vis_btn_Click()
Application.Visible = True
End Sub
