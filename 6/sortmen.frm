VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} sortmen 
   Caption         =   "Sort Manager"
   ClientHeight    =   6450
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3240
   OleObjectBlob   =   "sortmen.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "sortmen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btn_AddNewSplit_Click()
newsplitfrm.Show
End Sub

Private Sub btn_Addsplit_Click()
Dim PkgPrefix As Boolean

If Len(Me.AddsplitBox.Text) = 4 Or Len(Me.AddsplitBox.Text) = 5 Then
    PkgPrefix = False
Else:
    If Len(Me.AddsplitBox.Text) = 2 Or Len(Me.AddsplitBox.Text) = 1 Then
        PkgPrefix = True
    Else
        MsgBox ("You must enter a 1 or 2 digit Prefix" & vbNewLine & "or a 4 or 5 digit Suffix")
        Exit Sub
    End If
End If

Call AddSplits(Me.SplitSelectGui.Text, Me.AddsplitBox.Text, PkgPrefix)
Me.AddsplitBox.Value = ""
End Sub

Private Sub btn_DeleteSplit_Click()
Call SplitCentral.removeMstrSplit(Me.SplitSelectGui.Text)
Call SplitCentral.SortMenOpen
End Sub

Private Sub btn_removeSplit_Click()
Dim SplitID As String
Dim position As Integer
position = 0

For i = 0 To Me.Splitsdisplay.ListCount - 1
    If Me.Splitsdisplay.Selected(i) Then
        SplitID = Me.Splitsdisplay.List(i)
        position = i + 5
        Exit For
    End If
Next i

If position = 0 Then Exit Sub
Call RemoveSubSplit(Me.SplitSelectGui.Text, SplitID, position)

Me.Splitsdisplay.Clear
loadsplits (Me.SplitSelectGui.Text)
End Sub

Private Sub btn_SaveAndClose_Click()

Call FunctionModule.UpdateSplitList

For Each W In Application.Workbooks
 W.Save
Next W

sortmen.Hide
Call Module1.OpenBDG
End Sub

Private Sub CommandButton1_Click()
If sortmen.SplitSelectGui = "" Then Exit Sub
Dim newDest As String
newDest = InputBox("Enter the new Destination to use for this split", "New Destination")
If Len(newDest) >= 6 Or Len(newDest) < 3 Then
    MsgBox ("Please enter the locations 3, 4, or 5 digit destination" & vbNewLine & _
            "ex. 'MEM' , 'MEMH' , 'PHXRT'")
    Exit Sub
End If
Call ChangeDest(sortmen.SplitSelectGui.Text, newDest)
End Sub

Private Sub CommandButton2_Click()
Call FunctionModule.UpdateSplitList

For Each W In Application.Workbooks
 W.Save
Next W

sortmen.Hide
Call Module1.OpenBDG
End Sub

Private Sub SplitSelectGui_Change()
'If Me.SplitSelectGui.Value = "" Then
'    Me.Height = 126.75
'Else
'    Me.Height = 360
'End If
Me.Splitsdisplay.Clear
loadsplits (Me.SplitSelectGui.Text)
End Sub



Private Sub UserForm_Initialize()

Call SplitCentral.SortMenOpen

End Sub
