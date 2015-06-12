VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} newsplitfrm 
   Caption         =   "New Split Form"
   ClientHeight    =   2355
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3960
   OleObjectBlob   =   "newsplitfrm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "newsplitfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btn_closeNewSplit_Click()
If Me.BooPrefix = True And Me.BooSuffix = True Then
    MsgBox ("Please select only Prefix or Suffix")
    Exit Sub
End If
If Me.BooPrefix = False And Me.BooSuffix = False Then
    MsgBox ("Please select Prefix or Suffix")
    Exit Sub
End If

If Me.BooPrefix = True Then
    Call SplitCentral.addMstrSplit(Me.SplitName, Me.SplitDest, True)
End If
If Me.BooSuffix = True Then
    Call SplitCentral.addMstrSplit(Me.SplitName, Me.SplitDest, False)
End If

Me.SplitDest = ""
Me.SplitName = ""
Call SplitCentral.SortMenOpen

Call FunctionModule.UpdateSplitList

Me.Hide
End Sub
