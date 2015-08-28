<<<<<<< HEAD
Attribute VB_Name = "SplitCentral"
=======
>>>>>>> origin/master

Dim excelcol As Integer
Dim excelrow As Integer

Sub SortMenOpen()
sortmen.SplitSelectGui.Clear
excelcol = 2
Do Until Sheet6.Cells(2, excelcol).Value = ""
    sortmen.SplitSelectGui.AddItem (Sheet6.Cells(2, excelcol).Value)
    excelcol = excelcol + 1
Loop
If sortmen.Visible = False Then
    sortmen.Show
End If
End Sub

Sub loadsplits(SplitName As String)
excelrow = 5
excelcol = 2
Do Until Sheet6.Cells(2, excelcol).text = SplitName
    If Sheet6.Cells(2, excelcol).text = "" Then
        sortmen.Height = 126.75
        Exit Sub
    End If
    excelcol = excelcol + 1
Loop
sortmen.textDestination.Value = Sheet6.Cells(4, excelcol).Value
Do Until Sheet6.Cells(excelrow, excelcol) = ""
    sortmen.Splitsdisplay.AddItem Sheet6.Cells(excelrow, excelcol)
    excelrow = excelrow + 1
Loop
End Sub

Sub AddSplits(mstrSplit As String, subSplit As String, isPrefix As Boolean)
excelcol = 2
excelrow = 5
subSplit = UCase(subSplit)
Do Until Sheet6.Cells(2, excelcol).text = mstrSplit
    excelcol = excelcol + 1
Loop
x = Sheet6.Cells(3, excelcol).text
If Sheet6.Cells(3, excelcol).text <> isPrefix Then
    MsgBox ("You cannot mix Prefix and Suffix in 1 split")
    Exit Sub
End If
Do Until Sheet6.Cells(excelrow, excelcol) = ""
    If Sheet6.Cells(excelrow, excelcol) = subSplit Then
        MsgBox ("URSA code already exists for this split")
        Exit Sub
    End If
    excelrow = excelrow + 1
Loop
 Sheet6.Cells(excelrow, excelcol) = subSplit
sortmen.Splitsdisplay.AddItem subSplit
End Sub

Sub RemoveSubSplit(MasterID As String, SplitID As String, position As Integer)
excelrow = position
excelcol = 2
Do Until Sheet6.Cells(2, excelcol).Value = MasterID
    excelcol = excelcol + 1
Loop
 Sheet6.Cells(excelrow, excelcol).Delete shift:=xlUp
End Sub

Sub ChangeDest(MasterID As String, newDest As String)
excelcol = 2
Do Until Sheet6.Cells(2, excelcol) = MasterID
    excelcol = excelcol + 1
Loop
 Sheet6.Cells(4, excelcol) = newDest
sortmen.Splitsdisplay.Clear
loadsplits (sortmen.SplitSelectGui.text)
End Sub

Sub removeMstrSplit(MasterID As String)
excelcol = 2
If MasterID = "Local" Then
    MsgBox ("Deleting the local sort is not allowed." & vbNewLine & "You can generate a new local sort by clicking the 'Generate Local Sort' button.")
    Exit Sub
End If
Do Until Sheet6.Cells(2, excelcol).Value = MasterID
    If Sheet6.Cells(2, excelcol).Value = "" Then
        MsgBox ("Error occured please reload split manager and try again")
        Exit Sub
    End If
    excelcol = excelcol + 1
Loop
 Sheet6.Cells(2, excelcol).EntireColumn.Delete
End Sub

Sub addMstrSplit(MasterID As String, Dest As String, isPrefix As Boolean)
excelcol = 3
Do Until Sheet6.Cells(2, excelcol).text = ""
    If Sheet6.Cells(2, excelcol).text = MasterID Then
        MsgBox ("A split with this name already exists!" & vbNewLine & _
                "If you want to modify an existing split just select it from the drop down list")
        Exit Sub
    End If
    excelcol = excelcol + 1
Loop
 Sheet6.Cells(2, excelcol).Value = MasterID
 Sheet6.Cells(3, excelcol).Value = isPrefix
 Sheet6.Cells(4, excelcol).Value = Dest
Call SplitCentral.SortMenOpen
End Sub

Function doIgoHere(SplitName As String, URSA As String, isPrefix As Boolean)
'all splits are assumed guilty until proven innocent!
verdict = False

excelcol = 3
Do Until Sheet6.Cells(2, excelcol) = SplitName
    If Sheet6.Cells(2, excelcol) = "" Then
        MsgBox ("Problem assigning" & vbNewLine & _
            "Splitname :" & SplitName & " does not exist!")
        Exit Function
    End If
    excelcol = excelcol + 1
Loop

If isPrefix = True Then
    URSA = Left(URSA, 2)
    If Sheet6.Cells(3, excelcol) <> isPrefix Then MsgBox ("Prefix/Suffix does not match split")
Else
    URSA = Trim(Mid(URSA, 3, 5))
    If Sheet6.Cells(3, excelcol) <> isPrefix Then MsgBox ("Prefix/Suffix does not match split")
End If

'trying to match first 2 letters
ERow = 5
Do Until Sheet6.Cells(ERow, excelcol) = URSA
    If Sheet6.Cells(ERow, excelcol) = "" Then GoTo check1char
    ERow = ERow + 1
Loop
verdict = True
'if we exit we have found a match so we'll skip next part
GoTo aftercharcheck

'checking 1 char string for match
check1char:
URSA = Left(URSA, 1)
ERow = 5
Do Until Sheet6.Cells(ERow, excelcol) = URSA
    If Sheet6.Cells(ERow, excelcol) = "" Then GoTo aftercharcheck
    ERow = ERow + 1
Loop
verdict = True

aftercharcheck:
doIgoHere = verdict
End Function

Function isPrefix(MasterID As String)
excelcol = 2

Do Until Sheet6.Cells(2, excelcol) = MasterID
    If Sheet6.Cells(2, excelcol) = "" Then
        MsgBox ("Could not find split: " & MasterID & ". Please refresh splits.")
        Exit Function
    End If
    excelcol = excelcol + 1
Loop

'if we exit then we have found split name
'return value of isPrefix on worksheet
isPrefix = Sheet6.Cells(3, excelcol)

End Function

Sub clearLocal()
'local col is always 2
col = 2
row = 4
Sheet6.Cells(row, col).Value = ""
row = row + 1
Do Until Sheet6.Cells(row, col) = ""
    Sheet6.Cells(row, col) = ""
    row = row + 1
Loop
End Sub


<<<<<<< HEAD


=======
>>>>>>> origin/master
