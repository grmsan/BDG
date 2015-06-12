Attribute VB_Name = "FlexAssign"

Dim host As Variant
Dim cannum As Variant
Dim canType As Variant
Dim canSplit As Variant
Dim canDest As Variant
Dim ADGfind As Variant
Dim IDGfind As Variant
Dim pieces As Integer

Dim retval As Variant
Private Declare Function MessageBox _
Lib "User32" Alias "MessageBoxA" _
(ByVal hWnd As Long, _
ByVal lpText As String, _
ByVal lpCaption As String, _
ByVal wType As Long) _
As Long

Option Compare Text

Sub FlexAssignDirectory(Optional can As String = "ALL")
ChDir "C:\"
Set host = CreateObject("BZwhll.whllobj")
retval = host.OpenSession(0, 11, "fdx3270.zmd", 30, 1)
retval = host.Connect("K")
Set Wnd = host.Window()
Wnd.Caption = "Auto Assign in Progress"
host.waitready 1, 51
If can = "ALL" Then
    Call setupAssignArrays
Else
    cannum = Array(BORG.txt_canNum.Text)
    canSplit = Array(BORG.combo_splitName.Text)
    canDest = Array(BORG.txt_Dest.Text)
    canType = Array(BORG.combo_hazType.Text)
End If

ADGfind = Array("1.4", "2.1", "3", "4.", "5", "8")
IDGfind = Array("2.2", "6.", "7", "9")
Call DGscreenChooser("Assign", host)

Dim i As Integer
Dim hazFilter As String
tempval = UBound(cannum, 1)
For i = 0 To (UBound(cannum, 1))
    Select Case canType(i)
        Case "ADG"
            hazFilter = "A"
        Case "IDG"
            hazFilter = "I"
        Case "ALL"
            hazFilter = " "
        Case Else
            hazFilter = " "
    End Select
    Dim x As String
    x = canSplit(i)
    If isSplitLocal(x) = True Then
        Call SuffixAssign(i, hazFilter)
    ElseIf isSplitLocal(x) = False Then
        Call PrefixAssign(i, hazFilter)
    Else 'something has gone horribly wrong....
        MsgBox ("Error occured Please restart BDG")
        Exit Sub
    End If
Next
'If BORG.booDelIce = True Then Call DeleteIce
If can = "ALL" Then Call isAnythingLeft
BORG.labelUpdater.Caption = "Finished assigning " & pieces & " shipment(s)"
Call GhostAssign.DGscreenChooser("close", host)
End Sub
Sub setupAssignArrays()

Dim c_cannums As New Collection
Dim c_cansplits As New Collection
Dim c_candests As New Collection
Dim c_cantypes As New Collection
row = 3
Do While Sheet4.Cells(row, 1) <> ""
   c_cannums.Add Sheet4.Cells(row, 1) '  dynamically add value to the end
   c_cansplits.Add Sheet4.Cells(row, 2)
   c_candests.Add Sheet4.Cells(row, 3)
   c_cantypes.Add Sheet4.Cells(row, 4)
   row = row + 1
Loop

cannum = toArray(c_cannums) 'convert collection to an array
canSplit = toArray(c_cansplits)
canDest = toArray(c_candests)
canType = toArray(c_cantypes)
End Sub

Function isUrsaLocal(URSA As String)
    ERow = 5
    Do Until Sheet6.Cells(ERow, 2).Value = ""
        If Sheet6.Cells(ERow, 2).Value = URSA Then
            isUrsaLocal = True
            Exit Function
        End If
        ERow = ERow + 1
    Loop
    isUrsaLocal = False
End Function

Function isSplitLocal(MasterID As String)
    If MasterID = "" Then Exit Function
    Ecol = 3
    Do Until Sheet6.Cells(2, Ecol).Value = ""
        If Sheet6.Cells(2, Ecol).Value = MasterID Then
            isSplitLocal = Not (Sheet6.Cells(3, Ecol).Value)
            Exit Function
        End If
        Ecol = Ecol + 1
    Loop
    MsgBox ("not able to find if " & MasterID & " is a local split" & vbNewLine & "error occured in Function isSplitLocal")
End Function

'Function AssignScrn()
'host.sendkey "@C"                       'clears screen in IMS
'host.sendkey "asap@e"                   'types ASAP and enters command
'host.waitready 1, 51
'host.sendkey "68@e"                     'enters 26 for dg training
'host.waitready 1, 51
'host.sendkey "assign@e"                 'enters assign into first field to bring us to assign screen
'host.waitready 1, 51
'host.sendkey BORG.Location.Text      'inputs the location ID in DGinput into station
'If BORG.printerID <> "" Then host.writescreen BORG.printerID, 21, 32
'host.sendkey "@e"                       'sends enter key to bring us finally to Assign Screen
'host.waitready 1, 51
'
'host.readscreen check, 35, 3, 25
'
'If InStr(1, check, "VIEW ALL DG") > 1 Then host.sendkey "@2"
'
'End Function

Sub SuffixAssign(i As Integer, hazFilter As String)
Ecol = 3

Do Until Sheet6.Cells(2, Ecol) = canSplit(i)
    If Sheet6.Cells(2, Ecol).Value = "" Then
        MsgBox ("could not find split " & canSplit(i) & "for can " & cannum(i))
        Exit Sub
    End If
    Ecol = Ecol + 1
Loop

ERow = 5
Do Until Sheet6.Cells(ERow, Ecol) = ""
    host.writescreen "     ", 5, 38
    host.writescreen Sheet6.Cells(ERow, Ecol).Text, 5, 38
    host.writescreen hazFilter, 6, 45
    host.sendkey "@e"
    host.waitready 1, 51

ErrorChecker

    bluerow = 10
    host.readscreen miscdata, 13, bluerow, 5
    Do Until Trim(miscdata) = ""
CheckingPage:
    host.readscreen miscdata, 13, bluerow, 5
        If Right(miscdata, 2) <> "RT" Then
            If Trim(miscdata) <> "" Then
                host.writescreen "A", bluerow, 2
                pieces = pieces + 1
            ElseIf bluerow = 19 Then
                host.writescreen "          ", 7, 24
                host.writescreen cannum(i), 7, 24
                host.writescreen "    ", 7, 53
                host.writescreen canDest(i), 7, 53
                host.sendkey "@e"
                host.waitready 1, 51
                Call FlexAssign.ErrorChecker
                bluerow = 10
                GoTo CheckingPage
            Else
                host.writescreen "          ", 7, 24
                host.writescreen cannum(i), 7, 24
                host.writescreen "    ", 7, 53
                host.writescreen canDest(i), 7, 53
                host.sendkey "@e"
                host.waitready 1, 51
                Call FlexAssign.ErrorChecker
            End If
        End If
        bluerow = bluerow + 1
    Loop
    ERow = ERow + 1
Loop
End Sub

Sub PrefixAssign(i As Integer, hazFilter As String)
    Ecol = 3
    
    Do Until Sheet6.Cells(2, Ecol) = canSplit(i)
        If Sheet6.Cells(2, Ecol).Value = "" Then
            MsgBox ("could not find split " & canSplit(i) & "for can " & cannum(i))
            Exit Sub
        End If
        Ecol = Ecol + 1
    Loop
    
    ERow = 5
    Do Until Sheet6.Cells(ERow, Ecol) = ""
        host.writescreen "  ", 5, 28
        host.writescreen Sheet6.Cells(ERow, Ecol).Text, 5, 28
        host.writescreen hazFilter, 6, 45
        host.sendkey "@e"
        host.waitready 1, 51

ErrorChecker

        bluerow = 10
        host.readscreen miscdata, 13, bluerow, 5
        Do Until Trim(miscdata) = ""
CheckingPagePrefix:
        host.readscreen miscdata, 13, bluerow, 5
            If Right(miscdata, 2) <> "RT" Then
                If isUrsaLocal(Trim(Right(miscdata, 5))) <> True Then
                    If Trim(miscdata) <> "" Then
                        host.writescreen "A", bluerow, 2
                        pieces = pieces + 1
                    ElseIf bluerow = 19 Then
                        host.writescreen "          ", 7, 24
                        host.writescreen cannum(i), 7, 24
                        host.writescreen "    ", 7, 53
                        host.writescreen canDest(i), 7, 53
                        host.sendkey "@e"
                        host.waitready 1, 51
                        Call FlexAssign.ErrorChecker
                        bluerow = 10
                        GoTo CheckingPagePrefix
                    Else
                        host.writescreen "          ", 7, 24
                        host.writescreen cannum(i), 7, 24
                        host.writescreen "    ", 7, 53
                        host.writescreen canDest(i), 7, 53
                        host.sendkey "@e"
                        host.waitready 1, 51
                        Call FlexAssign.ErrorChecker
                    End If
                End If
            End If
            bluerow = bluerow + 1
        Loop
        ERow = ERow + 1
    Loop

End Sub

Sub isAnythingLeft()
host.writescreen "Close ", 2, 17
host.sendkey "@e"
host.waitready 1, 51
host.writescreen "Assign", 2, 17
host.sendkey "@e"
host.waitready 1, 51
leftover = 0
row = 10
Do Until row = 18
    host.readscreen miscdata, 18, row, 51
    If Trim(miscdata) <> "" Then leftover = leftover + 1
    row = row + 1
Loop

If leftover <> 0 Then
    MsgBox ("You have pieces left over after AutoSort" & vbNewLine & _
        "Please view packages in assign screen to determine what to do with them" & _
        vbNewLine & leftover & " pieces at least")
End If

End Sub

Sub DeleteIce()
''Refresh our screen to clear filters
host.writescreen "assign", 2, 17
host.sendkey "@e"
host.waitready 1, 51
''set filters to filter ICE only shipments
host.writescreen "C", 6, 45
host.writescreen "Deleteship", 7, 24
host.sendkey "@e"
host.waitready 1, 51

Data = "tempdata"
row = 10
Do Until Trim(Data) = ""
    host.readscreen Data, 15, row, 5
    If Trim(Data) <> "" Then
        host.writescreen "a", row, 2
    ElseIf Trim(Data) = "" Then
        host.sendkey "@e"
        host.waitready 1, 51
        If row = 18 Then row = 10
    End If
    row = row + 1
Loop
End Sub

Function ErrorChecker()
host.readscreen errorMisc, 3, 24, 2
If errorMisc = "091" Then
    host.sendkey "@4"
    host.waitready 1, 51
End If
If errorMisc = "INV" Then 'invalid container error
    MsgBox ("invalid container")
End If
End Function
