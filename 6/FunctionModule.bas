Attribute VB_Name = "FunctionModule"

Function AssignITSetup()

Call splitsetup(BORG.canSplit1)
Call splitsetup(BORG.canSplit2)
Call splitsetup(BORG.canSplit3)
Call splitsetup(BORG.canSplit4)
Call splitsetup(BORG.canSplit5)
Call splitsetup(BORG.canSplit6)
Call splitsetup(BORG.canSplit7)

BORG.CanType1.AddItem ""
BORG.CanType1.AddItem "ADG"
BORG.CanType1.AddItem "IDG"
BORG.CanType1.AddItem "ALL"

BORG.CanType2.AddItem ""
BORG.CanType2.AddItem "ADG"
BORG.CanType2.AddItem "IDG"
BORG.CanType2.AddItem "ALL"

BORG.CanType3.AddItem ""
BORG.CanType3.AddItem "ADG"
BORG.CanType3.AddItem "IDG"
BORG.CanType3.AddItem "ALL"

BORG.CanType4.AddItem ""
BORG.CanType4.AddItem "ADG"
BORG.CanType4.AddItem "IDG"
BORG.CanType4.AddItem "ALL"

BORG.CanType5.AddItem ""
BORG.CanType5.AddItem "ADG"
BORG.CanType5.AddItem "IDG"
BORG.CanType5.AddItem "ALL"

BORG.CanType6.AddItem ""
BORG.CanType6.AddItem "ADG"
BORG.CanType6.AddItem "IDG"
BORG.CanType6.AddItem "ALL"

BORG.CanType7.AddItem ""
BORG.CanType7.AddItem "ADG"
BORG.CanType7.AddItem "IDG"
BORG.CanType7.AddItem "ALL"

End Function
Sub splitsetup(cannum As Object)
Dim excelcol As Integer
excelcol = 3

With cannum
    .Clear
    .AddItem ""
Do Until Sheet6.Cells(2, excelcol).Text = ""
    .AddItem Sheet6.Cells(2, excelcol)
    excelcol = excelcol + 1
Loop
End With
End Sub
Function CanSave()

Sheet3.Cells(3, 16).Value = BORG.CanNum1.Value
Sheet3.Cells(4, 16).Value = BORG.CanNum2.Value
Sheet3.Cells(5, 16).Value = BORG.CanNum3.Value
Sheet3.Cells(6, 16).Value = BORG.CanNum4.Value
Sheet3.Cells(7, 16).Value = BORG.CanNum5.Value
Sheet3.Cells(8, 16).Value = BORG.CanNum6.Value
Sheet3.Cells(9, 16).Value = BORG.CanNum7.Value

Sheet3.Cells(3, 17).Value = BORG.canDest1.Value
Sheet3.Cells(4, 17).Value = BORG.canDest2.Value
Sheet3.Cells(5, 17).Value = BORG.canDest3.Value
Sheet3.Cells(6, 17).Value = BORG.canDest4.Value
Sheet3.Cells(7, 17).Value = BORG.canDest5.Value
Sheet3.Cells(8, 17).Value = BORG.canDest6.Value
Sheet3.Cells(9, 17).Value = BORG.canDest7.Value

Sheet3.Cells(3, 18).Value = BORG.CanType1.Value
Sheet3.Cells(4, 18).Value = BORG.CanType2.Value
Sheet3.Cells(5, 18).Value = BORG.CanType3.Value
Sheet3.Cells(6, 18).Value = BORG.CanType4.Value
Sheet3.Cells(7, 18).Value = BORG.CanType5.Value
Sheet3.Cells(8, 18).Value = BORG.CanType6.Value
Sheet3.Cells(9, 18).Value = BORG.CanType7.Value

Sheet3.Cells(3, 19).Value = BORG.canSplit1.Value
Sheet3.Cells(4, 19).Value = BORG.canSplit2.Value
Sheet3.Cells(5, 19).Value = BORG.canSplit3.Value
Sheet3.Cells(6, 19).Value = BORG.canSplit4.Value
Sheet3.Cells(7, 19).Value = BORG.canSplit5.Value
Sheet3.Cells(8, 19).Value = BORG.canSplit6.Value
Sheet3.Cells(9, 19).Value = BORG.canSplit7.Value

End Function
Function AssignRecover()
BORG.CanNum1.Value = Sheet3.Cells(3, 16).Value
BORG.CanNum2.Value = Sheet3.Cells(4, 16).Value
BORG.CanNum3.Value = Sheet3.Cells(5, 16).Value
BORG.CanNum4.Value = Sheet3.Cells(6, 16).Value
BORG.CanNum5.Value = Sheet3.Cells(7, 16).Value
BORG.CanNum6.Value = Sheet3.Cells(8, 16).Value
BORG.CanNum7.Value = Sheet3.Cells(9, 16).Value

BORG.canDest1.Value = Sheet3.Cells(3, 17).Value
BORG.canDest2.Value = Sheet3.Cells(4, 17).Value
BORG.canDest3.Value = Sheet3.Cells(5, 17).Value
BORG.canDest4.Value = Sheet3.Cells(6, 17).Value
BORG.canDest5.Value = Sheet3.Cells(7, 17).Value
BORG.canDest6.Value = Sheet3.Cells(8, 17).Value
BORG.canDest7.Value = Sheet3.Cells(9, 17).Value

BORG.CanType1.Value = Sheet3.Cells(3, 18).Value
BORG.CanType2.Value = Sheet3.Cells(4, 18).Value
BORG.CanType3.Value = Sheet3.Cells(5, 18).Value
BORG.CanType4.Value = Sheet3.Cells(6, 18).Value
BORG.CanType5.Value = Sheet3.Cells(7, 18).Value
BORG.CanType6.Value = Sheet3.Cells(8, 18).Value
BORG.CanType7.Value = Sheet3.Cells(9, 18).Value

BORG.canSplit1.Value = Sheet3.Cells(3, 19).Value
BORG.canSplit2.Value = Sheet3.Cells(4, 19).Value
BORG.canSplit3.Value = Sheet3.Cells(5, 19).Value
BORG.canSplit4.Value = Sheet3.Cells(6, 19).Value
BORG.canSplit5.Value = Sheet3.Cells(7, 19).Value
BORG.canSplit6.Value = Sheet3.Cells(8, 19).Value
BORG.canSplit7.Value = Sheet3.Cells(9, 19).Value

End Function


Function Classfind(Raw As String)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''Uses strings within excel to find class and subrisk'''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If InStr(1, Raw, "RADIOACTIVE MATERIAL, EXCEPTED PACKAGE") > 1 Then Exit Function
classposition = 0
hazclass = ""

Subend = 1

SubClass = Array(", 1.4B, ", ", 1.4C, ", ", 1.4D, ", ", 1.4E, ", ", 1.4G, ", _
    ", 1.4S, ", ", 2.1, ", ", 2.2, ", ", 3, ", ", 4.1, ", ", 4.2, ", ", 4.3, ", _
    ", 5.1, ", ", 5.2, ", ", 6.1, ", ", 6.2, ", ", 7, ", ", 8, ", ", 9, ", ", 1.4B(", _
    ", 1.4C(", ", 1.4D(", ", 1.4E(", ", 1.4G(", ", 1.4S(", ", 2.1(", ", 2.2(", _
    ", 3(", ", 4.1(", ", 4.2(", ", 4.3(", ", 5.1(", ", 5.2(", ", 6.1(", ", 6.2(", _
    ", 7(", ", 8(", ", 9(")

x = 0
Do Until classposition > 1 Or x > 37
    classposition = InStr(1, Raw, SubClass(x))
    If classposition > 1 Then
        classposition = classposition + 1
            If x > 18 Then
            Do Until endcheck = ")"
                endcheck = Mid(Raw, classposition + Subend, 1)
                If endcheck = ")" Then Exit Do
                Subend = Subend + 1
            Loop
            Else
                Classes = Array("1.4B", "1.4C", "1.4D", "1.4E", "1.4G", "1.4S", "2.1", "2.2", "3", _
                    "4.1", "4.2", "4.3", "5.1", "5.2", "6.1", "6.2", "7", "8", "9")
                hazclass = Classes(x)
                Exit Do
            End If
        hazclass = Mid(Raw, classposition + 1, classposition - (classposition - Subend))
    End If
x = x + 1
Loop

'UGLY code
Sheet3.Cells(16, 6).Value = classposition - 1
Sheet3.Cells(16, 5).Value = hazclass

End Function

Function PSNfind(Raw As String)
If InStr(1, Raw, "RADIOACTIVE MATERIAL, EXCEPTED PACKAGE") > 1 Then
    If Left(Raw, 2) = "RQ" Then
        RQ = "RQ - "
    Else
        RQ = ""
    End If
    Sheet3.Cells(16, 4).Value = RQ & "RADIOACTIVE MATERIAL, EXCEPTED PACKAGE"
    PSNfind = RQ & "RADIOACTIVE MATERIAL, EXCEPTED PACKAGE"
    Exit Function
End If
Start = 8
classposition = Sheet3.Cells(16, 6).Value
RQ = ""
If Left(Raw, 2) = "RQ" Then
    Start = Start + 4
    RQ = "RQ - "
    End If
If classposition = -1 Then classposition = 80

PSN = (RQ + Mid(Raw, Start, (classposition - Start)))

Sheet3.Cells(16, 4).Value = RQ + PSN
PSNfind = Trim(PSN)
    
End Function

Function PGfind(Raw As String)
If InStr(1, Raw, "RADIOACTIVE MATERIAL, EXCEPTED PACKAGE") > 1 Then
    PG = "X"
    Sheet3.Cells(16, 7).Value = PG
    Exit Function
End If
PG = ""
PGpos = 0
If PGpos = 0 Then
    If InStr(1, Raw, ", III,") > 1 Then
        PGpos = InStr(1, Raw, ", III,")
        PG = "III"
    End If
End If

If PGpos = 0 Then
    If InStr(1, Raw, ", II,") > 1 Then
        PGpos = InStr(1, Raw, ", II,")
        PG = "II"
    End If
End If

If PGpos = 0 Then
    If InStr(1, Raw, ", I,") > 1 Then
        PGpos = InStr(1, Raw, ", I,")
        PG = "I"
    End If
End If

If PG = "" Then
    PG = "X"
    PGpos = Sheet3.Cells(16, 6).Value
    End If
Sheet3.Cells(16, 7).Value = PG
Sheet3.Cells(15, 7).Value = PGpos


End Function

Function WTfind(Raw As String)
If InStr(1, Raw, "RADIOACTIVE MATERIAL, EXCEPTED PACKAGE") > 1 Then
    WTfind = Array("EQ", "EQ")
    Exit Function
End If

WT = 0
x = 0
UM = 0
Start = 0
Last = 0
classposition = Sheet3.Cells(16, 6).Value

If Sheet3.Cells(16, 7).Value <> "X" Then
Start = Sheet3.Cells(15, 7).Value
Else: Start = Sheet3.Cells(16, 6).Value + Len(Sheet3.Cells(16, 5))
End If

x = InStr(Start + 1, Raw, ",")
Start = x
'x = InStr(Start, Raw, ",")
'Start = x
Last = InStr(Start + 1, Raw, ",")
WTUM = Mid(Raw, Start + 2, (Last - Start) - 2)

spaceSearch = " "
x = InStr(1, WTUM, spaceSearch)
y = Len(WTUM)
WT = Left(WTUM, x - 1)
UM = Right(WTUM, y - x)

WTfind = Array(WT, UM)

End Function

Function Num_Pcs(Raw As String)
On Error GoTo errorout
x = InStr(1, Raw, " PIECE")
y = 1
Do Until commacheck = ","
    commacheck = Mid(Raw, x - y, 1)
    y = y + 1
Loop

Num_Pcs = Mid(Raw, x - y + 3, y - 3)
Exit Function

errorout:
'yes it's the easy way out.... but it's prolly right more often than not..
Num_Pcs = 1

End Function

Function GhostList() As Boolean

On Error GoTo ErrOut

Dim index As Integer
Dim inList As Boolean
maxRows = GetMaxRow

For i = 3 To maxRows - 1
    inList = False
    For index = 0 To BORG.ghostCombo.ListCount - 1
        If Sheet1.Cells(i, 21) = CStr(BORG.ghostCombo.List(index)) Then
            inList = True
            Exit For
        End If
    Next index
    If inList = False And Sheet1.Cells(i, 21) <> "" Then BORG.ghostCombo.AddItem Sheet1.Cells(i, 21)
Next
GhostList = True
Exit Function

ErrOut:
GhostList = False
End Function

Function setLogoutTime()
Dim curDate As Date
curDate = Now
iday = Day(curDate)
ihour = Hour(curDate)
imonth = Month(curDate)
iyear = Year(curDate)

Sheet3.Cells(4, 4).Value = iday
Sheet3.Cells(5, 4).Value = ihour
Sheet3.Cells(6, 4).Value = imonth
Sheet3.Cells(7, 4).Value = iyear
End Function

Function Clear_old_cans()
Dim curDate As Date
curDate = Now
iday = Day(curDate)
ihour = Hour(curDate)
imonth = Month(curDate)
iyear = Year(curDate)

If iyear - Sheet3.Cells(7, 4).Value <> 0 Or _
   iday - Sheet3.Cells(4, 4).Value <> 0 Or _
   imonth - Sheet3.Cells(6, 4).Value <> 0 Then
    Call BORG.btn_clearCans_Click
    Call FunctionModule.CanSave
ElseIf Abs(ihour - Sheet3.Cells(5, 4)) >= 6 Then
    Call BORG.Clear_Assign_Click
    Call FunctionModule.CanSave
End If

Sheet3.Cells(4, 4).Value = iday
Sheet3.Cells(5, 4).Value = ihour
Sheet3.Cells(6, 4).Value = imonth
Sheet3.Cells(7, 4).Value = iyear
End Function

Function UpdateCanList()
'Update can list upon call
'"DATA!A3:D(last row)" is goal for str
finalstr = "DATA!A3:E"
row = 3
Do Until Sheet4.Cells(row, 1) & Sheet4.Cells(row, 2) & Sheet4.Cells(row, 3) & Sheet4.Cells(row, 4) = ""
    row = row + 1
Loop
lastrow = Trim(str(row))
finalstr = finalstr & lastrow
BORG.listCan.RowSource = finalstr

End Function

Function UpdateSplitList()
'for borg UI
'update the split combo box to include all splits to user
'should be run anytime a split is added/removed/modified
BORG.combo_splitName.Clear
col = 2
Do Until Sheet6.Cells(2, col) = ""
    BORG.combo_splitName.AddItem Sheet6.Cells(2, col).Text
    col = col + 1
Loop
End Function

Function RetrieveOptions()
BORG.EmpNum = Sheet4.Cells(2, 8)
BORG.PW_remember = Sheet4.Cells(4, 8)

If BORG.PW_remember = True Then
    BORG.PasswordBox = Sheet4.Cells(3, 8)
Else:
    Sheet4.Cells(4, 8) = ""
End If

BORG.Location = Sheet4.Cells(5, 8)
BORG.printerID = Sheet4.Cells(6, 8)
BORG.StationSort = Sheet4.Cells(7, 8)
BORG.Can_flight = Sheet4.Cells(8, 8)
BORG.phx_Food = Sheet4.Cells(9, 8)
BORG.PrintQ = Sheet4.Cells(10, 8)
BORG.booMoreControls = Sheet4.Cells(11, 8)
BORG.booGhostShow = Sheet4.Cells(12, 8)
End Function

Function SaveOptions()
Sheet4.Cells(2, 8) = BORG.EmpNum
Sheet4.Cells(4, 8) = BORG.PW_remember

If BORG.PW_remember = True Then
    Sheet4.Cells(3, 8) = BORG.PasswordBox
Else:
    Sheet4.Cells(4, 8) = ""
End If

Sheet4.Cells(5, 8) = BORG.Location
Sheet4.Cells(6, 8) = BORG.printerID
Sheet4.Cells(7, 8) = BORG.StationSort
Sheet4.Cells(8, 8) = BORG.Can_flight
Sheet4.Cells(9, 8) = BORG.phx_Food
Sheet4.Cells(10, 8) = BORG.PrintQ
Sheet4.Cells(11, 8) = BORG.booMoreControls
Sheet4.Cells(12, 8) = BORG.booGhostShow
End Function

Function clearSVConCans()
row = 3
Do While Sheet4.Cells(row, 5) <> ""
    Sheet4.Cells(row, 5) = "--"
row = row + 1
Loop

End Function
