Dim host As Variant
Dim row As Integer
Dim excelrow As Integer
Option Compare Text

Sub Directions(host As Variant, excelrow As Integer)

'On Error GoTo errout

Dim special As Integer
special = 0
If Sheet1.Cells(excelrow, 16).Value > 0 Then
    Call VAWB.VerifyOP(host)
    special = 1
    End If
If Sheet1.Cells(excelrow, 14).Value > 0 Then
    Call VAWB.VerifyALPK(host)
    special = 1
    End If
If special = 0 Then Call VAWB.VerifyAWB(host)

Call VAWB.VAWB_Origin(host, excelrow)
Call VAWB.Assembly(host, excelrow)

Dim row As Integer
row = 17

Dim unpos As Integer
Dim clspos As Integer

Dim raw As String
raw = ""

Do Until Sheet3.Cells(row, 2).Value = ""
    raw = Sheet3.Cells(row, 2).Text
    unpos = VAWB_UN_RQ(raw, excelrow, row, special)
    clspos = VAWB.VAWB_Class(raw, excelrow, row, special)
    Call VAWB.VAWB_PSN(raw, excelrow, row, unpos, clspos, special)
    
    If special = 1 Then
        Call VAWB.VAWB_PG(raw, excelrow, row, clspos)
        Call VAWB.VAWB_WT(raw, excelrow, row, host, special, clspos)
        Call VAWB.NumPcs(raw, excelrow, row, special)
    End If
    
    If BORG.Can_flight.Value = True Then Call VAWB.canflight(host, excelrow, row, special)
    
    row = row + 1
Loop
Exit Sub

errout:
If Err.Number = 424 Then
    Set host = ReturnHost
    Call VAWB.Directions(host, excelrow)
Else
    MsgBox ("Unhandled Error In Directions : " & Err.Number & vbNewLine _
    & "Desc: " & Err.Description & vbNewLine _
    & "source: " & Err.Source _
    & "help context: " & Err.HelpContext)
End If
End Sub

Sub VerifyOP(host As Variant)

End Sub

Sub VerifyALPK(host As Variant)

End Sub

Sub VerifyAWB(host As Variant)

End Sub
Sub Assembly(host As Variant, excelrow As Integer)
'On Error GoTo errout

Sheet3.Rows("16:99").Clear
Sheet3.Cells(15, 7).Clear

Call GrabAWBlines(host)

row = 17
TwoRow = 16
Do Until Sheet3.Cells(row, 1).Value = ""
    If Trim(Sheet3.Cells(row, 1).Text) = "" Then Exit Do
    x = InStr(1, Sheet3.Cells(row, 1).Value, "RQ")
    If x <> 6 Then x = InStr(1, Sheet3.Cells(row, 1).Value, "UN")
    If x <> 6 And x <> 10 Then x = InStr(1, Sheet3.Cells(row, 1).Value, "ID8000")
    
    If x = 6 Or x = 10 Then
        TwoRow = TwoRow + 1
        Sheet3.Cells(TwoRow, 2) = Trim(Sheet3.Cells(row, 1))
    End If
    
    If x = 0 Then
        Sheet3.Cells(TwoRow, 2) = Sheet3.Cells(TwoRow, 2).Text + " " + Trim(Sheet3.Cells(row, 1).Text)
    End If
    
    Sheet3.Cells(row, 1).Clear
    row = row + 1
Loop

PCS = 0.0001
TwoRow = TwoRow - 16

If TwoRow <> 1 Then
    Do Until TwoRow = 1
        Sheet1.Rows(excelrow).Offset(1).EntireRow.Insert
        Sheet1.Cells(excelrow + 1, 14).Value = Sheet1.Cells(excelrow, 14).Value + PCS
        Sheet1.Cells(excelrow + 1, 15).Value = Sheet1.Cells(excelrow, 15).Value + PCS
        Sheet1.Cells(excelrow + 1, 16).Value = Sheet1.Cells(excelrow, 16).Value + PCS
        Sheet1.Cells(excelrow + 1, 17).Value = Sheet1.Cells(excelrow, 17).Value + PCS
        Sheet1.Cells(excelrow + 1, 1).Value = Sheet1.Cells(excelrow, 1).Value
        Sheet1.Cells(excelrow + 1, 2).Value = Sheet1.Cells(excelrow, 2).Value
        Sheet1.Cells(excelrow + 1, 3).Value = Sheet1.Cells(excelrow, 3).Value
        Sheet1.Cells(excelrow + 1, 6).Value = Sheet1.Cells(excelrow, 6).Value
        Sheet1.Cells(excelrow + 1, 12).Value = Sheet1.Cells(excelrow, 12).Value
        Sheet1.Cells(excelrow + 1, 13).Value = Sheet1.Cells(excelrow, 13).Value
        Sheet1.Cells(excelrow + 1, 23).Value = Sheet1.Cells(excelrow, 23).Value
        PCS = PCS + 0.0001
        TwoRow = TwoRow - 1
    Loop
End If
Exit Sub

errout:
If Err.Number = 424 Then
    Set host = ReturnHost
    Call VAWB.Assembly(host, excelrow)
Else
    MsgBox ("Unhandled Error In Assembly : " & Err.Number & vbNewLine _
    & "Desc: " & Err.Description & vbNewLine _
    & "source: " & Err.Source _
    & "help context: " & Err.HelpContext)
End If

End Sub 'end assembly sub

Sub VAWB_Origin(host As Variant, excelrow As Integer)
'On Error GoTo errout

host.readscreen Origin, 5, 4, 24

Sheet1.Cells(excelrow, 2).Value = Trim(Origin) 'Grab origin station of piece. Or at least who entered the bloody thing
    If Trim(Origin) = "PHXR" Then Sheet1.Cells(excelrow, 12).Value = 1
    If Trim(Origin) = "MSCA" Then Sheet1.Cells(excelrow, 12).Value = 2
    If Trim(Origin) = "LUFA" Then Sheet1.Cells(excelrow, 12).Value = 3
    If Trim(Origin) = "SCFA" Then Sheet1.Cells(excelrow, 12).Value = 4
    If Trim(Origin) = "ZSYA" Then Sheet1.Cells(excelrow, 12).Value = 5
    If Sheet1.Cells(excelrow, 12) = "" Then Sheet1.Cells(excelrow, 12).Value = 6

Exit Sub

errout:
If Err.Number = 424 Then
    Set host = ReturnHost
    Call VAWB_Origin(host, excelrow)
Else
    MsgBox ("Unhandled Error In VAWB origin : " & Err.Number & vbNewLine _
    & "Desc: " & Err.Description & vbNewLine _
    & "source: " & Err.Source _
    & "help context: " & Err.HelpContext)
End If

End Sub 'end origin sub
Sub canflight(host As Variant, excelrow As Integer, row As Integer, special As Integer)

getCanFlight = CanFlightBulk(host)

can = getCanFlight(0)
flight = getCanFlight(1)

If excelrow = 0 Then excelrow = Sheet3.Cells(2, 1).Value

If special = 1 Then
    Sheet1.Cells(excelrow + (row - 17), 19).Value = can
    Sheet1.Cells(excelrow + (row - 17), 20).Value = flight
Else:
    Sheet1.Cells(excelrow, 19).Value = can
    Sheet1.Cells(excelrow, 20).Value = flight
End If

End Sub

Function CanFlightBulk(host As Variant)
'On Error GoTo errout

r = 13
c = 8
BORG.Location.Text = UCase(BORG.Location.Text)
Do Until r = 22
    host.readscreen phxrchk, 4, r, c
    host.readscreen orgchk, Len(BORG.Location), 4, 24
    If orgchk = BORG.Location Then
        retCan = BORG.Location
        retFlight = BORG.Location
        CanFlightBulk = Array(retCan, retFlight)
        Exit Function
    End If
    If phxrchk = BORG.Location.Text Then
        host.readscreen can, 10, r, 14
        host.readscreen flightTruck, 5, r, 35
        retCan = can
        retFlight = flightTruck
        CanFlightBulk = Array(retCan, retFlight)
        Exit Function
    End If
    r = r + 1
    If r = 22 Then
        retCan = "Unknown"
        retFlight = "Unknown"
        CanFlightBulk = Array(retCan, retFlight)
        Exit Function
    End If

Loop
Exit Function

errout:
If Err.Number = 424 Then
    Set host = ReturnHost
    Call CanFlightBulk(host)
Else
    MsgBox ("Unhandled Error In canflightbulk : " & Err.Number & vbNewLine _
    & "Desc: " & Err.Description & vbNewLine _
    & "source: " & Err.Source _
    & "help context: " & Err.HelpContext)
End If

End Function

Sub NumPcs(raw As String, excelrow As Integer, row As Integer, special As Integer)
PCS = Num_Pcs(raw)

If special = 1 Then
   Sheet1.Cells(excelrow + (row - 17), 9).Value = PCS
Else:
   Sheet1.Cells(excelrow, 9).Value = PCS
End If
End Sub

Function Num_Pcs(raw As String)
On Error GoTo errorout
x = InStr(1, raw, " PIECE")
If x < 1 Then
    Num_Pcs = 1
    Exit Function
End If

y = 1
Do Until commacheck = ","
    commacheck = Mid(raw, x - y, 1)
    y = y + 1
Loop

Num_Pcs = Mid(raw, x - y + 3, y - 3)
Exit Function

errorout:
'yes it's the easy way out.... but it's prolly right more often than not..
Num_Pcs = 1

End Function

Sub VAWB_WT(raw As String, excelrow As Integer, row As Integer, host As Variant, special As Integer, clspos As Integer)
'On Error GoTo errout

If InStr(1, raw, "RADIOACTIVE") >= 1 Then
    If InStr(1, raw, "EXCEPTED") = 0 Then
        host.readscreen TInum, 6, 4, 46
        bqloc = InStr(1, raw, "BQ") - 3
        Do Until spacefind = " "
            spacefind = Mid(raw, bqloc, 1)
            bqloc = bqloc - 1
        Loop
        endloc = InStr(1, raw, "BQ") + 2
        bbqnum = Trim(Mid(raw, bqloc, (endloc - bqloc)))
    Else:
        TInum = "EQ"
    End If
        
    x = Array(TInum, "TI")
Else
    x = WTfind(raw, clspos)
End If
 
If special = 1 Then
    Sheet1.Cells(excelrow + (row - 17), 11).Value = x(1)
    Sheet1.Cells(excelrow + (row - 17), 10).Value = x(0)
Else:
    Sheet1.Cells(excelrow, 11).Value = x(1)
    Sheet1.Cells(excelrow, 10).Value = x(0)
End If

Exit Sub
errout:
If Err.Number = 424 Then
    Set host = ReturnHost
    Call VAWB_WT(raw, excelrow, row, host, special, clspos)
Else
    MsgBox ("Unhandled Error In VAWB_WT : " & Err.Number & vbNewLine _
    & "Desc: " & Err.Description & vbNewLine _
    & "source: " & Err.Source _
    & "help context: " & Err.HelpContext)
End If
End Sub

Function WTfind(raw As String, clspos As Integer)
If InStr(1, raw, "RADIOACTIVE") >= 1 And InStr(1, raw, "EXCEPTED") > 1 Then
    WTfind = Array("EQ", "EQ")
    Exit Function
End If

WT = 0
x = 0
UM = 0
Start = 0
last = 0

Start = clspos
If Start <= 0 Then Start = 1

If InStr(Start, raw, " L, ") > 1 Then
    UM = "L"
    last = InStr(Start, raw, " L, ")
ElseIf InStr(Start, raw, " KG, ") > 1 Then
    UM = "KG"
    last = InStr(Start, raw, " KG, ")
ElseIf InStr(Start, raw, " KG G, ") > 1 Then
    UM = " KG G"
    last = InStr(Start, raw, " KG G, ")
ElseIf InStr(Start, raw, " G G, ") > 1 Then
    UM = "G G"
    last = InStr(Start, raw, " G G, ")
ElseIf InStr(Start, raw, " G, ") > 1 Then
    UM = "G"
    last = InStr(Start, raw, " G, ")
ElseIf InStr(Start, raw, " ML, ") > 1 Then
    UM = "ML"
    last = InStr(Start, raw, " ML, ")
Else
    UM = ""
    last = Len(raw)
End If

spacecheck = ""
first = last - 1
Do Until spacecheck = " "
    spacecheck = Mid(raw, first, 1)
    first = first - 1
Loop
WT = Trim(Mid(raw, first + 2, (last - first) - 2))
WTfind = Array(WT, UM)

End Function

Sub VAWB_PG(raw As String, excelrow As Integer, row As Integer, clspos As Integer)
    PG = PGfind(raw, clspos)
    If special = 1 Then
        Sheet1.Cells(excelrow + (row - 17), 8).Value = PG(0)
    Else:
        Sheet1.Cells(excelrow, 8).Value = PG(0)
    End If
End Sub

Function PGfind(raw As String, clspos As Integer)
pgpos = 0
PG = "X"

If InStr(1, raw, "RADIOACTIVE MATERIAL, EXCEPTED PACKAGE") > 1 Then
    PGfind = Array(PG, pgpos)
    Exit Function
End If

If InStr(1, raw, ", III,") > clspos Then
    pgpos = InStr(1, raw, ", III,")
    PG = "III"
ElseIf InStr(1, raw, ", II,") > clspos Then
    pgpos = InStr(1, raw, ", II,")
    PG = "II"
ElseIf InStr(1, raw, ", I,") > clspos Then
    pgpos = InStr(1, raw, ", I,")
    PG = "I"
End If
    
PGfind = Array(PG, pgpos)
End Function

Sub VAWB_PSN(raw As String, excelrow As Integer, row As Integer, unpos As Integer, clspos As Integer, special As Integer)
    PSN = PSNfind(raw, clspos)
    If special = 1 Then
        Sheet1.Cells(excelrow + (row - 17), 5).Value = PSN
    Else:
        Sheet1.Cells(excelrow, 5).Value = PSN
    End If
End Sub

Function PSNfind(raw As String, clspos As Integer) As String
If InStr(1, raw, "RADIOACTIVE MATERIAL, EXCEPTED PACKAGE") > 1 Then
    If Left(raw, 2) = "RQ" Then
        RQ = "RQ - "
    Else
        RQ = ""
    End If
    Sheet3.Cells(16, 4).Value = RQ & "RADIOACTIVE MATERIAL, EXCEPTED PACKAGE"
    PSNfind = RQ & "RADIOACTIVE MATERIAL, EXCEPTED PACKAGE"
    Exit Function
End If
Start = 9
RQ = ""
If Left(raw, 2) = "RQ" Then
    Start = Start + 4
    RQ = "RQ - "
    End If
'If classposition <= -1 Then classposition = 80
If clspos <= 0 Then
    hazcl = Classfind(raw)
    If hazcl(0) = "" Then
        PSNfind = "PSN FIND ERROR"
        Exit Function
    Else
        clspos = hazcl(1)
    End If
End If
PSN = (RQ + Mid(raw, Start, (clspos - Start) - 1))
'Sheet3.Cells(16, 4).Value = RQ + PSN
PSNfind = Trim(PSN)
    
End Function

Function VAWB_UN_RQ(raw As String, excelrow As Integer, row As Integer, special As Integer) As Integer
'On Error GoTo errout:
Start = 1
    If Left(raw, 2) = "RQ" Then
        Start = Start + 4
    End If
    UNnum = Mid(raw, Start, 6)
    If special = 1 Then
        Sheet1.Cells(excelrow + (row - 17), 4).Value = UNnum
    Else:
        Sheet1.Cells(excelrow, 4).Value = UNnum
    End If
    
    VAWB_UN_RQ = Start + 6
    
Exit Function

errout:
If Err.Number = 424 Then
    Set host = ReturnHost
    Call VAWB_UN_RQ(raw, excelrow, row, special)
Else
    MsgBox ("Unhandled Error In VAWB UN RQ : " & Err.Number & vbNewLine _
    & "Desc: " & Err.Description & vbNewLine _
    & "source: " & Err.Source _
    & "help context: " & Err.HelpContext)
End If
End Function

Function VAWB_Class(raw As String, excelrow As Integer, row As Integer, special As Integer) As Integer
Start = 1
    If Left(Sheet3.Cells(row, 2).Value, 2) = "RQ" Then
        Start = Start + 4
    End If
    
    RADchk = InStr(1, raw, "EXCEPTED PACKAGE")
    If RADchk >= 1 Then
        classinfo = Array("0", RADchk)
        hazclass = classinfo(0)
        GoTo hazclassAssign
    End If
    
    classinfo = Classfind(raw)
    hazclass = classinfo(0)
    
    VAWB_Class = classinfo(1)
     
hazclassAssign:
    If special = 1 Then
        Sheet1.Cells(excelrow + (row - 17), 7).Value = hazclass
    Else:
        Sheet1.Cells(excelrow, 7).Value = hazclass
    End If
    VAWB_Class = classinfo(1)
End Function

Function Classfind(raw As String)
'On Error GoTo errout:
Subend = 1

Subclass = Array(", 1.4B, ", ", 1.4C, ", ", 1.4D, ", ", 1.4E, ", ", 1.4G, ", _
    ", 1.4S, ", ", 2.1, ", ", 2.2, ", ", 3, ", ", 4.1, ", ", 4.2, ", ", 4.3, ", _
    ", 5.1, ", ", 5.2, ", ", 6.1, ", ", 6.2, ", ", 7, ", ", 8, ", ", 9, ", ", 1.4B(", _
    ", 1.4C(", ", 1.4D(", ", 1.4E(", ", 1.4G(", ", 1.4S(", ", 2.1(", ", 2.2(", _
    ", 3(", ", 4.1(", ", 4.2(", ", 4.3(", ", 5.1(", ", 5.2(", ", 6.1(", ", 6.2(", _
    ", 7(", ", 8(", ", 9(")

For Each class In Subclass
    classposition = InStr(1, raw, class)
    If classposition > 1 Then
        classposition = classposition + 1
        If InStr(1, class, "(") >= 1 Then
            Do Until endcheck = ")"
                endcheck = Mid(raw, classposition + Subend, 1)
                If endcheck = ")" Then Exit Do
                Subend = Subend + 1
            Loop
            hazclass = Mid(raw, classposition + 1, classposition - (classposition - Subend))
        Else
            hazclass = class
        End If
        Exit For
    End If
Next
hazclass = Trim(Replace(hazclass, ",", ""))

Classfind = Array(hazclass, classposition)
Exit Function
errout:
Classfind = Array("0", 0)
If Err.Number = 424 Then
    'shouldn't happen no host command called
Else
    MsgBox ("Unhandled Error In classfind : " & Err.Number & vbNewLine _
    & "Desc: " & Err.Description & vbNewLine _
    & "source: " & Err.Source _
    & "help context: " & Err.HelpContext)
End If


End Function

