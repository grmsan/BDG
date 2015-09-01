
Dim row As Integer
Dim excelrow As Integer
Option Compare Text

Sub Directions(excelrow As Integer)
'On Error GoTo errout

Dim special As Integer
special = 0
If Sheet1.Cells(excelrow, 16).Value > 0 Then
    Call VAWB.VerifySpecial(excelrow, True)
    special = 1
    End If
If Sheet1.Cells(excelrow, 14).Value > 0 Then
    Call VAWB.VerifySpecial(excelrow, False)
    special = 1
    End If
If special = 0 Then Call VAWB.VerifyAWB(excelrow)

Call VAWB.VAWB_Origin(excelrow)
Call VAWB.Assembly(excelrow)

Dim row As Integer
row = 17

Dim unpos As Integer
Dim clspos As Integer

Dim raw As String
raw = ""

Do Until Sheet3.Cells(row, 2).Value = ""
    raw = Sheet3.Cells(row, 2).text
    unpos = VAWB_UN_RQ(raw, excelrow, row, special)
    clspos = VAWB.VAWB_Class(raw, excelrow, row, special)
    Call VAWB.VAWB_PSN(raw, excelrow, row, unpos, clspos, special)
    Call VAWB.VAWB_WT(raw, excelrow, row, special, clspos)
    If special = 1 Then
        Call VAWB.VAWB_PG(raw, excelrow, row, clspos)
        
        Call VAWB.NumPcs(raw, excelrow, row, special)
    End If
    
    If BORG.Can_flight.Value = True Then Call VAWB.canflight(excelrow, row, special)
    
    row = row + 1
Loop
Exit Sub

errout:
    MsgBox ("Unhandled Error In Directions : " & Err.Number & vbNewLine _
    & "Desc: " & Err.Description & vbNewLine _
    & "source: " & Err.Source _
    & "help context: " & Err.HelpContext)
End Sub

Sub VerifySpecial(excelrow As Integer, OP As Boolean)
On Error GoTo errout

If OP = True Then
    datarow = 16
Else
    datarow = 14
End If

shipend = 0
shipbegin = 0
verify = 0
opcheck = 0
OP_ID = 0
OPpcs = 0
URSAcheck = 0
firstrun = 0
startTime = Minute(Time()) + (0.01 * Second(Time()))

GoTo verifyingOP 'let's first check and see if we got lucky with whatever shows up first...

SearchVAWB_OP:
firstrun = 1

Do Until verify = 1 Or (shipbegin = 1 And shipend = 1)
    lastshipment = BZreadscreen(3, 24, 2)
    If lastshipment = "305" Then
        shipbegin = 1 'we've reached the beginning
    ElseIf lastshipment = "306" Then
        shipend = 1 'we've reached the end
    End If
    
    If shipbegin = 0 Then
        Call BZsendKey("@1")
    ElseIf shipend = 0 Then
        Call BZsendKey("@2")
    Else
        Exit Sub 'we fubbed up
    End If
    
verifyingOP:
    opcheck = BZreadscreen(3, 24, 2)
    URSAcheck = BZreadscreen(8, 4, 35)
    OP_ID = BZreadscreen(3, 4, 66)
    OPpcs = BZreadscreen(3, 4, 77)
    
    If OP_ID <> "   " And OPpcs <> "   " Then
        OP_ID = CInt(OP_ID)
        OPpcs = CInt(OPpcs)
        u = Trim(Sheet1.Cells(excelrow, 6).Value)
        ul = Trim(URSAcheck)
        uu = Sheet1.Cells(excelrow, datarow).Value
        uul = OP_ID
        uuu = Sheet1.Cells(excelrow, datarow + 1).Value
        uuul = OPpcs
        If (Trim(Sheet1.Cells(excelrow, 6).Value) = Trim(URSAcheck)) _
        And (Sheet1.Cells(excelrow, datarow).Value = OP_ID) Then
            verify = 1
            Exit Do
        End If
    End If
    
    'curtime = Minute(Time()) + (0.01 * Second(Time()))
    'If Abs(curtime - startTime) > 0.1 Then Exit Sub 'we are taking too long.... move it along
        
    If verify = 0 And firstrun = 0 Then
        GoTo SearchVAWB_OP 'we did not find what we wanted lets start searching....
    End If
Loop

Exit Sub
errout:
If Err.Number = 13 Then 'trying to int a str
    MsgBox ("error 13 in Verify OP " & Err.Description)
End If
End Sub

Sub VerifyAWB(excelrow)
On Error GoTo errout

URSAcheck = 0
RQcheck = 0
UNcheck = 0
awbcheck = 0

bluerow = 6
lineread = 0

shipend = 0
shipbegin = 0
verify = 0

firstrun = 0
'startTime = Minute(Time()) + (0.01 * Second(Time()))

GoTo verifyingAWB 'let's first check and see if we got lucky with whatever shows up first...

SearchVAWB:
firstrun = 1

Do Until verify = 1 Or (shipbegin = 1 And shipend = 1)
    lastshipment = BZreadscreen(3, 24, 2)
    If lastshipment = "305" Then
        shipbegin = 1 'we've reached the beginning
    ElseIf lastshipment = "306" Then
        shipend = 1 'we've reached the end
    End If
    
    If shipbegin = 0 Then
        Call BZsendKey("@1")
    ElseIf shipend = 0 Then
        Call BZsendKey("@2")
    Else
        Exit Sub 'we fubbed up
    End If

verifyingAWB:
    col = 6
    awbcheck = BZreadscreen(12, 4, 6)
    If awbcheck = Sheet1.Cells(excelrow, 1).text Then
    
        'URSAcheck = BZreadscreen(8, 4, 35)
        row = 12
        loc_check = ""
        next_check = ""
        Do Until row = 22 Or (loc_check = BORG.Location And next_check = "")
            row = row + 1
            loc_check = Trim(BZreadscreen(5, row, 2))
            next_check = Trim(BZreadscreen(5, row + 1, 2))
            
            If loc_check = "" Then Exit Do
        Loop
        If (loc_check = BORG.Location And next_check = "") Then
            lineinfo = BZreadscreen(38, row, 2)
            If InStr(1, lineinfo, "MAINT") >= 1 Or InStr(1, lineinfo, "DELETE") >= 1 Then
                verify = 0
            Else
                verify = 1
            End If
        Else
            verify = 0 'pass
        End If
    End If
Loop

Exit Sub
errout:
If Err.Number = 13 Then 'trying to int a str
    MsgBox ("error 13 in Verify OP " & Err.Description)
End If
End Sub
Sub Assembly(excelrow As Integer)
'On Error GoTo errout

Sheet3.Rows("16:99").Clear
Sheet3.Cells(15, 7).Clear

Call GrabAWBlines

row = 17
TwoRow = 16
Do Until Sheet3.Cells(row, 1).Value = ""
    If Trim(Sheet3.Cells(row, 1).text) = "" Then Exit Do
    x = InStr(1, Sheet3.Cells(row, 1).Value, "RQ")
    If x <> 6 Then x = InStr(1, Sheet3.Cells(row, 1).Value, "UN")
    If x <> 6 And x <> 10 Then x = InStr(1, Sheet3.Cells(row, 1).Value, "ID8000")
    
    If x = 6 Or x = 10 Then
        TwoRow = TwoRow + 1
        Sheet3.Cells(TwoRow, 2) = Trim(Sheet3.Cells(row, 1))
    End If
    
    If x = 0 Then
        Sheet3.Cells(TwoRow, 2) = Sheet3.Cells(TwoRow, 2).text + " " + Trim(Sheet3.Cells(row, 1).text)
    End If
    
    Sheet3.Cells(row, 1).Clear
    row = row + 1
Loop

PCS = 0.0001
TwoRow = TwoRow - 16

If TwoRow <> 1 Then
    Do Until TwoRow <= 1
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
    MsgBox ("Unhandled Error In Assembly : " & Err.Number & vbNewLine _
    & "Desc: " & Err.Description & vbNewLine _
    & "source: " & Err.Source _
    & "help context: " & Err.HelpContext)

End Sub 'end assembly sub

Sub VAWB_Origin(excelrow As Integer)
'On Error GoTo errout
Origin = BZreadscreen(5, 4, 24)

Sheet1.Cells(excelrow, 2).Value = Trim(Origin) 'Grab origin station of piece. Or at least who entered the bloody thing
    If Trim(Origin) = "PHXR" Then Sheet1.Cells(excelrow, 12).Value = 1
    If Trim(Origin) = "MSCA" Then Sheet1.Cells(excelrow, 12).Value = 2
    If Trim(Origin) = "LUFA" Then Sheet1.Cells(excelrow, 12).Value = 3
    If Trim(Origin) = "SCFA" Then Sheet1.Cells(excelrow, 12).Value = 4
    If Trim(Origin) = "ZSYA" Then Sheet1.Cells(excelrow, 12).Value = 5
    If Sheet1.Cells(excelrow, 12) = "" Then Sheet1.Cells(excelrow, 12).Value = 6

Exit Sub

errout:
    MsgBox ("Unhandled Error In VAWB origin : " & Err.Number & vbNewLine _
    & "Desc: " & Err.Description & vbNewLine _
    & "source: " & Err.Source _
    & "help context: " & Err.HelpContext)

End Sub 'end origin sub
Sub canflight(excelrow As Integer, row As Integer, special As Integer)

getCanFlight = CanFlightBulk()

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

Function CanFlightBulk()
'On Error GoTo errout
Dim r As Integer
Dim c As Integer

loc_check = ""
col_check = BZreadscreen(10, 13, 42)
If Trim(col_check) = "" Then
    c = 8
Else
    c = 48
End If

BORG.Location.text = UCase(BORG.Location.text)
orgchk = BZreadscreen(Len(BORG.Location), 4, 24)
looper:
r = 22
Do Until r = 12
    loc_check = Trim(BZreadscreen(5, r, c))
    If orgchk = BORG.Location Then
        retCan = BORG.Location
        retFlight = BORG.Location
        CanFlightBulk = Array(retCan, retFlight)
        Exit Function
    End If
    If loc_check = BORG.Location.text Then
        can = BZreadscreen(10, r, c + 6)
        flighttruck = BZreadscreen(5, r, c + 27)
        If Trim(flighttruck) = "" Then
            flighttruck = Trim(BZreadscreen(5, r, c - 6))
        End If
        retCan = Trim(can)
        retFlight = flighttruck
        CanFlightBulk = Array(retCan, retFlight)
        Exit Function
    End If
    r = r - 1
    If r = 12 Then
        If c = 8 Then
            retCan = "Unknown"
            retFlight = "Unknown"
            CanFlightBulk = Array(retCan, retFlight)
            Exit Function
        Else
            c = 8
            GoTo looper
        End If
    End If
Loop
Exit Function

errout:
    MsgBox ("Unhandled Error In canflightbulk : " & Err.Number & vbNewLine _
    & "Desc: " & Err.Description & vbNewLine _
    & "source: " & Err.Source _
    & "help context: " & Err.HelpContext)

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

Sub VAWB_WT(raw As String, excelrow As Integer, row As Integer, special As Integer, clspos As Integer)
'On Error GoTo errout

If InStr(1, raw, "RADIOACTIVE") >= 1 Then
    If InStr(1, raw, "EXCEPTED") = 0 Then
        TInum = BZreadscreen(6, 4, 46)
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
    MsgBox ("Unhandled Error In VAWB_WT : " & Err.Number & vbNewLine _
    & "Desc: " & Err.Description & vbNewLine _
    & "source: " & Err.Source _
    & "help context: " & Err.HelpContext)

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
    MsgBox ("Unhandled Error In VAWB UN RQ : " & Err.Number & vbNewLine _
    & "Desc: " & Err.Description & vbNewLine _
    & "source: " & Err.Source _
    & "help context: " & Err.HelpContext)

End Function

Function VAWB_Class(raw As String, excelrow As Integer, row As Integer, special As Integer) As Integer
Start = 1
    If Left(Sheet3.Cells(row, 2).Value, 2) = "RQ" Then
        Start = Start + 4
    End If
    
    RADchk = InStr(1, raw, "EXCEPTED PACKAGE")
    If RADchk >= 1 Then
        classinfo = Array("0", RADchk)
        HazClass = classinfo(0)
        GoTo hazclassAssign
    End If
    
    classinfo = Classfind(raw)
    HazClass = classinfo(0)
    
    VAWB_Class = classinfo(1)
     
hazclassAssign:
    If special = 1 Then
        Sheet1.Cells(excelrow + (row - 17), 7).Value = HazClass
    Else:
        Sheet1.Cells(excelrow, 7).Value = HazClass
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
            HazClass = Mid(raw, classposition + 1, classposition - (classposition - Subend))
        Else
            HazClass = class
        End If
        Exit For
    End If
Next

If InStr(1, HazClass, "(") >= 1 Then
    'MsgBox (hazclass)
Else
    HazClass = Trim(Replace(HazClass, ",", ""))
End If

Classfind = Array(HazClass, classposition)
Exit Function
errout:
    Classfind = Array("0", 0)
    
    MsgBox ("Unhandled Error In classfind : " & Err.Number & vbNewLine _
    & "Desc: " & Err.Description & vbNewLine _
    & "source: " & Err.Source _
    & "help context: " & Err.HelpContext)
End Function



