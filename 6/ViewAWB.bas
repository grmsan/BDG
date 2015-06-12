Attribute VB_Name = "ViewAWB"
Dim host As Variant
Dim row As Integer
Dim excelrow As Integer
Option Compare Text

Sub Directions(host As Variant)
excelrow = Sheet3.Cells(2, 1).Value
row = 17
'''''''''''''''Step 1''''''''''''''''
''''Section verifies package vs excel
'if OP send to VerifyOP sub
'if ALPK send to VerifyALPK sub
'if not either OP or ALPK then send to VerifyAWB sub
'''''''''''''''Step 2''''''''''''''''
''''checks next line for continuity with current line
'check how many pieces in this shipment set piece count to maxpieces
'check that line has UN or RQ and grab line info
''grab below previous line
'''determine if no data exsists on that line at all
'''determine if it's a new line with a new RQ/UN number or...
'''determine if it's part of the previous shipment or....
''''''''''''''Step 3'''''''''''''''''
''''Section starts grabbing information for updated more accurate list
'send info to VAWB_Origin to grab origin ID (assign number to station id for later sort)
'sendto VAWB_UN_RQ sub
'sendto VAWB_Class sub
'sendto VAWB_PSN sub
'sendto VAWB_PG sub
'sendto VAWB_WT sub
'if maxpieces = currentpiece then go to step 2
''''''''''''''Step 4'''''''''''''''''
'send line of information to EQfix
  
host.waitready 1, 51

Special = 0
If Sheet1.Cells(excelrow, 16).Value > 0 Then
    'Call ViewAWB.VerifyOP(host)
    Special = 1
    End If
If Sheet1.Cells(excelrow, 14).Value > 0 Then
    'Call ViewAWB.VerifyALPK(host)
    Special = 1
    End If
'If Special = 0 Then Call ViewAWB.VerifyAWB(host)

Call ViewAWB.VAWB_Origin(host)
Call ViewAWB.Assembly(host)

If BORG.Can_flight.Value = True Then Call ViewAWB.CanFlightBulk(host)

row = 17
Do Until Sheet3.Cells(row, 2).Value = ""

Call ViewAWB.VAWB_UN_RQ
Call ViewAWB.VAWB_Class
Call ViewAWB.VAWB_PSN
If BORG.Can_flight.Value = True Then Call ViewAWB.canflight

'If Sheet3.Cells(16, 1).Value <> "Normal" Then
    Call ViewAWB.VAWB_PG
    Call ViewAWB.VAWB_WT(host)
If Special = 1 Then Call ViewAWB.NumPcs
'End If

Increment:
row = row + 1
Loop

End Sub

Sub VAWB_UN_RQ()
Start = 1
    If Left(Sheet3.Cells(row, 2).Value, 2) = "RQ" Then
        Start = Start + 4
    End If
     
     UNnum = Mid(Sheet3.Cells(row, 2).Value, Start, 6)
     If Sheet3.Cells(16, 1).Value <> "Normal" Then
        Sheet1.Cells(excelrow + (row - 17), 4).Value = UNnum
        Else: Sheet1.Cells(excelrow, 4).Value = UNnum
        End If
End Sub 'end un rq sub

Sub VAWB_Class()
Start = 1
    If Left(Sheet3.Cells(row, 2).Value, 2) = "RQ" Then
    Start = Start + 4
    End If
    rawData = Sheet3.Cells(row, 2).Value
    If InStr(1, rawData, "RADIOACTIVE MATERIAL, EXCEPTED PACKAGE") >= 1 Then
        Sheet3.Cells(16, 5).Value = "0"
        GoTo hazclassAssign
    End If
    Classfind (Sheet3.Cells(row, 2).Value)
     
hazclassAssign:
     hazclass = Sheet3.Cells(16, 5).Value
     If Sheet3.Cells(16, 1).Value <> "Normal" Then
        Sheet1.Cells(excelrow + (row - 17), 7).Value = hazclass
        Else: Sheet1.Cells(excelrow, 7).Value = hazclass
        End If
End Sub 'end class sub

Sub VAWB_PSN()
    PSN = PSNfind(Sheet3.Cells(row, 2).Value)
    If Sheet3.Cells(16, 1).Value <> "Normal" Then
       Sheet1.Cells(excelrow + (row - 17), 5).Value = PSN
    Else: Sheet1.Cells(excelrow, 5).Value = PSN
    End If
End Sub 'end psn sub

Sub VAWB_PG()
PGfind (Sheet3.Cells(row, 2).Value)
PG = Sheet3.Cells(16, 7).Value

If Sheet3.Cells(16, 1).Value <> "Normal" Then
   Sheet1.Cells(excelrow + (row - 17), 8).Value = PG
Else:
   Sheet1.Cells(excelrow, 8).Value = PG
End If

End Sub

Sub VAWB_WT(host As Variant)
On Error GoTo exitWT

'u = Sheet3.Cells(17, 2).Value

If Sheet3.Cells(16, 5).Value = "7" Or InStr(1, Sheet3.Cells(16, 5), "7(") >= 1 Then
    host.readscreen TInum, 6, 4, 46
    bbqnum = Mid(myr, rad2 + 3, rad - rad2 + 2)
    x = Array(TInum, "TI")
Else
    x = WTfind(Sheet3.Cells(row, 2).Value)
End If

Sheet3.Cells(16, 9).Value = x(0) 'sets WT
Sheet3.Cells(16, 10).Value = x(1) 'sets UM
 
If Sheet3.Cells(16, 1).Value <> "Normal" Then
    Sheet1.Cells(excelrow + (row - 17), 11).Value = x(1)
    Sheet1.Cells(excelrow + (row - 17), 10).Value = x(0)
Else:
    Sheet1.Cells(excelrow, 11).Value = x(1)
    Sheet1.Cells(excelrow, 10).Value = x(0)
End If

exitWT:

End Sub 'end pg sub

Sub NumPcs()
PCS = Num_Pcs(Sheet3.Cells(row, 2).Value)

If Sheet3.Cells(16, 1).Value <> "Normal" Then
   Sheet1.Cells(excelrow + (row - 17), 9).Value = PCS
Else:
   Sheet1.Cells(excelrow, 9).Value = PCS
End If
End Sub 'end pcs sub

Sub canflight()

can = Sheet3.Cells(4, 1)
flight = Sheet3.Cells(4, 2)

If excelrow = 0 Then excelrow = Sheet3.Cells(2, 1).Value
'u = Sheet3.Cells(16, 1).Value
If Sheet3.Cells(16, 1).Value <> "Normal" Then
    Sheet1.Cells(excelrow + (row - 17), 19).Value = can
    Sheet1.Cells(excelrow + (row - 17), 20).Value = flight
Else:
    Sheet1.Cells(excelrow, 19).Value = can
    Sheet1.Cells(excelrow, 20).Value = flight
End If

End Sub

Sub CanFlightBulk(Optional host As Variant)
r = 13
c = 8
BORG.Location.Text = UCase(BORG.Location.Text)
Do Until r = 22

    host.readscreen phxrchk, 4, r, c
    host.readscreen orgchk, Len(BORG.Location), 4, 24
    If orgchk = BORG.Location Then
        Sheet3.Cells(4, 1).Value = BORG.Location
        Sheet3.Cells(4, 2).Value = BORG.Location
        Exit Sub
    End If
    If phxrchk = BORG.Location.Text Then
        host.readscreen can, 10, r, 14
        host.readscreen flightTruck, 5, r, 35
        Sheet3.Cells(4, 1).Value = can
        Sheet3.Cells(4, 2).Value = flightTruck
        Exit Sub
    End If
    r = r + 1
    If r = 22 Then
        Sheet3.Cells(4, 1).Value = "Unknown"
        Sheet3.Cells(4, 2).Value = "Unknown"
    End If

Loop

End Sub
Sub VerifyAWB(Optional host As Variant)
verify = 0
RQcheck = 0
URSAcheck = 0
UNcheck = 0
awbcheck = 0
bluerow = 6
lineread = 0
firstrun = 0
If excelrow = 0 Then
    excelrow = Sheet3.Cells(2, 1).Value
End If
If Trim(Sheet1.Cells(excelrow, 4).Text) = "UN1845" Then
    Sheet1.Cells(excelrow, 5).Value = "Dry Ice"
    Exit Sub
End If

GoTo VAWBMainLoop

SearchVAWB:
firstrun = 1
Do Until lastshipment = "305"
    host.sendkey "@1"
    host.waitready 1, 51
    host.readscreen lastshipment, 3, 24, 2
Loop

Do Until lastshipment = "306"
host.readscreen lastshipment, 3, 24, 2

VAWBMainLoop:
    col = 6
    host.readscreen awbcheck, 4, 4, 14
    host.readscreen URSAcheck, 8, 4, 35
    host.readscreen lineread, 80, bluerow, 1
    host.readscreen lineread2, 80, bluerow + 1, 1
    host.readscreen RQcheck, 2, 6, 6 'Checking for RQ
    host.readscreen normCheck, 19, 4, 61
    norm = Trim(normCheck) 'should = ""

    myLine = Trim(lineread) & " " & Trim(lineread2)

    If RQcheck = "RQ" Then 'if RQ is found then move UN starting position over accordingly
        col = col + 4 'Moving col over for UNcheck to work properly
    Else: col = 6
    End If

    host.readscreen UNcheck, 6, bluerow, col

    x = (" " & Sheet1.Cells(excelrow, 9).Value & " PIECE")
    pcs_chk = InStr(1, myLine, x)
        If pcs_chk = 0 Then
            itercheck = 0
            pcsrow = 13
            Do Until Trim(misc) = Trim(Sheet1.Cells(excelrow, 9))
                host.readscreen misc, 3, pcsrow, 26
                pcs_chk = InStr(1, misc, Sheet1.Cells(excelrow, 9).Value)
                If pcs_chk >= 1 Then Exit Do
                pcsrow = pcsrow + 1
            Loop
        End If

    WTchk = InStr(1, myLine, Sheet1.Cells(excelrow, 10).Value)
        'U = Sheet1.Cells(ExcelRow, 10).Value
        If WTchk = 0 Then
            host.readscreen lineread3, 78, bluerow + 2, 1
            WTchk = InStr(1, lineread3, Sheet1.Cells(excelrow, 10).Value)
        End If

PSNTEMP = Trim(Sheet1.Cells(excelrow, 5).Value)
If PSNTEMP = "RADIOACTIV" Or PSNTEMP = "Radioactive, Excepted Qty" Then PSNTEMP = "RADIOACTIVE"
    PSNchk = InStr(1, lineread, PSNTEMP)
    If PSNchk > 1 Then
        If awbcheck = Sheet1.Cells(excelrow, 3).Text Then  'checks last 4 awb with those on assign screen
            UNTEMP = Sheet1.Cells(excelrow, 4).Value
            If Sheet1.Cells(excelrow, 4).Value = UNcheck Then 'checks UNnumber
                If Trim(Sheet1.Cells(excelrow, 6).Value) = Trim(URSAcheck) Then 'checks ursa from that from assign screen
                    If (WTchk >= 1 Or PSNTEMP = "RADIOACTIVE") And norm = "" Then
                        'y = locCHK(host)
                        'If y = True Then
                            verify = 1 'if all checks passed then we have verified our shipment
                            Exit Sub
                        'End If
                    End If
                End If
            End If
        End If
    End If
    If verify = 0 Then
        host.sendkey "@2"
        host.waitready 1, 51
        If lastshipment = "306" Then GoTo SearchVAWB
        If firstrun = 0 Then GoTo SearchVAWB
    End If
Loop
End Sub 'end verify awb sub

Sub VerifyOP(Optional host As Variant)
verify = 0
OPcheck = 0
OP_ID = 0
OPpcs = 0
URSAcheck = 0
firstrun = 0
GoTo verifyingOP

SearchVAWB_OP:
firstrun = 1
    Do Until lastshipment = "305"
        host.sendkey "@1"
        host.waitready 1, 51
        host.readscreen lastshipment, 3, 24, 2
    Loop

Do Until lastshipment = "306"
verifyingOP:
    host.readscreen lastshipment, 3, 24, 2
    host.readscreen OPcheck, 2, 4, 63
    host.readscreen URSAcheck, 8, 4, 35
    host.readscreen OP_ID, 3, 4, 66
    host.readscreen OPpcs, 3, 4, 77
    If OPcheck <> "  " Then
        OP_ID = CInt(OP_ID)
        OPpcs = CInt(OPpcs)
'        u = Sheet1.Cells(excelrow, 16).Value
'        uu = Sheet1.Cells(excelrow, 6).Value
'        uuu = Sheet1.Cells(excelrow, 17).Value
        If (Trim(Sheet1.Cells(excelrow, 6).Value) = Trim(URSAcheck)) And _
           (Sheet1.Cells(excelrow, 16).Value = OP_ID) And _
           (Sheet1.Cells(excelrow, 17).Value = OPpcs) Then
                verify = 1
                Exit Do
        End If
    End If
If verify = 0 Then
    host.sendkey "@2"
    host.waitready 1, 51
    If lastshipment = "306" Then GoTo SearchVAWB_OP
    If firstrun = 0 Then GoTo SearchVAWB_OP:
End If

Loop

End Sub 'end verifyOP sub

Sub VerifyALPK(Optional host As Variant)
verify = 0
ALPKcheck = 0
ALPK_ID = 0
ALPKpcs = 0
URSAcheck = 0
firstrun = 0

GoTo VerifyingAWB_ALPK
SearchVAWB:
firstrun = 1
    Do Until lastshipment = "305"
    host.sendkey "@1"
    host.waitready 1, 51
    host.readscreen lastshipment, 3, 24, 2
    Loop
    
Do Until lastshipment = "306"
VerifyingAWB_ALPK:
    host.readscreen lastshipment, 3, 24, 2
    host.readscreen ALPKcheck, 4, 4, 63
    host.readscreen URSAcheck, 8, 4, 35
    host.readscreen ALPK_ID, 3, 4, 66
    host.readscreen ALPKpcs, 3, 4, 77
    If ALPK_ID <> "   " Then
        ALPK_ID = CInt(ALPK_ID)
        ALPKpcs = CInt(ALPKpcs)
        URSAcheck = Trim(URSAcheck)
        If (Trim(Sheet1.Cells(excelrow, 6).Value) = Trim(URSAcheck)) And _
           (Sheet1.Cells(excelrow, 14).Value = ALPK_ID) And _
           (Sheet1.Cells(excelrow, 15).Value = ALPKpcs) Then
                verify = 1
                Exit Do
        End If
    End If
    
    If verify = 0 Then
        host.sendkey "@2"
        host.waitready 1, 51
        If lastshipment = "306" Then GoTo VerifyingAWB_ALPK
        If firstrun = 0 Then GoTo VerifyingAWB_ALPK
    End If
Loop
End Sub 'end verifyALPKN1 sub
Sub Assembly(host As Variant)
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

End Sub 'end assembly sub
Sub VAWB_Origin(Optional host As Variant)
host.readscreen Origin, 5, 4, 24

Sheet1.Cells(excelrow, 2).Value = Trim(Origin) 'Grab origin station of piece. Or at least who entered the bloody thing
    If Trim(Origin) = "PHXR" Then Sheet1.Cells(excelrow, 12).Value = 1
    If Trim(Origin) = "MSCA" Then Sheet1.Cells(excelrow, 12).Value = 2
    If Trim(Origin) = "LUFA" Then Sheet1.Cells(excelrow, 12).Value = 3
    If Trim(Origin) = "SCFA" Then Sheet1.Cells(excelrow, 12).Value = 4
    If Trim(Origin) = "ZSYA" Then Sheet1.Cells(excelrow, 12).Value = 5
    If Sheet1.Cells(excelrow, 12) = "" Then Sheet1.Cells(excelrow, 12).Value = 6

End Sub 'end origin sub
Function GrabAWBlines(host As Variant)
host.readscreen whatisthis, 4, 4, 61
If Trim(whatisthis) = "" Then
    whatisthis = "Normal"
Else 'if not a normal piece get ID and PC count and put them in excel
    host.readscreen idnum, 3, 4, 66
    host.readscreen PCS, 3, 4, 77
    Sheet3.Cells(16, 2).Value = idnum
    Sheet3.Cells(16, 3).Value = PCS
End If

Sheet3.Cells(16, 1).Value = whatisthis
bluerow = 6
ERow = 17
host.readscreen readline, 80, bluerow, 1
linedata = "temp to not exit do"
Do Until linedata = ""
    host.readscreen readline, 80, bluerow, 1
    linedata = Trim(readline)
    If linedata = "" Then Exit Do
    Sheet3.Cells(ERow, 1).Value = readline
    host.readscreen miscdata, 3, 24, 2
    If miscdata = "490" And bluerow = 11 Then
        host.sendkey "@8"
        host.waitready 1, 51
        bluerow = 5
    End If
    bluerow = bluerow + 1
    ERow = ERow + 1
    
Loop

End Function
