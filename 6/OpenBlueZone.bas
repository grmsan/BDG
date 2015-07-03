Dim excelrow As Integer
Dim host As Variant
Dim retval As Variant
Private Declare Function MessageBox _
Lib "User32" Alias "MessageBoxA" _
(ByVal hWnd As Long, _
ByVal lpText As String, _
ByVal lpCaption As String, _
ByVal wType As Long) _
As Long

Option Compare Text

Sub BZcloseSessions()

Set BZ = CreateObject("BZWhll.WhllObj")

With BZ
    .waitready 1, 51
    .CloseSession 0, 11
    BORG.loginStatusOff.Visible = True
BORG.labelUpdater.Caption = "Closing Previous Sesson..."
Application.Wait Now + TimeValue("00:00:01")
End With

End Sub
Sub BZOpenSession()
ChDir "C:\"
Set host = CreateObject("BZwhll.whllobj")

retval = host.OpenSession(0, 11, "fdx3270.zmd", 30, 1)
host.WaitCursor 1, 9, 1, 1
retval = host.Connect("K")
Set Wnd = host.Window()
Wnd.Visible = True
Wnd.Caption = "BDG is Searching"
Wnd.State = 0 ' 0 restore, 1 minimize, 2 maximize
If (retval) Then
    host.MsgBox "Error connecting to Session K!", 48
    End If
host.waitready 1, 500

EmployeeNumber = BORG.EmpNum
IMSPassword = BORG.PasswordBox
    
Call OpenBlueZone.Terminal_check(host)

ResultCode = host.Connect("K")
If (ResultCode <> 0) Then
    OpenBlueZone.BZcloseSessions
    x = MsgBox("Error!" & vbNewLine & "Unable to connect to bluezone!" _
        & vbNewLine & "Please try and log in again.", vbCritical, "Error!")
    Exit Sub
End If

host.readscreen fedex, 15, 8, 33
iter = 0
Do Until fedex = "FEDERAL EXPRESS"
    host.waitready 1, 51
    host.readscreen fedex, 15, 8, 33
    iter = iter + 1
    If iter >= 25 Then
        OpenBlueZone.BZcloseSessions
        x = MsgBox("Error!" & vbNewLine & "Unable to connect to bluezone!" _
            & vbNewLine & "Please try and log in again.", vbCritical, "Error!")
        Exit Sub
    End If
Loop


host.writescreen "stsa", 9, 1
host.sendkey "@E"
host.waitready 1, 51

host.readscreen fedex, 35, 1, 23
iter = 0
Do Until fedex = "F E D E R A L  E X P R E S S  I M S"
    host.waitready 1, 51
    host.readscreen fedex, 35, 1, 23
    iter = iter + 1
    If iter >= 25 Then
        OpenBlueZone.BZcloseSessions
        x = MsgBox("Error!" & vbNewLine & "Unable to connect to bluezone!" _
            & vbNewLine & "Please try and log in again.", vbCritical, "Error!")
        Exit Sub
    End If
Loop


host.writescreen EmployeeNumber, 7, 15
host.writescreen IMSPassword, 7, 43
host.sendkey "@E"
host.waitready 1, 51


host.readscreen Enter, 5, 14, 15
iter = 0
Do Until Enter = "ENTER"
    host.waitready 1, 51
    host.readscreen fedex, 35, 1, 23
    iter = iter + 1
    If iter >= 25 Then
        OpenBlueZone.BZcloseSessions
        MsgBox "Incorrect Login Credentials"
        Exit Sub
    End If
Loop

BORG.labelUpdater.Caption = "Going to AutoDG..."
'host.sendkey "26" 'DG training for testing comment this whole line before publishing
host.sendkey "68" 'Live AutoDG
host.sendkey "@E"
host.waitready 1, 51

host.readscreen DGscreen, 12, 2, 32



host.sendkey "assign"
host.writescreen BORG.Location.Text, 19, 44
If BORG.printerID <> "" Then host.writescreen BORG.printerID.Text, 21, 32
host.sendkey "@E"
host.waitready 1, 51
    
Call DGscreenChooser("close", host)
Call GrabCloseScreen

BORG.labelUpdater.Caption = "BDG is now connected to Bluezone Session."
BORG.tgl_btnLogin.Value = False
BORG.loginStatusOff.Visible = False
End Sub

Function GrabReconcile(excelrow As Integer, Optional host As Variant = Empty) As Integer

If TypeName(host) <> "IWhllObj" Then
    ChDir "C:\"
    Set host = CreateObject("BZwhll.whllobj")
    retval = host.OpenSession(0, 11, "fdx3270.zmd", 30, 1)
    host.WaitCursor 1, 9, 1, 1
    retval = host.Connect("K")
    Set Wnd = host.Window() ' Makes the window invisible.....
End If

'Call OpenBlueZone.Terminal_check(host)
Dim bluerow As Integer
bluerow = 6
host.readscreen cannum, 10, 4, 9

Dim SeqFinished As String
host.readscreen SeqFinished, 26, 24, 2
BORG.labelUpdater.Caption = "Doing work in the Reconcile Screen"

Do Until SeqFinished = "018-LAST PAGE IS DISPLAYED"
    host.readscreen SeqFinished, 26, 24, 2
    BORG.labelUpdater.Caption = "Doing work in the Assign Screen..." & "Grabbing " & (excelrow - 3) & " Pieces"
    
    'host.readscreen autoDGcheck, 1, bluerow, 72
    host.readscreen fullinfo, 68, bluerow, 5
    If Right(fullinfo, 1) = "X" Then
        Sheet1.Cells(excelrow, 13).Value = cannum
        
        'awbfull = Replace(awbfull, "-", "")
        
        awbfull = Replace(Left(fullinfo, 14), "-", "")
        Sheet1.Cells(excelrow, 3).Value = Right(awbfull, 4)
        Sheet1.Cells(excelrow, 1).Value = awbfull
        
        'If Trim(awbfull) = "" Then Exit Do
        
        UNnum = Mid(fullinfo, 27, 6)
        If UNnum = "******" Then UNnum = "Overpack"
        Sheet1.Cells(excelrow, 4).Value = UNnum
        
        PSN = Mid(fullinfo, 34, 10)
        Sheet1.Cells(excelrow, 5).Value = Trim(PSN)
        
        URSA = Mid(fullinfo, 17, 8)
        Sheet1.Cells(excelrow, 6).Value = Trim(URSA)
        
        hazclass = Mid(fullinfo, 45, 4)
        If hazclass = "****" Then hazclass = "Ovrpk"
        Sheet1.Cells(excelrow, 7).Value = Trim(hazclass)
        
        PackingGroup = Mid(fullinfo, 50, 3)
        If PackingGroup = "***" Then PackingGroup = "Ovrk"
        If PackingGroup = "   " Then PackingGroup = "X"
        Sheet1.Cells(excelrow, 8).Value = Trim(PackingGroup)
        
        'should be only 1 piece
        Sheet1.Cells(excelrow, 9).Value = 1
        
        'reconcile has no weight
        
        APio = Mid(fullinfo, 34, 6)
        If APio = "ALPKN1" Then
            'host.readscreen APnum, 3, bluerow, 45
            APnum = Mid(fullinfo, 41, 3)
            Sheet1.Cells(excelrow, 14).Value = Trim(APnum)
            Sheet1.Cells(excelrow, 15).Value = 1
        ElseIf APio = "OVRPCK" Then
            'host.readscreen OPnum, 3, bluerow, 45
            OPnum = Mid(fullinfo, 41, 3)
            Sheet1.Cells(excelrow, 16).Value = Trim(OPnum)
            Sheet1.Cells(excelrow, 17).Value = 1
        End If
        
        excelrow = excelrow + 1
        
    End If
    If bluerow >= 21 Then
        host.sendkey "@8"
        host.waitready 1, 51
        bluerow = 6
        host.readscreen SeqFinished, 26, 24, 2
    End If
        
    bluerow = bluerow + 1
Loop

Sheet1.Columns("A:A").NumberFormat = "000000000000"
Sheet1.Columns("C:C").NumberFormat = "0000"
Sheet1.Columns("J:J").NumberFormat = "0.00000"
'excelrow = excelrow - 1
Sheet3.Cells(3, 1).Value = excelrow
Sheet3.Cells(2, 1).Value = excelrow

GrabReconcile = excelrow

End Function


Sub GoViewAWB(host As Variant, excelrow As Integer)
'This section is for finding the origin of a package.
'Also for filling out information regarding to All packed in One's and Overpack
'check rq - check UN - check PSN - check class (Subclass)- check PG - check amount- check unit of measure - piece count

BORG.labelUpdater.Caption = "Doing work in the View Airway Bill Screen..."
Call DGscreenChooser("ViewAWB", host)

Maximum = GetMaxRow - 2


Call Module4.EQfix
Do Until excelrow = 2
    If Sheet1.Cells(excelrow, 1) > 1 Then
        BORG.labelUpdater.Caption = "Doing work in the View Airway Bill Screen..." & Maximum - (excelrow - 3) & " of " & Maximum
        host.writescreen Sheet1.Cells(excelrow, 1).Text, 3, 6
        host.sendkey "@E"
        host.waitready 1, 51
        Sheet3.Cells(2, 1).Value = excelrow
        host.readscreen miscdata, 3, 24, 2
        host.waitready 1, 51
        If miscdata = "142" Or miscdata = "145" Then
            MsgBox ("no data for " & Sheet1.Cells(excelrow, 1).Text & vbNewLine & _
                    "ExcelRow = " & excelrow)
        Else
            Call VAWB.Directions(host, excelrow)
        End If
    End If
    excelrow = excelrow - 1
Loop



End Sub

Sub CloseSession()

BORG.labelUpdater.Caption = "Closing IMS..."
host.CloseSession 0, 11
Sheet3.Cells(2, 4).Value = Time()
BORG.labelUpdater.Caption = "Done!"

End Sub



Function Terminal_check(host As Variant)
terminal = ""
host.readscreen terminal, 17, 1, 19

If terminal = "TERMINAL INACTIVE" Then
    MsgBox ("Terminal Inactive Error" & vbNewLine & "Re-run BDG")
    host.CloseSession 0, 11
    End
End If

End Function

Function GrabCloseScreen()
BORG.CanSelectGUI.Visible = False
cannum = ""
STA = ""
Status = ""

row = 8
i = 0
col = Array(6, 33, 60)
x = 0
With BORG.CanSelectGUI
    .AddItem Trim(cannum)
    .Column(1, x) = Trim(STA)
    .Column(2, x) = Status
End With

addingCans:

x = x + 1
host.readscreen inClose, 21, 2, 29
If inClose = "CLOSE/REOPEN ULD/BULK" Then
    Do Until cannum = "          "
        host.readscreen cannum, 10, row, col(i)
        host.readscreen STA, 5, row, col(i) + 11
        host.readscreen Status, 1, row, col(i) + 18
        If cannum = "          " Then
            BORG.CanSelectGUI.Locked = False
            Exit Do
        End If
        Sheet3.Cells(x + 3, 12).Value = cannum
        Sheet3.Cells(x + 3, 13).Value = STA
        Sheet3.Cells(x + 3, 14).Value = Status
        With BORG.CanSelectGUI
            .AddItem Trim(cannum)
            .Column(1, x) = Trim(STA)
            .Column(2, x) = Status
        End With
            If i = 2 Then
                i = 0
                row = row + 1
            Else: i = i + 1
            End If
        x = x + 1
    Loop
'potentional code for memphis integration
'for now loading cans takes far far far too long
'other options will need to be looked at.
'If row = 19 Then
'    row = 6
'    host.sendkey "@8"
'    host.waitready 1, 51
'    GoTo addingCans
'End If

End If
BORG.CanSelectGUI.Visible = True
End Function
Function closedcancheck()
col = Array(3, 30, 57)
closerow = 8
i = 0
Do Until BORG.cannum.Value = Sheet3.Cells(x, 12).Value
    x = x + 1
Loop
If Sheet3.Cells(x, 14).Value = "C" Then
    host.readscreen can, 10, closerow, col(i)
End If
End Function

Function openclosedcan(can As Variant)
cannum = ""
STA = ""
Status = ""
row = 8
i = 0
col = Array(6, 33, 60)
x = 0
host.readscreen inClose, 21, 2, 29
If inClose = "CLOSE/REOPEN ULD/BULK" Then
    Do Until cannum = "          "
        host.readscreen cannum, 10, row, col(i)
        host.readscreen STA, 5, row, col(i) + 11
        host.readscreen Status, 1, row, col(i) + 18
        If Trim(cannum) = can Then
            If Status = "C" Or Status = "R" Then
                host.writescreen "O", row, col(i) - 3
                host.sendkey "@e"
                host.waitready 1, 51
                host.readscreen ErrorCode, 3, 24, 2
                If ErrorCode = "469" Then
                    host.readscreen ErrorCode, 25, 24, 20
                    MsgBox (ErrorCode & vbNewLine & "HAS NOT DEPARTED ORIGIN LOCATION")
                    host.writescreen " ", row, col(i) - 3
                End If
                If ErrorCode = "057" Then
                    MsgBox (cannum & " opened successfully")
                End If
            Else:
                MsgBox (cannum & " is already open")
            End If
         End If
            If i = 2 Then
                i = 0
                row = row + 1
            Else: i = i + 1
            End If
        x = x + 1
    Loop
End If
End Function

Function UnassignCan(can As Variant)
cannum = ""
STA = ""
Status = ""
row = 8
i = 0
col = Array(6, 33, 60)
x = 0
host.readscreen inClose, 21, 2, 29
If inClose = "CLOSE/REOPEN ULD/BULK" Then
    Do Until cannum = "          "
        host.readscreen cannum, 10, row, col(i)
        host.readscreen STA, 5, row, col(i) + 11
        host.readscreen Status, 1, row, col(i) + 18
        If Trim(cannum) = can Then
            If Status = "O" Then
                host.writescreen "U", row, col(i) - 3
                host.sendkey "@e"
                host.waitready 1, 51
                host.readscreen ErrorCode, 3, 24, 2
                If ErrorCode = "469" Then
                    host.readscreen ErrorCode, 25, 24, 20
                    MsgBox (ErrorCode & vbNewLine & "HAS NOT DEPARTED ORIGIN LOCATION")
                End If
                Exit Function
            Else
                MsgBox (Trim(cannum) & " is not open." & vbNewLine & _
                        "Please open " & Trim(cannum) & " and try again.")
            End If
        End If
            If i = 2 Then
                i = 0
                row = row + 1
            Else: i = i + 1
            End If
        x = x + 1
    Loop
End If

End Function

Function CloseCan(can As Variant)
cannum = ""
STA = ""
Status = ""
row = 8
i = 0
col = Array(6, 33, 60)
x = 0
host.readscreen inClose, 21, 2, 29
If inClose = "CLOSE/REOPEN ULD/BULK" Then
    Do Until cannum = "          "
    host.readscreen cannum, 10, row, col(i)
    host.readscreen STA, 5, row, col(i) + 11
    host.readscreen Status, 1, row, col(i) + 18
    If Trim(cannum) = can Then
        If Status = "O" Then
            host.writescreen "C", row, col(i) - 3
            host.sendkey "@e"
            host.waitready 1, 51
            host.readscreen ErrorCode, 3, 24, 2
            If ErrorCode = "279" Then
                MsgBox ("Can is already Closed!")
                host.writescreen " ", row, col(i) - 3
            End If
            If ErrorCode = "057" Then
                MsgBox (cannum & " opened successfully")
            End If
            If ErrorCode = "470" Then
                host.readscreen Error, 50, 24, 20
                MsgBox (Error & vbNewLine & "Flight has not yet arrived in the system yet")
            End If
            If ErrorCode = "068" Then
                host.sendkey "ym"
                host.sendkey "@e"
                host.waitready 1, 51
                host.readscreen ErrorCode, 3, 24, 2
            End If
            If ErrorCode = "084" Then
                printer = Application.InputBox("Enter the printer you wish to print the manifest from", "Printer Select", , , , , , 2)
                host.sendkey printer
                host.sendkey "@e"
                host.waitready 1, 51
            End If
            host.readscreen ErrorCode, 3, 24, 2
            If ErrorCode = "083" Then
                MsgBox (can & " has been closed successfully and manifest has been sent to printer " & printer)
            End If
        Else:
            MsgBox (cannum & " is already Closed")
        End If
     End If
        If i = 2 Then
            i = 0
            row = row + 1
        Else: i = i + 1
        End If
    x = x + 1
    Loop
End If
host.sendkey "@3"
End Function

Function ReconcileCan(can As String, Optional host As Variant)

cannum = ""
STA = ""
Status = ""
row = 8
i = 0
col = Array(6, 33, 60)
x = 0
host.readscreen inClose, 21, 2, 29
If inClose = "CLOSE/REOPEN ULD/BULK" Then
    Do Until cannum = "          "
        host.readscreen cannum, 10, row, col(i)
        host.readscreen STA, 5, row, col(i) + 11
        host.readscreen Status, 1, row, col(i) + 18
        If Trim(cannum) = can Then
            If Status = "O" Or Status = "R" Then
                host.writescreen "R", row, col(i) - 3
                host.sendkey "@e"
                host.waitready 1, 51
                host.readscreen ErrorCode, 3, 24, 2
                If ErrorCode = "547" Then
                    host.readscreen ErrorCode, 25, 24, 20
                    MsgBox (ErrorCode & vbNewLine & "Asset is Closed and must be open to be reconciled")
                End If
                Exit Function
            Else
                MsgBox (Trim(cannum) & " is not open." & vbNewLine & _
                        "Please open " & Trim(cannum) & " and try again.")
            End If
        End If
            If i = 2 Then
                i = 0
                row = row + 1
            Else: i = i + 1
            End If
        x = x + 1
    Loop
End If
End Function

Function AddIce(can As Variant)
cannum = ""
STA = ""
Status = ""
row = 8
i = 0
col = Array(6, 33, 60)
x = 0
userinput = 0
If can = "none" Then
    can = Application.InputBox("Enter the can/bulk to enter your ice into", "Can entry", , , , , , 2)
    STA = Application.InputBox("Where is " & can & " going?" & vbNewLine & _
        "PHXR, MEMH, OAKH, etc etc", "Station Select", , , , , , 2)
    userinput = 1
    host.writescreen "ICE", 2, 17
    host.sendkey "@e"
    host.waitready 1, 51
    host.writescreen can, 7, 24
    host.writescreen STA, 7, 50
    ice = Application.InputBox("How much ice is in " & can & " can?", "Ice Entry", , , , , , 1)
    host.writescreen ice, 13, 24
    host.sendkey "@6"
    host.waitready 1, 51
    host.readscreen Error, 3, 24, 2
    If Error = "053" Then
        MsgBox (ice & " ice has been assigned to " & cannum)
        host.sendkey "@3"
    End If
    If Error = "028" Then
        STA = Application.InputBox("Where is " & can & " going?" & vbNewLine & _
        "PHXR, MEMH, OAKH, etc etc", "Station Select", , , , , , 2)
        host.sendkey "@3"
        Exit Function
    End If
End If
'host.readscreen inClose, 32, 2, 25
'If inClose = "CLOSEOUT ULD/BULK - CONFIRMATION" Then
host.readscreen inClose, 21, 2, 29
If inClose = "CLOSE/REOPEN ULD/BULK" And userinput = 0 Then
    Do Until cannum = "          "
    host.readscreen cannum, 10, row, col(i)
    host.readscreen STA, 5, row, col(i) + 11
    host.readscreen Status, 1, row, col(i) + 18
    If Trim(cannum) = can Then
        If Status = "O" Then
            host.writescreen "ICE", 2, 17
            host.sendkey "@e"
            host.waitready 1, 51
            host.writescreen cannum, 7, 24
            host.writescreen STA, 7, 50
            ice = Application.InputBox("How much ice is in this can?", "Ice Entry", , , , , , 1)
            host.writescreen ice, 13, 24
            host.sendkey "@6"
            host.waitready 1, 51
            host.readscreen Error, 3, 24, 2
            If Error = "053" Then
                MsgBox (ice & " ice has been assigned to " & cannum)
                Exit Do
            End If
            
        Else:
            MsgBox (cannum & " is closed and needs to be open to assign ICE.")
        End If
     End If
        If i = 2 Then
            i = 0
            row = row + 1
        Else: i = i + 1
        End If
    x = x + 1
    Loop
End If
host.sendkey "@3"
End Function

Function locCHK(host As Variant)
mylocation = UCase(BORG.Location)
r = 13
c = 8
host.readscreen org, Len(BORG.Location), 4, 24
If org = BORG.Location Then
    locCHK = True
    Exit Function
End If

host.readscreen Location, 5, r, c
Do Until r = 22
host.readscreen Location, 5, r, c
If Trim(Location) = mylocation Then
    locCHK = True
    Exit Function
End If
r = r + 1
Loop
End Function

Sub generateLocalList(stationID As String)

host.readscreen miscdata, 30, 1, 25
If InStr(1, miscdata, "DANGEROUS GOODS SYSTEM") >= 1 Then
    'only triggers if in the DG (68) system
    host.writescreen "viewursa", 2, 17
    host.sendkey "@e"
Else 'take us to the DG system
    host.sendkey "@C"                       'clears screen in IMS
    host.sendkey "asap@e"                   'types ASAP and enters command
    host.waitready 1, 51
    host.sendkey "68@e"                     'change to  26 for dg training - 68 for live dg
    host.waitready 1, 51
    host.sendkey "viewursa"                 'enters assign into first field to bring us to assign screen
    host.sendkey "@e"
    host.waitready 1, 51
    host.sendkey BORG.Location.Text      'inputs the location ID in DGinput into station
    If BORG.printerID <> "" Then host.writescreen BORG.printerID, 21, 32
    host.sendkey "@e"                       'sends enter key to bring us finally to Assign Screen
    host.waitready 1, 51
End If

'if everthing looks good up to hear let's be sure to clear our last local list...
Call SplitCentral.clearLocal

enterStationID:
stationID = Left(stationID, 3)
host.writescreen stationID, 4, 11
host.sendkey "@E"
host.waitready 1, 51

host.readscreen miscdata, 3, 24, 2
If miscdata = "005" Then
    stationID = InputBox("Error occured." & vbNewLine & "Please enter your 3 digit ramp ID", "Invalid Destination")
    GoTo enterStationID
End If
miscdata = "temp"
Sheet6.Cells(4, 2).Value = stationID

ERow = 5
col = 2
row = 8
Do Until miscdata = "   "
    host.readscreen miscdata, 3, row, col
    If Trim(miscdata) = "" Then Exit Do
    miscdata = miscdata + "A"
    Sheet6.Cells(ERow, 2).Value = miscdata
    ERow = ERow + 1
    col = col + 6
    If col > 75 Then
        col = 2
        row = row + 1
    End If
Loop

MsgBox ("Your local sort list has been generated." & vbNewLine & _
        "Please look over the list in the Sort Menu and confirm your splits." _
        & vbNewLine & vbNewLine & _
        "BDG generates the list form the ViewURSA screen in AutoDG. However the ViewURSA has some incorrect data and it's important to view the list yourself.")
End Sub

