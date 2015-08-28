Attribute VB_Name = "OpenBlueZone"

Dim excelrow As Integer
Option Compare Text

Sub BZOpenSession()
BZinit 'initialize and connect to bluezone session
res = DGscreenChooser("close")
If res = False Then Exit Sub
'Call BZLogin(BORG.EmpNum, BORG.PasswordBox)
Call GrabCloseScreen

BORG.labelUpdater.Caption = "BDG is now connected to Bluezone Session."
BORG.tgl_btnLogin.Value = False
BORG.loginStatusOff.Visible = False
End Sub

Function GrabReconcile(excelrow As Integer) As Integer
Dim bluerow As Integer
bluerow = 6

cannum = BZreadscreen(10, 4, 9)


Dim SeqFinished As String
SeqFinished = BZreadscreen(26, 24, 2)

BORG.labelUpdater.Caption = "Doing work in the Reconcile Screen"

Do Until SeqFinished = "018-LAST PAGE IS DISPLAYED"
    SeqFinished = BZreadscreen(26, 24, 2)
    BORG.labelUpdater.Caption = "Doing work in the Assign Screen..." & "Grabbing " & (excelrow - 3) & " Pieces"
    fullinfo = BZreadscreen(68, bluerow, 5)
    If Right(fullinfo, 1) = "X" Then
        Sheet1.Cells(excelrow, 13).Value = cannum
        awbfull = Replace(Left(fullinfo, 14), "-", "")
        Sheet1.Cells(excelrow, 3).Value = Right(awbfull, 4)
        Sheet1.Cells(excelrow, 1).Value = awbfull
        
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
        
        APiO = Mid(fullinfo, 34, 6)
        If APiO = "ALPKN1" Then
            APnum = Mid(fullinfo, 41, 3)
            Sheet1.Cells(excelrow, 14).Value = Trim(APnum)
            Sheet1.Cells(excelrow, 15).Value = 1
        ElseIf APiO = "OVRPCK" Then
            OPnum = Mid(fullinfo, 41, 3)
            Sheet1.Cells(excelrow, 16).Value = Trim(OPnum)
            Sheet1.Cells(excelrow, 17).Value = 1
        End If
        
        excelrow = excelrow + 1
        
    End If
    If bluerow >= 21 Then
        Call BZsendKey("@8")
        bluerow = 6
        SeqFinished = BZreadscreen(26, 24, 2)
    End If
        
    bluerow = bluerow + 1
Loop

Sheet1.Columns("A:A").NumberFormat = "000000000000"
Sheet1.Columns("C:C").NumberFormat = "0000"
Sheet1.Columns("J:J").NumberFormat = "0.00000"

GrabReconcile = excelrow

End Function


Sub GoViewAWB(excelrow As Integer)
'This section is for finding the origin of a package.
'Also for filling out information regarding to All packed in One's and Overpack
'check rq - check UN - check PSN - check class (Subclass)- check PG - check amount- check unit of measure - piece count

BORG.labelUpdater.Caption = "Doing work in the View Airway Bill Screen..."
Call DGscreenChooser("ViewAWB")

Maximum = GetMaxRow - 2

Call Module4.EQfix
Do Until excelrow = 2
    If Sheet1.Cells(excelrow, 1) > 1 Then
        BORG.labelUpdater.Caption = "Doing work in the View Airway Bill Screen..." & Maximum - (excelrow - 3) & " of " & Maximum
        Call BZwritescreen(Sheet1.Cells(excelrow, 1).text, 3, 6)
        Call BZsendKey("@E")
        Sheet3.Cells(2, 1).Value = excelrow
        miscdata = BZreadscreen(3, 24, 2)
        If miscdata = "142" Or miscdata = "145" Then
            MsgBox ("no data for " & Sheet1.Cells(excelrow, 1).text & vbNewLine & _
                    "ExcelRow = " & excelrow)
        Else
            Call VAWB.Directions(excelrow)
        End If
    End If
    excelrow = excelrow - 1
Loop

End Sub

Function GrabCloseScreen()
BORG.CanSelectGUI.Visible = False
cannum = ""
STA = ""
Status = ""
Dim row As Integer
row = 8


i = 0
col = Array(6, 33, 60)
Dim intCol As Integer
x = 0
'With BORG.CanSelectGUI
'    .AddItem Trim(cannum)
'    .Column(1, x) = Trim(STA)
'    .Column(2, x) = Status
'End With

addingCans:

x = x + 1
inClose = BZreadscreen(21, 2, 29)
If inClose = "CLOSE/REOPEN ULD/BULK" Then
    Do Until cannum = "          "
        intCol = col(i)
        cannum = BZreadscreen(10, row, intCol)
        If Trim(cannum) = "" Then Exit Do
        STA = BZreadscreen(5, row, intCol + 11)
        Status = BZreadscreen(1, row, intCol + 18)
        If cannum = "          " Then
            BORG.CanSelectGUI.Locked = False
            Exit Do
        End If
        Sheet3.Cells(x + 2, 12).Value = cannum
        Sheet3.Cells(x + 2, 13).Value = STA
        Sheet3.Cells(x + 2, 14).Value = Status
'        With BORG.CanSelectGUI
'            .AddItem Trim(cannum)
'            .Column(1, x) = Trim(STA)
'            .Column(2, x) = Status
'        End With
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
rowsrcSTR = "VARIABLES!L3:N" & Trim(str(x + 1)) 'N15""
BORG.CanSelectGUI.RowSource = rowsrcSTR
End If
BORG.CanSelectGUI.Visible = True
End Function
Function closedcancheck()
col = Array(3, 30, 57)
Dim closerow As Integer
closerow = 8
i = 0
Do Until BORG.cannum.Value = Sheet3.Cells(x, 12).Value
    x = x + 1
Loop
If Sheet3.Cells(x, 14).Value = "C" Then
    Dim tempcol As Integer
    tempcol = col(i)
    can = BZreadscreen(10, closerow, tempcol)
End If
End Function

Function openclosedcan(can As Variant)
Dim tempcol As Integer
Dim row As Integer

cannum = ""
STA = ""
Status = ""
row = 8
i = 0
col = Array(6, 33, 60)
x = 0
inClose = BZreadscreen(21, 2, 29)
If inClose = "CLOSE/REOPEN ULD/BULK" Then
    Do Until cannum = "          "
        tempcol = col(i)
        cannum = BZreadscreen(10, row, tempcol)
        STA = BZreadscreen(5, row, tempcol + 11)
        Status = BZreadscreen(2, row, tempcol + 18)
        If Trim(cannum) = can Then
            If Status = "C" Or Status = "R" Then
                Call BZwritescreen("O", row, tempcol - 3)
                Call BZsendKey("@E")
                ErrorCode = BZreadscreen(3, 24, 2)
                If ErrorCode = "469" Then
                    ErrorCode = BZreadscreen(3, 24, 2)
                    BORG.labelUpdater.Caption = ErrorCode & vbNewLine & "HAS NOT DEPARTED ORIGIN LOCATION"
                    Call BZwritescreen(" ", row, tempcol - 3)
                End If
                If ErrorCode = "057" Then
                    BORG.labelUpdater.Caption = cannum & " opened successfully"
                End If
            Else:
                BORG.labelUpdater.Caption = cannum & " is already open"
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
BORG.labelUpdater.Caption = "Succesfully Opened " & can & "."

End Function

Function UnassignCan(can As Variant)
Dim tempcol As Integer
Dim row As Integer

cannum = ""
STA = ""
Status = ""
row = 8
i = 0
col = Array(6, 33, 60)
x = 0
inClose = BZreadscreen(21, 2, 29)
If inClose = "CLOSE/REOPEN ULD/BULK" Then
    Do Until cannum = "          "
        tempcol = col(i)
        cannum = BZreadscreen(10, row, tempcol)
        STA = BZreadscreen(5, row, tempcol + 11)
        Status = BZreadscreen(1, row, tempcol + 18)
        If Trim(cannum) = can Then
            If Status = "O" Then
                tempcol = col(i)
                Call BZwritescreen("U", row, tempcol - 3)
                Call BZsendKey("@E")
                ErrorCode = BZreadscreen(3, 24, 2)
                If ErrorCode = "469" Then
                    ErrorCode = BZreadscreen(25, 24, 20)
                    BORG.labelUpdater.Caption = "ERROR: " & ErrorCode & vbNewLine & "HAS NOT DEPARTED ORIGIN LOCATION"
                End If
                Exit Function
            Else
                BORG.labelUpdater.Caption = "ERROR: " & Trim(cannum) & " is not open." & _
                        "Please open " & Trim(cannum) & " and try again."
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

BORG.labelUpdater.Caption = "Succesfully Unassigned freight from " & can & "."

End Function

Function CloseCan(can As Variant)
Dim tempcol As Integer
Dim row As Integer

cannum = ""
STA = ""
Status = ""
row = 8
i = 0
col = Array(6, 33, 60)
x = 0
inClose = BZreadscreen(21, 2, 29)
If inClose = "CLOSE/REOPEN ULD/BULK" Then
    Do Until cannum = "          "
    tempcol = col(i)
    cannum = BZreadscreen(10, row, tempcol)
    STA = BZreadscreen(5, row, tempcol + 11)
    Status = BZreadscreen(1, row, tempcol + 18)
    If Trim(cannum) = can Then
        If Status = "O" Then
            Call BZwritescreen("C", row, tempcol - 3)
            Call BZsendKey("@e")
            ErrorCode = BZreadscreen(3, 24, 2)
            If ErrorCode = "279" Then
                BORG.labelUpdater.Caption = "Can is already Closed!"
                Call BZwritescreen(" ", row, tempcol - 3)
            ElseIf ErrorCode = "057" Then
                BORG.labelUpdater.Caption = cannum & " opened successfully"
            ElseIf ErrorCode = "470" Then
                ErrorCode = BZreadscreen(50, 24, 20)
                BORG.labelUpdater.Caption = Error & vbNewLine & "Flight has not yet arrived in the system yet"
            ElseIf ErrorCode = "068" Then
                Call BZsendKey("YM")
                Call BZsendKey("@E")
                ErrorCode = BZreadscreen(3, 24, 2)
            ElseIf ErrorCode = "084" Then
                Dim printer As String
                printer = Application.InputBox("Enter the printer you wish to print the manifest from", "Printer Select", , , , , , 2)
                Call BZsendKey(printer)
                Call BZsendKey("@E")
            ElseIf ErrorCode = "548" Then
                BORG.labelUpdater.Caption = cannum & " needs to be reconciled before closed."
            End If
            ErrorCode = BZreadscreen(3, 24, 2)
            If ErrorCode = "083" Then
                BORG.labelUpdater.Caption = can & " has been closed successfully and manifest has been sent to printer " & printer
            End If
        Else:
            BORG.labelUpdater.Caption = cannum & " is already Closed"
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
Call BZsendKey("@3")
End Function

Function ReconcileCan(can As String, Optional host As Variant)
Dim colint As Integer
Dim row As Integer

cannum = ""
STA = ""
Status = ""
row = 8
i = 0
col = Array(6, 33, 60)
x = 0
inClose = BZreadscreen(21, 2, 29)
If inClose = "CLOSE/REOPEN ULD/BULK" Then
    Do Until cannum = "          "
        colint = col(i)
        cannum = BZreadscreen(10, row, colint)
        STA = BZreadscreen(5, row, colint + 11)
        Status = BZreadscreen(1, row, colint + 18)
        If Trim(cannum) = can Then
            If Status = "O" Or Status = "R" Then
                Call BZwritescreen("R", row, colint - 3)
                Call BZsendKey("@E")
                ErrorCode = BZreadscreen(3, 24, 2)
                If ErrorCode = "547" Then
                    ErrorCode = BZreadscreen(3, 24, 2)
                    BORG.labelUpdater.Caption = ErrorCode & vbNewLine & "Asset is Closed and must be open to be reconciled"
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

Function AddIce(can As String)
Dim row As Integer
Dim colint As Integer
Dim cannum As String
Dim STA As String
Dim ice As String

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
    DGscreenChooser ("ICE")
    Call BZwritescreen(can, 7, 24)
    Call BZwritescreen(STA, 7, 50)
    ice = Application.InputBox("How much ice is in " & can & " can?", "Ice Entry", , , , , , 1)
    Call BZwritescreen(ice, 13, 24)
    Call BZsendKey("@6")
    
    read_error = BZreadscreen(3, 24, 2)
    If read_error = "053" Then
        MsgBox (ice & " ice has been assigned to " & cannum)
        Call BZsendKey("@3")
    ElseIf read_error = "028" Then
        STA = Application.InputBox("Where is " & can & " going?" & vbNewLine & _
        "PHXR, MEMH, OAKH, etc etc", "Station Select", , , , , , 2)
        Call BZsendKey("@3")
        Exit Function
    End If
End If

inClose = BZreadscreen(21, 2, 29)
If inClose = "CLOSE/REOPEN ULD/BULK" And userinput = 0 Then
    Do Until cannum = "          "
    colint = Int(col(i))
    cannum = BZreadscreen(10, row, colint)
    STA = BZreadscreen(5, row, colint + 11)
    Status = BZreadscreen(1, row, colint + 18)
    
    If Trim(cannum) = can Then
        If Status = "O" Then
            Call DGscreenChooser("ICE")
            Call BZwritescreen(cannum, 7, 24)
            Call BZwritescreen(STA, 7, 50)
            ice = Application.InputBox("How much ice is in this can?", "Ice Entry", , , , , , 1)
            Call BZwritescreen(ice, 13, 24)
            Call BZsendKey("@6")
            read_error = BZreadscreen(3, 24, 2)
            If read_error = "053" Then
                BORG.labelUpdater.Caption = ice & " ice has been assigned to " & cannum
                Exit Do
            End If
            
        Else:
            BORG.labelUpdater.Caption = cannum & " is closed and needs to be open to assign ICE."
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
Call BZsendKey("@3")
End Function

Function locCHK(host As Variant)
mylocation = UCase(BORG.Location)
r = 13
c = 8
org = BZreadscreen(Len(BORG.Location), 4, 24)
If org = BORG.Location Then
    locCHK = True
    Exit Function
End If

Location = BZreadscreen(5, r, c)
Do Until r = 22
Location = BZreadscreen(5, r, c)
If Trim(Location) = mylocation Then
    locCHK = True
    Exit Function
End If
r = r + 1
Loop
End Function

Sub generateLocalList(stationID As String)

miscdata = BZreadscreen(30, 1, 25)
DGscreenChooser ("viewursa")
'if everthing looks good up to hear let's be sure to clear our last local list...
Call SplitCentral.clearLocal

enterStationID:
stationID = Left(stationID, 3)
Call BZwritescreen(stationID, 4, 11)
Call BZsendKey("@E")
miscdata = BZreadscreen(3, 24, 2)
If miscdata = "005" Then
    stationID = InputBox("Error occured." & vbNewLine & "Please enter your 3 digit ramp ID", "Invalid Destination")
    GoTo enterStationID
End If
miscdata = "temp"
Sheet6.Cells(4, 2).Value = stationID

ERow = 5
Dim col As Integer
Dim row As Integer
col = 2
row = 8
Do Until miscdata = "   "
    miscdata = BZreadscreen(3, row, col)
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





