Attribute VB_Name = "GhostAssign"
Dim canarr() As Variant
Dim splarr() As Variant
Dim destarr() As Variant
Dim typearr() As Variant
Dim excelrow As Integer
Dim ExcelStartRow As Integer
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

Sub setupArrays()

Dim cannums As New Collection
Dim cansplits As New Collection
Dim candests As New Collection
Dim cantypes As New Collection

row = 3
Do While Sheet4.Cells(row, 1) <> ""
   cannums.Add Sheet4.Cells(row, 1) '  dynamically add value to the end
   cansplits.Add Sheet4.Cells(row, 2)
   candests.Add Sheet4.Cells(row, 3)
   cantypes.Add Sheet4.Cells(row, 4)
   row = row + 1
Loop

canarr = toArray(cannums) 'convert collection to an array
splarr = toArray(cansplits)
destarr = toArray(candests)
typearr = toArray(cantypes)


End Sub

Function toArray(col As Collection)
  Dim arr() As Variant
  ReDim arr(0 To col.Count - 1) As Variant
  For i = 1 To col.Count
      arr(i - 1) = col(i)
  Next
  toArray = arr

End Function


Sub bzConnect()

ChDir "C:\"
Set host = CreateObject("BZwhll.whllobj")
retval = host.OpenSession(0, 11, "fdx3270.zmd", 30, 1)
host.WaitCursor 1, 9, 1, 1
retval = host.Connect("K")
Set Wnd = host.Window()
Wnd.Visible = True

Wnd.Caption = "BDG v4.0 is Searching"
If (retval) Then
    host.MsgBox "Error connecting to Session K!", 48
End If
Call Module1.DEL
Call GhostAssign.ErrorChecker(host)
host.readscreen misc, 25, 8, 25
host.readscreen misctwo, 42, 14, 15
If Trim(misc) = "FEDERAL EXPRESS" Then
    Call GhostAssign.Login
ElseIf misctwo = "ENTER THE NUMBER OF THE DESIRED SYSTEM ===>" Then
Else
    Call GhostAssign.GrabUnassigned
End If
End Sub
Sub Login()
host.sendkey "IMS"
host.sendkey "@E"
host.waitready 1, 150
host.readscreen setupData, 37, 1, 23

If Trim(setupData) <> "F E D E R A L  E X P R E S S  I M S" Then
    host.sendkey "@3"
    Call GhostAssign.Login
End If

host.writescreen BORG.EmpNum, 7, 15
host.writescreen BORG.PasswordBox, 7, 43
host.sendkey "@E"
host.waitready 1, 51

host.sendkey "68"
host.sendkey "@E"
host.waitready 1, 51
host.readscreen DGscreenInfo, 22, 1, 28
If DGscreenInfo = "DANGEROUS GOODS SYSTEM" Then
    host.writescreen BORG.Location.Text, 19, 44
    If Trim(BORG.printerID.Text) <> "" Then host.writescreen BORG.printerID.Text, 21, 32
    Call GhostAssign.GrabUnassigned(3)
End If

End Sub

Sub GrabUnassigned(ERow As Integer)

If TypeName(host) <> "IWhllObj" Then
    ChDir "C:\"
    Set host = CreateObject("BZwhll.whllobj")
    retval = host.OpenSession(0, 11, "fdx3270.zmd", 30, 1)
    host.WaitCursor 1, 9, 1, 1
    retval = host.Connect("K")
    Set Wnd = host.Window() ' Makes the window invisible.....
End If

Call DGscreenChooser("Assign", host)
host.readscreen AssignViewType, 15, 3, 32
If AssignViewType <> "UNASSIGNED VIEW" Then
    host.sendkey "@2"
    host.waitready 1, 51
End If

excelrow = ERow

Call Module1.SETUP
BORG.labelUpdater.Caption = "Doing work in the Assign Screen..."

excelrow = GhostAssign.GrabAssign(host, "A", excelrow)
excelrow = GhostAssign.GrabAssign(host, "I", excelrow)

'setup format and variables for VAWB section
Sheet1.Columns("A:A").NumberFormat = "000000000000"
Sheet1.Columns("C:C").NumberFormat = "0000"
Sheet1.Columns("J:J").NumberFormat = "0.00000"
excelrow = excelrow - 1
Sheet3.Cells(3, 1).Value = excelrow
Sheet3.Cells(2, 1).Value = excelrow
'finish setting stuff up for VAWB

Call OpenBlueZone.GoViewAWB(host, excelrow)

End Sub

Sub GrabAssigned(ERow As Integer, can As String, Optional host As Variant = Empty)

'close screen to find can
Call GhostAssign.DGscreenChooser("Close", host)

'type R and go through new reconcile screen grabbing AWB's
Call OpenBlueZone.ReconcileCan(can, host)
'confirm can num at reconcile screen with can given
host.readscreen asset, 10, 4, 9
If Trim(asset) <> can Then
    MsgBox ("error occured in GrabAsssigned sub in GhostAssign Module")
    Exit Sub
End If

Dim lastrow As Integer
excelrow = OpenBlueZone.GrabReconcile(3, host)

'go to Vawb section like normal
host.sendkey "@3"
host.waitready 1, 51
Call DGscreenChooser("viewawb", host)
'Call OpenBlueZone.GoViewAWB(host, lastrow)

'setup format and variables for VAWB section
Sheet1.Columns("A:A").NumberFormat = "000000000000"
Sheet1.Columns("C:C").NumberFormat = "0000"
Sheet1.Columns("J:J").NumberFormat = "0.00000"

'what is this???
'excelrow = excelrow - 1
Sheet3.Cells(3, 1).Value = excelrow
Sheet3.Cells(2, 1).Value = excelrow
'finish setting stuff up for VAWB

Call OpenBlueZone.GoViewAWB(host, excelrow)

End Sub
Function DGscreenChooser(menu As String, host As Variant)

If TypeName(host) <> "IWhllObj" Then
    ChDir "C:\"
    Set host = CreateObject("BZwhll.whllobj")
    retval = host.OpenSession(0, 11, "fdx3270.zmd", 30, 1)
    host.WaitCursor 1, 9, 1, 1
    retval = host.Connect("K")
    Set Wnd = host.Window()
End If
    
host.readscreen DGscreenInfo, 50, 1, 20
If InStr(1, DGscreenInfo, "DANGEROUS GOODS SYSTEM") >= 1 Then
    host.writescreen menu, 2, 17
    host.sendkey "@E"
    host.waitready 1, 51
Else
    host.sendkey "@C"                       'clears screen in IMS
    host.sendkey "asap@e"                   'types ASAP and enters command
    host.waitready 1, 51
    host.readscreen miscdata, 32, 1, 2
    If miscdata = "ASAP COMMAND IS UNKNOWN TO VTAM." Then
        Call GhostAssign.Login
    End If
    host.sendkey "68@e"                     'enters 26 for dg training
    host.waitready 1, 51
    host.writescreen menu, 2, 17              'enters assign into first field to bring us to assign screen
    host.writescreen BORG.Location.Text, 19, 44   'inputs the location ID in DGinput into station
    If BORG.printerID <> "" Then host.writescreen BORG.printerID, 21, 32
    host.sendkey "@e"                       'sends enter key to bring us finally to Assign Screen
    host.waitready 1, 51
End If
host.readscreen retCode, 3, 24, 2
If retCode = "136" Then
    host.writescreen BORG.Location.Text, 19, 44
End If

End Function
Function ErrorChecker(host As Variant)

terminal = ""
host.readscreen terminal, 80, 1, 1
If InStr(1, terminal, "APPLICATION NOT") > 1 Then
    host.CloseSession 0, 11
    Call GhostAssign.Login
End If
If InStr(1, terminal, "TERMINAL INACTIVE") > 1 Then
    MsgBox ("Terminal Inactive Error" & vbNewLine & "Re-run BDG")
    host.CloseSession 0, 11
    GoTo RestartSession
End If

Exit Function

RestartSession:
    Call GhostAssign.bzConnect
    
End Function

Function GrabAssign(host As Variant, Optional haztype As String = " ", Optional excelrow As Integer = 3) As Integer
If TypeName(host) = Empty Then
    ChDir "C:\"
    Set host = CreateObject("BZwhll.whllobj")
    retval = host.OpenSession(0, 11, "fdx3270.zmd", 30, 1)
    host.WaitCursor 1, 9, 1, 1
    retval = host.Connect("K")
    Set Wnd = host.Window() ' Makes the window invisible.....
End If

Call DGscreenChooser("assign", host)

host.readscreen miscdata, 15, 3, 32
If miscdata <> "UNASSIGNED VIEW" Then host.sendkey "@2"

row = 10

host.writescreen haztype, 6, 45
host.sendkey "@E"
host.waitready 1, 51
host.readscreen SeqFinished, 26, 24, 2

Do Until SeqFinished = "018-LAST PAGE IS DISPLAYED"
    If Left(SeqFinished, 3) = 256 Then Exit Do
    host.readscreen CanAssigned, 10, row, 26
    If Trim(CanAssigned) = "" Then
        host.readscreen awbfull, 12, row, 5
        If Trim(awbfull) = "" Then Exit Do
        Sheet1.Cells(excelrow, 1).Value = awbfull
        Sheet1.Cells(excelrow, 3).Value = Right(awbfull, 4) 'get last 4 for our filter
    
        Sheet1.Cells(excelrow, 23).Value = haztype
        CanAssign = "UnAssigned"
        
        Sheet1.Cells(excelrow, 13).Value = CanAssign
        BORG.labelUpdater.Caption = "Doing work in the Assign Screen..." & "Grabbing " & (excelrow - 2) & " Pieces"
        host.readscreen UNnum, 4, row, 44
        If UNnum = "******" Then UNnum = "Overpack"
        Sheet1.Cells(excelrow, 4).Value = UNnum
        
        host.readscreen PSN, 9, row, 49
        Sheet1.Cells(excelrow, 5).Value = PSN
        
        host.readscreen URSA, 8, row, 18
        Sheet1.Cells(excelrow, 6).Value = Trim(URSA)
        
        host.readscreen hazclass, 4, row, 59
        If hazclass = "***" Then hazclass = "Ovrpk"
        Sheet1.Cells(excelrow, 7).Value = hazclass
        
        host.readscreen PackingGroup, 3, row, 64
        If PackingGroup = "***" Then PackingGroup = "Ovrk"
        Sheet1.Cells(excelrow, 8).Value = PackingGroup
        
        'piece should now always be 1
        Sheet1.Cells(excelrow, 9).Value = 1
        
        host.readscreen Weight, 10, row, 68
        Sheet1.Cells(excelrow, 10).Value = Weight
        
        host.readscreen UnitofMeasure, 2, row, 79
        Sheet1.Cells(excelrow, 11).Value = UnitofMeasure
        
        host.readscreen APio, 6, row, 49
        If APio = "ALPKN1" Then
            host.readscreen APnum, 3, row, 55
            Sheet1.Cells(excelrow, 14).Value = APnum
            Sheet1.Cells(excelrow, 15).Value = 1
        ElseIf APio = "OVRPCK" Then
            host.readscreen OPnum, 3, row, 55
            Sheet1.Cells(excelrow, 16).Value = OPnum
            Sheet1.Cells(excelrow, 17).Value = 1
        End If
    End If
    
    excelrow = excelrow + 1
    row = row + 1
    
    If row >= 18 Then
        host.sendkey "@8"
        host.waitready 1, 51
        row = 10
        host.readscreen SeqFinished, 26, 24, 2
    End If
Loop 'do until grabbing stuff from assign screen END

Sheet1.Columns("A:A").NumberFormat = "000000000000"
Sheet1.Columns("C:C").NumberFormat = "0000"
Sheet1.Columns("J:J").NumberFormat = "0.00000"
'excelrow = excelrow - 1
Sheet3.Cells(3, 1).Value = excelrow
Sheet3.Cells(2, 1).Value = excelrow

'If cannum <> "unassigned" Then
'    Call OpenBlueZone.GoViewAWB(host, excelrow)
'End If
GrabAssign = excelrow
End Function

Sub GhostSort()

maxRow = GetMaxRow
Call GhostAssign.setupArrays
'If canarr() = Empty Then Exit Sub
For i = 0 To (UBound(canarr, 1) - LBound(canarr, 1))
    ERow = 5
    Ecol = 3
    Dim typeHaz As String
    typeHaz = typearr(i)
    setType = SetTypeFilter(typeHaz) 'convert skynet text to filterable text
    
    'set filter to include only pieces that match type of canType(i)
    Sheet1.Range("$A$2:$X$2").AutoFilter Field:=23, Criteria1:=setType
    
    Do Until Sheet6.Cells(2, Ecol) = splarr(i)
        Ecol = Ecol + 1
    Loop
    
    typeHaz = splarr(i)
    If typeHaz = "" Then GoTo nextIteration
    Dim isLocal As Boolean
    isLocal = isSplitLocal(typeHaz)
    If isLocal = False Then
        FilterOutLocal 'hides local stuff
    Else 'shows local stuff
        Sheet1.Range("$A$2:$X$2").AutoFilter Field:=6
    End If
    
    
    Do Until Sheet6.Cells(ERow, Ecol) = ""
        Dim tempSplit As String
        tempSplit = Sheet6.Cells(ERow, Ecol).Text
        Call setSplitFilter(tempSplit, isLocal)
        
        For row = 3 To maxRow - 1
            'Use Hidden property to check if filtered or not
            If Sheet1.Cells(row, 1).EntireRow.Hidden = False Then
                Sheet1.Cells(row, 21).Value = canarr(i)
                Sheet1.Cells(row, 22).Value = destarr(i)
            End If
        Next
        ERow = ERow + 1
    Loop
nextIteration:
assignedCansHidden = hideAssignedPcs()
Next
Call GhostAssign.filterClear
End Sub


Function GetMaxRow() As Integer
    maxRow = 3
    Do Until Sheet1.Cells(maxRow, 1).Value = ""
        maxRow = maxRow + 1
    Loop
    GetMaxRow = maxRow
End Function

Function SetCanFilter(can As String) As Boolean
On Error GoTo ErrOut:
    Sheet1.Range("$A$2:$X$2").AutoFilter Field:=21, Criteria1:=can, Operator:=xlAnd
    SetCanFilter = True
    Exit Function
    
ErrOut:
    SetCanFilter = False
End Function

Function hideAssignedPcs() As Boolean
On Error GoTo ErrOut:
    Sheet1.Range("$A$2:$X$2").AutoFilter Field:=21, Criteria1:="=", Operator:=xlAnd
    hideAssignedPcs = True
    Exit Function
    
ErrOut:
    hideAssignedPcs = False
End Function

Function setSplitFilter(Split As String, isLocal As Boolean) As Boolean
On Error GoTo ErrOut:
    If isLocal = False Then
        splitFilter = "=" + Split + "*"
        'Sheet1.Range("$A$2:$X$2").AutoFilter Field:=6, Criteria1:="=N*", Operator:=xlAnd
        Sheet1.Range("$A$2:$X$2").AutoFilter Field:=6, Criteria1:=splitFilter
        setSplitFilter = True
        Exit Function
    Else 'we are doing a suffix split
        splitFilter = "=*" + Split
        Sheet1.Range("$A$2:$X$2").AutoFilter Field:=6, Criteria1:=splitFilter
        setSplitFilter = True
        Exit Function
    End If
ErrOut:
    setSplitFilter = False
End Function

Function SetTypeFilter(haztype As String)

    Select Case haztype
    Case "ADG"
        ans = "A"
    Case "IDG"
        ans = "I"
    Case "ALL"
        ans = "=*"
    End Select
    
    SetTypeFilter = ans
    
End Function


Function FilterOutLocal()
splshtRow = 5
splitcol = 2

shtOnerow = 3

Do Until Sheet1.Cells(shtOnerow, 1) = ""
    ursaCHK = Trim(Right(Sheet1.Cells(shtOnerow, 6), 5))
    splshtRow = 5
    Do Until Sheet6.Cells(splshtRow, 2) = ""
        If Sheet6.Cells(splshtRow, 2) = ursaCHK Then
            Sheet1.Cells(shtOnerow, 24) = "L"
        End If
    splshtRow = splshtRow + 1
    Loop
shtOnerow = shtOnerow + 1
Loop

Sheet1.Range("$A$2:$X$2").AutoFilter Field:=24, Criteria1:="<>L"
End Function
    'Sheet1.Range("$A$2:$X$2").AutoFilter
    '"=*inwa"
    'Sheet1.Range("$A$2:$X$2").AutoFilter Field:=14, Criteria1:="=N*"
    'Sheet1.Range("$A$2:$X$2").AutoFilter Field:=23, Criteria1:="A"
    'Sheet1.Range("$A$2:$X$2").AutoFilter Field:=6, Operator:=xlAnd SortOn:=xlSortOnValues, Order:=xlAscending
Sub filterClear()
    Sheet1.Range("$A$2:$X$2").AutoFilter
    Sheet1.Range("$A$2:$X$2").AutoFilter
End Sub


Sub filterCanSort(cannum As String)
Call GhostAssign.filterClear
Call Module1.DELmanifestSheet

myFilter = "=" + cannum
Sheet1.Range("$A$2:$X$2").AutoFilter Field:=21, Criteria1:=myFilter
ActiveWorkbook.Worksheets("Sheet1").AutoFilter.Sort.SortFields.Add Key:=Range("C2:C999"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
    xlSortTextAsNumbers
With ActiveWorkbook.Worksheets("Sheet1").AutoFilter.Sort
    .Header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With

If BORG.StationSort.Value = True Then
        ActiveWorkbook.Worksheets("Sheet1").Sort.SortFields.Add Key:=Range("B3:B999") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    Call Module4.HAZ_LIST_w_Station
ElseIf BORG.Can_flight.Value = True Then
ActiveWorkbook.Worksheets("Sheet1").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Sheet1").AutoFilter.Sort.SortFields.Add Key:=Range _
        ("C2"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortTextAsNumbers
    With ActiveWorkbook.Worksheets("Sheet1").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveWorkbook.Worksheets("Sheet1").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Sheet1").AutoFilter.Sort.SortFields.Add Key:=Range _
        ("S2"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortTextAsNumbers
    With ActiveWorkbook.Worksheets("Sheet1").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 3
    Sheet1.Range("$A$2:$X$2").AutoFilter Field:=21, Criteria1:=cannum
    Call Module4.HAZ_LIST_w_flightInfo
Else
    Call Module4.HAZ_LIST
End If

Sheet2.Cells(2, 6) = UCase(cannum) + "  -  (Not Assigned)"
Call Module4.QoL_stuff
Call Module1.printFun

End Sub

Sub dupFind()

myRows = GetMaxRow
If myRows = 4 Then Exit Sub 'nothing to find dup data on....

duprun:

dups_found = 0
Call GhostAssign.filterClear
Call GhostAssign.sortingsub
eRows = 3

Do Until Sheet1.Cells(eRows, 1).Text = ""
megastring = ""
    For i = 1 To 22
        megastring = megastring & Sheet1.Cells(eRows, i).Text
    Next
    If oldstring = megastring And Sheet1.Cells(eRows, 23) <> Sheet1.Cells(eRows - 1, 23) Then
        'MsgBox ("we found a dupe!")
        dups_found = dups_found + 1
        BORG.labelUpdater.Caption = "Removing Dupe!"
        If Sheet1.Cells(eRows, 23).Text = "I" Then
            'delete duplicate I row
            Sheet1.Rows(eRows).Delete shift:=xlUp
        ElseIf Sheet1.Cells(eRows - 1, 23).Text = "I" Then
            'delete duplicate I row from previous megastring
            Sheet1.Rows(eRows - 1).Delete shift:=xlUp
        Else
            'something is wrong
            bla = MsgBox("dup find did not work!", vbCritical)
            Exit Sub
        End If
    End If
    oldstring = megastring
    eRows = eRows + 1
Loop

If dups_found > 0 Then
    GoTo duprun
End If

End Sub


Sub sortingsub()
    ActiveWorkbook.Worksheets("Sheet1").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Sheet1").AutoFilter.Sort.SortFields.Add Key:=Range _
        ("N2"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortTextAsNumbers
    With ActiveWorkbook.Worksheets("Sheet1").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveWorkbook.Worksheets("Sheet1").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Sheet1").AutoFilter.Sort.SortFields.Add Key:=Range _
        ("P2"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortTextAsNumbers
    With ActiveWorkbook.Worksheets("Sheet1").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveWorkbook.Worksheets("Sheet1").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Sheet1").AutoFilter.Sort.SortFields.Add Key:=Range _
        ("A2"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortTextAsNumbers
    With ActiveWorkbook.Worksheets("Sheet1").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub


