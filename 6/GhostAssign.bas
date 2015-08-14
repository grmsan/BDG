
Dim canarr() As Variant
Dim splarr() As Variant
Dim destarr() As Variant
Dim typearr() As Variant
Dim excelrow As Integer
Dim ExcelStartRow As Integer

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

Sub GrabUnassigned(ERow As Integer)

Call DGscreenChooser("Assign")
assignViewType = BZreadscreen(15, 3, 32)
If assignViewType <> "UNASSIGNED VIEW" Then
    Call BZsendKey("@2")
End If

Dim excelrow As Integer
excelrow = ERow

Call Module1.SETUP
BORG.labelUpdater.Caption = "Doing work in the Assign Screen..."

excelrow = GhostAssign.GrabAssign("A", excelrow)
excelrow = GhostAssign.GrabAssign("I", excelrow)

'setup format and variables for VAWB section
Sheet1.Columns("A:A").NumberFormat = "000000000000"
Sheet1.Columns("C:C").NumberFormat = "0000"
Sheet1.Columns("J:J").NumberFormat = "0.00000"
excelrow = excelrow - 1
Sheet3.Cells(3, 1).Value = excelrow
Sheet3.Cells(2, 1).Value = excelrow
'finish setting stuff up for VAWB

Call OpenBlueZone.GoViewAWB(excelrow)

End Sub

Sub GrabAssigned(ERow As Integer, can As String)

'close screen to find can
Call DGscreenChooser("Close")
'type R and go through new reconcile screen grabbing AWB's
Call OpenBlueZone.ReconcileCan(can)
'confirm can num at reconcile screen with can given
asset = BZreadscreen(10, 4, 9)
If Trim(asset) <> can Then
    MsgBox ("error occured in GrabAsssigned sub in GhostAssign Module")
    Exit Sub
End If

Dim lastrow As Integer
excelrow = GrabReconcile(3)
Call BZsendKey("@3")
Call DGscreenChooser("viewawb")

'setup format and variables for VAWB section
Sheet1.Columns("A:A").NumberFormat = "000000000000"
Sheet1.Columns("C:C").NumberFormat = "0000"
Sheet1.Columns("J:J").NumberFormat = "0.00000"

'what is this???
'excelrow = excelrow - 1
Sheet3.Cells(3, 1).Value = excelrow
Sheet3.Cells(2, 1).Value = excelrow
'finish setting stuff up for VAWB

Call OpenBlueZone.GoViewAWB(excelrow)

End Sub

Function ErrorChecker()

terminal = ""
terminal = BZreadscreen(80, 1, 1)
If InStr(1, terminal, "APPLICATION NOT") > 1 Then
    'host.CloseSession 0, 11
End If
If InStr(1, terminal, "TERMINAL INACTIVE") > 1 Then
    MsgBox ("Terminal Inactive Error" & vbNewLine & "Re-run BDG")
    'host.CloseSession 0, 11
    GoTo RestartSession
End If

Exit Function

RestartSession:
    Call GhostAssign.bzConnect
    
End Function

Function GrabAssign(Optional haztype As String = " ", Optional excelrow As Integer = 3) As Integer

Call DGscreenChooser("assign")

miscdata = BZreadscreen(15, 3, 34)
If miscdata <> "UNASSIGNED VIEW" Then
    Call BZsendKey("@2")
End If

CanAssign = "UnAssigned"
Dim row As Integer
row = 10

Call BZwritescreen(haztype, 6, 45)
Call BZsendKey("@E")
SeqFinished = BZreadscreen(26, 24, 2)


Do Until SeqFinished = "018-LAST PAGE IS DISPLAYED"
    If Left(SeqFinished, 3) = "256" Then Exit Do
    'canassigned = BZreadscreen(10, row, 26)
    fullinfo = BZreadscreen(76, row, 5)
    If Trim(canassigned) = "" Then
        awbfull = BZreadscreen(12, row, 5)
        If Trim(awbfull) = "" Then Exit Do
        Sheet1.Cells(excelrow, 1).Value = awbfull
        Sheet1.Cells(excelrow, 3).Value = Right(awbfull, 4) 'get last 4 for our filter
        Sheet1.Cells(excelrow, 23).Value = haztype
        Sheet1.Cells(excelrow, 13).Value = CanAssign
        
        BORG.labelUpdater.Caption = "Doing work in the Assign Screen..." & "Grabbing " & (excelrow - 2) & " Pieces"
        UNnum = "UN" & Mid(fullinfo, 40, 4)
        
        If UNnum = "UN****" Then UNnum = "Overpack"
        Sheet1.Cells(excelrow, 4).Value = UNnum
        
        PSN = Mid(fullinfo, 45, 9)
        Sheet1.Cells(excelrow, 5).Value = PSN
        
        URSA = Mid(fullinfo, 14, 8)
        Sheet1.Cells(excelrow, 6).Value = Trim(URSA)
        
        hazclass = Mid(fullinfo, 55, 4)
        If hazclass = "*** " Then hazclass = "Ovrpk"
        Sheet1.Cells(excelrow, 7).Value = hazclass
              
        PackingGroup = Mid(fullinfo, 60, 3)
        If PackingGroup = "***" Then PackingGroup = "Ovrk"
        Sheet1.Cells(excelrow, 8).Value = PackingGroup
        
        'piece should now always be 1
        Sheet1.Cells(excelrow, 9).Value = 1

        Weight = Mid(fullinfo, 64, 10)
        Sheet1.Cells(excelrow, 10).Value = Weight
        
        UnitofMeasure = Mid(fullinfo, 75, 2)
        Sheet1.Cells(excelrow, 11).Value = UnitofMeasure
        
        APiO = Mid(fullinfo, 45, 6)
        If APiO = "ALPKN1" Then
            APnum = Mid(fullinfo, 51, 3)
            Sheet1.Cells(excelrow, 14).Value = APnum
            Sheet1.Cells(excelrow, 15).Value = 1
        ElseIf APiO = "OVRPCK" Then
            OPnum = Mid(fullinfo, 51, 3)
            Sheet1.Cells(excelrow, 16).Value = OPnum
            Sheet1.Cells(excelrow, 17).Value = 1
        End If
    End If
    
    excelrow = excelrow + 1
    row = row + 1
    
    If row >= 18 Then
        Call BZsendKey("@8")
        row = 10
        SeqFinished = BZreadscreen(26, 24, 2)
    End If
Loop 'do until grabbing stuff from assign screen END

Sheet1.Columns("A:A").NumberFormat = "000000000000"
Sheet1.Columns("C:C").NumberFormat = "0000"
Sheet1.Columns("J:J").NumberFormat = "0.00000"
'excelrow = excelrow - 1

'what are these for? VVVVVVVVVVVVVV
Sheet3.Cells(3, 1).Value = excelrow
Sheet3.Cells(2, 1).Value = excelrow

GrabAssign = excelrow
End Function

Sub GhostSort()

maxRow = GetMaxRow
Call GhostAssign.setupArrays
'If canarr() = Empty Then Exit Sub
For i = 0 To (UBound(canarr, 1) - LBound(canarr, 1))
    ERow = 5
    ecol = 3
    Dim typeHaz As String
    typeHaz = typearr(i)
    setType = SetTypeFilter(typeHaz) 'convert skynet text to filterable text
    
    'set filter to include only pieces that match type of canType(i)
    Sheet1.Range("$A$2:$X$2").AutoFilter Field:=23, Criteria1:=setType
    
    Do Until Sheet6.Cells(2, ecol) = splarr(i)
        ecol = ecol + 1
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
    
    
    Do Until Sheet6.Cells(ERow, ecol) = ""
        Dim tempSplit As String
        tempSplit = Sheet6.Cells(ERow, ecol).text
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
On Error GoTo errout:
    Sheet1.Range("$A$2:$X$2").AutoFilter Field:=21, Criteria1:=can, Operator:=xlAnd
    SetCanFilter = True
    Exit Function
    
errout:
    SetCanFilter = False
End Function

Function hideAssignedPcs() As Boolean
On Error GoTo errout:
    Sheet1.Range("$A$2:$X$2").AutoFilter Field:=21, Criteria1:="=", Operator:=xlAnd
    hideAssignedPcs = True
    Exit Function
    
errout:
    hideAssignedPcs = False
End Function

Function setSplitFilter(Split As String, isLocal As Boolean) As Boolean
On Error GoTo errout:
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
errout:
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

Do Until Sheet1.Cells(eRows, 1).text = ""
megastring = ""
    For i = 1 To 22
        megastring = megastring & Sheet1.Cells(eRows, i).text
    Next
    If oldstring = megastring And Sheet1.Cells(eRows, 23) <> Sheet1.Cells(eRows - 1, 23) Then
        'MsgBox ("we found a dupe!")
        dups_found = dups_found + 1
        BORG.labelUpdater.Caption = "Removing Dupe!"
        If Sheet1.Cells(eRows, 23).text = "I" Then
            'delete duplicate I row
            Sheet1.Rows(eRows).Delete shift:=xlUp
        ElseIf Sheet1.Cells(eRows - 1, 23).text = "I" Then
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




