Attribute VB_Name = "Module4"
Sub EQfix()
EQrow = 3
Do Until Sheet1.Cells(EQrow, 1).Value < 1
If Sheet1.Cells(EQrow, 4) = "UN2911" Or Sheet1.Cells(EQrow, 4) = "UN2910" Or _
    Sheet1.Cells(EQrow, 4) = "UN2909" Or Sheet1.Cells(EQrow, 4) = "UN2908" Then
        Sheet1.Cells(EQrow, 5).Value = "Radioactive, Excepted Qty"
        Sheet1.Cells(EQrow, 7).Value = "EQ"
        Sheet1.Cells(EQrow, 10).Value = "EQ"
        Sheet1.Cells(EQrow, 11).Value = "EQ"
End If
EQrow = EQrow + 1
Loop
End Sub
Sub APOPfix()
row = 3
col = 14
Do Until Sheet1.Cells(row, 1).Value = ""
col = 14
    Do Until col = 18
        If Sheet1.Cells(row, col).Value < 1 Then Sheet1.Cells(row, col).Value = ""
        col = col + 1
    Loop
row = row + 1
Loop
End Sub


Sub SORT_MACRO()
Attribute SORT_MACRO.VB_ProcData.VB_Invoke_Func = " \n14"
Sheet1.Columns("A:A").NumberFormat = "000000000000"
Sheet1.Columns("C:C").NumberFormat = "0000"
        ActiveWorkbook.Worksheets("Sheet1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Sheet1").Sort.SortFields.Add Key:=Range("L3:L9999") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        
If BORG.StationSort = True Then
    ActiveWorkbook.Worksheets("Sheet1").Sort.SortFields.Add Key:=Range("B3:B9999") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
End If

If BORG.Can_flight.Value = True Then
    ActiveWorkbook.Worksheets("Sheet1").Sort.SortFields.Add Key:=Range("S3:S9999") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
End If

    ActiveWorkbook.Worksheets("Sheet1").Sort.SortFields.Add Key:=Range("C3:C9999") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("Sheet1").Sort.SortFields.Add Key:=Range("N3:N9999") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("Sheet1").Sort.SortFields.Add Key:=Range("P3:P9999") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Sheet1").Sort
        .SetRange Range("A2:T9999")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    

'If BORG.StationSort = True Then
'    Call Module4.HAZ_LIST_w_Station
'ElseIf BORG.Can_flight = True Then
'    Call Module4.HAZ_LIST_w_flightInfo
'Else
    Call Module4.HAZ_LIST
'End If

End Sub

Sub HAZ_LIST_w_Station()
Dim currentOrg As String
x = 3
col = 2
y = 3
TwoRow = 6
oldURSA = ""
'OpenBlueZone.ExcelRow = 95
Sheet2.Cells(2, 6).Value = UCase(BORG.CanSelectGUI.Value)
Do Until Sheet1.Cells(y, 1) = ""
    y = y + 1
Loop

Dim URSA As String
Dim str As String
str = "Incoming DG from "
stars = "  *****************  "

Do Until x = y
    If Sheet1.Cells(x, 1).EntireRow.Hidden = False Then
        URSA = Sheet1.Cells(x, 2)
            If URSA = oldURSA Then 'if ursa has not changed
            '--------------- if statement for All-Packed-In-One's ---------'
                If Sheet1.Cells(x, 14).Value = CInt(Sheet1.Cells(x, 14).Value) And Sheet1.Cells(x, 14).Value >= 1 Then
                    Sheet2.Rows(TwoRow).Offset(1).EntireRow.Insert
                    Sheet2.Cells(TwoRow, 4).Value = stars & "  All-Packed-In-One # " & Sheet1.Cells(x, 14).Value & stars
                    Sheet2.Cells(TwoRow, 8).Value = Sheet1.Cells(x, 15).Value
                    Sheet2.Range(Cells(TwoRow, 4), Cells(TwoRow, 7)).Merge
                    Sheet2.Cells(TwoRow, 4).HorizontalAlignment = xlCenter
                    Sheet2.Cells(TwoRow, 3).Value = Sheet1.Cells(x, 6).Value
                    Sheet2.Cells(TwoRow, 2).Value = Sheet1.Cells(x, 1).Value 'AWB
                    Sheet2.Cells(TwoRow, 2).NumberFormat = "0000-0000-0000"
                    Sheet2.Cells(TwoRow, 9).Value = "    *** "
                    Sheet2.Cells(TwoRow, 10).Value = "  ** "
                    TwoRow = TwoRow + 1
                End If 'end of ALPKN1
                '--------------- if statement for Overpack ---------'
                debug1 = Sheet1.Cells(x, 16).Value
                debug2 = CInt(Sheet1.Cells(x, 16).Value)
                debug3 = Sheet1.Cells(x, 16).Value >= 1
                If Sheet1.Cells(x, 16).Value = CInt(Sheet1.Cells(x, 16).Value) And Sheet1.Cells(x, 16).Value >= 1 Then
                    Sheet2.Rows(TwoRow).Offset(1).EntireRow.Insert
                    Sheet2.Cells(TwoRow, 4).Value = stars & "  Overpack # " & Sheet1.Cells(x, 16).Value & stars
                    Sheet2.Cells(TwoRow, 8).Value = Sheet1.Cells(x, 17).Value
                    Sheet2.Range(Cells(TwoRow, 4), Cells(TwoRow, 7)).Merge
                    Sheet2.Cells(TwoRow, 4).HorizontalAlignment = xlCenter
                    Sheet2.Cells(TwoRow, 3).Value = Sheet1.Cells(x, 6).Value
                    Sheet2.Cells(TwoRow, 2).Value = Sheet1.Cells(x, 1).Value 'AWB
                    Sheet2.Cells(TwoRow, 2).NumberFormat = "0000-0000-0000"
                    Sheet2.Cells(TwoRow, 9).Value = "    *** "
                    Sheet2.Cells(TwoRow, 10).Value = "  ** "
                    TwoRow = TwoRow + 1
                End If 'end of OP
                
                
                '------if for stuff that is not OP or ALPKN1----------'
                If Sheet1.Cells(x, 14) + Sheet1.Cells(x, 16) = 0 Then
                    Sheet2.Cells(TwoRow, 2).Value = Sheet1.Cells(x, 1).Value 'AWB
                    Sheet2.Cells(TwoRow, 2).NumberFormat = "0000-0000-0000"
                    Sheet2.Cells(TwoRow, 3).Value = Sheet1.Cells(x, 6).Value 'URSA DEST
                Else 'merge and copy stuff
                    Sheet2.Range(Cells(TwoRow, 1), Cells(TwoRow, 3)).Merge
                End If
                'filling out other stuff
                Sheet2.Cells(TwoRow, 4).Value = Sheet1.Cells(x, 4).Value 'UN#
                Sheet2.Cells(TwoRow, 5).Value = Sheet1.Cells(x, 7).Value 'Class
                Sheet2.Cells(TwoRow, 6).Value = Sheet1.Cells(x, 5).Value 'PSN
                Sheet2.Cells(TwoRow, 7).Value = Sheet1.Cells(x, 8).Value 'PG
                Sheet2.Cells(TwoRow, 8).Value = Sheet1.Cells(x, 9).Value 'Pcs
                Sheet2.Cells(TwoRow, 9).Value = Sheet1.Cells(x, 10).Value 'WT/Amt
                Sheet2.Cells(TwoRow, 10).Value = Sheet1.Cells(x, 11).Value 'UM
                x = x + 1 'increase row on sheet1 by +1
            Else 'if ursa has changed
    '---------------------------------------------------------
                Worksheets("CanManifest").Select
                Sheet2.Range(Cells(TwoRow, 1), Cells(TwoRow, col + 8)).Select
                With Selection
                    Selection.Merge
                    Selection.Font.Bold = True
                    Selection.Font.Underline = xlUnderlineStyleSingle
                    .HorizontalAlignment = xlCenter
                    .VerticalAlignment = xlBottom
                    .WrapText = False
                    .Orientation = 0
                    .AddIndent = False
                    .IndentLevel = 0
                    .ShrinkToFit = False
                    .ReadingOrder = xlContext
                    .MergeCells = True
                With Selection.Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .ThemeColor = xlThemeColorDark1
                    .TintAndShade = -0.149998474074526
                    .PatternTintAndShade = 0
                End With
                End With
                '-------------- merge and put incoming dg from "station ursa"
                Sheet2.Range(Cells(TwoRow, 1), Cells(TwoRow, col + 8)).Value = str + URSA
            End If 'If URSA = oldURSA end
            
    TwoRow = TwoRow + 1 'increase row usage on canManifest screen
    
    oldURSA = URSA 'changes old to new
End If
Loop 'loop from x=y

Sheet2.Columns("E:E").HorizontalAlignment = xlCenter
End Sub

Sub HAZ_LIST()
Dim currentOrg As String
x = 3
col = 2
y = 3
TwoRow = 6
oldURSA = ""
Sheet2.Cells(2, 6).Value = UCase(BORG.CanSelectGUI.Value)
Do Until Sheet1.Cells(y, 1) = ""
y = y + 1
'roadrunner
Loop
Dim URSA As String
stars = " ***************** "

Do Until x = y
If Sheet1.Cells(x, 1).EntireRow.Hidden = False Then
    URSA = Sheet1.Cells(x, 2)
        '--------------- if statement for All-Packed-In-One's ---------'
        If Sheet1.Cells(x, 14).Value = CInt(Sheet1.Cells(x, 14).Value) And Sheet1.Cells(x, 14).Value >= 1 Then
            Sheet2.Rows(TwoRow).Offset(1).EntireRow.Insert
            Sheet2.Cells(TwoRow, 4).Value = stars & "  All-Packed-In-One # " & Sheet1.Cells(x, 14).Value & stars
            Sheet2.Cells(TwoRow, 8).Value = Sheet1.Cells(x, 15).Value
            Sheet2.Range(Cells(TwoRow, 4), Cells(TwoRow, 7)).Merge
            Sheet2.Cells(TwoRow, 4).HorizontalAlignment = xlCenter
            Sheet2.Cells(TwoRow, 3).Value = Sheet1.Cells(x, 6).Value
            Sheet2.Cells(TwoRow, 2).Value = Sheet1.Cells(x, 1).Value 'AWB
            Sheet2.Cells(TwoRow, 2).NumberFormat = "0000-0000-0000"
            Sheet2.Cells(TwoRow, 9).Value = "    *** "
            Sheet2.Cells(TwoRow, 10).Value = "  ** "
            TwoRow = TwoRow + 1
        End If 'end of ALPKN1
        '--------------- if statement for Overpack ---------'
       
        If Sheet1.Cells(x, 16).Value = CInt(Sheet1.Cells(x, 16).Value) And Sheet1.Cells(x, 16).Value >= 1 Then
            Sheet2.Rows(TwoRow).Offset(1).EntireRow.Insert
            Sheet2.Cells(TwoRow, 4).Value = stars & "  Overpack # " & Sheet1.Cells(x, 16).Value & stars
            Sheet2.Cells(TwoRow, 8).Value = Sheet1.Cells(x, 17).Value
            Sheet2.Range(Cells(TwoRow, 4), Cells(TwoRow, 7)).Merge
            Sheet2.Cells(TwoRow, 4).HorizontalAlignment = xlCenter
            Sheet2.Cells(TwoRow, 3).Value = Sheet1.Cells(x, 6).Value
            Sheet2.Cells(TwoRow, 2).Value = Sheet1.Cells(x, 1).Value 'AWB
            Sheet2.Cells(TwoRow, 2).NumberFormat = "0000-0000-0000"
            Sheet2.Cells(TwoRow, 9).Value = "    *** "
            Sheet2.Cells(TwoRow, 10).Value = "  ** "
            TwoRow = TwoRow + 1
        End If 'end of OP
        
        '------if for stuff that is not OP or ALPKN1----------'
        If Sheet1.Cells(x, 14) + Sheet1.Cells(x, 16) = 0 Then
            Sheet2.Cells(TwoRow, 2).Value = Sheet1.Cells(x, 1).Value 'AWB
            Sheet2.Cells(TwoRow, 2).NumberFormat = "0000-0000-0000"
            Sheet2.Cells(TwoRow, 3).Value = Sheet1.Cells(x, 6).Value 'URSA DEST
        Else 'merge and copy stuff
            Sheet2.Range(Cells(TwoRow, 1), Cells(TwoRow, 3)).Merge
        End If
        
        'filling out other stuff
        Sheet2.Cells(TwoRow, 4).Value = Sheet1.Cells(x, 4).Value 'UN#
        Sheet2.Cells(TwoRow, 5).Value = Sheet1.Cells(x, 7).Value 'Class
        Sheet2.Cells(TwoRow, 6).Value = Sheet1.Cells(x, 5).Value 'PSN
        Sheet2.Cells(TwoRow, 7).Value = Sheet1.Cells(x, 8).Value 'PG
        Sheet2.Cells(TwoRow, 8).Value = Sheet1.Cells(x, 9).Value 'Pcs
        Sheet2.Cells(TwoRow, 9).Value = Sheet1.Cells(x, 10).Value 'WT/Amt
        Sheet2.Cells(TwoRow, 10).Value = Sheet1.Cells(x, 11).Value 'UM
        x = x + 1 'increase row on sheet1 by +1
        
TwoRow = TwoRow + 1 'increase row usage on canManifest screen
End If
Loop 'loop from x=y
Sheet2.Columns("E:E").HorizontalAlignment = xlCenter
End Sub

Sub HAZ_LIST_w_flightInfo()
Dim currentOrg As String
x = 3
col = 2
y = 3
TwoRow = 6
oldURSA = ""

Sheet2.Cells(2, 6).Value = UCase(BORG.CanSelectGUI.Value)
Do Until Sheet1.Cells(y, 1) = ""
    y = y + 1
Loop

Dim URSA As String
Dim str As String
str = "DG coming from can "
str2 = " on "
stars = "  *****************  "

Do Until x = y
If Sheet1.Cells(x, 1).EntireRow.Hidden = False Then
    URSA = Sheet1.Cells(x, 19)
        If URSA = oldURSA Then 'if ursa has not changed
        '--------------- if statement for All-Packed-In-One's ---------'
            If Sheet1.Cells(x, 14).Value = CInt(Sheet1.Cells(x, 14).Value) And Sheet1.Cells(x, 14).Value >= 1 Then
                Sheet2.Rows(TwoRow).Offset(1).EntireRow.Insert
                Sheet2.Cells(TwoRow, 4).Value = stars & "  All-Packed-In-One # " & Sheet1.Cells(x, 14).Value & stars
                Sheet2.Cells(TwoRow, 8).Value = Sheet1.Cells(x, 15).Value
                Sheet2.Range(Cells(TwoRow, 4), Cells(TwoRow, 7)).Merge
                Sheet2.Cells(TwoRow, 4).HorizontalAlignment = xlCenter
                Sheet2.Cells(TwoRow, 3).Value = Sheet1.Cells(x, 6).Value
                Sheet2.Cells(TwoRow, 2).Value = Sheet1.Cells(x, 1).Value 'AWB
                Sheet2.Cells(TwoRow, 2).NumberFormat = "0000-0000-0000"
                Sheet2.Cells(TwoRow, 9).Value = "    *** "
                Sheet2.Cells(TwoRow, 10).Value = "  ** "
                TwoRow = TwoRow + 1
            End If 'end of ALPKN1
            '--------------- if statement for Overpack ---------'
            If Sheet1.Cells(x, 16).Value = CInt(Sheet1.Cells(x, 16).Value) And Sheet1.Cells(x, 16).Value >= 1 Then
                Sheet2.Rows(TwoRow).Offset(1).EntireRow.Insert
                Sheet2.Cells(TwoRow, 4).Value = stars & "  Overpack # " & Sheet1.Cells(x, 16).Value & stars
                Sheet2.Cells(TwoRow, 8).Value = Sheet1.Cells(x, 17).Value
                Sheet2.Range(Cells(TwoRow, 4), Cells(TwoRow, 7)).Merge
                Sheet2.Cells(TwoRow, 4).HorizontalAlignment = xlCenter
                Sheet2.Cells(TwoRow, 3).Value = Sheet1.Cells(x, 6).Value
                Sheet2.Cells(TwoRow, 2).Value = Sheet1.Cells(x, 1).Value 'AWB
                Sheet2.Cells(TwoRow, 2).NumberFormat = "0000-0000-0000"
                Sheet2.Cells(TwoRow, 9).Value = "    *** "
                Sheet2.Cells(TwoRow, 10).Value = "  ** "
                TwoRow = TwoRow + 1
            End If 'end of OP
            
            
            '------if for stuff that is not OP or ALPKN1----------'
            If Sheet1.Cells(x, 14) + Sheet1.Cells(x, 16) = 0 Then
            Sheet2.Cells(TwoRow, 2).Value = Sheet1.Cells(x, 1).Value 'AWB
            Sheet2.Cells(TwoRow, 2).NumberFormat = "0000-0000-0000"
            Sheet2.Cells(TwoRow, 3).Value = Sheet1.Cells(x, 6).Value 'URSA DEST
            Else 'merge and copy stuff
            Sheet2.Range(Cells(TwoRow, 1), Cells(TwoRow, 3)).Merge
            End If
            'filling out other stuff
            Sheet2.Cells(TwoRow, 4).Value = Sheet1.Cells(x, 4).Value 'UN#
            Sheet2.Cells(TwoRow, 5).Value = Sheet1.Cells(x, 7).Value 'Class
            Sheet2.Cells(TwoRow, 6).Value = Sheet1.Cells(x, 5).Value 'PSN
            Sheet2.Cells(TwoRow, 7).Value = Sheet1.Cells(x, 8).Value 'PG
            Sheet2.Cells(TwoRow, 8).Value = Sheet1.Cells(x, 9).Value 'Pcs
            Sheet2.Cells(TwoRow, 9).Value = Sheet1.Cells(x, 10).Value 'WT/Amt
            Sheet2.Cells(TwoRow, 10).Value = Sheet1.Cells(x, 11).Value 'UM
            x = x + 1 'increase row on sheet1 by +1
        Else 'if ursa has changed
'---------------------------------------------------------
        Worksheets("CanManifest").Select
        Sheet2.Range(Cells(TwoRow, 1), Cells(TwoRow, col + 8)).Select
            With Selection
            Selection.Merge
            Selection.Font.Bold = True
            Selection.Font.Underline = xlUnderlineStyleSingle
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlBottom
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = True
            With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorDark1
            .TintAndShade = -0.149998474074526
            .PatternTintAndShade = 0
            End With
            End With
            '-------------- merge and put incoming dg from "station ursa"
            Sheet2.Range(Cells(TwoRow, 1), Cells(TwoRow, col + 8)).Value = str + URSA + str2 + Sheet1.Cells(x, 20).Text
        End If 'If URSA = oldURSA end
        
    TwoRow = TwoRow + 1 'increase row usage on canManifest screen
    
    oldURSA = URSA 'changes old to new
Else: x = x + 1
End If
Loop 'loop from x=y

Sheet2.Columns("E:E").HorizontalAlignment = xlCenter
End Sub

Sub gasCount()
row = 6
Gas = 0
Do Until (Sheet2.Cells(row, 4)) & (Sheet2.Cells(row + 1, 4)) & (Sheet2.Cells(row + 1, 4)) = ""
If InStr(1, Sheet2.Cells(row, 5).Value, "2.2") <> 0 Then
    If Trim(Sheet2.Cells(row, 10).Value) = "KG" Then
        Gas = Gas + ((Sheet2.Cells(row, 9).Value) * (Sheet2.Cells(row, 8).Value))
    Else
        x = (Sheet2.Cells(row, 9).Value) / 100
        Gas = Gas + x * (Sheet2.Cells(row, 8).Value)
    End If
End If

row = row + 1
Loop
If Gas > 0 Then
    Sheet2.Cells(4, 1).Value = "    (Total 2.2 in can = " & Gas & " KG)    "
End If
End Sub

Sub pieceCount()
TotalPieces = 0
row = 6

Do Until Sheet2.Cells(row, 4).Text + Sheet2.Cells(row + 1, 4).Text + Sheet2.Cells(row + 2, 4).Text = ""
n = Sheet2.Cells(row, 4).Text + Sheet2.Cells(row + 1, 4).Text + Sheet2.Cells(row + 2, 4).Text
u = Right(Sheet2.Cells(row, 2).Value, 4)
If u <> "" Then
    If CInt(u) > 1 Then
        TotalPieces = TotalPieces + Sheet2.Cells(row, 8).Value
    End If
End If
row = row + 1
Loop
If TotalPieces >= 1 Then
    Sheet2.Cells(4, 1).Value = Sheet2.Cells(4, 1) & "   (Total Pieces in Can = " & TotalPieces & ")    "
End If
End Sub

Sub QoL_stuff()

'once manifest is created.. populate it with QoL stuff
BORG.labelUpdater.Caption = "Counting Gas"
Call Module4.gasCount
BORG.labelUpdater.Caption = "Counting Pieces"
Call Module4.pieceCount

End Sub
