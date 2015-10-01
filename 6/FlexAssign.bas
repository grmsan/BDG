
Dim cannum As Variant
Dim cantype As Variant
Dim cansplit As Variant
Dim candest As Variant
Dim ADGfind As Variant
Dim IDGfind As Variant
Dim pieces As Integer
Dim retval As Variant

Option Compare Text

Sub FlexAssignDirectory(Optional can As String = "ALL")
If can = "ALL" Then
    Call setupAssignArrays
Else
    cannum = Array(BORG.txt_canNum.text)
    cansplit = Array(BORG.combo_splitName.text)
    candest = Array(BORG.txt_Dest.text)
    cantype = Array(BORG.combo_hazType.text)
End If

ADGfind = Array("1.4", "2.1", "3", "4.", "5", "8")
IDGfind = Array("2.2", "6.", "7", "9")
Call DGscreenChooser("Assign")

Dim i As Integer
Dim hazFilter As String
tempval = UBound(cannum, 1)
For i = 0 To (UBound(cannum, 1))
    Select Case cantype(i)
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
    x = cansplit(i)
    If isSplitLocal(x) = True Then
        Call SuffixAssign(i, hazFilter)
    ElseIf isSplitLocal(x) = False Then
        Call PrefixAssign(i, hazFilter)
    Else 'something has gone horribly wrong....
        MsgBox ("Error occured Please restart BDG")
        Exit Sub
    End If
Next

If can = "ALL" Then Call isAnythingLeft
BORG.labelUpdater.Caption = "Finished assigning " & pieces & " shipment(s)"
Call DGscreenChooser("close")
End Sub
Sub setupAssignArrays()

Dim c_cannums As New Collection
Dim c_cansplits As New Collection
Dim c_candests As New Collection
Dim c_cantypes As New Collection
mytypes = Array("ADG", "ALL", "IDG")
t = 0

addtocollections:
row = 3
Do While Sheet4.Cells(row, 1) <> ""
    If Trim(Sheet4.Cells(row, 4)) = mytypes(t) Then
       c_cannums.Add CStr(Sheet4.Cells(row, 1)) '  dynamically add value to the end
       c_cansplits.Add CStr(Sheet4.Cells(row, 2))
       c_candests.Add CStr(Sheet4.Cells(row, 3))
       c_cantypes.Add CStr(Sheet4.Cells(row, 4))
    End If
   row = row + 1
Loop

If t < 2 Then
    t = t + 1
    GoTo addtocollections
End If


cannum = toArray(c_cannums) 'convert collection to an array
cansplit = toArray(c_cansplits)
candest = toArray(c_candests)
cantype = toArray(c_cantypes)
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
    ecol = 3
    Do Until Sheet6.Cells(2, ecol).Value = ""
        If Sheet6.Cells(2, ecol).Value = MasterID Then
            isSplitLocal = Not (Sheet6.Cells(3, ecol).Value)
            Exit Function
        End If
        ecol = ecol + 1
    Loop
    MsgBox ("not able to find if " & MasterID & " is a local split" & vbNewLine & "error occured in Function isSplitLocal")
End Function

Sub SuffixAssign(i As Integer, hazFilter As String)
ecol = 3

Do Until Sheet6.Cells(2, ecol) = cansplit(i)
    If Sheet6.Cells(2, ecol).Value = "" Then
        MsgBox ("could not find split " & cansplit(i) & "for can " & cannum(i))
        Exit Sub
    End If
    ecol = ecol + 1
Loop

ERow = 5
Do Until Sheet6.Cells(ERow, ecol) = ""
    Call BZwritescreen("     ", 5, 38)
    Call BZwritescreen(Sheet6.Cells(ERow, ecol).text, 5, 38)
    Call BZwritescreen(hazFilter, 6, 45)
    Call BZsendKey("@e")

ErrorChecker
    Dim bluerow As Integer
    Dim tempstr As String
    bluerow = 10
    miscdata = BZreadscreen(8, bluerow, 18)
    Do Until Trim(miscdata) = ""
CheckingPage:
    miscdata = BZreadscreen(8, bluerow, 18)
        If Right(miscdata, 2) <> "RT" Then
            If Trim(miscdata) <> "" Then
                Call BZwritescreen("A", bluerow, 2)
                pieces = pieces + 1
            ElseIf bluerow = 19 Then
                Call BZwritescreen("          ", 7, 24)
                tempstr = cannum(i)
                Call BZwritescreen(tempstr, 7, 24)
                Call BZwritescreen("    ", 7, 53)
                tempstr = candest(i)
                Call BZwritescreen(tempstr, 7, 53)
                Call BZsendKey("@e")
                Call FlexAssign.ErrorChecker
                Call bulkOveride(CStr(cannum(i)), CStr(cansplit(i)), CStr(candest(i)), CStr(cantype(i)))
                bluerow = 10
                GoTo CheckingPage
            Else
                Call BZwritescreen("          ", 7, 24)
                tempstr = cannum(i)
                Call BZwritescreen(tempstr, 7, 24)
                Call BZwritescreen("    ", 7, 53)
                tempstr = candest(i)
                Call BZwritescreen(tempstr, 7, 53)
                Call BZsendKey("@e")
                Call FlexAssign.ErrorChecker
                Call bulkOveride(CStr(cannum(i)), CStr(cansplit(i)), CStr(candest(i)), CStr(cantype(i)))
            End If
        End If
        bluerow = bluerow + 1
    Loop
    ERow = ERow + 1
Loop
End Sub

Sub PrefixAssign(i As Integer, hazFilter As String)
    Dim bluerow As Integer
    Dim tempstr As String
    ecol = 3
    ignored = 0
    Do Until Sheet6.Cells(2, ecol) = cansplit(i)
        If Sheet6.Cells(2, ecol).Value = "" Then
            MsgBox ("could not find split " & cansplit(i) & "for can " & cannum(i))
            Exit Sub
        End If
        ecol = ecol + 1
    Loop
    ERow = 5
    Do Until Sheet6.Cells(ERow, ecol) = ""
        Call BZwritescreen("  ", 5, 28)
        Call BZwritescreen(Sheet6.Cells(ERow, ecol).text, 5, 28)
        Call BZwritescreen(hazFilter, 6, 45)
        Call BZsendKey("@e")
ErrorChecker
        bluerow = 10
        miscdata = BZreadscreen(8, bluerow, 18)
        Do Until Trim(miscdata) = ""
CheckingPagePrefix:
        
        miscdata = BZreadscreen(8, bluerow, 18)
            If Right(miscdata, 2) <> "RT" Then
                If isUrsaLocal(Trim(Right(miscdata, 5))) <> True Then
                    If Trim(miscdata) <> "" Then
                        Call BZwritescreen("A", bluerow, 2)
                        pieces = pieces + 1
                    ElseIf bluerow = 19 Then
                        Call BZwritescreen("          ", 7, 24)
                        tempstr = cannum(i)
                        Call BZwritescreen(tempstr, 7, 24)
                        Call BZwritescreen("    ", 7, 53)
                        tempstr = candest(i)
                        Call BZwritescreen(tempstr, 7, 53)
                        Call BZsendKey("@e")
                        ignored = 0
                        Call FlexAssign.ErrorChecker
                        tempcannum = CStr(cannum(i))
                        
                        Call bulkOveride(CStr(cannum(i)), CStr(cansplit(i)), CStr(candest(i)), CStr(cantype(i)))
                        bluerow = 10
                        GoTo CheckingPagePrefix
                    Else
                        Call BZwritescreen("          ", 7, 24)
                        tempstr = cannum(i)
                        Call BZwritescreen(tempstr, 7, 24)
                        Call BZwritescreen("    ", 7, 53)
                        tempstr = candest(i)
                        Call BZwritescreen(tempstr, 7, 53)
                        Call BZsendKey("@e")
                        ignored = 0
                        Call FlexAssign.ErrorChecker
                        Call bulkOveride(CStr(cannum(i)), CStr(cansplit(i)), CStr(candest(i)), CStr(cantype(i)))
                    End If
                Else
                    ignored = ignored + 1
                End If
            Else
                ignored = ignored + 1
            End If
            If ignored = 9 Then
                Call BZsendKey("@8")
            End If
            bluerow = bluerow + 1
        Loop
        ERow = ERow + 1
    Loop
End Sub

Sub isAnythingLeft()
Dim row As Integer

Call BZwritescreen("Close ", 2, 17)
Call BZsendKey("@e")
Call BZwritescreen("Assign", 2, 17)
Call BZsendKey("@e")

leftover = 0
row = 10
Do Until row = 18
    miscdata = BZreadscreen(18, row, 51)
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

Call BZwritescreen("assign", 2, 17)
Call BZsendKey("@e")
Call BZwritescreen("C", 6, 45)
Call BZwritescreen("Deleteship", 7, 24)
Call BZsendKey("@e")

Data = "tempdata"
row = 10
Do Until Trim(Data) = ""
    Data = BZreadscreen(15, row, 5)
    If Trim(Data) <> "" Then
        Call BZwritescreen("a", row, 2)
    ElseIf Trim(Data) = "" Then
        Call BZsendKey("@e")
        If row = 18 Then row = 10
    End If
    row = row + 1
Loop
End Sub

Function ErrorChecker()
errormisc = BZreadscreen(3, 24, 2)
If errormisc = "091" Then
    Call BZsendKey("@4")
ElseIf errormisc = "095" Then 'bulk doesn't exist
    oldbulk = BZreadscreen(9, 7, 24)
    Call BZwritescreen("bulk*", 7, 24)
    Call BZsendKey("@E")
    cannum(i) = BZreadscreen(9, 7, 24)
    datarow = 3
    Do Until Sheet4.Cells(datarow, 1) = oldbulk And _
        Sheet4.Cells(datarow, 2) = cansplit(i) And _
        Sheet4.Cells(datarow, 3) = candest(i) And _
        Sheet4.Cells(datarow, 4) = cantype(i)
        If Sheet4.Cells(datarow, 1) = "" Then Exit Do
        datarow = datarow + 1
    Loop
    Sheet4.Cells(datarow, 1) = cannum(i)
ElseIf errormisc = "INV" Then 'invalid container error
    MsgBox ("invalid container")
End If
End Function

Function bulkOveride(cannum As String, cansplit As String, candest As String, cantype As String)
'bulkoveride(cannum(i), cansplit(i), candest(i), cantype(i))

If cannum = "BULK*" Then
    cannum = BZreadscreen(9, 7, 24)
    datarow = 3
    Do Until Sheet4.Cells(datarow, 1) = "BULK*" And _
        Sheet4.Cells(datarow, 2) = cansplit And _
        Sheet4.Cells(datarow, 3) = candest And _
        Sheet4.Cells(datarow, 4) = cantype
        datarow = datarow + 1
    Loop
Sheet4.Cells(datarow, 1) = cannum
End If

End Function




