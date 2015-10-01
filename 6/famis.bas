Option Compare Text

Sub famislogin()
famislogingui.Show

retrylogin:

Call BZsendKey("@c", True)

miscdata = BZreadscreen(11, 3, 32)
If miscdata = "FAMIS LOGON" Then GoTo famislogin

Call BZsendKey("45@e", True)

famislogin:
    Call BZwritescreen(famislogingui.EmpNum, 5, 39)
    Call BZwritescreen(famislogingui.famispassword, 6, 39)
    Call BZsendKey("@e")

readerror = BZreadscreen(78, 23, 2)
If InStr(1, readerror, "FEDEX ID MISSING/INVALID") > 1 Or _
   InStr(1, readerror, "ENTERED PASSWORD DOES NOT MATCH PERSONNEL DATABASE") > 1 Then
   MsgBox ("Your Famis Username/Password were invalid" & vbNewLine & "Please try again")
    famislogingui.famispassword = ""
    famislogingui.Show
   If famislogingui.famispassword = "" Then
    Call DGscreenChooser("close")
    GrabCloseScreen
    Exit Sub
   Else
    Call BZwritescreen("          ", 5, 39)
    GoTo retrylogin
   End If
   Exit Sub
End If

whatmenu = BZreadscreen(12, 4, 30)
Call BZsendKey("@3", True)
Call BZsendKey("1", True)
Call BZsendKey("@e", True)
Call BZsendKey("8", True)
Call BZsendKey("@e", True)

ULDprimMenu:
Call BZwritescreen("3", 2, 11, True)
Call BZwritescreen(BORG.Location, 4, 12, True)
Call BZsendKey("@E", True)

readerror = BZreadscreen(5, 1, 13)
If readerror = "FS180" Then GoTo ULDprimMenu

ULDdamagedMenu:
Call BZsendKey("2@E", True)
readerror = BZreadscreen(6, 1, 13)
If readerror <> "FS1832" Then GoTo ULDdamagedMenu

Dim row As Integer
row = 3
Dim bluerow As Integer
bluerow = 9

For i = 0 To 7
    Call BZwritescreen("          ", i + 9, 2)
Next

Do While Sheet4.Cells(row, 1) <> ""
    If InStr(1, "BULK", Sheet4.Cells(row, 1)) <> 0 Then
        row = row + 1
    Else
        Call BZwritescreen(Sheet4.Cells(row, 1).text, bluerow, 2)
        If bluerow = 16 Or Sheet4.Cells(row + 1, 1) = "" Then
            Call BZsendKey("@E", True)
            Call famischeckcans
            bluerow = 8
        End If
        row = row + 1
        bluerow = bluerow + 1
    End If
Loop

BORG.labelUpdater.Caption = "Finished checking cans..."

End Sub
Sub famischeckcans()
screencheck = BZreadscreen(6, 1, 13)
If screencheck <> "FS1832" Then Exit Sub

Dim bluerow As Integer
For bluerow = 9 To 16
    miscdata = BZreadscreen(70, bluerow, 2)
    If Trim(miscdata) = "" Then Exit Sub
    If Trim(Right(miscdata, 60)) = "" Then
        x = MsgBox(Trim(Left(miscdata, 10)) & " does not exist in System. Please re-check can number.", vbCritical, "Can Doesn't Exist!")
    Else
        cannum = Trim(Left(miscdata, 10))
        svc = Mid(miscdata, 15, 2)
        isLocal = InStr(1, miscdata, Left(BORG.Location.text, 3))
        If isLocal <= 0 Then
            Call MsgBox(cannum & " is not currently at your location in FAMIS. Please re-check can number to verify correct entry.", vbCritical, "ERROR: can not at location")
        End If
        Dim row As Integer
        row = 3
        Do Until Sheet4.Cells(row, 1) = cannum
            row = row + 1
        Loop
        Sheet4.Cells(row, 5) = svc
        
    End If
Next

End Sub
