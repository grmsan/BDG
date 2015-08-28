Attribute VB_Name = "famis"
Dim host As Variant
Private Declare Function MessageBox _
Lib "User32" Alias "MessageBoxA" _
(ByVal hWnd As Long, _
ByVal lpText As String, _
ByVal lpCaption As String, _
ByVal wType As Long) _
As Long

Option Compare Text

Sub famislogin()
famislogingui.Show
ChDir "C:\"
Set host = CreateObject("BZwhll.whllobj")
retval = host.OpenSession(0, 11, "fdx3270.zmd", 30, 1)
host.WaitCursor 1, 9, 1, 1
retval = host.Connect("K")

Set Wnd = host.Window() ' Makes the window invisible.....
Wnd.Visible = True
host.waitready 1, 51
retrylogin:
host.sendkey "@c"
host.waitready 1, 51

host.readscreen miscdata, 11, 3, 32
If miscdata = "FAMIS LOGON" Then GoTo famislogin

host.sendkey "45@e"
host.waitready 1, 51

famislogin:
    host.writescreen famislogingui.empnum, 5, 39
    host.writescreen famislogingui.famispassword, 6, 39
    host.sendkey "@e"
    host.waitready 1, 51

host.readscreen readerror, 78, 23, 2
If InStr(1, readerror, "FEDEX ID MISSING/INVALID") > 1 Or _
   InStr(1, readerror, "ENTERED PASSWORD DOES NOT MATCH PERSONNEL DATABASE") > 1 Then
   MsgBox ("Your Famis Username/Password were invalid" & vbNewLine & "Please try again")
    famislogingui.famispassword = ""
    famislogingui.Show
   If famislogingui.famispassword = "" Then
    Call DGscreenChooser("close", host)
    GrabCloseScreen
    Exit Sub
   Else
    host.writescreen "          ", 5, 39
    GoTo retrylogin
   End If
   Exit Sub
End If

host.readscreen whatmenu, 12, 4, 30

host.sendkey "@3"
host.waitready 1, 51
host.sendkey "1"
host.waitready 1, 51
host.sendkey "@e"
host.waitready 1, 51
host.sendkey "8"
host.waitready 1, 51
host.sendkey "@e"
host.waitready 1, 51

ULDprimMenu:
host.writescreen "6", 2, 11
host.waitready 1, 51
host.writescreen BORG.Location, 4, 12
host.waitready 1, 51
host.sendkey "@E"
host.waitready 1, 51

host.readscreen readerror, 5, 1, 13
If readerror = "FS180" Then GoTo ULDprimMenu

ULDdamagedMenu:
host.sendkey "2@E"
host.waitready 1, 51
host.sendkey "@E"
host.waitready 1, 51
host.readscreen readerror, 6, 1, 13
If readerror = "FS186 " Then GoTo ULDdamagedMenu

Dim row As Integer
row = 3

Do While Sheet4.Cells(row, 1) <> ""
    Call famiscancheck(Sheet4.Cells(row, 1).text, row)
    row = row + 1
Loop

BORG.labelUpdater.Caption = "Finished checking cans..."

End Sub

Sub famiscancheck(can As String, row As Integer)

mytext = can

CanEntertime:
host.writescreen "           ", 4, 16
host.writescreen mytext, 4, 16
host.sendkey "@e"
host.readscreen readerror, 11, 5, 18

If readerror = "SERVICEABLE" Then
    Sheet4.Cells(row, 5) = "SV"
ElseIf readerror = "DAMGD TRUCK" Then
    Sheet4.Cells(row, 5) = "TO"
ElseIf Trim(readerror) = "" Then
    host.readscreen miscdata, 80, 24, 1
    If InStr(1, miscdata, "Please Verify") >= 1 Then
        mytext = InputBox(can & " is an invalid can please check can number and reenter below", "Invalid Can")
        If mytext = "" Then Exit Sub
        Sheet4.Cells(row, 1) = mytext
        GoTo CanEntertime
    End If
Else
    Sheet4.Cells(row, 5) = "NU"
End If



End Sub

Sub famisDestCheck()

host.sendkey "@3"
host.waitready 1, 151
host.sendkey "@3"
host.waitready 1, 151

host.sendkey "3"
host.sendkey "@E"
host.waitready 1, 151
host.sendkey "2"
host.sendkey "@E"
host.waitready 1, 151


row = 3

destcheckStart:
bluerow = 9
Do Until Sheet4.Cells(row, 1) = ""
    host.writescreen Sheet4.Cells(row, 1).text, bluerow, 2
    If row > 8 Then
        host.sendkey "@e"
        host.waitready 1, 51
        
        GoTo destcheckStart
    End If
bluerow = bluerow + 1
row = row + 1
Loop

host.sendkey "@E"
host.waitready 1, 51

badcans = ""

notAvail = False
bluerow = 9
Do Until bluerow >= 17
    host.readscreen mydata, 80, bluerow, 1
    If InStr(1, mydata, Left(BORG.Location.text, 3)) = 0 And Trim(mydata) <> "" Then
        notAvail = True
        host.readscreen badcan, 11, bluerow, 2
        badcans = badcans & vbNewLine & badcan
    End If
    bluerow = bluerow + 1
Loop

If notAvail = True Then
    MsgBox ("Error with cans in Famis." & vbNewLine & "Current cans are not at location set in login section: " & badcans)
End If
End Sub


