Dim host As Object
Public bz_connected As Boolean

Option Compare Text

Function BZinit()
Set host = openBlueZoneSession
End Function

Function openBlueZoneSession() As Object

ChDir "C:\"
Set host = CreateObject("BZwhll.whllobj")
retval = host.OpenSession(0, 11, "fdx3270.zmd", 30, 1)
host.Connect ("K")
'host.WaitCursor 1, 9, 1, 1
Set Wnd = host.Window()

Wnd.Caption = "BDG Window"
Wnd.State = 0 ' 0 restore, 1 minimize, 2 maximize
Wnd.Visible = BORG.Bluezone_Vis.Value
host.waitready 1, 500

bz_connected = True
Set openBlueZoneSession = host


End Function

Function BZreadscreen(length As Integer, x As Integer, y As Integer, Optional wait As Boolean = False) As String
On Error GoTo erroutread
Dim loopcheck As Integer
loopcheck = 0
read:
Dim BZdata As String
BZdata = ""
BZmodule.host.readscreen BZdata, length, x, y
If wait = True Then host.waitready 1, 51
BZreadscreen = BZdata
Exit Function
erroutread:
    Set host = openBlueZoneSession
    loopcheck = loopcheck + 1
    If loopcheck >= 5 Then
        Exit Function
    End If
    GoTo read
    
End Function
Sub testcaller()
Set host = openBlueZoneSession

host.readscreen text, 12, 10, 10
Call BZwritescreen("text", 11, 25)
x = BZreadscreen(5, 5, 5)
Call BZsendKey("@C")

End Sub
Function BZwritescreen(text As String, x As Integer, y As Integer, Optional wait As Boolean = False)
On Error GoTo erroutwrite
Dim loopcheck As Integer
loopcheck = 0
writeme:
If TypeName(host) = "IWhllObj" Then
    host.writescreen text, x, y
    If wait = True Then host.waitready 1, 51
Else
    MsgBox ("error" & Err.Number & " in bzwritescreen")
End If
Exit Function
erroutwrite:
    Set host = openBlueZoneSession
    loopcheck = loopcheck + 1
    If loopcheck >= 5 Then
        Exit Function
    End If
    GoTo writeme
End Function

Function BZsendKey(text As String, Optional wait As Boolean = True)
On Error GoTo erroutSend
Dim loopcheck As Integer
loopcheck = 0
pushkey:
host.sendkey text
If wait = True Then host.waitready 1, 51
Exit Function
erroutSend:
    Set host = openBlueZoneSession
    loopcheck = loopcheck + 1
    If loopcheck >= 5 Then
        Exit Function
    End If
    GoTo pushkey
End Function

Sub BZgotoAUTOdg()
'checks to see if we are connected to a bluezone session if so
'regardless of current position in system will get us to the DG section of the mainframe display
If BZmodule.BZConnected() Then
    
End If
End Sub

Function BZLogin(empnum As String, password As String) As Boolean
'Call BZsendKey("@C")
'Call BZsendKey("STSA@E", True)
Call BZsendKey("ims@E", True)

fedex = BZreadscreen(35, 1, 23)
iter = 0
Do Until fedex = "F E D E R A L  E X P R E S S  I M S"
    fedex = BZreadscreen(35, 1, 23, True)
    iter = iter + 1
    If iter >= 25 Then
        BZmodule.BZcloseSessions
        x = MsgBox("Error!" & vbNewLine & "Unable to connect to bluezone!" _
            & vbNewLine & "Please try and log in again.", vbCritical, "Error!")
        Exit Function
    End If
Loop

Call BZwritescreen(empnum, 7, 15)
Call BZwritescreen(password, 7, 43)
password = ""
Call BZsendKey("@E", True)
readerror = BZreadscreen(80, 24, 2)
If InStr(1, readerror, "INCORRECT PASSWORD ENTERED") Then
    BZmodule.BZcloseSessions
    x = MsgBox("Incorrect Login Credentials", vbCritical, "Incorrect Password")
    BZLogin = False
    Exit Function
End If
Enter = BZreadscreen(5, 14, 15)
iter = 0
Do Until Enter = "ENTER"
    fedex = BZreadscreen(35, 1, 23, True)
    iter = iter + 1
    If iter >= 25 Then
        BZmodule.BZcloseSessions
        BZLogin = False
        Exit Function
    End If
Loop

BZLogin = True
End Function

Function DGscreenChooser(menu As String) As Boolean
'On Error GoTo erroutScreenChoice
DGscreenInfo = BZreadscreen(50, 1, 20)
If InStr(1, DGscreenInfo, "DANGEROUS GOODS SYSTEM") >= 1 Then
    dgscreeninfo2 = BZreadscreen(50, 2, 20)
    If InStr(1, dgscreeninfo2, "SCAN RECONCILIATION SCREEN") > 1 Then
        Call BZsendKey("@3")
    End If
    
    Call BZwritescreen(menu, 2, 17)
    Call BZsendKey("@E")
Else
    Call BZsendKey("@C", True) 'clears screen in IMS
    Call BZsendKey("asap@e", True) 'types ASAP and enters command
    miscdata = BZreadscreen(32, 1, 2)
    If miscdata = "ASAP COMMAND IS UNKNOWN TO VTAM." Or miscdata = "APPLICATION NOT ACTIVE.         " Then
        res = BZLogin(BORG.empnum, BORG.PasswordBox)
        If res = False Then
            DGscreenChooser = False
            Exit Function
        End If
    End If
    Call BZsendKey("68") 'enter 26 for dg training
    Call BZsendKey("@E", True)
    Call BZwritescreen(menu, 2, 17) 'enters assign into first field to bring us to assign screen
    Call BZwritescreen(BORG.Location.text, 19, 44) 'inputs the location ID in DGinput into station
    If BORG.printerID <> "" Then Call BZwritescreen(BORG.printerID.text, 21, 32)
    Call BZsendKey("@e", True) 'sends enter key to bring us finally to Assign Screen
End If

retCode = BZreadscreen(3, 24, 2)
If retCode = "136" Then
    Call BZwritescreen(BORG.Location.text, 19, 44)
End If
DGscreenChooser = True
Exit Function

erroutScreenChoice:
MsgBox (Err.Number & " error occured in dgscreenchooser sub")
DGscreenChooser = False
End Function
Function BZConnected() As Boolean
If TypeName(host) = "" Then
    terminal = ""
    host.readscreen terminal, 80, 1, 1
    If InStr(1, terminal, "TERMINAL INACTIVE") > 1 Then
        CloseSession (host)
        BZConnected = False
    Else
        BZConnected = True
    End If
Else
    BZConnected = False
End If
End Function
Sub CloseSession()
BORG.labelUpdater.Caption = "Closing IMS..."
host.CloseSession 0, 11
BORG.labelUpdater.Caption = "Done!"
End Sub

Sub BZcloseSessions()

If host Is Nothing Then Exit Sub
Set host = openBlueZoneSession
With host
    .waitready 1, 51
    .CloseSession 0, 11
End With
BORG.labelUpdater.Caption = "Closing Previous Sesson..."
Application.wait Now + TimeValue("00:00:01")

End Sub





