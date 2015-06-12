Attribute VB_Name = "AutoUpdate"
Public Const strThisVer As String = "5.0"


Public Const strFileName As String = "BDG"

Sub AddModuleToProject(modName As String)
    Dim VBProj As VBIDE.VBProject
    Dim VBComp As VBIDE.VBComponent
    Set VBProj = ActiveWorkbook.VBProject
    
    Set VBComp = VBProj.VBComponents.Add(vbext_ct_StdModule)
    VBComp.Name = modName
End Sub
Sub SETUP()
    On Error GoTo exitsub
   'allows use of VBIDE
    ThisWorkbook.VBProject.References.AddFromGuid GUID:="{0002E157-0000-0000-C000-000000000046}", _
        Major:=5, Minor:=3
exitsub:
End Sub

Sub Updater()
Call AutoUpdate.SETUP
Dim myURL As String
Dim modulenames As Collection
Set modulenames = getModuleNames
Dim tempStr As String
For Each Item In modulenames
    tempStr = Item
    myURL = modAddress(tempStr)
    BDGdata = getBDGdata(myURL)
    If BDGdata = "Module Doesn't Exist" Then
        'skip it
    Else
        BDGdata = "'" & BDGdata
        Call remakeModule(tempStr, BDGdata)
    End If
Next
    
    'now to update the version number within this module itself....
    strCurVer = getBDGdata("https://raw.githubusercontent.com/grmsan/Learning-Repo/master/BDG/BDGversion")
    Dim VBProj As VBIDE.VBProject
    Dim VBComp As VBIDE.VBComponent
    Dim CodeMod As VBIDE.CodeModule
    Dim LineNum As Long
    Set VBProj = ActiveWorkbook.VBProject
    Set VBComp = VBProj.VBComponents("AutoUpdate")
    Set CodeMod = VBComp.CodeModule
    With CodeMod
        .ReplaceLine 1, "Public Const strThisVer As String = " & """" & strCurVer
        .DeleteLines 2, 1
    End With
    
x = MsgBox("BDG is now up to date!", vbInformation)

End Sub

Sub remakeModule(myModule As String, BDGdata As Variant)
    Dim VBProj As VBIDE.VBProject
    Dim VBComp As VBIDE.VBComponent
    Dim CodeMod As VBIDE.CodeModule
    Dim LineNum As Long
    Const DQUOTE = """" ' one " character
    
    On Error GoTo createModule
    
    Set VBProj = ActiveWorkbook.VBProject
    Set VBComp = VBProj.VBComponents(myModule)
    
    Set CodeMod = VBComp.CodeModule

    With CodeMod
        .DeleteLines 1, .CountOfLines
        .InsertLines 1, BDGdata
    End With
    Exit Sub
createModule:
AddModuleToProject (myModule)
myCMD = InputBox("Adding module " & myModule & " to BDG." & vbNewLine & "If message keeps appearing type stop to stop BDG")
If myCMD = "stop" Then Exit Sub
Call remakeModule(myModule, BDGdata)
End Sub

Function modAddress(myModule As String) As String
    modAddress = "https://raw.githubusercontent.com/grmsan/Learning-Repo/master/BDG/" & myModule & ".bas"
End Function


Function CreateModList()
    Dim myFile As String
    Dim VBProj As VBIDE.VBProject
    Dim VBComp As VBIDE.VBComponent
    
    Set VBProj = ActiveWorkbook.VBProject
    Dim col As Collection
    Set col = New Collection
    For Each Item In VBProj.VBComponents
        col.Add Item.Name
    Next
    
    'myFile = Application.DefaultFilePath & "\Modules.txt"
    myFile = "C:\Learning-Repo\BDG\Modules.txt"
    Open myFile For Output As #1
    
    For Each Item In col
            Write #1, Item
    Next
    
    Close #1
End Function

Function currentversion() As Boolean
    Dim strThisVer As String
    Dim strCurVer As String
    Dim nThisVer As Double
    Dim nCurVer As Double
    
    On Error GoTo ErrOut
    strCurVer = getBDGdata("https://raw.githubusercontent.com/grmsan/Learning-Repo/master/BDG/BDGversion")
    
    nCurVer = Val(Trim(strCurVer))
    nThisVer = Val(AutoUpdate.strThisVer)
 
    If nCurVer <= nThisVer Then
        currentversion = True
        MsgBox ("up to date!")
    Else
        MsgBox ("not up to date")
        currentversion = False
    End If
    Exit Function

ErrOut:
    MsgBox ("Error occured while checking for newest version of BDG")

End Function

Function getBDGdata(URL As String) As String
    'If Error("80072ee7") Then GoTo errhandler
    
    Dim xmlHTTP
    Set xmlHTTP = CreateObject("MSXML2.ServerXMLHTTP")
    
    'xmlHTTP.Open "POST", "https://raw.githubusercontent.com/grmsan/Learning-Repo/master/BDG/BDGversion", False
    'xmlHTTP.Open "POST", "https://raw.githubusercontent.com/grmsan/Learning-Repo/master/BDG/Modules.bas", False
    xmlHTTP.Open "POST", URL, False
    xmlHTTP.send "Doesn't matter what I put here, response always the same"
    
    If xmlHTTP.responseText = "Not Found" Then
        getBDGdata = "Module Doesn't Exist"
    Else
        getBDGdata = xmlHTTP.responseText
    End If
    Exit Function
    
errhandler:
    MsgBox ("Trouble connecting with internet")
End Function


Function getModuleNames() As Collection
Dim modstr As Collection
Set modstr = New Collection
modData = getBDGdata("https://raw.githubusercontent.com/grmsan/Learning-Repo/master/BDG/Modules.txt")
mystr = Split(modData, """")
For Each Item In mystr
    If Len(Item) > 1 Then
        modstr.Add Item
    End If
Next
Set getModuleNames = modstr
End Function
