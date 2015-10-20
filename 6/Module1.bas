
Sub getstarted()
Sheets("CanManifest").Activate
Module1.DEL
Application.Visible = False
Module1.OpenBDG
End Sub

Sub printFun()
'Print the Manifest using the default printer
    Application.ActiveWindow.DisplayGridlines = True
    Application.PrintCommunication = True
    ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True, IgnorePrintAreas:=False
End Sub

Sub DEL()
'Clear up old manifest data
    Sheet1.Rows("3:9999").Clear
    Sheet2.Rows("6:9999").Clear
    Sheet3.Rows("16:9999").Clear
    Sheet2.Cells(4, 1) = ""
    'Selection.Delete Shift:=xlUp
    'Sheet2.Cells(1, 1).Select
End Sub
Sub deletecans()
    myrow = 3
    Do Until Sheet3.Cells(myrow, 12) = ""
        Sheet3.Cells(myrow, 12).Clear
        Sheet3.Cells(myrow, 13).Clear
        Sheet3.Cells(myrow, 14).Clear
        myrow = myrow + 1
    Loop
End Sub
Sub OpenBDG()

Load BORG
BORG.Show

End Sub

Sub SETUP()
'
' SETUP Macro
' SET UP THE CATEGORIES FOR FILTERING AUTODG DATA
'

    Sheet1.Cells(2, 1).Value = "Full AWB"
    Sheet1.Cells(2, 2).Value = "Station"
    Sheet1.Cells(2, 3).Value = "AWBfour"
    Sheet1.Cells(2, 4).Value = "UN#"
    Sheet1.Cells(2, 5).Value = "PSN"
    Sheet1.Cells(2, 6).Value = "URSA"
    Sheet1.Cells(2, 7).Value = "Class"
    Sheet1.Cells(2, 8).Value = "PG"
    Sheet1.Cells(2, 9).Value = "Pcs"
    Sheet1.Cells(2, 10).Value = "WT/Amt"
    Sheet1.Cells(2, 11).Value = "Units"
    Sheet1.Cells(2, 12).Value = "Station ID"
    Sheet1.Cells(2, 13).Value = "Can Assigned"
    Sheet1.Cells(2, 14).Value = "APio ID"
    Sheet1.Cells(2, 15).Value = "APio Pcs"
    Sheet1.Cells(2, 16).Value = "OP ID"
    Sheet1.Cells(2, 17).Value = "OP Pcs"
    Sheet1.Cells(2, 18).Value = "oldIgnore"
    Sheet1.Cells(2, 19).Value = "Arrival Can"
    Sheet1.Cells(2, 20).Value = "Arriving on Route"
    Sheet1.Cells(2, 21).Value = "Predict Assign Can"
    Sheet1.Cells(2, 22).Value = "Predict Assign Dest"
    Sheet1.Cells(2, 23).Value = "hazType"
    Sheet1.Cells(2, 24).Value = "isLocal"
    Sheet1.Cells(2, 25).Value = ""
End Sub

Sub DELmanifestSheet() 'Clear up old manifest data on SortMen screen
    Sheet2.Rows("6:9999").Clear
    Sheet2.Cells(4, 1) = ""
End Sub
