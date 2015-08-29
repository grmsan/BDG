Sub p()

raw = "tool"

x = InStr(1, raw, "EXCEPTED")
'Call ASGN023(host)
End Sub

Sub ASGN023(host As Variant, Optional cannum As String = "unassigned")
'MsgBox ("")
Call DGscreenChooser("ASGN023", host)
Dim SeqFinished As String
Dim bluerow As Integer
Call Module1.SETUP

excelrow = GetMaxRow
bluerow = 10

If cannum <> "unassigned" Then
    host.sendkey "@2"
    host.waitready 1, 51
    host.sendkey "@E"
    host.waitready 1, 51
End If

host.readscreen SeqFinished, 26, 24, 2
BORG.labelUpdater.Caption = "Doing work in the Assign Screen..."
Do Until SeqFinished = "018-LAST PAGE IS DISPLAYED"
    host.readscreen SeqFinished, 26, 24, 2
        If canassigned = cannum Then
            BORG.labelUpdater.Caption = "Doing work in the Assign Screen..." & "Grabbing " & (excelrow - 3) & " Pieces"
            host.writescreen "#", bluerow, 2
            host.sendkey "@e"
            host.waitready 1, 51
            host.readscreen awbfour, 4, bluerow, 5
            If awbfour = "    " Then Exit Do
            Sheet1.Cells(excelrow, 3).Value = awbfour
            host.readscreen UNnum, 6, bluerow, 36
                If UNnum = "******" Then UNnum = "Overpack"
            Sheet1.Cells(excelrow, 4).Value = UNnum
            host.readscreen PSN, 10, bluerow, 43
            Sheet1.Cells(excelrow, 5).Value = PSN
            host.readscreen URSA, 8, bluerow, 10
            Sheet1.Cells(excelrow, 6).Value = Trim(URSA)
            host.readscreen hazclass, 4, bluerow, 54
                If hazclass = "***" Then hazclass = "Ovrpk"
            Sheet1.Cells(excelrow, 7).Value = hazclass
            host.readscreen PackingGroup, 3, bluerow, 59
                If PackingGroup = "***" Then PackingGroup = "Ovrk"
            Sheet1.Cells(excelrow, 8).Value = PackingGroup
            host.readscreen pieces, 3, bluerow, 64
            Sheet1.Cells(excelrow, 9).Value = pieces
            
            host.readscreen Weight, 10, bluerow, 68
            Sheet1.Cells(excelrow, 10).Value = Weight
            host.readscreen UnitofMeasure, 2, bluerow, 79
            Sheet1.Cells(excelrow, 11).Value = UnitofMeasure
            
            host.readscreen FullAWB, 12, 24, 21
                If oldawb = FullAWB Then
                    host.waitready 1, 75
                    host.writescreen "#", row, 2
                    host.sendkey "@e"
                    host.waitready 1, 75
                    host.readscreen FullAWB, 12, 24, 21
               End If
            oldawb = FullAWB
            Sheet1.Cells(excelrow, 1).Value = FullAWB
            Sheet1.Cells(excelrow, 13).Value = cannum
            
            host.readscreen APiO, 6, bluerow, 43
            If APiO = "ALPKN1" Then
                host.readscreen APnum, 3, bluerow, 50
                Sheet1.Cells(excelrow, 14).Value = APnum
                host.readscreen APpcs, 3, bluerow, 64
                Sheet1.Cells(excelrow, 15).Value = APpcs
            ElseIf APiO = "OVRPCK" Then
                host.readscreen OPnum, 3, bluerow, 50
                Sheet1.Cells(excelrow, 16).Value = OPnum
                host.readscreen OPpcs, 3, bluerow, 64
                Sheet1.Cells(excelrow, 17).Value = APpcs
            End If

            excelrow = excelrow + 1
        End If
If bluerow >= 18 Then
    host.sendkey "@8"
    host.waitready 1, 51
    bluerow = 10
    host.readscreen SeqFinished, 26, 24, 2
End If
    
bluerow = bluerow + 1
Loop

Sheet1.Columns("A:A").NumberFormat = "000000000000"
Sheet1.Columns("C:C").NumberFormat = "0000"
Sheet1.Columns("J:J").NumberFormat = "0.00000"

End Sub
