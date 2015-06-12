VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} miniform 
   Caption         =   "MiniMode"
   ClientHeight    =   630
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   1890
   OleObjectBlob   =   "miniform.frx":0000
   ShowModal       =   0   'False
End
Attribute VB_Name = "miniform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandButton1_Click()

End Sub

Private Sub btn_miniClose_Click()
Call BORG.btnClose_Click
Unload Me

End Sub

Private Sub Image1_Click()
Me.Hide
BORG.Show
End Sub
