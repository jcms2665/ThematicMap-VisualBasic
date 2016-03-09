VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Hoja7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub CommandButton1_Click()
Application.ScreenUpdating = False
Call PDF
ActiveWorkbook.Close savechanges:=False
Application.ScreenUpdating = True
End Sub

Private Sub CommandButton2_Click()
Application.ScreenUpdating = False
Call Reiniciar
Application.ScreenUpdating = True
End Sub

Private Sub CommandButton3_Click()
Application.ScreenUpdating = False

'Call Reiniciar
ActiveWorkbook.Close savechanges:=False
Application.ScreenUpdating = True
End Sub

Private Sub CommandButton4_Click()
UserForm1.Show
End Sub

Private Sub CommandButton5_Click()
UserForm2.Show
End Sub
