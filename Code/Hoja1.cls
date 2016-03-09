VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Hoja1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub CommandButton1_Click()
Application.ScreenUpdating = False


Sheets("CTASAS").Visible = True
Sheets("CTASAS").Select
Sheets("CTASAS (2)").Visible = True
Sheets("DATOS").Visible = False
Call all

Application.ScreenUpdating = True
End Sub

Private Sub OptionButton1_Click()
Application.ScreenUpdating = False
'Manualmente
Sheets("DATOS").Range("K11:L14").Select
Selection.ClearContents

Sheets("DATOS").Range("j21:L21").Select
Selection.ClearContents

Sheets("DATOS").Select
Range("k23").Value = 1
Application.ScreenUpdating = True
End Sub

Private Sub OptionButton2_Click()
Application.ScreenUpdating = False
'Automático
Sheets("DATOS").Range("k11").Value = Sheets("T_DATOS").Range("S15").Value
Sheets("DATOS").Range("l11").Value = Sheets("T_DATOS").Range("T15").Value


Sheets("DATOS").Range("k12").Value = Sheets("T_DATOS").Range("S16").Value
Sheets("DATOS").Range("l12").Value = Sheets("T_DATOS").Range("T16").Value

Sheets("DATOS").Range("k13").Value = Sheets("T_DATOS").Range("S17").Value
Sheets("DATOS").Range("l13").Value = Sheets("T_DATOS").Range("T17").Value

Sheets("DATOS").Range("k14").Value = Sheets("T_DATOS").Range("S18").Value
Sheets("DATOS").Range("l14").Value = Sheets("T_DATOS").Range("T18").Value

Sheets("DATOS").Range("j21:L21").Select
Selection.ClearContents


Sheets("DATOS").Select
Range("k23").Value = 0
Application.ScreenUpdating = True

End Sub
