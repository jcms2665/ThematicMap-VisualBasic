VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm3 
   Caption         =   "Instrucciones Generales"
   ClientHeight    =   6330
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6930
   OleObjectBlob   =   "UserForm3.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "UserForm3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
Application.ScreenUpdating = False

UserForm3.Hide
Sheets("DATOS").Select
Sheets("DATOS").OptionButton1.Value = True

For i = 7 To 38
                Cells(i, 5).Value = 0
Next i
    Range("J21:L21").Select
    Selection.ClearContents
    Range("L24").Select

Application.ScreenUpdating = True
End Sub
Private Sub UserForm_Initialize()
Application.ScreenUpdating = False
Sheets("T_DATOS").Visible = False
Sheets("CTASAS (2)").Visible = False
Sheets("CTASAS").Visible = False

Sheets("DATOS").Visible = True
    Range("k11:L14").Select
    Selection.ClearContents
Application.ScreenUpdating = True
End Sub
