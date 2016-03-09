Attribute VB_Name = "Módulo1"
Sub all()
Application.ScreenUpdating = False

Sheets("DATOS").Visible = True
Sheets("DATOS").Select
u = Range("k23").Value

If u = 0 Then
        Sheets("T_DATOS").Visible = True
        Sheets("T_DATOS").Select
        q0 = 0
        q1 = Range("r5").Value
        q2 = Range("r6").Value
        q3 = Range("r7").Value
        q4 = Range("r8").Value
Else
        Sheets("DATOS").Select
        q0 = Range("k11").Value
        q1 = Range("L11").Value
        q2 = Range("l12").Value
        q3 = Range("l13").Value
        q4 = Range("l14").Value
End If

k = 0
For i = 5 To 36
                X = Cells(i, 7).Value
                
                f = 0
                If X >= q0 And X <= q1 Then
                            f = 1
                ElseIf X > q1 And X <= q2 Then
                            f = 2
                ElseIf X > q2 And X <= q3 Then
                            f = 3
                ElseIf X > q3 And X <= q4 Then
                            f = 4
                End If
                
                k = i - 4
                
                With ThisWorkbook.Sheets("CTASAS").Shapes("AutoShape " & k)
                        If f = 1 Then
                                    .Fill.ForeColor.RGB = rgbLightYellow
                        ElseIf f = 2 Then
                                    .Fill.ForeColor.RGB = rgbLime
                        ElseIf f = 3 Then
                                    .Fill.ForeColor.RGB = rgbGreen
                        ElseIf f = 4 Then
                                    .Fill.ForeColor.RGB = rgbOlive
                        End If
                End With


                With ThisWorkbook.Sheets("CTASAS (2)").Shapes("AutoShape " & k)
                        If f = 1 Then
                                    .Fill.ForeColor.RGB = rgbLightYellow
                        ElseIf f = 2 Then
                                    .Fill.ForeColor.RGB = rgbLime
                        ElseIf f = 3 Then
                                    .Fill.ForeColor.RGB = rgbGreen
                        ElseIf f = 4 Then
                                    .Fill.ForeColor.RGB = rgbOlive
                        End If
                End With


Next i

     
Sheets("T_DATOS").Visible = False
Sheets("DATOS").Visible = False

Sheets("CTASAS").Select
Sheets("CTASAS (2)").Visible = False
MsgBox "Se ha generado el mapa"
Application.ScreenUpdating = True


End Sub
Sub PDF()
Application.ScreenUpdating = False
Sheets("DATOS").Visible = True
Sheets("DATOS").Select
u = Range("J23").Value
Sheets("CTASAS (2)").Visible = True
ruta = ActiveWorkbook.Path
Sheets("CTASAS (2)").Select

    Sheets("CTASAS (2)").Copy
     ActiveWorkbook.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
        ruta & u & ".pdf", Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=True
            Application.DisplayAlerts = False
                ActiveWindow.Close
            Application.DisplayAlerts = True
    ActiveWorkbook.Sheets("CTASAS").Select
Sheets("CTASAS (2)").Visible = False
Sheets("CTASAS").Select
Application.ScreenUpdating = False
End Sub
Sub Reiniciar()
Application.ScreenUpdating = False
Sheets("DATOS").Visible = True
Sheets("CTASAS").Visible = False
Sheets("DATOS").Select
For i = 7 To 38
                Cells(i, 5).Value = 0
Next i
    Range("J21:L21").Select
    Selection.ClearContents
    Range("L24").Select
Sheets("DATOS").OptionButton1.Value = True
Application.ScreenUpdating = True
End Sub

