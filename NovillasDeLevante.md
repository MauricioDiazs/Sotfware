Private Sub Button_Eliminar_Click()
' Elimina filas segun un criterio
Dim i As Integer, nfilas As Integer
nfilas = ActiveSheet.Cells(1, 2).CurrentRegion.Rows.Count
' Obtiene el criterio de eliminación
qCrit = InputBox("Ingrese el criterio de eliminación")
' Obtiene columna para criterio de eliminación
Eliminacion_Formulario.Show
auxVerificar = Eliminacion_Formulario.Combo_Columna.Value
bandera = True
bandera2 = True
Select Case auxVerificar
    Case "ID"
        qCol = "B"
        qCrit = CDbl(qCrit)
    Case "Foto"
        qCol = "C"
    Case "Nombre"
        qCol = "D"
    Case "Raza"
        qCol = "E"
    Case "Fecha de nacimiento"
        qCol = "F"
        qCrit = CDate(qCrit)
        qCrit = Format(qCrit, "mm/dd/yyyy")
        bandera = False
    Case "Madre"
        qCol = "G"
    Case "Padre"
        qCol = "H"
    Case "Peso"
        qCol = "I"
    Case "Promedio de leche"
        qCol = "J"
        qCrit = CDbl(qCrit)
    Case "Pezones"
        qCol = "K"
    Case "# de crias"
        qCol = "L"
        qCrit = CDbl(qCrit)
    Case Else
        MsgBox ("La columna ingresada no es valida.")
        bandera2 = False
End Select
' Evalua la columna con el criterio de eliminación
contador = 0
For i = nfilas And bandera2 To 2 Step -1
    Cells(i, qCol).Select
        If Cells(i, qCol) = qCrit And bandera Then
            ActiveCell.EntireRow.Select
            Selection.Delete
            contador = contador + 1
        Else
        If Not bandera Then
            If CDate(Cells(i, qCol)) = qCrit Then
                ActiveCell.EntireRow.Select
                Selection.Delete
                contador = contador + 1
            End If
        End If
        End If
Next i
' Verifica si hubo coincidencias
If contador > 0 Then
MsgBox ("Eliminación exitosa.")
Else
If bandera2 Then
MsgBox ("El criterio de eliminación no coincide con ninguna fila.")
End If
End If
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
' Carga la imagen para visualizarse en el formulario
If Target.Row = 1 Then
Foto_Presente.Visible = False
Else
If Not Len(ActiveWorkbook.Path & "\Recursos\" & ActiveSheet.Cells(Target.Row, 3)) = Len(ActiveWorkbook.Path & "\Recursos\") Then
Foto_Presente.Visible = True
Foto_Presente.Picture = LoadPicture(ActiveWorkbook.Path & "\Recursos\" & ActiveSheet.Cells(Target.Row, 3))
Else
Foto_Presente.Visible = False
End If
End If
End Sub
