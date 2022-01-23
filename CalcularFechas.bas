Attribute VB_Name = "CalcularFechas"
Public Function CalcularAños() As String
    'Call SetRecordset(rstFecha, "Select Comprobante,Fecha From Carga where right(Comprobante,2) = " & Right(strEjercicio, 2) & " Order by Fecha Desc")
    If rstFecha.BOF = False Then
        With rstFecha
            .MoveFirst
            intMeses = Format(.Fields("Fecha"), "mm")
        End With
    Else
        intMeses = 12
    End If
    Set rstFecha = Nothing
End Function
