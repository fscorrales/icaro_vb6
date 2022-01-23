Attribute VB_Name = "Busqueda"
Public Function BuscarTexto(Texto As String, Busqueda As String) As Boolean

    Dim i As Integer
    
    i = InStr(1, Texto, Busqueda, vbTextCompare)
    
    If i > 0 Then
        BuscarTexto = True
    Else
        BuscarTexto = False
    End If

End Function

Public Function SeekNoMatch(ValorBuscado As String, NombreTabla As String) As Boolean
    
    SeekNoMatch = False
    Set rstBuscarIcaro = New ADODB.Recordset
    rstBuscarIcaro.Open NombreTabla, dbIcaro, adOpenDynamic, adLockReadOnly
    With rstBuscarIcaro
        .Index = rstBuscarIcaro.Fields.Item(0).Name
        .Seek "=", ValorBuscado
    End With
    If rstBuscarIcaro.EOF = True Then
        SeekNoMatch = True
    End If
    rstBuscarIcaro.Close
    Set rstBuscarIcaro = Nothing
    
End Function

Public Function FindNoMatch(ValorBuscado As String, SQL As String, CampoBusqueda As String) As Boolean
    
    FindNoMatch = False
    Set rstBuscarIcaro = New ADODB.Recordset
    rstBuscarIcaro.Open SQL, dbIcaro, adOpenDynamic, adLockOptimistic
    rstBuscarIcaro.Find CampoBusqueda & "='" & ValorBuscado & "'"
    If rstBuscarIcaro.EOF = True Then
        FindNoMatch = True
    End If
    rstBuscarIcaro.Close
    Set rstBuscarIcaro = Nothing
    
End Function

Public Function SQLNoMatch(SQL As String) As Boolean
    
    SQLNoMatch = False
    Set rstBuscarIcaro = New ADODB.Recordset
    rstBuscarIcaro.Open SQL, dbIcaro, adOpenForwardOnly, adLockOptimistic
    If (rstBuscarIcaro.EOF And rstBuscarIcaro.BOF) Then
        SQLNoMatch = True
    End If
    rstBuscarIcaro.Close
    Set rstBuscarIcaro = Nothing
    
End Function

