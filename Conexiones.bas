Attribute VB_Name = "Conexiones"
Public dbIcaro As ADODB.Connection
Public rstListadoIcaro As ADODB.Recordset
Public rstBuscarIcaro As ADODB.Recordset
Public rstRegistroIcaro As ADODB.Recordset

Public dbAFIP As ADODB.Connection
Public rstRegistroAFIP As ADODB.Recordset
Public rstRegistroICARO2 As ADODB.Recordset

Public Sub Conectar()
    
    Set dbIcaro = New ADODB.Connection
    dbIcaro.CursorLocation = adUseClient
    dbIcaro.Mode = adModeUnknown
    dbIcaro.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Icaro.mdb"
    dbIcaro.Open

End Sub

Public Sub Desconectar()
    
    dbIcaro.Close
    Set dbIcaro = Nothing

End Sub

Public Sub DesconectarRecordset()

    Set rstRegistroIcaro = Nothing
    Set rstBuscarIcaro = Nothing
    Set rstListadoIcaro = Nothing
    
End Sub
