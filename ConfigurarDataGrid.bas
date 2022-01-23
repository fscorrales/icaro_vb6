Attribute VB_Name = "ConfigurarDataGrid"
Public Sub ConfigurardgListadoGasto()

    With ListadoImputacion.dgListado
        .Clear
        .Cols = 10
        .Rows = 2
        .TextMatrix(0, 0) = "Comprobante"
        .TextMatrix(0, 1) = "Fecha"
        .TextMatrix(0, 2) = "Fuente"
        .TextMatrix(0, 3) = "CUIT"
        .TextMatrix(0, 4) = "Obra"
        .TextMatrix(0, 5) = "Certificado"
        .TextMatrix(0, 6) = "Importe"
        .TextMatrix(0, 7) = "Fondo Reparo"
        .TextMatrix(0, 8) = "Cuenta"
        .TextMatrix(0, 9) = "Avance"
        .ColWidth(0) = 1100
        .ColWidth(1) = 1000
        .ColWidth(2) = 1
        .ColWidth(3) = 1
        .ColWidth(4) = 4000
        .ColWidth(5) = 1
        .ColWidth(6) = 1000
        .ColWidth(7) = 1
        .ColWidth(8) = 1
        .ColWidth(9) = 1
        .FixedCols = 0
        .FixedRows = 1
        .FocusRect = flexFocusHeavy
        .HighLight = flexHighlightWithFocus
        .SelectionMode = flexSelectionByRow
        .ColAlignment(0) = 1
        .ColAlignment(1) = 7
        .ColAlignment(4) = 1
        .ColAlignment(6) = 7
    End With
    
End Sub

Public Sub CargardgListadoGasto(Optional ByVal Comprobante As String = "", _
Optional ByVal Fecha As String = "", Optional ByVal BuscarComprobante As String = 0, _
Optional ByVal Obra As String = "", Optional ByVal CUITContratista As String = "")

    Dim i As String
    Dim FilaBuscada As String
    Dim SQL As String
    
    FilaBuscada = 0
    i = 0
    ListadoImputacion.dgListado.Rows = 2
    Set rstListadoIcaro = New ADODB.Recordset
    If Comprobante = "" And Fecha = "" And Obra = "" And CUITContratista = "" Then
        SQL = ""
    ElseIf Fecha <> "" Then
        Select Case Len(Fecha)
        Case Is = 0
            SQL = ""
        Case Is = 4
            SQL = "Where Year(Fecha) ='" & Fecha & "'"
        Case Is = 7
            SQL = "Where Year(Fecha) ='" & Right(Fecha, 4) _
            & "' And Format(Month(Fecha),'00') ='" & Left(Fecha, 2) & "'"
        Case Is = 10
            SQL = "Where Year(Fecha) ='" & Right(Fecha, 4) _
            & "' And Format(Month(Fecha),'00') ='" & Mid(Fecha, 4, 2) _
            & "' And Format(Day(Fecha),'00') ='" & Left(Fecha, 2) & "'"
        End Select
    ElseIf Comprobante <> "" Then
        SQL = "Where Comprobante Like '%" & Comprobante & "%'"
    ElseIf Obra <> "" Then
        SQL = "Where Obra = '" & Obra & "'"
    ElseIf CUITContratista <> "" Then
        SQL = "Where CUIT Like '%" & CUITContratista & "%'"
    End If
    SQL = "SELECT * FROM CARGA " & SQL & " ORDER BY Fecha DESC,Comprobante DESC"
    rstListadoIcaro.Open SQL, dbIcaro, adOpenDynamic, adLockOptimistic
    With ListadoImputacion.dgListado
        If rstListadoIcaro.BOF = False Then
            rstListadoIcaro.MoveFirst
            While rstListadoIcaro.EOF = False
                i = i + 1
                .RowHeight(i) = 300
                If BuscarComprobante <> "0" Then
                    If BuscarComprobante = rstListadoIcaro!Comprobante Then
                        FilaBuscada = i
                    End If
                End If
                .TextMatrix(i, 0) = rstListadoIcaro!Comprobante
                .TextMatrix(i, 1) = rstListadoIcaro!Fecha
                .TextMatrix(i, 2) = rstListadoIcaro!Fuente
                .TextMatrix(i, 3) = rstListadoIcaro!CUIT
                .TextMatrix(i, 4) = rstListadoIcaro!Obra
                If Len(rstListadoIcaro!Certificado) <> "0" Then
                    .TextMatrix(i, 5) = rstListadoIcaro!Certificado
                Else
                    .TextMatrix(i, 5) = "0"
                End If
                .TextMatrix(i, 6) = FormatNumber(rstListadoIcaro!Importe, 2)
                .TextMatrix(i, 7) = FormatNumber(rstListadoIcaro!FondoDeReparo, 2)
                .TextMatrix(i, 8) = rstListadoIcaro!Cuenta
                .TextMatrix(i, 9) = rstListadoIcaro!Avance
                rstListadoIcaro.MoveNext
                .Rows = .Rows + 1
            Wend
        End If
        .Rows = .Rows - 1
        If FilaBuscada <> "0" Then
            .TopRow = FilaBuscada
            .Row = FilaBuscada
        End If
    End With
    rstListadoIcaro.Close
    Set rstListadoIcaro = Nothing
    
End Sub

Public Sub ConfigurardgImputacion()

    With ListadoImputacion.dgImputacion
        .Clear
        .Cols = 3
        .Rows = 2
        .TextMatrix(0, 0) = "Imputacion"
        .TextMatrix(0, 1) = "Partida"
        .TextMatrix(0, 2) = "Importe"
        .ColWidth(0) = 1100
        .ColWidth(1) = 1000
        .ColWidth(2) = 1000
        .FixedCols = 0
        .FixedRows = 1
        .FocusRect = flexFocusHeavy
        .HighLight = flexHighlightWithFocus
        .SelectionMode = flexSelectionByRow
        .ColAlignment(0) = 1
        .ColAlignment(1) = 1
        .ColAlignment(2) = 7
    End With
    
End Sub

Public Sub CargardgImputacion(Numero As String)

    Dim i As Integer
    Dim SQL As String
    Dim SumaTotal As Double
    
    i = 0
    ListadoImputacion.dgImputacion.Rows = 2
    Set rstListadoIcaro = New ADODB.Recordset
    SQL = "Select Imputacion,Partida,Importe FROM Imputacion Where Comprobante = " & "'" & Numero & "'"
    rstListadoIcaro.Open SQL, dbIcaro, adOpenDynamic, adLockOptimistic
    With ListadoImputacion.dgImputacion
        If rstListadoIcaro.EOF = False Then
            rstListadoIcaro.MoveFirst
            While rstListadoIcaro.EOF = False
                i = i + 1
                .RowHeight(i) = 300
                .TextMatrix(i, 0) = rstListadoIcaro!Imputacion
                .TextMatrix(i, 1) = rstListadoIcaro!Partida
                .TextMatrix(i, 2) = FormatNumber(rstListadoIcaro!Importe, 2)
                SumaTotal = SumaTotal + rstListadoIcaro!Importe
                rstListadoIcaro.MoveNext
                .Rows = .Rows + 1
            Wend
            ListadoImputacion.txtTotalImputacion.Text = FormatNumber(SumaTotal, 2)
        End If
        .Rows = .Rows - 1
    End With
    rstListadoIcaro.Close
    Set rstListadoIcaro = Nothing
    SumaTotal = 0
    
End Sub

Public Sub ConfigurardgRetencion()

    With ListadoImputacion.dgRetencion
        .Clear
        .Cols = 2
        .Rows = 2
        .TextMatrix(0, 0) = "Código"
        .TextMatrix(0, 1) = "Importe"
        .ColWidth(0) = 2100
        .ColWidth(1) = 1000
        .FixedCols = 0
        .FixedRows = 1
        .FocusRect = flexFocusHeavy
        .HighLight = flexHighlightWithFocus
        .SelectionMode = flexSelectionByRow
        .ColAlignment(0) = 1
        .ColAlignment(1) = 7
    End With
    
End Sub

Public Sub CargardgRetencion(Numero As String)

    Dim i As Integer
    Dim SQL As String
    Dim SumaTotal As Double
    
    i = 0
    ListadoImputacion.dgRetencion.Rows = 2
    Set rstListadoIcaro = New ADODB.Recordset
    SQL = "Select Codigo,Importe FROM Retenciones Where Comprobante = " & "'" & Numero & "'"
    rstListadoIcaro.Open SQL, dbIcaro, adOpenDynamic, adLockOptimistic

    With ListadoImputacion.dgRetencion
        If rstListadoIcaro.BOF = False Then
            rstListadoIcaro.MoveFirst
            While rstListadoIcaro.EOF = False
                i = i + 1
                .RowHeight(i) = 300
                .TextMatrix(i, 0) = rstListadoIcaro!Codigo
                .TextMatrix(i, 1) = FormatNumber(rstListadoIcaro!Importe, 2)
                SumaTotal = SumaTotal + rstListadoIcaro!Importe
                rstListadoIcaro.MoveNext
                .Rows = .Rows + 1
            Wend
        ListadoImputacion.txtTotalRetencion.Text = FormatNumber(SumaTotal, 2)
        End If
        .Rows = .Rows - 1
    End With
    rstListadoIcaro.Close
    Set rstListadoIcaro = Nothing
    SumaTotal = 0
    
End Sub

Public Sub ConfigurardgProveedores()

    With ListadoProveedores.dgProveedores
        .Clear
        .Cols = 2
        .Rows = 2
        .TextMatrix(0, 0) = "CUIT"
        .TextMatrix(0, 1) = "Descripción"
        .ColWidth(0) = 1200
        .ColWidth(1) = 5300
        .FixedCols = 0
        .FocusRect = flexFocusHeavy
        .HighLight = flexHighlightWithFocus
        .SelectionMode = flexSelectionByRow
        .AllowUserResizing = flexResizeColumns
        .ColAlignment(0) = 1
        .ColAlignment(1) = 1
    End With
    
End Sub

Public Sub CargardgProveedores(Optional ByVal CUIT As String = "", _
Optional ByVal Nombre As String = "")
    
    Dim i As Integer
    Dim SQL As String
    i = 0
    
    ListadoProveedores.dgProveedores.Rows = 2
    Set rstListadoIcaro = New ADODB.Recordset
    If CUIT = "" And Nombre = "" Then
        SQL = ""
    ElseIf Nombre <> "" Then
        SQL = "Where Nombre Like '%" & Nombre & "%'"
    ElseIf CUIT <> "" Then
        SQL = "Where CUIT Like '%" & CUIT & "%'"
    End If
    SQL = "SELECT * FROM CUIT " & SQL & " ORDER BY Nombre"
    rstListadoIcaro.Open SQL, dbIcaro, adOpenDynamic, adLockOptimistic
    With ListadoProveedores.dgProveedores
        If rstListadoIcaro.BOF = False Then
            rstListadoIcaro.MoveFirst
            While rstListadoIcaro.EOF = False
                i = i + 1
                .RowHeight(i) = 300
                .TextMatrix(i, 0) = rstListadoIcaro!CUIT
                .TextMatrix(i, 1) = rstListadoIcaro!Nombre
                rstListadoIcaro.MoveNext
                .Rows = .Rows + 1
            Wend
        End If
        .Rows = .Rows - 1
    End With
    rstListadoIcaro.Close
    Set rstListadoIcaro = Nothing
    
End Sub

Public Sub ConfigurardgCuentasBancarias()
    
    With ListadoCuentasBancarias.dgCuentasBancarias
        .Clear
        .Cols = 2
        .Rows = 2
        .TextMatrix(0, 0) = "Cuenta"
        .TextMatrix(0, 1) = "Descripción"
        .ColWidth(0) = 1200
        .ColWidth(1) = 3400
        .FixedCols = 0
        .FocusRect = flexFocusHeavy
        .HighLight = flexHighlightWithFocus
        .SelectionMode = flexSelectionByRow
        .AllowUserResizing = flexResizeColumns
        .ColAlignment(0) = 4
        .ColAlignment(1) = 1
    End With
    
End Sub

Public Sub CargardgCuentasBancarias()
    
    Dim i As Integer
    Dim SQL As String
    i = 0
    
    ListadoCuentasBancarias.dgCuentasBancarias.Rows = 2
    Set rstListadoIcaro = New ADODB.Recordset
    SQL = "SELECT * FROM CUENTAS ORDER BY Cuenta"
    rstListadoIcaro.Open SQL, dbIcaro, adOpenDynamic, adLockOptimistic
    With ListadoCuentasBancarias.dgCuentasBancarias
        If rstListadoIcaro.BOF = False Then
            rstListadoIcaro.MoveFirst
            While rstListadoIcaro.EOF = False
                i = i + 1
                .RowHeight(i) = 300
                .TextMatrix(i, 0) = rstListadoIcaro!Cuenta
                .TextMatrix(i, 1) = rstListadoIcaro!Descripcion
                rstListadoIcaro.MoveNext
                .Rows = .Rows + 1
            Wend
        End If
        .Rows = .Rows - 1
    End With
    rstListadoIcaro.Close
    Set rstListadoIcaro = Nothing
    
End Sub

Public Sub ConfigurardgPartidas()
    
    With ListadoPartidas.dgPartidas
        .Clear
        .Cols = 2
        .Rows = 2
        .TextMatrix(0, 0) = "Partida"
        .TextMatrix(0, 1) = "Descripción"
        .ColWidth(0) = 600
        .ColWidth(1) = 4000
        .FixedCols = 0
        .FocusRect = flexFocusHeavy
        .HighLight = flexHighlightWithFocus
        .SelectionMode = flexSelectionByRow
        .AllowUserResizing = flexResizeColumns
        .ColAlignment(0) = 4
        .ColAlignment(1) = 1
    End With
    
End Sub

Public Sub CargardgPartidas()
    
    Dim i As Integer
    Dim SQL As String
    i = 0
    
    ListadoPartidas.dgPartidas.Rows = 2
    Set rstListadoIcaro = New ADODB.Recordset
    SQL = "Select * From PARTIDAS Order By PARTIDA"
    rstListadoIcaro.Open SQL, dbIcaro, adOpenDynamic, adLockOptimistic
    With ListadoPartidas.dgPartidas
        If rstListadoIcaro.BOF = False Then
            rstListadoIcaro.MoveFirst
            While rstListadoIcaro.EOF = False
                i = i + 1
                .RowHeight(i) = 300
                .TextMatrix(i, 0) = rstListadoIcaro!Partida
                .TextMatrix(i, 1) = rstListadoIcaro!Descripcion
                rstListadoIcaro.MoveNext
                .Rows = .Rows + 1
            Wend
        End If
        .Rows = .Rows - 1
    End With
    rstListadoIcaro.Close
    Set rstListadoIcaro = Nothing
    
End Sub

Public Sub ConfigurardgFuentes()
    
    With ListadoFuentes.dgFuentes
        .Clear
        .Cols = 2
        .Rows = 2
        .TextMatrix(0, 0) = "Fuente"
        .TextMatrix(0, 1) = "Descripción"
        .ColWidth(0) = 600
        .ColWidth(1) = 4000
        .FixedCols = 0
        .FocusRect = flexFocusHeavy
        .HighLight = flexHighlightWithFocus
        .SelectionMode = flexSelectionByRow
        .AllowUserResizing = flexResizeColumns
        .ColAlignment(0) = 4
        .ColAlignment(1) = 1
    End With
    
End Sub

Public Sub CargardgFuentes()
    
    Dim i As Integer
    Dim SQL As String
    i = 0
    
    ListadoFuentes.dgFuentes.Rows = 2
    Set rstListadoIcaro = New ADODB.Recordset
    SQL = "Select * From FUENTES Order By FUENTE"
    rstListadoIcaro.Open SQL, dbIcaro, adOpenDynamic, adLockOptimistic
    With ListadoFuentes.dgFuentes
        If rstListadoIcaro.BOF = False Then
            rstListadoIcaro.MoveFirst
            While rstListadoIcaro.EOF = False
                i = i + 1
                .RowHeight(i) = 300
                .TextMatrix(i, 0) = rstListadoIcaro!Fuente
                .TextMatrix(i, 1) = rstListadoIcaro!Descripcion
                rstListadoIcaro.MoveNext
                .Rows = .Rows + 1
            Wend
        End If
        .Rows = .Rows - 1
    End With
    rstListadoIcaro.Close
    Set rstListadoIcaro = Nothing
    
End Sub

Public Sub ConfigurardgCodigoRetenciones()
    
    With ListadoRetenciones.dgRetenciones
        .Clear
        .Cols = 2
        .Rows = 2
        .TextMatrix(0, 0) = "Código"
        .TextMatrix(0, 1) = "Descripción"
        .ColWidth(0) = 600
        .ColWidth(1) = 4000
        .FixedCols = 0
        .FocusRect = flexFocusHeavy
        .HighLight = flexHighlightWithFocus
        .SelectionMode = flexSelectionByRow
        .AllowUserResizing = flexResizeColumns
        .ColAlignment(0) = 4
        .ColAlignment(1) = 1
    End With
    
End Sub

Public Sub CargardgCodigoRetenciones()
    
    Dim i As Integer
    Dim SQL As String
    i = 0
    
    ListadoRetenciones.dgRetenciones.Rows = 2
    Set rstListadoIcaro = New ADODB.Recordset
    SQL = "Select * From CODIGOSRETENCIONES Order By CODIGO"
    rstListadoIcaro.Open SQL, dbIcaro, adOpenDynamic, adLockOptimistic
    With ListadoRetenciones.dgRetenciones
        If rstListadoIcaro.BOF = False Then
            rstListadoIcaro.MoveFirst
            While rstListadoIcaro.EOF = False
                i = i + 1
                .RowHeight(i) = 300
                .TextMatrix(i, 0) = rstListadoIcaro!Codigo
                .TextMatrix(i, 1) = rstListadoIcaro!Descripcion
                rstListadoIcaro.MoveNext
                .Rows = .Rows + 1
            Wend
        End If
        .Rows = .Rows - 1
    End With
    rstListadoIcaro.Close
    Set rstListadoIcaro = Nothing
    
End Sub

Public Sub ConfigurardgLocalidades()
    
    With ListadoLocalidades.dgLocalidades
        .Clear
        .Cols = 1
        .Rows = 2
        .TextMatrix(0, 0) = "Localidad"
        .ColWidth(0) = 2000
        .FixedCols = 0
        .FocusRect = flexFocusHeavy
        .HighLight = flexHighlightWithFocus
        .SelectionMode = flexSelectionByRow
        .AllowUserResizing = flexResizeColumns
        .ColAlignment(0) = 1
    End With
    
End Sub

Public Sub CargardgLocalidades()
    
    Dim i As Integer
    Dim SQL As String
    i = 0
    
    ListadoLocalidades.dgLocalidades.Rows = 2
    Set rstListadoIcaro = New ADODB.Recordset
    SQL = "Select * From LOCALIDADES Order By LOCALIDAD"
    rstListadoIcaro.Open SQL, dbIcaro, adOpenDynamic, adLockOptimistic
    With ListadoLocalidades.dgLocalidades
        If rstListadoIcaro.BOF = False Then
            rstListadoIcaro.MoveFirst
            While rstListadoIcaro.EOF = False
                i = i + 1
                .RowHeight(i) = 300
                .TextMatrix(i, 0) = rstListadoIcaro!Localidad
                rstListadoIcaro.MoveNext
                .Rows = .Rows + 1
            Wend
        End If
        .Rows = .Rows - 1
    End With
    rstListadoIcaro.Close
    Set rstListadoIcaro = Nothing
    
End Sub

Public Sub ConfigurardgObras()
    
    With ListadoObras.dgObras
        .Clear
        .Cols = 2
        .Rows = 2
        .TextMatrix(0, 0) = "Descripción"
        .TextMatrix(0, 1) = "Imputación"
        .ColWidth(0) = 6500
        .ColWidth(1) = 1500
        .FixedCols = 0
        .FocusRect = flexFocusHeavy
        .HighLight = flexHighlightWithFocus
        .SelectionMode = flexSelectionByRow
        .AllowUserResizing = flexResizeColumns
        .ColAlignment(0) = 1
        .ColAlignment(1) = 1
    End With
    
End Sub

Public Sub CargardgObras(Optional ByVal DescripcionObra As String = "", _
Optional ByVal CUITContratista As String = "")
    
    Dim i As Integer
    Dim SQL As String
    i = 0
    
    ListadoObras.dgObras.Rows = 2
    Set rstListadoIcaro = New ADODB.Recordset
    If DescripcionObra = "" And CUITContratista = "" Then
        SQL = ""
    ElseIf DescripcionObra <> "" Then
        SQL = "Where Descripcion Like '%" & DescripcionObra & "%'"
    ElseIf CUITContratista <> "" Then
        SQL = "Where CUIT Like '%" & CUITContratista & "%'"
    End If
    SQL = "SELECT * FROM OBRAS " & SQL & " ORDER BY Descripcion"
    rstListadoIcaro.Open SQL, dbIcaro, adOpenDynamic, adLockOptimistic
    With ListadoObras.dgObras
        If rstListadoIcaro.BOF = False Then
            rstListadoIcaro.MoveFirst
            While rstListadoIcaro.EOF = False
                i = i + 1
                .RowHeight(i) = 300
                .TextMatrix(i, 0) = rstListadoIcaro!Descripcion
                .TextMatrix(i, 1) = rstListadoIcaro!Imputacion & "-" & rstListadoIcaro!Partida
                rstListadoIcaro.MoveNext
                .Rows = .Rows + 1
            Wend
        End If
        .Rows = .Rows - 1
    End With
    rstListadoIcaro.Close
    Set rstListadoIcaro = Nothing
    
End Sub

Public Sub ConfigurardgAutocargaCertificados()
    
    With AutocargaCertificados.dgAutocargaCertificado
        .Clear
        .Cols = 12
        .Rows = 2
        .TextMatrix(0, 0) = "Proveedor"
        .TextMatrix(0, 1) = "Obra"
        .TextMatrix(0, 2) = "Cert"
        .TextMatrix(0, 3) = "Monto Bruto"
        .TextMatrix(0, 4) = "Retenciones"
        .TextMatrix(0, 5) = "Período"
        .ColWidth(0) = 3000
        .ColWidth(1) = 4000
        .ColWidth(2) = 500
        .ColWidth(3) = 1200
        .ColWidth(4) = 1200
        .ColWidth(5) = 700
        .ColWidth(6) = 1
        .ColWidth(7) = 1
        .ColWidth(8) = 1
        .ColWidth(9) = 1
        .ColWidth(10) = 1
        .ColWidth(11) = 1
        .FixedCols = 0
        .FocusRect = flexFocusHeavy
        .HighLight = flexHighlightWithFocus
        .SelectionMode = flexSelectionByRow
        .AllowUserResizing = flexResizeColumns
        .ColAlignment(0) = 1
        .ColAlignment(1) = 1
        .ColAlignment(2) = 7
        .ColAlignment(3) = 7
        .ColAlignment(4) = 7
        .ColAlignment(5) = 7
    End With
    
End Sub

Public Sub CargardgAutocargaCertificados()
    
    Dim i As Integer
    Dim SQL As String
    i = 0
    
    AutocargaCertificados.dgAutocargaCertificado.Rows = 2
    Set rstListadoIcaro = New ADODB.Recordset
    SQL = "Select * From CERTIFICADOS Where Comprobante = """" Order by Periodo Desc, Proveedor Asc"
    rstListadoIcaro.Open SQL, dbIcaro, adOpenDynamic, adLockOptimistic
    With AutocargaCertificados.dgAutocargaCertificado
        If rstListadoIcaro.BOF = False Then
            rstListadoIcaro.MoveFirst
            While rstListadoIcaro.EOF = False
                i = i + 1
                .RowHeight(i) = 300
                .TextMatrix(i, 0) = rstListadoIcaro!Proveedor
                .TextMatrix(i, 1) = rstListadoIcaro!Obra
                .TextMatrix(i, 2) = rstListadoIcaro!Certificado
                .TextMatrix(i, 3) = rstListadoIcaro!MontoBruto
                .TextMatrix(i, 4) = rstListadoIcaro!INVICO + rstListadoIcaro!Ganancias _
                + rstListadoIcaro!SUSS + rstListadoIcaro!LP + rstListadoIcaro!IB
                .TextMatrix(i, 5) = Format(rstListadoIcaro!Periodo, "mmm" & "/" & "yy")
                .TextMatrix(i, 6) = rstListadoIcaro!FondoDeReparo
                .TextMatrix(i, 7) = rstListadoIcaro!IB
                .TextMatrix(i, 8) = rstListadoIcaro!LP
                .TextMatrix(i, 9) = rstListadoIcaro!SUSS
                .TextMatrix(i, 10) = rstListadoIcaro!Ganancias
                .TextMatrix(i, 11) = rstListadoIcaro!INVICO
                rstListadoIcaro.MoveNext
                .Rows = .Rows + 1
            Wend
        End If
        .Rows = .Rows - 1
    End With
    
    rstListadoIcaro.Close
    Set rstListadoIcaro = Nothing
        
    With AutocargaCertificados
        i = .dgAutocargaCertificado.Row
        If i = 0 Then
        'If Len(.dgAutocargaCertificado.TextMatrix(1, 0)) = "0" Then
            .cmdAgregarCertificado.Enabled = False
            .cmdEliminarCertificado.Enabled = False
        Else
            .cmdAgregarCertificado.Enabled = True
            .cmdEliminarCertificado.Enabled = True
        'End If
        End If
    End With
    
End Sub

Public Sub ConfigurardgAutocargaEPAM()
    
    With AutocargaEPAM.dgAutocargaEPAM
        .Clear
        .Cols = 10
        .Rows = 2
        .TextMatrix(0, 0) = "Código"
        .TextMatrix(0, 1) = "Monto Bruto"
        .TextMatrix(0, 2) = "Retenciones"
        .TextMatrix(0, 3) = "Inicio"
        .TextMatrix(0, 4) = "Cierre"
        .ColWidth(0) = 4000
        .ColWidth(1) = 1200
        .ColWidth(2) = 1200
        .ColWidth(3) = 1000
        .ColWidth(4) = 1000
        .ColWidth(5) = 1
        .ColWidth(6) = 1
        .ColWidth(7) = 1
        .ColWidth(8) = 1
        .ColWidth(9) = 1
        .FixedCols = 0
        .FocusRect = flexFocusHeavy
        .HighLight = flexHighlightWithFocus
        .SelectionMode = flexSelectionByRow
        .AllowUserResizing = flexResizeColumns
        .ColAlignment(0) = 1
        .ColAlignment(1) = 7
        .ColAlignment(2) = 7
        .ColAlignment(3) = 7
        .ColAlignment(4) = 7
    End With
    
End Sub

Public Sub CargardgAutocargaEPAM()
    
    Dim i As Integer
    Dim SQL As String
    i = 0
    
    AutocargaEPAM.dgAutocargaEPAM.Rows = 2
    Set rstListadoIcaro = New ADODB.Recordset
    SQL = "Select * From EPAM Where Comprobante = """"  Order by Inicio Desc, Codigo Asc"
    rstListadoIcaro.Open SQL, dbIcaro, adOpenDynamic, adLockOptimistic
    With AutocargaEPAM.dgAutocargaEPAM
        If rstListadoIcaro.BOF = False Then
            rstListadoIcaro.MoveFirst
            While rstListadoIcaro.EOF = False
                i = i + 1
                .RowHeight(i) = 300
                .TextMatrix(i, 0) = rstListadoIcaro!Codigo
                .TextMatrix(i, 1) = rstListadoIcaro!MontoBruto
                .TextMatrix(i, 2) = FormatNumber(rstListadoIcaro!Ganancias + rstListadoIcaro!Sellos _
                + rstListadoIcaro!IB + rstListadoIcaro!LP + rstListadoIcaro!SUSS, 2)
                .TextMatrix(i, 3) = rstListadoIcaro!Inicio
                .TextMatrix(i, 4) = rstListadoIcaro!Cierre
                .TextMatrix(i, 5) = rstListadoIcaro!Ganancias
                .TextMatrix(i, 6) = rstListadoIcaro!Sellos
                .TextMatrix(i, 7) = rstListadoIcaro!IB
                .TextMatrix(i, 8) = rstListadoIcaro!LP
                .TextMatrix(i, 9) = rstListadoIcaro!SUSS
                rstListadoIcaro.MoveNext
                .Rows = .Rows + 1
            Wend
        End If
        .Rows = .Rows - 1
    End With
    
    rstListadoIcaro.Close
    Set rstListadoIcaro = Nothing
    
    With AutocargaEPAM
        If .dgAutocargaEPAM.Rows < 2 Then
            .cmdAgregarEPAM.Enabled = False
            .cmdEliminarEPAM.Enabled = False
            .cmdPagarEPAM.Enabled = False
        Else
            .cmdAgregarEPAM.Enabled = True
            .cmdEliminarEPAM.Enabled = True
            .cmdPagarEPAM.Enabled = True
        End If
    End With
    
End Sub

Public Sub ConfigurardgEPAMConsolidado()
    
    With EPAMConsolidado.dgEPAMConsolidado
        .Clear
        .Cols = 5
        .Rows = 2
        .TextMatrix(0, 0) = "Código"
        .TextMatrix(0, 1) = "Imputación Presupuestaria"
        .TextMatrix(0, 2) = "T. Pagado"
        .TextMatrix(0, 3) = "T. Imputado"
        .TextMatrix(0, 4) = "Importe a Certificar"
        .ColWidth(0) = 4000
        .ColWidth(1) = 2100
        .ColWidth(2) = 1100
        .ColWidth(3) = 1100
        .ColWidth(4) = 1700
        .FixedCols = 0
        .FocusRect = flexFocusHeavy
        .HighLight = flexHighlightWithFocus
        .SelectionMode = flexSelectionByRow
        .AllowUserResizing = flexResizeColumns
        .ColAlignment(0) = 1
        .ColAlignment(1) = 4
        .ColAlignment(2) = 7
        .ColAlignment(3) = 7
        .ColAlignment(4) = 7
    End With
    
End Sub

Public Sub CargardgEPAMConsolidado()
    
    Dim i As Integer
    Dim SQL As String
    i = 0
    
    EPAMConsolidado.dgEPAMConsolidado.Rows = 2
    Set rstListadoIcaro = New ADODB.Recordset
    SQL = "Select * From DATOSFIJOSEPAM Order by Codigo"
    rstListadoIcaro.Open SQL, dbIcaro, adOpenDynamic, adLockOptimistic
    With EPAMConsolidado.dgEPAMConsolidado
        If rstListadoIcaro.BOF = False Then
            rstListadoIcaro.MoveFirst
            Set rstBuscarIcaro = New ADODB.Recordset
            While rstListadoIcaro.EOF = False
                i = i + 1
                .RowHeight(i) = 300
                .TextMatrix(i, 0) = rstListadoIcaro!Codigo
                SQL = "Select OBRAS.Imputacion, OBRAS.Partida From OBRAS Where OBRAS.Descripcion = '" & .TextMatrix(i, 0) & "'"
                rstBuscarIcaro.Open SQL, dbIcaro, adOpenDynamic, adLockOptimistic
                If rstBuscarIcaro.BOF = False Then
                    rstBuscarIcaro.MoveFirst
                    .TextMatrix(i, 1) = rstBuscarIcaro!Imputacion & "-" & rstBuscarIcaro!Partida
                Else
                    .TextMatrix(i, 1) = "Error en BASE DATOS"
                End If
                rstBuscarIcaro.Close
                If Len(rstListadoIcaro!Pagado) <> "0" Then
                    .TextMatrix(i, 2) = rstListadoIcaro!Pagado
                Else
                    .TextMatrix(i, 2) = "0"
                End If
                SQL = "Select Sum(EPAM.MontoBruto) As Total From EPAM " _
                & "Where EPAM.Comprobante <> """" And EPAM.Codigo = '" & .TextMatrix(i, 0) & "'"
                rstBuscarIcaro.Open SQL, dbIcaro, adOpenDynamic, adLockOptimistic
                If rstBuscarIcaro.BOF = False Then
                    rstBuscarIcaro.MoveFirst
                    If Len(rstBuscarIcaro.Fields("Total")) <> "0" Then
                        .TextMatrix(i, 2) = .TextMatrix(i, 2) + rstBuscarIcaro.Fields("Total")
                    Else
                        .TextMatrix(i, 2) = .TextMatrix(i, 2)
                    End If
                End If
                rstBuscarIcaro.Close
                If Len(rstListadoIcaro!Imputado) <> "0" Then
                    .TextMatrix(i, 3) = rstListadoIcaro!Imputado
                Else
                    .TextMatrix(i, 3) = De_Num_a_Tx_01(0, , 2)
                End If
                SQL = "Select Sum(Carga.Importe) As Total From Carga " _
                & "Where Carga.Obra = '" & .TextMatrix(i, 0) & "'"
                rstBuscarIcaro.Open SQL, dbIcaro, adOpenDynamic, adLockOptimistic
                If rstBuscarIcaro.BOF = False Then
                    rstBuscarIcaro.MoveFirst
                    If Len(rstBuscarIcaro!Total) <> "0" Then
                        .TextMatrix(i, 3) = .TextMatrix(i, 3) + rstBuscarIcaro!Total
                    Else
                        .TextMatrix(i, 3) = .TextMatrix(i, 3)
                    End If
                End If
                rstBuscarIcaro.Close
                If (De_Txt_a_Num_01(.TextMatrix(i, 2)) - De_Txt_a_Num_01(.TextMatrix(i, 3))) > 0 Then
                    .TextMatrix(i, 4) = De_Txt_a_Num_01(.TextMatrix(i, 2)) - De_Txt_a_Num_01(.TextMatrix(i, 3))
                Else
                    .TextMatrix(i, 4) = De_Num_a_Tx_01(0, , 2)
                End If
                rstListadoIcaro.MoveNext
                .Rows = .Rows + 1
            Wend
        End If
        .Rows = .Rows - 1
    End With
    
    rstListadoIcaro.Close
    Set rstListadoIcaro = Nothing
    Set rstBuscarIcaro = Nothing
    
End Sub

Public Sub ConfigurardgResolucionesEPAM()
    
    With EPAMConsolidado.dgResolucionesEPAM
        .Clear
        .Cols = 8
        .Rows = 2
        .TextMatrix(0, 0) = "Descripción"
        .TextMatrix(0, 1) = "Expediente"
        .TextMatrix(0, 2) = "Resolución"
        .TextMatrix(0, 3) = "Presupuesto"
        .TextMatrix(0, 4) = "Pagado"
        .TextMatrix(0, 5) = "Imputado"
        .TextMatrix(0, 6) = "Decisión a Tomar"
        .TextMatrix(0, 7) = "Límite"
        .ColWidth(0) = 2800
        .ColWidth(1) = 1200
        .ColWidth(2) = 1000
        .ColWidth(3) = 1100
        .ColWidth(4) = 1100
        .ColWidth(5) = 1100
        .ColWidth(6) = 1700
        .ColWidth(7) = 1100
        .FixedCols = 0
        .FocusRect = flexFocusHeavy
        .HighLight = flexHighlightWithFocus
        .SelectionMode = flexSelectionByRow
        .AllowUserResizing = flexResizeColumns
        .ColAlignment(0) = 1
        .ColAlignment(1) = 1
        .ColAlignment(2) = 1
        .ColAlignment(3) = 7
        .ColAlignment(4) = 7
        .ColAlignment(5) = 7
        .ColAlignment(6) = 1
        .ColAlignment(7) = 7
    End With
    
End Sub

Public Sub CargardgResolucionesEPAM(Codigo As String)
    
    Dim i As Integer
    Dim SQL As String
    i = 0
    
    EPAMConsolidado.dgResolucionesEPAM.Rows = 2
    Set rstListadoIcaro = New ADODB.Recordset
    SQL = "Select * From DATOSFIJOSEPAM Where Codigo =" & "'" & Codigo & "'"
    rstListadoIcaro.Open SQL, dbIcaro, adOpenDynamic, adLockOptimistic
    With EPAMConsolidado.dgResolucionesEPAM
        If rstListadoIcaro.BOF = False Then
            rstListadoIcaro.MoveFirst
            Set rstBuscarIcaro = New ADODB.Recordset
            While rstListadoIcaro.EOF = False
                i = i + 1
                .RowHeight(i) = 300
                .TextMatrix(i, 0) = "Presupuesto Inicial"
                If Len(rstListadoIcaro!Expediente) <> "0" Then
                    .TextMatrix(i, 1) = rstListadoIcaro!Expediente
                Else
                    .TextMatrix(i, 1) = "FALTA EXP."
                End If
                If Len(rstListadoIcaro!Resolucion) <> "0" Then
                    .TextMatrix(i, 2) = rstListadoIcaro!Resolucion
                Else
                    .TextMatrix(i, 2) = "FALTA RES."
                End If
                If Len(rstListadoIcaro!Presupuestado) <> 0 Then
                    .TextMatrix(i, 3) = rstListadoIcaro!Presupuestado
                Else
                    .TextMatrix(i, 3) = "0"
                End If
                SQL = "Select EPAM.Codigo, Sum(EPAM.MontoBruto) As Suma From EPAM " _
                & "Inner Join DatosFijosEpam On EPAM.Codigo = DatosFijosEpam.Codigo " _
                & "Where EPAM.Comprobante <> """" and  EPAM.Codigo = '" & Codigo & "' " _
                & "Group by EPAM.Codigo Order by EPAM.Codigo"
                rstBuscarIcaro.Open SQL, dbIcaro, adOpenDynamic, adLockOptimistic
                If rstBuscarIcaro.BOF = False Then
                    .TextMatrix(i, 4) = rstBuscarIcaro!Suma + rstListadoIcaro!Pagado
                Else
                    .TextMatrix(i, 4) = rstListadoIcaro!Pagado
                End If
                rstBuscarIcaro.Close
                SQL = "Select Sum(Importe) as Suma From Carga " _
                & "Where Obra =" & "'" & Codigo & "'"
                rstBuscarIcaro.Open SQL, dbIcaro, adOpenDynamic, adLockOptimistic
                If rstBuscarIcaro!Suma <> 0 Then
                    .TextMatrix(i, 5) = rstBuscarIcaro!Suma + rstListadoIcaro!Imputado
                Else
                    .TextMatrix(i, 5) = rstListadoIcaro!Imputado
                End If
                rstBuscarIcaro.Close
                SQL = "Select * From RESOLUCIONESEPAM where Codigo =" & "'" & Codigo & "'"
                rstBuscarIcaro.Open SQL, dbIcaro, adOpenDynamic, adLockOptimistic
                If rstBuscarIcaro.BOF = False Then
                    .Rows = .Rows + 1
                    rstBuscarIcaro.MoveFirst
                    While rstBuscarIcaro.EOF = False
                        i = i + 1
                        .RowHeight(i) = 300
                        If Len(rstBuscarIcaro!Descripcion) <> "0" Then
                            .TextMatrix(i, 0) = rstBuscarIcaro!Descripcion
                        Else
                            .TextMatrix(i, 0) = "SIN DEFINIR"
                        End If
                        If Len(rstBuscarIcaro!Expediente) <> "0" Then
                            .TextMatrix(i, 1) = rstBuscarIcaro!Expediente
                        Else
                            .TextMatrix(i, 1) = "FALTA EXP."
                        End If
                        If Len(rstBuscarIcaro!Resolucion) <> "0" Then
                            .TextMatrix(i, 2) = rstBuscarIcaro!Resolucion
                        Else
                            .TextMatrix(i, 2) = "FALTA RES."
                        End If
                        If Len(rstBuscarIcaro!Importe) <> 0 Then
                            .TextMatrix(i, 3) = rstBuscarIcaro!Importe
                        Else
                            .TextMatrix(i, 3) = "0"
                        End If
                        If rstBuscarIcaro!Imputado <> 0 Then
                            .TextMatrix(i, 4) = rstBuscarIcaro!Imputado
                            .TextMatrix(i, 5) = rstBuscarIcaro!Imputado
                            .TextMatrix(1, 4) = De_Txt_a_Num_01(.TextMatrix(1, 4)) - rstBuscarIcaro!Imputado
                            .TextMatrix(1, 5) = De_Txt_a_Num_01(.TextMatrix(1, 5)) - rstBuscarIcaro!Imputado
                        Else
                            .TextMatrix(i, 4) = "0"
                            .TextMatrix(i, 5) = "0"
                        End If
                        DecisionATomar (i)
                        rstBuscarIcaro.MoveNext
                        .Rows = .Rows + 1
                    Wend
                    rstBuscarIcaro.Close
                    .Rows = .Rows - 1
                End If
                rstListadoIcaro.MoveNext
                .Rows = .Rows + 1
            Wend
            DecisionATomar ("1")
        End If
        .Rows = .Rows - 1
    End With
    
    rstListadoIcaro.Close
    Set rstListadoIcaro = Nothing
    Set rstBuscarIcaro = Nothing

End Sub

Public Sub DecisionATomar(Fila As String)
    
    With EPAMConsolidado.dgResolucionesEPAM
        Select Case De_Txt_a_Num_01(.TextMatrix(Fila, 3)) - De_Txt_a_Num_01(.TextMatrix(Fila, 4))
        Case Is = 0
            If (De_Txt_a_Num_01(.TextMatrix(Fila, 4)) - De_Txt_a_Num_01(.TextMatrix(Fila, 5))) = 0 Then
                .TextMatrix(Fila, 6) = "No Pagar, No Imputar"
                .TextMatrix(Fila, 7) = "0"
            ElseIf (De_Txt_a_Num_01(.TextMatrix(Fila, 4)) - De_Txt_a_Num_01(.TextMatrix(Fila, 5))) > 0 Then
                .TextMatrix(Fila, 6) = "No Pagar y Certificar"
                .TextMatrix(Fila, 7) = De_Num_a_Tx_01(De_Txt_a_Num_01(.TextMatrix(Fila, 4)) - De_Txt_a_Num_01(.TextMatrix(Fila, 5)))
            ElseIf (De_Txt_a_Num_01(.TextMatrix(Fila, 4)) - De_Txt_a_Num_01(.TextMatrix(Fila, 5))) < 0 Then
                .TextMatrix(Fila, 6) = "No Pagar. Verificar Presp."
                .TextMatrix(Fila, 7) = "0"
            End If
        Case Is > 0
            If (De_Txt_a_Num_01(.TextMatrix(Fila, 4)) - De_Txt_a_Num_01(.TextMatrix(Fila, 5))) = 0 Then
                .TextMatrix(Fila, 6) = "Pagar e Imputar"
                .TextMatrix(Fila, 7) = De_Num_a_Tx_01(De_Txt_a_Num_01(.TextMatrix(Fila, 3)) - De_Txt_a_Num_01(.TextMatrix(Fila, 4)))
            ElseIf (De_Txt_a_Num_01(.TextMatrix(Fila, 4)) - De_Txt_a_Num_01(.TextMatrix(Fila, 5))) > 0 Then
                .TextMatrix(Fila, 6) = "Pagar e Imputar"
                .TextMatrix(Fila, 7) = De_Num_a_Tx_01(De_Txt_a_Num_01(.TextMatrix(Fila, 3)) - De_Txt_a_Num_01(.TextMatrix(Fila, 4)))
            ElseIf (De_Txt_a_Num_01(.TextMatrix(Fila, 4)) - De_Txt_a_Num_01(.TextMatrix(Fila, 5))) < 0 And (De_Txt_a_Num_01(.TextMatrix(Fila, 3)) - De_Txt_a_Num_01(.TextMatrix(Fila, 5))) >= 0 Then
                .TextMatrix(Fila, 6) = "Pagar y No Imputar"
                .TextMatrix(Fila, 7) = De_Num_a_Tx_01(De_Txt_a_Num_01(.TextMatrix(Fila, 5)) - De_Txt_a_Num_01(.TextMatrix(Fila, 4)))
            ElseIf (De_Txt_a_Num_01(.TextMatrix(Fila, 4)) - De_Txt_a_Num_01(.TextMatrix(Fila, 5))) < 0 And (De_Txt_a_Num_01(.TextMatrix(Fila, 3)) - De_Txt_a_Num_01(.TextMatrix(Fila, 5))) < 0 Then
                .TextMatrix(Fila, 6) = "Pagar y No Imputar"
                .TextMatrix(Fila, 7) = De_Num_a_Tx_01(De_Txt_a_Num_01(.TextMatrix(Fila, 3)) - De_Txt_a_Num_01(.TextMatrix(Fila, 4)))
            End If
        Case Is < 0
            If (De_Txt_a_Num_01(.TextMatrix(Fila, 4)) - De_Txt_a_Num_01(.TextMatrix(Fila, 5))) = 0 Then
                .TextMatrix(Fila, 6) = "No Pagar, No Imputar"
                .TextMatrix(Fila, 7) = "0"
            ElseIf (De_Txt_a_Num_01(.TextMatrix(Fila, 4)) - De_Txt_a_Num_01(.TextMatrix(Fila, 5))) > 0 Then
                .TextMatrix(Fila, 6) = "No Pagar y Certificar"
                .TextMatrix(Fila, 7) = De_Num_a_Tx_01(De_Txt_a_Num_01(.TextMatrix(Fila, 4)) - De_Txt_a_Num_01(.TextMatrix(Fila, 5)))
            ElseIf (De_Txt_a_Num_01(.TextMatrix(Fila, 4)) - De_Txt_a_Num_01(.TextMatrix(Fila, 5))) < 0 Then
                .TextMatrix(Fila, 6) = "No Pagar. Verificar Presp."
                .TextMatrix(Fila, 7) = "0"
            End If
        End Select
    End With
    Fila = ""
    
End Sub

Public Sub ConfigurardgEstructura()
    
    With EstructuraProgramatica.dgEstructura
        .Clear
        .Cols = 2
        .Rows = 2
        .TextMatrix(0, 0) = "Imputación"
        .TextMatrix(0, 1) = "Descripción"
        .ColWidth(0) = 1500
        .ColWidth(1) = 6500
        .FixedCols = 0
        .FocusRect = flexFocusHeavy
        .HighLight = flexHighlightWithFocus
        .SelectionMode = flexSelectionByRow
        .AllowUserResizing = flexResizeColumns
        .ColAlignment(0) = 4
        .ColAlignment(1) = 1
    End With
    
End Sub

Public Sub CargardgEstructura(Optional ByVal Jerarquia As String = "0")
    
    Dim i As Integer
    Dim SQL As String
    i = 0
    
    EstructuraProgramatica.dgEstructura.Rows = 2
    Set rstListadoIcaro = New ADODB.Recordset
    With EstructuraProgramatica
        Select Case Len(Jerarquia)
        Case Is = 1
            SQL = "Select * From PROGRAMAS Order by Programa Asc"
            .cmdAgregar.Caption = "Agregar Programa (F6)"
            .cmdEditar.Caption = "Editar Programa (F9)"
            .cmdEliminar.Caption = "Eliminar Programa (F8)"
            .cmdIngresar.Caption = "Ver Subprograma (F4)"
            .cmdIngresar.Enabled = True
            .cmdVolver.Enabled = False
            .frmListado.Caption = "Listado de Programas"
        Case Is = 2
            SQL = "Select SUBPROGRAMA, DESCRIPCION From SUBPROGRAMAS " _
            & "Where PROGRAMA = '" & Jerarquia & " ' Order by SUBPROGRAMA Asc"
            .cmdAgregar.Caption = "Agregar Subprograma (F6)"
            .cmdEditar.Caption = "Editar Subprograma (F9)"
            .cmdEliminar.Caption = "Eliminar Subprograma (F8)"
            .cmdIngresar.Caption = "Ver Proyectos (F4)"
            .cmdVolver.Enabled = True
            .frmListado.Caption = "Listado de Subprogramas"
        Case Is = 5
            SQL = "Select PROYECTO, DESCRIPCION From PROYECTOS " _
            & "Where SUBPROGRAMA = '" & Jerarquia & " ' Order by PROYECTO Asc"
            .cmdAgregar.Caption = "Agregar Proyecto (F6)"
            .cmdEditar.Caption = "Editar Proyecto (F9)"
            .cmdEliminar.Caption = "Eliminar Proyecto (F8)"
            .cmdIngresar.Caption = "Ver Actividad (F4)"
            .cmdVolver.Enabled = True
            .frmListado.Caption = "Listado de Proyectos"
        Case Is = 8
            SQL = "Select ACTIVIDAD, DESCRIPCION From ACTIVIDADES " _
            & "Where PROYECTO = '" & Jerarquia & " ' Order by ACTIVIDAD Asc"
            .cmdAgregar.Caption = "Agregar Actividad (F6)"
            .cmdEditar.Caption = "Editar Actividad (F9)"
            .cmdEliminar.Caption = "Eliminar Actividad (F8)"
            .cmdIngresar.Caption = "Ver Obras Específicas (F4)"
            .cmdVolver.Enabled = True
            .frmListado.Caption = "Listado de Actividades"
        Case Is = 11
            SQL = "Select IMPUTACION, PARTIDA, DESCRIPCION From OBRAS " _
            & "Where IMPUTACION = '" & Jerarquia & " ' Order by DESCRIPCION Asc"
            .cmdAgregar.Caption = "Agregar Obra (F6)"
            .cmdEditar.Caption = "Editar Obra (F9)"
            .cmdEliminar.Caption = "Eliminar Obra (F8)"
            .cmdIngresar.Caption = ""
            .cmdIngresar.Enabled = False
            .cmdVolver.Enabled = True
            .frmListado.Caption = "Listado de Obras"
        End Select
    End With
    
    rstListadoIcaro.Open SQL, dbIcaro, adOpenDynamic, adLockOptimistic
    With EstructuraProgramatica.dgEstructura
        If rstListadoIcaro.BOF = False Then
            rstListadoIcaro.MoveFirst
            While rstListadoIcaro.EOF = False
                i = i + 1
                .RowHeight(i) = 300
                If Len(Jerarquia) = 11 Then
                    .TextMatrix(i, 0) = rstListadoIcaro.Fields("Imputacion") & "-" & rstListadoIcaro.Fields("Partida")
                Else
                    .TextMatrix(i, 0) = rstListadoIcaro.Fields(0)
                End If
                If Len(rstListadoIcaro.Fields("Descripcion")) <> "0" Then
                    .TextMatrix(i, 1) = rstListadoIcaro.Fields("Descripcion")
                Else
                    .TextMatrix(i, 1) = ""
                End If
                rstListadoIcaro.MoveNext
                .Rows = .Rows + 1
            Wend
            .Rows = .Rows - 1
        End If
    End With
    rstListadoIcaro.Close
    Set rstListadoIcaro = Nothing
    With EstructuraProgramatica
        If Len(.dgEstructura.TextMatrix(1, 0)) = "0" Then
            .cmdEditar.Enabled = False
            .cmdEliminar.Enabled = False
            .cmdIngresar.Enabled = False
        Else
            .cmdEditar.Enabled = True
            .cmdEliminar.Enabled = True
            .cmdIngresar.Enabled = True
        End If
    End With

End Sub

Sub ConfigurardgEjecucionMensual(CantidadDeMeses As Integer)

    Dim x As Integer
    x = 1
    
    With EjecucionMensual.dgEjecucion
        .Clear
        .Cols = (CantidadDeMeses + 3)
        .Rows = 2
        .TextMatrix(0, 0) = "Actividad"
        .TextMatrix(0, 1) = "Partida"
        .TextMatrix(0, 2) = "Total Anual"
        .ColWidth(0) = 1000
        .ColWidth(1) = 600
        .ColWidth(2) = 1000
        .ColAlignment(0) = 1
        .ColAlignment(1) = 7
        .ColAlignment(2) = 7
        .FixedCols = 0
        .FocusRect = flexFocusHeavy
        .HighLight = flexHighlightWithFocus
        .AllowUserResizing = flexResizeColumns
        For x = 3 To (CantidadDeMeses + 2)
            .TextMatrix(0, x) = Format(Format("01/0" & (x - 2) & "/2009", "mmmm"), ">")
            .ColWidth(x) = 1000
            .ColAlignment(x) = 7
        Next x
    End With
    
    x = 0
    
End Sub

Sub CargardgEjecucionMensual(CantidadDeMeses As Integer)

    Dim i As Integer
    i = 0
    Dim x As Integer
    x = 0
    Dim SQL As String
    
    EjecucionMensual.dgEjecucion.Rows = 2
    Set rstListadoIcaro = New ADODB.Recordset
    SQL = "TRANSFORM Sum(IMPUTACION.Importe) AS SumaDeImporte SELECT ACTIVIDADES.Actividad, " _
    & "IMPUTACION.Partida, Sum(IMPUTACION.Importe) AS [Total Anual] FROM " _
    & "CARGA INNER JOIN (ACTIVIDADES INNER JOIN IMPUTACION ON ACTIVIDADES.Actividad=IMPUTACION.Imputacion) " _
    & "ON CARGA.Comprobante=IMPUTACION.Comprobante Where (((Year(CARGA.Fecha)) =" & strEjercicio & ")) " _
    & "GROUP BY ACTIVIDADES.Actividad, IMPUTACION.Partida, Year(Carga.Fecha) ORDER BY Format(CARGA!Fecha,""mm"") " _
    & "PIVOT Format(CARGA!Fecha,""mm"")"
    rstListadoIcaro.Open SQL, dbIcaro, adOpenDynamic, adLockOptimistic
    With EjecucionMensual.dgEjecucion
        If rstListadoIcaro.BOF = False Then
            rstListadoIcaro.MoveFirst
            While rstListadoIcaro.EOF = False
                i = i + 1
                .RowHeight(i) = 300
                For x = 0 To (CantidadDeMeses + 2)
                    If x <= 1 Then
                        If Len(rstListadoIcaro.Fields(x)) <> "0" Then
                            .TextMatrix(i, x) = rstListadoIcaro.Fields(x)
                        Else
                            .TextMatrix(i, x) = "0"
                        End If
                    Else
                        If Len(rstListadoIcaro.Fields(x)) <> "0" Then
                            .TextMatrix(i, x) = FormatNumber(rstListadoIcaro.Fields(x), 2)
                        Else
                            .TextMatrix(i, x) = "0"
                        End If
                    End If
                Next x
                rstListadoIcaro.MoveNext
                .Rows = .Rows + 1
            Wend
            .Rows = .Rows - 1
        End If
    End With
    
    rstListadoIcaro.Close
    Set rstListadoIcaro = Nothing
    i = 0
    x = 0
    SQL = ""
        
        
End Sub

Public Sub ConfigurardgComprobantesEjecucionMensual()

    With EjecucionMensual.dgComprobantes
        .Clear
        .Cols = 5
        .Rows = 2
        .TextMatrix(0, 0) = "Comprobante"
        .TextMatrix(0, 1) = "Fecha"
        .TextMatrix(0, 2) = "Obra"
        .TextMatrix(0, 3) = "Importe"
        .TextMatrix(0, 4) = "Avance"
        .ColWidth(0) = 1200
        .ColWidth(1) = 1000
        .ColWidth(2) = 3500
        .ColWidth(3) = 1000
        .ColWidth(4) = 1000
        .FixedCols = 0
        .FocusRect = flexFocusHeavy
        .HighLight = flexHighlightWithFocus
        .SelectionMode = flexSelectionByRow
        .ColAlignment(0) = 1
        .ColAlignment(1) = 1
        .ColAlignment(2) = 1
        .ColAlignment(3) = 7
        .ColAlignment(4) = 7
    End With
    
End Sub

Public Sub CargardgComprobantesEjecucionMensual(Imputacion As String)
    
    Dim i As Integer
    i = 0
    Dim SQL As String
    
    EjecucionMensual.dgComprobantes.Rows = 2
    Set rstListadoIcaro = New ADODB.Recordset
    SQL = "Select IMPUTACION.Comprobante, IMPUTACION.Importe, IMPUTACION.Partida, CARGA.Fecha, " _
    & "CARGA.Obra, CARGA.Avance From CARGA Inner Join IMPUTACION on CARGA.Comprobante = IMPUTACION.Comprobante " _
    & "Where Imputacion = " & "'" & Left(Imputacion, 11) & "' And Right(CARGA.Comprobante, 2) = '" & Right(strEjercicio, 2) _
    & "' And Partida = " & "'" & Right(Imputacion, 3) & "' Order by IMPUTACION.Comprobante Desc, Fecha Desc"
    rstListadoIcaro.Open SQL, dbIcaro, adOpenDynamic, adLockOptimistic
    With EjecucionMensual.dgComprobantes
        If rstListadoIcaro.BOF = False Then
            rstListadoIcaro.MoveFirst
            While rstListadoIcaro.EOF = False
                i = i + 1
                .RowHeight(i) = 300
                .TextMatrix(i, 0) = rstListadoIcaro!Comprobante
                .TextMatrix(i, 1) = rstListadoIcaro!Fecha
                .TextMatrix(i, 2) = rstListadoIcaro!Obra
                .TextMatrix(i, 3) = FormatNumber(rstListadoIcaro!Importe, 2)
                .TextMatrix(i, 4) = rstListadoIcaro!Avance * 100
                rstListadoIcaro.MoveNext
                .Rows = .Rows + 1
            Wend
            .Rows = .Rows - 1
        End If
    End With

    rstListadoIcaro.Close
    Set rstListadoIcaro = Nothing
    i = 0
    SQL = ""
    
End Sub

Public Sub ConfigurardgObrasTerminadas()
    
    With ObrasTerminadasPorEstructura.dgObrasTerminadas
        .Clear
        .Cols = 13
        .Rows = 2
        .TextMatrix(0, 0) = "Imputacion"
        .TextMatrix(0, 1) = "Descripcion"
        .TextMatrix(0, 2) = "Monto Contrato"
        .TextMatrix(0, 3) = "Adicional"
        .TextMatrix(0, 4) = "Total Contrato"
        .TextMatrix(0, 5) = "Ejec. Anterior"
        .TextMatrix(0, 6) = "Ejec. Actual"
        .TextMatrix(0, 7) = "Ejec. Total"
        .TextMatrix(0, 8) = "En Curso"
        .TextMatrix(0, 9) = "Term. Anterior"
        .TextMatrix(0, 10) = "Term. Actual"
        .TextMatrix(0, 11) = "Rec. Anterior"
        .TextMatrix(0, 12) = "Rec. Actual"
        .ColWidth(0) = 1300
        .ColWidth(1) = 4800
        .ColWidth(2) = 1200
        .ColWidth(3) = 1200
        .ColWidth(4) = 1200
        .ColWidth(5) = 1200
        .ColWidth(6) = 1200
        .ColWidth(7) = 1200
        .ColWidth(8) = 1200
        .ColWidth(9) = 1200
        .ColWidth(10) = 1200
        .ColWidth(11) = 1200
        .ColWidth(12) = 1200
        .FixedCols = 1
        .ColAlignment(0) = 1
        .ColAlignment(1) = 1
        .ColAlignment(2) = 7
        .ColAlignment(3) = 7
        .ColAlignment(4) = 7
        .ColAlignment(5) = 7
        .ColAlignment(6) = 7
        .ColAlignment(7) = 7
        .ColAlignment(8) = 7
        .ColAlignment(9) = 7
        .ColAlignment(10) = 7
        .ColAlignment(11) = 7
        .ColAlignment(12) = 7
        .FocusRect = flexFocusHeavy
        .HighLight = flexHighlightWithFocus
        .AllowUserResizing = flexResizeColumns
    End With
    
End Sub

Public Sub CargardgObrasTerminadas()
    
    Dim curMontoAcumulado As Currency
    Dim SQL As String
    
    ObrasTerminadasPorEstructura.dgObrasTerminadas.Rows = 2
    Set rstListadoIcaro = New ADODB.Recordset
    SQL = "SELECT OBRAS.Imputacion, OBRAS.Partida, ACTIVIDADES.Descripcion, Sum(CARGA.Importe) AS Ejecucion " _
    & "FROM ((((PROGRAMAS INNER JOIN SUBPROGRAMAS ON PROGRAMAS.Programa = SUBPROGRAMAS.Programa) " _
    & "INNER JOIN PROYECTOS ON SUBPROGRAMAS.Subprograma = PROYECTOS.Subprograma) " _
    & "INNER JOIN ACTIVIDADES ON PROYECTOS.Proyecto = ACTIVIDADES.Proyecto) " _
    & "INNER JOIN OBRAS ON ACTIVIDADES.Actividad = OBRAS.Imputacion) " _
    & "INNER JOIN CARGA ON OBRAS.Descripcion = CARGA.Obra " _
    & "WHERE (((Year(CARGA.Fecha)) <=" & strEjercicio & ")) " _
    & "GROUP BY OBRAS.Imputacion, OBRAS.Partida, ACTIVIDADES.Descripcion " _
    & "ORDER BY OBRAS.Imputacion, OBRAS.Partida, ACTIVIDADES.Descripcion"
    rstListadoIcaro.Open SQL, dbIcaro, adOpenDynamic, adLockOptimistic
    With ObrasTerminadasPorEstructura.dgObrasTerminadas
        If rstListadoIcaro.BOF = False Then
            rstListadoIcaro.MoveFirst
            Set rstBuscarIcaro = New ADODB.Recordset
            Set rstRegistroIcaro = New ADODB.Recordset
            While rstListadoIcaro.EOF = False
                i = i + 1
                .RowHeight(i) = 300
                'Imputación y Descripción GENERAL
                .TextMatrix(i, 0) = rstListadoIcaro.Fields("Imputacion") & "-" & rstListadoIcaro.Fields("Partida")
                .TextMatrix(i, 1) = rstListadoIcaro.Fields("Descripcion")
                'Montos de Contratos GENERAL
                SQL = "SELECT OBRAS.Imputacion, OBRAS.Partida, Sum(OBRAS.MontodeContrato) As MontoDeContrato, Sum(OBRAS.Adicional) As Adicional " _
                & "From OBRAS Where OBRAS.Imputacion = '" & rstListadoIcaro.Fields("Imputacion") & "' " _
                & "And OBRAS.Partida = '" & rstListadoIcaro.Fields("Partida") & "' " _
                & "Group By OBRAS.Imputacion, OBRAS.Partida"
                rstBuscarIcaro.Open SQL, dbIcaro, adOpenDynamic, adLockOptimistic
                .TextMatrix(i, 2) = Format(rstBuscarIcaro.Fields("MontoDeContrato"), "#,##0.00")
                .TextMatrix(i, 3) = Format(rstBuscarIcaro.Fields("Adicional"), "#,##0.00")
                .TextMatrix(i, 4) = Format((rstBuscarIcaro.Fields("MontoDeContrato") + rstBuscarIcaro.Fields("Adicional")), "#,##0.00")
                rstBuscarIcaro.Close
                'Montos de Ejecución GENERAL
                SQL = "SELECT OBRAS.Imputacion, OBRAS.Partida, Sum(CARGA.Importe) AS Ejecucion " _
                & "FROM OBRAS INNER JOIN CARGA ON OBRAS.Descripcion = CARGA.Obra " _
                & "Where (((Year(CARGA.Fecha)) =" & strEjercicio & ")) " _
                & "And OBRAS.Imputacion = '" & rstListadoIcaro.Fields("Imputacion") & "' " _
                & "And OBRAS.Partida = '" & rstListadoIcaro.Fields("Partida") & "' " _
                & "GROUP BY OBRAS.Imputacion, OBRAS.Partida"
                rstBuscarIcaro.Open SQL, dbIcaro, adOpenDynamic, adLockOptimistic
                If rstBuscarIcaro.BOF = False Then
                    .TextMatrix(i, 6) = Format(rstBuscarIcaro.Fields("Ejecucion"), "#,##0.00")
                    .TextMatrix(i, 5) = Format((rstListadoIcaro.Fields("Ejecucion") - rstBuscarIcaro.Fields("Ejecucion")), "#,##0.00")
                Else
                    .TextMatrix(i, 6) = Format("0", "#,##0.00")
                    .TextMatrix(i, 5) = Format(rstListadoIcaro.Fields("Ejecucion"), "#,##0.00")
                End If
                rstBuscarIcaro.Close
                .TextMatrix(i, 7) = Format(rstListadoIcaro.Fields("Ejecucion"), "#,##0.00")
                'Monto de Ejecución GENERAL en Curso (Avance <> 100%)
                SQL = "SELECT CARGA.Obra, Sum(CARGA.Importe) AS Ejecucion " _
                & "FROM OBRAS INNER JOIN CARGA ON OBRAS.Descripcion = CARGA.Obra " _
                & "Where (((Year(CARGA.Fecha)) <=" & strEjercicio & ")) " _
                & "And OBRAS.Imputacion = '" & rstListadoIcaro.Fields("Imputacion") & "' " _
                & "And OBRAS.Partida = '" & rstListadoIcaro.Fields("Partida") & "' " _
                & "GROUP BY CARGA.Obra Having (((Max([CARGA].[Avance]) * 100) <> 100)) " _
                & "ORDER BY CARGA.Obra"
                rstBuscarIcaro.Open SQL, dbIcaro, adOpenDynamic, adLockOptimistic
                If rstBuscarIcaro.BOF = False Then
                    rstBuscarIcaro.MoveFirst
                    While rstBuscarIcaro.EOF = False
                        Debug.Print rstBuscarIcaro.Fields("Obra")
                        curMontoAcumulado = curMontoAcumulado + rstBuscarIcaro.Fields("Ejecucion")
                        rstBuscarIcaro.MoveNext
                    Wend
                    .TextMatrix(i, 8) = Format(curMontoAcumulado, "#,##0.00")
                    curMontoAcumulado = 0
                Else
                    .TextMatrix(i, 8) = Format("0", "#,##0.00")
                End If
                rstBuscarIcaro.Close
                'Monto Ejecución GENERAL Terminada en el Ejercicio Actual (Avance = 100% Ejercicio _
                Actual + Avance  <> 100% Ejercicios Anteriores)
                SQL = "SELECT CARGA.Obra As Descripcion, Sum(CARGA.Importe) AS Ejecucion " _
                & "FROM OBRAS INNER JOIN CARGA ON OBRAS.Descripcion = CARGA.Obra " _
                & "Where (((Year(CARGA.Fecha)) = " & strEjercicio & ")) " _
                & "And OBRAS.Imputacion = '" & rstListadoIcaro.Fields("Imputacion") & "' " _
                & "And OBRAS.Partida = '" & rstListadoIcaro.Fields("Partida") & "' " _
                & "GROUP BY CARGA.Obra Having (((Max([CARGA].[Avance]) * 100) = 100)) " _
                & "ORDER BY CARGA.Obra"
                rstBuscarIcaro.Open SQL, dbIcaro, adOpenDynamic, adLockOptimistic
                If rstBuscarIcaro.BOF = False Then
                    rstBuscarIcaro.MoveFirst
                    While rstBuscarIcaro.EOF = False
                        curMontoAcumulado = curMontoAcumulado + rstBuscarIcaro.Fields("Ejecucion")
                        SQL = "SELECT CARGA.Obra, Sum(CARGA.Importe) AS Ejecucion " _
                        & "FROM OBRAS INNER JOIN CARGA ON OBRAS.Descripcion = CARGA.Obra " _
                        & "Where (((Year(CARGA.Fecha)) < " & strEjercicio & ")) " _
                        & "And CARGA.Obra = '" & rstBuscarIcaro.Fields("Descripcion") & "' " _
                        & "GROUP BY CARGA.Obra Having (((Max([CARGA].[Avance]) * 100) <> 100))"
                        rstRegistroIcaro.Open SQL, dbIcaro, adOpenDynamic, adLockOptimistic
                        If rstRegistroIcaro.BOF = False Then
                            rstRegistroIcaro.MoveFirst
                            While rstRegistroIcaro.EOF = False
                                curMontoAcumulado = curMontoAcumulado + rstRegistroIcaro.Fields("Ejecucion")
                                rstRegistroIcaro.MoveNext
                            Wend
                        End If
                        rstRegistroIcaro.Close
                        rstBuscarIcaro.MoveNext
                    Wend
                    .TextMatrix(i, 10) = Format(curMontoAcumulado, "#,##0.00")
                    curMontoAcumulado = 0
                Else
                    .TextMatrix(i, 10) = Format("0", "#,##0.00")
                End If
                rstBuscarIcaro.Close
                'Monto Ejecución GENERAL Terminadas en Ejercicios Anteriores (se obtiene por difencia)
                .TextMatrix(i, 9) = Format((.TextMatrix(i, 7) - .TextMatrix(i, 8) - .TextMatrix(i, 10)), "#,##0.00")
                'DESAGREGANDO Datos Generales (NO SE DESAGREGA OBRAS GLOBO)
                'If rstListadoIcaro.Fields("Imputacion") <> "11-00-02-51" And rstListadoIcaro.Fields("Imputacion") <> "11-00-02-79" _
                And rstListadoIcaro.Fields("Imputacion") <> "11-00-02-82" And rstListadoIcaro.Fields("Imputacion") <> "11-00-02-96" _
                And rstListadoIcaro.Fields("Imputacion") <> "11-00-43-52" And rstListadoIcaro.Fields("Imputacion") <> "11-00-44-51" _
                And rstListadoIcaro.Fields("Imputacion") <> "11-00-49-51" And rstListadoIcaro.Fields("Imputacion") <> "12-00-01-52" _
                And rstListadoIcaro.Fields("Imputacion") <> "12-00-01-84" And rstListadoIcaro.Fields("Imputacion") <> "12-00-43-51" _
                And rstListadoIcaro.Fields("Imputacion") <> "12-00-51-80" Then
                    SQL = "SELECT OBRAS.MontoDeContrato, OBRAS.Adicional, OBRAS.Descripcion, Sum(CARGA.Importe) AS Ejecucion " _
                    & "FROM OBRAS INNER JOIN CARGA ON OBRAS.Descripcion = CARGA.Obra " _
                    & "WHERE (((Year(CARGA.Fecha)) <=" & strEjercicio & ")) " _
                    & "And OBRAS.Imputacion = '" & rstListadoIcaro.Fields("Imputacion") & "' " _
                    & "And OBRAS.Partida = '" & rstListadoIcaro.Fields("Partida") & "' " _
                    & "GROUP BY OBRAS.Descripcion, OBRAS.MontoDeContrato, OBRAS.Adicional " _
                    & "ORDER BY OBRAS.Descripcion"
                    rstBuscarIcaro.Open SQL, dbIcaro, adOpenDynamic, adLockOptimistic
                    rstBuscarIcaro.MoveFirst
                    While rstBuscarIcaro.EOF = False
                        .Rows = .Rows + 1
                        i = i + 1
                        .RowHeight(i) = 300
                        'Descripción DESAGREGADA de las Obras
                        .TextMatrix(i, 1) = rstBuscarIcaro.Fields("Descripcion")
                        'Montos de Contratos DESAGREGADOS
                        .TextMatrix(i, 2) = Format(rstBuscarIcaro.Fields("MontoDeContrato"), "#,##0.00")
                        .TextMatrix(i, 3) = Format(rstBuscarIcaro.Fields("Adicional"), "#,##0.00")
                        .TextMatrix(i, 4) = Format((rstBuscarIcaro.Fields("MontoDeContrato") + rstBuscarIcaro.Fields("Adicional")), "#,##0.00")
                        'Fecha de Alta de la Obra
                        SQL = "SELECT CARGA.Fecha FROM Carga " _
                        & "WHERE CARGA.Obra = '" & rstBuscarIcaro.Fields("Descripcion") & "' " _
                        & "ORDER BY Fecha Asc"
                        rstRegistroIcaro.Open SQL, dbIcaro, adOpenDynamic, adLockOptimistic
                        rstRegistroIcaro.MoveFirst
                        .TextMatrix(i, 0) = "Alta " & Right(rstRegistroIcaro.Fields("Fecha"), 4)
                        rstRegistroIcaro.Close
                        'Montos de Ejecución DESAGREGADO
                        SQL = "SELECT CARGA.Obra, Sum(CARGA.Importe) AS Ejecucion FROM CARGA " _
                        & "Where (((Year(CARGA.Fecha)) =" & strEjercicio & ")) " _
                        & "And CARGA.Obra = '" & rstBuscarIcaro.Fields("Descripcion") & "'" _
                        & "GROUP BY CARGA.Obra"
                        rstRegistroIcaro.Open SQL, dbIcaro, adOpenDynamic, adLockOptimistic
                        If rstRegistroIcaro.BOF = False Then
                            .TextMatrix(i, 6) = Format(rstRegistroIcaro.Fields("Ejecucion"), "#,##0.00")
                            .TextMatrix(i, 5) = Format((rstBuscarIcaro.Fields("Ejecucion") - rstRegistroIcaro.Fields("Ejecucion")), "#,##0.00")
                        Else
                            .TextMatrix(i, 6) = Format("0", "#,##0.00")
                            .TextMatrix(i, 5) = Format(rstBuscarIcaro.Fields("Ejecucion"), "#,##0.00")
                        End If
                        rstRegistroIcaro.Close
                        .TextMatrix(i, 7) = Format(rstBuscarIcaro.Fields("Ejecucion"), "#,##0.00")
                        'Montos de Obras en Curso y Terminadas DESAGREGADO
                        'Primero se determina si la obra esta en curso, en caso contrario _
                        & se evalúa cuando se terminó
                        SQL = "SELECT CARGA.Obra, Sum(CARGA.Importe) AS Ejecucion FROM CARGA " _
                        & "Where (((Year(CARGA.Fecha)) <=" & strEjercicio & ")) " _
                        & "And CARGA.Obra = '" & rstBuscarIcaro.Fields("Descripcion") & "'" _
                        & "GROUP BY CARGA.Obra HAVING (((Max([CARGA].[Avance]) * 100) <> 100))"
                        rstRegistroIcaro.Open SQL, dbIcaro, adOpenDynamic, adLockOptimistic
                        If rstRegistroIcaro.BOF = False Then
                            .TextMatrix(i, 8) = Format(rstRegistroIcaro.Fields("Ejecucion"), "#,##0.00")
                            .TextMatrix(i, 9) = Format("0", "#,##0.00")
                            .TextMatrix(i, 10) = Format("0", "#,##0.00")
                        Else
                            'En caso que la obra no esté en curso, se evalúa si se terminó anteriormente
                            .TextMatrix(i, 8) = Format("0", "#,##0.00")
                            rstRegistroIcaro.Close
                            SQL = "SELECT CARGA.Obra, Sum(CARGA.Importe) AS Ejecucion FROM CARGA " _
                            & "Where (((Year(CARGA.Fecha)) <" & strEjercicio & ")) " _
                            & "And CARGA.Obra = '" & rstBuscarIcaro.Fields("Descripcion") & "'" _
                            & "GROUP BY CARGA.Obra HAVING (((Max([CARGA].[Avance]) * 100) = 100))"
                            rstRegistroIcaro.Open SQL, dbIcaro, adOpenDynamic, adLockOptimistic
                            If rstRegistroIcaro.BOF = False Then
                                .TextMatrix(i, 9) = Format(rstRegistroIcaro.Fields("Ejecucion"), "#,##0.00")
                                If rstRegistroIcaro.Fields("Ejecucion") = rstBuscarIcaro.Fields("Ejecucion") Then
                                    .TextMatrix(i, 10) = Format("0", "#,##0.00")
                                Else
                                    .TextMatrix(i, 10) = Format((rstBuscarIcaro.Fields("Ejecucion") - rstRegistroIcaro.Fields("Ejecucion")), "#,##0.00")
                                End If
                            Else
                                .TextMatrix(i, 9) = Format("0", "#,##0.00")
                                .TextMatrix(i, 10) = Format(rstBuscarIcaro.Fields("Ejecucion"), "#,##0.00")
                            End If
                        End If
                        rstRegistroIcaro.Close

                        
                        
                        rstBuscarIcaro.MoveNext
                    Wend
                    rstBuscarIcaro.Close
                'End If
                rstListadoIcaro.MoveNext
                .Rows = .Rows + 1
            Wend
            .Rows = .Rows - 1
        End If
    End With
    rstListadoIcaro.Close
    Set rstListadoIcaro = Nothing
    Set rstBuscarIcaro = Nothing
    Set rstRegistroIcaro = Nothing

End Sub

Public Sub ConfigurardgCertificados100()
    
    With Certificados100.dgObrasTerminadas
        .Clear
        .Cols = 14
        .Rows = 2
        .TextMatrix(0, 0) = "Imputacion"
        .TextMatrix(0, 1) = "Descripcion"
        .TextMatrix(0, 2) = "Monto Contrato"
        .TextMatrix(0, 3) = "Adicional"
        .TextMatrix(0, 4) = "Total Contrato"
        .TextMatrix(0, 5) = "Ejec. Anterior"
        .TextMatrix(0, 6) = "Ejec. Actual"
        .TextMatrix(0, 7) = "Ejec. Total"
        .TextMatrix(0, 8) = "FR Anterior"
        .TextMatrix(0, 9) = "FR Actual"
        .TextMatrix(0, 10) = "FR Total"
        .TextMatrix(0, 11) = "Total Certif."
        .TextMatrix(0, 12) = "Diferencia"
        .TextMatrix(0, 13) = "Último Certif."
        .ColWidth(0) = 1300
        .ColWidth(1) = 6200
        .ColWidth(2) = 1200
        .ColWidth(3) = 1200
        .ColWidth(4) = 1200
        .ColWidth(5) = 1200
        .ColWidth(6) = 1200
        .ColWidth(7) = 1200
        .ColWidth(8) = 1200
        .ColWidth(9) = 1200
        .ColWidth(10) = 1200
        .ColWidth(11) = 1200
        .ColWidth(12) = 1200
        .ColWidth(13) = 1200
        .FixedCols = 1
        .ColAlignment(0) = 1
        .ColAlignment(1) = 1
        .ColAlignment(2) = 7
        .ColAlignment(3) = 7
        .ColAlignment(4) = 7
        .ColAlignment(5) = 7
        .ColAlignment(6) = 7
        .ColAlignment(7) = 7
        .ColAlignment(8) = 7
        .ColAlignment(9) = 7
        .ColAlignment(10) = 7
        .ColAlignment(11) = 7
        .ColAlignment(12) = 7
        .FocusRect = flexFocusHeavy
        .HighLight = flexHighlightWithFocus
        .AllowUserResizing = flexResizeColumns
    End With

End Sub

Public Sub CargardgCertificados100()

    Dim SQL As String

    Certificados100.dgObrasTerminadas.Rows = 2
    Set rstListadoIcaro = New ADODB.Recordset
    SQL = "SELECT OBRAS.Imputacion, OBRAS.Partida, OBRAS.Descripcion, Sum(CARGA.Importe) AS Ejecucion, Sum(CARGA.Fondodereparo) AS [Fondo de Reparo], OBRAS.Montodecontrato, OBRAS.Adicional " _
    & "FROM ((PROGRAMAS INNER JOIN (SUBPROGRAMAS INNER JOIN PROYECTOS ON SUBPROGRAMAS.Subprograma = PROYECTOS.Subprograma) ON PROGRAMAS.Programa = SUBPROGRAMAS.Programa) INNER JOIN ACTIVIDADES ON PROYECTOS.Proyecto = ACTIVIDADES.Proyecto) INNER JOIN (PARTIDAS INNER JOIN (OBRAS INNER JOIN CARGA ON OBRAS.Descripcion = CARGA.Obra) ON PARTIDAS.Partida = OBRAS.Partida) ON ACTIVIDADES.Actividad = OBRAS.Imputacion " _
    & "WHERE (((Year(CARGA.Fecha)) =" & strEjercicio & ")) " _
    & "GROUP BY OBRAS.Imputacion, OBRAS.Partida, OBRAS.Descripcion, OBRAS.Montodecontrato, OBRAS.Adicional " _
    & "HAVING (((Max([CARGA].[Avance]) * 100) = 100)) " _
    & "ORDER BY OBRAS.Imputacion, OBRAS.Partida, OBRAS.Descripcion"
    rstListadoIcaro.Open SQL, dbIcaro, adOpenDynamic, adLockOptimistic
    With Certificados100.dgObrasTerminadas
        If rstListadoIcaro.BOF = False Then
            rstListadoIcaro.MoveFirst
            Set rstBuscarIcaro = New ADODB.Recordset
            While rstListadoIcaro.EOF = False
                i = i + 1
                .RowHeight(i) = 300
                .TextMatrix(i, 0) = rstListadoIcaro.Fields("Imputacion") & "-" & rstListadoIcaro.Fields("Partida")
                .TextMatrix(i, 1) = rstListadoIcaro.Fields("Descripcion")
                .TextMatrix(i, 2) = Format(rstListadoIcaro.Fields("MontoDeContrato"), "#,##0.00")
                .TextMatrix(i, 3) = Format(rstListadoIcaro.Fields("Adicional"), "#,##0.00")
                .TextMatrix(i, 4) = Format((rstListadoIcaro.Fields("MontoDeContrato") + rstListadoIcaro.Fields("Adicional")), "#,##0.00")
                SQL = "SELECT OBRAS.Imputacion, OBRAS.Partida, OBRAS.Descripcion, Sum(CARGA.Importe) AS Ejecucion, Sum(CARGA.Fondodereparo) AS [Fondo de Reparo], OBRAS.Montodecontrato, OBRAS.Adicional " _
                & "FROM ((PROGRAMAS INNER JOIN (SUBPROGRAMAS INNER JOIN PROYECTOS ON SUBPROGRAMAS.Subprograma = PROYECTOS.Subprograma) ON PROGRAMAS.Programa = SUBPROGRAMAS.Programa) INNER JOIN ACTIVIDADES ON PROYECTOS.Proyecto = ACTIVIDADES.Proyecto) INNER JOIN (PARTIDAS INNER JOIN (OBRAS INNER JOIN CARGA ON OBRAS.Descripcion = CARGA.Obra) ON PARTIDAS.Partida = OBRAS.Partida) ON ACTIVIDADES.Actividad = OBRAS.Imputacion " _
                & "WHERE (((Year(CARGA.Fecha)) < " & strEjercicio & ")) and OBRAS.Descripcion = '" & rstListadoIcaro.Fields("Descripcion") & "' " _
                & "GROUP BY OBRAS.Imputacion, OBRAS.Partida, OBRAS.Descripcion, OBRAS.Montodecontrato, OBRAS.Adicional " _
                & "ORDER BY OBRAS.Imputacion, OBRAS.Partida, OBRAS.Descripcion"
                rstBuscarIcaro.Open SQL, dbIcaro, adOpenDynamic, adLockOptimistic
                If rstBuscarIcaro.BOF = False Then
                    rstBuscarIcaro.MoveFirst
                    .TextMatrix(i, 5) = Format(rstBuscarIcaro.Fields("Ejecucion"), "#,##0.00")
                    .TextMatrix(i, 6) = Format(rstListadoIcaro.Fields("Ejecucion"), "#,##0.00")
                    .TextMatrix(i, 7) = Format((rstListadoIcaro.Fields("Ejecucion") + rstBuscarIcaro.Fields("Ejecucion")), "#,##0.00")
                    .TextMatrix(i, 8) = Format(rstBuscarIcaro.Fields("Fondo de Reparo"), "#,##0.00")
                    .TextMatrix(i, 9) = Format(rstListadoIcaro.Fields("Fondo de Reparo"), "#,##0.00")
                    .TextMatrix(i, 10) = Format((rstListadoIcaro.Fields("Fondo de Reparo") + rstBuscarIcaro.Fields("Fondo de Reparo")), "#,##0.00")
                    .TextMatrix(i, 11) = Format((rstListadoIcaro.Fields("Ejecucion") + rstBuscarIcaro.Fields("Ejecucion") + rstListadoIcaro.Fields("Fondo de Reparo") + rstBuscarIcaro.Fields("Fondo de Reparo")), "#,##0.00")
                    .TextMatrix(i, 12) = Format((rstListadoIcaro.Fields("MontoDeContrato") + rstListadoIcaro.Fields("Adicional") - rstListadoIcaro.Fields("Ejecucion") - rstBuscarIcaro.Fields("Ejecucion") - rstListadoIcaro.Fields("Fondo de Reparo") - rstBuscarIcaro.Fields("Fondo de Reparo")), "#,##0.00")
                Else
                    .TextMatrix(i, 5) = Format("0", "#,##0.00")
                    .TextMatrix(i, 6) = Format(rstListadoIcaro.Fields("Ejecucion"), "#,##0.00")
                    .TextMatrix(i, 7) = Format(rstListadoIcaro.Fields("Ejecucion"), "#,##0.00")
                    .TextMatrix(i, 8) = Format("0", "#,##0.00")
                    .TextMatrix(i, 9) = Format(rstListadoIcaro.Fields("Fondo de Reparo"), "#,##0.00")
                    .TextMatrix(i, 10) = Format(rstListadoIcaro.Fields("Fondo de Reparo"), "#,##0.00")
                    .TextMatrix(i, 11) = Format((rstListadoIcaro.Fields("Ejecucion") + rstListadoIcaro.Fields("Fondo de Reparo")), "#,##0.00")
                    .TextMatrix(i, 12) = Format((rstListadoIcaro.Fields("MontoDeContrato") + rstListadoIcaro.Fields("Adicional") - rstListadoIcaro.Fields("Ejecucion") - rstListadoIcaro.Fields("Fondo de Reparo")), "#,##0.00")
                End If
                rstBuscarIcaro.Close
                SQL = "SELECT * " _
                & "FROM Carga " _
                & "WHERE (((Year(CARGA.Fecha)) =" & strEjercicio & ")) And Obra = " & "'" & rstListadoIcaro.Fields("Descripcion") & "' And Avance = 1 " _
                & "ORDER BY Comprobante Desc"
                rstBuscarIcaro.Open SQL, dbIcaro, adOpenDynamic, adLockOptimistic
                .TextMatrix(i, 13) = rstBuscarIcaro.Fields("Comprobante")
                rstBuscarIcaro.Close
                rstListadoIcaro.MoveNext
                .Rows = .Rows + 1
            Wend
            .Rows = .Rows - 1
        End If
    End With
    rstListadoIcaro.Close
    Set rstListadoIcaro = Nothing
    Set rstBuscarIcaro = Nothing
    
End Sub

Public Sub ConfigurardgControlRetenciones()
    
    With ListadoControlRetenciones.dgControlRetenciones
        .Clear
        .Cols = 14
        .Rows = 3
        .MergeCells = flexMergeRestrictColumns
        .MergeRow(0) = True
        .MergeCol(0) = True
        .MergeCol(1) = True
        .MergeCol(2) = True
        .MergeCol(3) = True
        .TextMatrix(0, 0) = "Proveedor"
        .TextMatrix(1, 0) = "Proveedor"
        .ColWidth(0) = 2000
        .TextMatrix(0, 1) = "Obra"
        .TextMatrix(1, 1) = "Obra"
        .ColWidth(1) = 2500
        .TextMatrix(0, 2) = "Mes"
        .TextMatrix(1, 2) = "Mes"
        .ColWidth(2) = 500
        .TextMatrix(0, 3) = "Nº CyO"
        .TextMatrix(1, 3) = "Nº CyO"
        .ColWidth(3) = 700
        .TextMatrix(0, 4) = "SUSS Const."
        .TextMatrix(1, 4) = "Retenido"
        .ColWidth(4) = 800
        .TextMatrix(0, 5) = "SUSS Const."
        .TextMatrix(1, 5) = "Calculado"
        .ColWidth(5) = 800
        .TextMatrix(0, 6) = "Ganancias"
        .TextMatrix(1, 6) = "Retenido"
        .ColWidth(6) = 800
        .TextMatrix(0, 7) = "Ganancias"
        .TextMatrix(1, 7) = "Calculado"
        .ColWidth(7) = 800
        .TextMatrix(0, 8) = "IIBB"
        .TextMatrix(1, 8) = "Retenido"
        .ColWidth(8) = 800
        .TextMatrix(0, 9) = "IIBB"
        .TextMatrix(1, 9) = "Calculado"
        .ColWidth(9) = 800
        .TextMatrix(0, 10) = "Sellos"
        .TextMatrix(1, 10) = "Retenido"
        .ColWidth(10) = 800
        .TextMatrix(0, 11) = "Sellos"
        .TextMatrix(1, 11) = "Calculado"
        .ColWidth(11) = 800
        .TextMatrix(0, 12) = "Tasa L.P."
        .TextMatrix(1, 12) = "Retenido"
        .ColWidth(12) = 800
        .TextMatrix(0, 13) = "Tasa L.P."
        .TextMatrix(1, 13) = "Calculado"
        .ColWidth(13) = 800
        .FixedCols = 0
        .FixedRows = 2
        .FocusRect = flexFocusHeavy
        .HighLight = flexHighlightWithFocus
        .SelectionMode = flexSelectionByRow
        .AllowUserResizing = flexResizeColumns
        .ColAlignmentFixed(0) = 4
        .ColAlignmentFixed(1) = 4
        .ColAlignmentFixed(2) = 4
        .ColAlignmentFixed(3) = 4
        .ColAlignmentFixed(4) = 4
        .ColAlignmentFixed(5) = 4
        .ColAlignmentFixed(6) = 4
        .ColAlignmentFixed(7) = 4
        .ColAlignmentFixed(8) = 4
        .ColAlignmentFixed(9) = 4
        .ColAlignmentFixed(10) = 4
        .ColAlignmentFixed(11) = 4
        .ColAlignmentFixed(12) = 4
        .ColAlignmentFixed(13) = 4
        .ColAlignment(0) = 1
        .ColAlignment(1) = 1
        .ColAlignment(2) = 1
        .ColAlignment(3) = 1
    End With
    
End Sub

Sub CargardgControlRetenciones(Proveedor As String, Ejercicio As String)
    
    Dim i As Long
    i = 1
    Dim SQL As String
    
    ConfigurardgControlRetenciones
    
    Set dbIcaro = New ADODB.Connection
    dbIcaro.CursorLocation = adUseClient
    dbIcaro.Mode = adModeUnknown
    dbIcaro.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\ICARO.mdb"
    dbIcaro.Open

    With ListadoControlRetenciones.dgControlRetenciones
        .Rows = 3
        Set rstRegistroIcaro = New ADODB.Recordset
        If Proveedor = "Todos los Proveedores" Then
            SQL = "Select CARGA.Comprobante, CARGA.Fecha, CARGA.Importe, CARGA.Obra, CARGA.CUIT, CUIT.Nombre, ACTIVIDADES.TipoDeObra, IMPUTACION.Partida From CUIT INNER JOIN (CARGA INNER JOIN (ACTIVIDADES INNER JOIN IMPUTACION ON ACTIVIDADES.Actividad = IMPUTACION.Imputacion) ON CARGA.Comprobante = IMPUTACION.Comprobante) ON CUIT.CUIT = CARGA.CUIT Where Year(Fecha) = '" & Ejercicio & "' And CUIT.CUIT NOT LIKE '30632351514' Order By CARGA.Comprobante"
            rstRegistroIcaro.Open SQL, dbIcaro, adOpenForwardOnly, adLockOptimistic
        Else
            SQL = "Select CARGA.Comprobante, CARGA.Fecha, CARGA.Importe, CARGA.Obra, CARGA.CUIT, CUIT.Nombre, ACTIVIDADES.TipoDeObra, IMPUTACION.Partida From CUIT INNER JOIN (CARGA INNER JOIN (ACTIVIDADES INNER JOIN IMPUTACION ON ACTIVIDADES.Actividad = IMPUTACION.Imputacion) ON CARGA.Comprobante = IMPUTACION.Comprobante) ON CUIT.CUIT = CARGA.CUIT Where Year(Fecha) = '" & Ejercicio & "' And Nombre = '" & Proveedor & "' Order By CARGA.Comprobante"
            rstRegistroIcaro.Open SQL, dbIcaro, adOpenForwardOnly, adLockOptimistic
        End If
        If rstRegistroIcaro.BOF = False Then
                Set rstRegistroICARO2 = New ADODB.Recordset
                rstRegistroIcaro.MoveFirst
                While rstRegistroIcaro.EOF = False
                    i = i + 1
                    .RowHeight(i) = 300
                    .TextMatrix(i, 0) = rstRegistroIcaro!Nombre
                    .TextMatrix(i, 1) = rstRegistroIcaro!Obra
                    .TextMatrix(i, 2) = Month(rstRegistroIcaro!Fecha)
                    .TextMatrix(i, 3) = rstRegistroIcaro!Comprobante
                    .TextMatrix(i, 4) = Format("0", "0.00")
                    .TextMatrix(i, 5) = Format("0", "0.00")
                    .TextMatrix(i, 6) = Format("0", "0.00")
                    .TextMatrix(i, 7) = Format("0", "0.00")
                    .TextMatrix(i, 8) = Format("0", "0.00")
                    .TextMatrix(i, 9) = Format("0", "0.00")
                    .TextMatrix(i, 10) = Format("0", "0.00")
                    .TextMatrix(i, 11) = Format("0", "0.00")
                    .TextMatrix(i, 12) = Format("0", "0.00")
                    .TextMatrix(i, 13) = Format("0", "0.00")
                    SQL = "Select * From RETENCIONES Where Comprobante = '" & rstRegistroIcaro!Comprobante & "' Order by Codigo"
                    rstRegistroICARO2.Open SQL, dbIcaro, adOpenForwardOnly, adLockOptimistic
                    If rstRegistroICARO2.BOF = False Then 'Cargar Retenciones Practicadas
                        rstRegistroICARO2.MoveFirst
                        While rstRegistroICARO2.EOF = False
                            Select Case rstRegistroICARO2!Codigo
                            Case "114"
                                .TextMatrix(i, 4) = Format(rstRegistroICARO2!Importe, "#0.00")
                            'Case "117"
                                '.TextMatrix(i, 4) = Format(rstRegistroICARO2!Importe, "#0.00")
                            Case "106"
                                .TextMatrix(i, 6) = Format(rstRegistroICARO2!Importe, "#0.00")
                            Case "113"
                                .TextMatrix(i, 6) = Format(rstRegistroICARO2!Importe, "#0.00")
                            Case "101"
                                .TextMatrix(i, 8) = Format(rstRegistroICARO2!Importe, "#0.00")
                            Case "110"
                                .TextMatrix(i, 8) = Format(rstRegistroICARO2!Importe, "#0.00")
                            Case "102"
                                .TextMatrix(i, 10) = Format(rstRegistroICARO2!Importe, "#0.00")
                            Case "111"
                                .TextMatrix(i, 10) = Format(rstRegistroICARO2!Importe, "#0.00")
                            Case "104"
                                .TextMatrix(i, 12) = Format(rstRegistroICARO2!Importe, "#0.00")
                            Case "112"
                                .TextMatrix(i, 12) = Format(rstRegistroICARO2!Importe, "#0.00")
                            End Select
                            rstRegistroICARO2.MoveNext
                        Wend
                    End If
                    rstRegistroICARO2.Close
                    .TextMatrix(i, 5) = Format(CalcularSUSS("CONSTRUCCION", rstRegistroIcaro!CUIT, rstRegistroIcaro!Fecha, rstRegistroIcaro!Importe, rstRegistroIcaro!TipoDeObra, rstRegistroIcaro!Partida, rstRegistroIcaro!Comprobante), "0.00")
                    .TextMatrix(i, 7) = Format(CalcularGanancias("LOCACION", rstRegistroIcaro!CUIT, rstRegistroIcaro!Fecha, rstRegistroIcaro!Importe, rstRegistroIcaro!TipoDeObra, rstRegistroIcaro!Partida, rstRegistroIcaro!Comprobante), "0.00")
                    rstRegistroIcaro.MoveNext
                    .Rows = .Rows + 1
                Wend
            Set rstRegistroICARO2 = Nothing
            .Rows = .Rows - 1
        End If
    End With
    rstRegistroIcaro.Close
    Set rstRegistroIcaro = Nothing
    dbIcaro.Close
    Set dbIcaro = Nothing
    
End Sub

Public Sub ConfigurardgTipoDeObra()

    With CargaTipoDeObra.dgReparaciones
        .Clear
        .Cols = 2
        .Rows = 2
        .TextMatrix(0, 0) = "Código"
        .ColWidth(0) = 1000
        .TextMatrix(0, 1) = "Descripción"
        .ColWidth(1) = 3500
        .FixedCols = 0
        .FocusRect = flexFocusHeavy
        .HighLight = flexHighlightWithFocus
        .SelectionMode = flexSelectionByRow
        .AllowUserResizing = flexResizeColumns
        .ColAlignment(0) = 1
        .ColAlignment(1) = 1
    End With
    
    With CargaTipoDeObra.dgNoCategorizadas
        .Clear
        .Cols = 2
        .Rows = 2
        .TextMatrix(0, 0) = "Código"
        .ColWidth(0) = 1000
        .TextMatrix(0, 1) = "Descripción"
        .ColWidth(1) = 3500
        .FixedCols = 0
        .FocusRect = flexFocusHeavy
        .HighLight = flexHighlightWithFocus
        .SelectionMode = flexSelectionByRow
        .AllowUserResizing = flexResizeColumns
        .ColAlignment(0) = 1
        .ColAlignment(1) = 1
    End With
    
    With CargaTipoDeObra.dgViviendas
        .Clear
        .Cols = 2
        .Rows = 2
        .TextMatrix(0, 0) = "Código"
        .ColWidth(0) = 1000
        .TextMatrix(0, 1) = "Descripción"
        .ColWidth(1) = 3500
        .FixedCols = 0
        .FocusRect = flexFocusHeavy
        .HighLight = flexHighlightWithFocus
        .SelectionMode = flexSelectionByRow
        .AllowUserResizing = flexResizeColumns
        .ColAlignment(0) = 1
        .ColAlignment(1) = 1
    End With

End Sub

Public Sub CargardgTipoDeObra()

    Dim i As Integer
    Dim SQL As String
    i = 0
    
    
    Set dbIcaro = New ADODB.Connection
    dbIcaro.CursorLocation = adUseClient
    dbIcaro.Mode = adModeUnknown
    dbIcaro.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\ICARO.mdb"
    dbIcaro.Open
    Set rstListadoIcaro = New ADODB.Recordset
    
    SQL = "Select * From ACTIVIDADES Where TipoDeObra = 'R' Order By Actividad"
    rstListadoIcaro.Open SQL, dbIcaro, adOpenDynamic, adLockOptimistic
    CargaTipoDeObra.dgReparaciones.Rows = 2
    With CargaTipoDeObra.dgReparaciones
        If rstListadoIcaro.BOF = False Then
            rstListadoIcaro.MoveFirst
            While rstListadoIcaro.EOF = False
                i = i + 1
                .RowHeight(i) = 300
                .TextMatrix(i, 0) = rstListadoIcaro!Actividad
                If Len(rstListadoIcaro!Descripcion) <> "0" Then
                    .TextMatrix(i, 1) = rstListadoIcaro!Descripcion
                Else
                    .TextMatrix(i, 1) = ""
                End If
                rstListadoIcaro.MoveNext
                .Rows = .Rows + 1
            Wend
        End If
        .Rows = .Rows - 1
    End With
    rstListadoIcaro.Close
    i = 0
    
    SQL = "Select * From ACTIVIDADES Where IsNull(TipoDeObra) = True Order By Actividad"
    rstListadoIcaro.Open SQL, dbIcaro, adOpenDynamic, adLockOptimistic
    CargaTipoDeObra.dgNoCategorizadas.Rows = 2
    With CargaTipoDeObra.dgNoCategorizadas
        If rstListadoIcaro.BOF = False Then
            rstListadoIcaro.MoveFirst
            While rstListadoIcaro.EOF = False
                i = i + 1
                .RowHeight(i) = 300
                .TextMatrix(i, 0) = rstListadoIcaro!Actividad
                If Len(rstListadoIcaro!Descripcion) <> "0" Then
                    .TextMatrix(i, 1) = rstListadoIcaro!Descripcion
                Else
                    .TextMatrix(i, 1) = ""
                End If
                rstListadoIcaro.MoveNext
                .Rows = .Rows + 1
            Wend
        End If
        .Rows = .Rows - 1
    End With
    rstListadoIcaro.Close
    i = 0
    
    SQL = "Select * From ACTIVIDADES Where TipoDeObra = 'V' Order By Actividad"
    rstListadoIcaro.Open SQL, dbIcaro, adOpenDynamic, adLockOptimistic
    CargaTipoDeObra.dgViviendas.Rows = 2
    With CargaTipoDeObra.dgViviendas
        If rstListadoIcaro.BOF = False Then
            rstListadoIcaro.MoveFirst
            While rstListadoIcaro.EOF = False
                i = i + 1
                .RowHeight(i) = 300
                .TextMatrix(i, 0) = rstListadoIcaro!Actividad
                If Len(rstListadoIcaro!Descripcion) <> "0" Then
                    .TextMatrix(i, 1) = rstListadoIcaro!Descripcion
                Else
                    .TextMatrix(i, 1) = ""
                End If
                rstListadoIcaro.MoveNext
                .Rows = .Rows + 1
            Wend
        End If
        .Rows = .Rows - 1
    End With
    rstListadoIcaro.Close
    i = 0
    
    Set rstListadoIcaro = Nothing
    dbIcaro.Close
    Set dbIcaro = Nothing

End Sub

Sub ConfigurardgEjecucionFuente10y13()
    
    Dim x As Integer
    x = 1
    With ListadoControlFuente10y13.dgEjecucionFuente10y13
        .Clear
        .Cols = 7
        .Rows = 2
        .TextMatrix(0, 0) = "Decreto"
        .ColWidth(0) = 1000
        .TextMatrix(0, 1) = "Fuente"
        .ColWidth(1) = 600
        .TextMatrix(0, 2) = "Total"
        .ColWidth(2) = 1000
        .TextMatrix(0, 3) = "2010"
        .ColWidth(3) = 1000
        .TextMatrix(0, 4) = "2011"
        .ColWidth(4) = 1000
        .TextMatrix(0, 5) = "2012"
        .ColWidth(5) = 1000
        .TextMatrix(0, 6) = "2013"
        .ColWidth(6) = 1000
        .FocusRect = flexFocusHeavy
        .HighLight = flexHighlightWithFocus
        .AllowUserResizing = flexResizeColumns
        .FixedCols = 0
        .ColAlignment(0) = 1
        .ColAlignment(1) = 7
        .ColAlignment(2) = 7
        .ColAlignment(3) = 7
        .ColAlignment(4) = 7
        .ColAlignment(5) = 7
        .ColAlignment(6) = 7
    End With
    
End Sub

Public Sub CargardgEjecucionPorFuente10y13()

    Dim i As Integer
    Dim SQL As String
    i = 0
    
    
    Set dbIcaro = New ADODB.Connection
    dbIcaro.CursorLocation = adUseClient
    dbIcaro.Mode = adModeUnknown
    dbIcaro.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\ICARO.mdb"
    dbIcaro.Open
    Set rstListadoIcaro = New ADODB.Recordset
    
    SQL = "Select * From CARGA Where TipoDeObra = 'R' Order By Actividad"
    rstListadoIcaro.Open SQL, dbIcaro, adOpenDynamic, adLockOptimistic
    CargaTipoDeObra.dgReparaciones.Rows = 2
    With CargaTipoDeObra.dgReparaciones
        If rstListadoIcaro.BOF = False Then
            rstListadoIcaro.MoveFirst
            While rstListadoIcaro.EOF = False
                i = i + 1
                .RowHeight(i) = 300
                .TextMatrix(i, 0) = rstListadoIcaro!Actividad
                If Len(rstListadoIcaro!Descripcion) <> "0" Then
                    .TextMatrix(i, 1) = rstListadoIcaro!Descripcion
                Else
                    .TextMatrix(i, 1) = ""
                End If
                rstListadoIcaro.MoveNext
                .Rows = .Rows + 1
            Wend
        End If
        .Rows = .Rows - 1
    End With
    rstListadoIcaro.Close
    i = 0
    Set rstListadoIcaro = Nothing
    dbIcaro.Close
    Set dbIcaro = Nothing
    
End Sub

Public Sub ConfigurarSincronizaciondgObrasSGF()
    
    With SincronizacionDescripcionSGF.dgObrasSGF
        .Clear
        .Cols = 3
        .Rows = 2
        .TextMatrix(0, 0) = "Proveedor"
        .TextMatrix(0, 1) = "Obra"
        .TextMatrix(0, 2) = "Período"
        .ColWidth(0) = 3000
        .ColWidth(1) = 8500
        .ColWidth(2) = 1500
        .FixedCols = 0
        .FocusRect = flexFocusHeavy
        .HighLight = flexHighlightWithFocus
        .SelectionMode = flexSelectionByRow
        .AllowUserResizing = flexResizeColumns
        .ColAlignment(0) = 1
        .ColAlignment(1) = 1
        .ColAlignment(1) = 2
    End With
    
End Sub

Public Sub CargarSincronizaciondgObrasSGF()
    
    Dim i As Integer
    Dim SQL As String
    i = 0
    
    SincronizacionDescripcionSGF.dgObrasSGF.Rows = 2
    Set rstListadoIcaro = New ADODB.Recordset
    SQL = "Select Proveedor, Obra, Last(Periodo) As Periodo From CERTIFICADOS " _
        & "Where Obra Not In (Select Descripcion From Obras)" _
        & "Group by Obra, Proveedor Order by Proveedor"
    rstListadoIcaro.Open SQL, dbIcaro, adOpenDynamic, adLockOptimistic
    With SincronizacionDescripcionSGF.dgObrasSGF
        If rstListadoIcaro.BOF = False Then
            rstListadoIcaro.MoveFirst
            While rstListadoIcaro.EOF = False
                i = i + 1
                .RowHeight(i) = 300
                .TextMatrix(i, 0) = rstListadoIcaro!Proveedor
                .TextMatrix(i, 1) = rstListadoIcaro!Obra
                .TextMatrix(i, 2) = Format(Month(rstListadoIcaro!Periodo), "00") & "/" & Year(rstListadoIcaro!Periodo)
                rstListadoIcaro.MoveNext
                .Rows = .Rows + 1
            Wend
        End If
        .Rows = .Rows - 1
    End With
    
    rstListadoIcaro.Close
    Set rstListadoIcaro = Nothing
    
End Sub

Public Sub ConfigurarSincronizaciondgObrasICARO()
    
    With SincronizacionDescripcionSGF.dgObrasICARO
        .Clear
        .Cols = 3
        .Rows = 2
        .TextMatrix(0, 0) = "Proveedor"
        .TextMatrix(0, 1) = "Obra"
        .TextMatrix(0, 2) = "Estructura"
        .ColWidth(0) = 3000
        .ColWidth(1) = 8500
        .ColWidth(2) = 1500
        .FixedCols = 0
        .FocusRect = flexFocusHeavy
        .HighLight = flexHighlightWithFocus
        .SelectionMode = flexSelectionByRow
        .AllowUserResizing = flexResizeColumns
        .ColAlignment(0) = 1
        .ColAlignment(1) = 1
        .ColAlignment(2) = 2
    End With
    
End Sub

Public Sub CargarSincronizaciondgObrasICARO(ObraSGF As String)
    
    Dim i As Integer
    Dim SQL As String
    i = 0
    
    SincronizacionDescripcionSGF.dgObrasICARO.Rows = 2
    Set rstListadoIcaro = New ADODB.Recordset
    SQL = "Select CUIT.Nombre, OBRAS.Descripcion, OBRAS.Imputacion, OBRAS.Partida " _
        & "From CUIT Inner Join OBRAS On CUIT.Cuit = OBRAS.Cuit " _
        & "Where OBRAS.Descripcion Like '" & ObraSGF & "%' " _
        & "Group by OBRAS.Descripcion, CUIT.Nombre, OBRAS.Imputacion, OBRAS.Partida Order by CUIT.Nombre"
    rstListadoIcaro.Open SQL, dbIcaro, adOpenDynamic, adLockOptimistic
    With SincronizacionDescripcionSGF.dgObrasICARO
        If rstListadoIcaro.BOF = False Then
            rstListadoIcaro.MoveFirst
            While rstListadoIcaro.EOF = False
                i = i + 1
                .RowHeight(i) = 300
                .TextMatrix(i, 0) = rstListadoIcaro!Nombre
                .TextMatrix(i, 1) = rstListadoIcaro!Descripcion
                .TextMatrix(i, 2) = rstListadoIcaro!Imputacion & "-" & rstListadoIcaro!Partida
                rstListadoIcaro.MoveNext
                .Rows = .Rows + 1
            Wend
        End If
        .Rows = .Rows - 1
    End With
    
    rstListadoIcaro.Close
    Set rstListadoIcaro = Nothing
    
End Sub

Public Sub CargardgRetencionesPorCtaCte(FechaInicio As String, FechaFin As String)

    Dim intFila As Integer
    Dim intColumna As Integer
    Dim SQL As String
    Dim SQLWhere As String
    Dim intNCampos As Integer
    Dim intNRegistros As Integer
    Dim datFechaInicio As Date
    Dim datFechaFin As Date
    
    Set rstListadoIcaro = New ADODB.Recordset
    datFechaInicio = DateTime.DateSerial(Right(FechaInicio, 4), Mid(FechaInicio, 4, 2), Left(FechaInicio, 2))
    datFechaFin = DateTime.DateSerial(Right(FechaFin, 4), Mid(FechaFin, 4, 2), Left(FechaFin, 2))
    If datFechaInicio > datFechaFin Then
        MsgBox "Debe ingresar una Fecha Inicio inferior a la Fecha Fin", vbCritical + vbOKOnly, "FECHAS INCORRECTAS"
        RetencionesPorCtaCte.txtFechaDesde.SetFocus
        Exit Sub
    Else
        SQLWhere = "CARGA.Fecha >= #" & Format(datFechaInicio, "MM/DD/YYYY") & "# " _
        & "AND CARGA.Fecha <= #" & Format(datFechaFin, "MM/DD/YYYY") & "# "
    End If

'    Select Case Mes
'    Case Is = "1er Semestre"
'        SQLWhere = "Year(CARGA.Fecha) = " & Ano & " And Month(CARGA.Fecha) <= 6 "
'    Case Is = "2do Semestre"
'        SQLWhere = "Year(CARGA.Fecha) = " & Ano & " And Month(CARGA.Fecha) >= 7 "
'    Case Is = "Año Completo"
'        SQLWhere = "Year(CARGA.Fecha) = " & Ano & " "
'    Case Else
'        SQLWhere = "Year(CARGA.Fecha) = " & Ano & " And Month(CARGA.Fecha) = " & Mes & " "
'    End Select
    SQL = "TRANSFORM Sum(RETENCIONES.Importe) AS SumaDeImporte " _
    & "SELECT CARGA.Cuenta, Sum(RETENCIONES.Importe) AS [Total Retenciones] " _
    & "FROM CODIGOSRETENCIONES INNER JOIN (CARGA INNER JOIN RETENCIONES ON CARGA.Comprobante = RETENCIONES.Comprobante) ON CODIGOSRETENCIONES.Codigo = RETENCIONES.Codigo " _
    & "Where " & SQLWhere _
    & "GROUP BY CARGA.Cuenta " _
    & "PIVOT CODIGOSRETENCIONES.Codigo"
    rstListadoIcaro.Open SQL, dbIcaro, adOpenDynamic, adLockOptimistic
    
'    Debug.Print "N° Campos: " & rstListadoIcaro.Fields.Count
'    Debug.Print "N° Registros: " & rstListadoIcaro.RecordCount
    
    intNRegistros = rstListadoIcaro.RecordCount
    
    With RetencionesPorCtaCte.dgRetencionesPorCtaCte
        'Verificamos que la consulta tenga registros
        If intNRegistros > 0 Then
            'Procedemos a dar formato a la tabla
            intNCampos = rstListadoIcaro.Fields.Count
            .Clear
            .Cols = intNCampos + 1
            .Rows = 2
            For intColumna = 0 To (intNCampos - 1)
                .TextMatrix(0, intColumna) = rstListadoIcaro.Fields.Item(intColumna).Name
                .ColWidth(intColumna) = 1400
                .ColAlignment(intColumna) = 7
            Next
            .ColAlignment(0) = 1
            .TextMatrix(0, intNCampos) = "Total Bruto"
            .ColAlignment(intNCampos) = 7
            .ColWidth(intNCampos) = 1400
            .FixedCols = 0
            .FocusRect = flexFocusHeavy
            .HighLight = flexHighlightWithFocus
            .SelectionMode = flexSelectionByRow
            .AllowUserResizing = flexResizeColumns
            'Procedemos a rellenar la tabla
            intFila = 0
            rstListadoIcaro.MoveFirst
            While rstListadoIcaro.EOF = False
                intFila = intFila + 1
                .RowHeight(intFila) = 300
                For intColumna = 0 To (intNCampos - 1)
                    If Len(rstListadoIcaro.Fields.Item(intColumna)) <> "0" Then
                        .TextMatrix(intFila, intColumna) = rstListadoIcaro.Fields.Item(intColumna)
                    Else
                        .TextMatrix(intFila, intColumna) = "0"
                    End If
                Next
                Set rstBuscarIcaro = New ADODB.Recordset
                SQL = "Select CARGA.Cuenta, Sum(CARGA.Importe) AS TotalBruto From CARGA " _
                & "Where " & SQLWhere & "And CARGA.Cuenta = '" & rstListadoIcaro.Fields.Item(0) & "' " _
                & "Group by CARGA.Cuenta"
                rstBuscarIcaro.Open SQL, dbIcaro, adOpenDynamic, adLockOptimistic
                .TextMatrix(intFila, intNCampos) = rstBuscarIcaro!TotalBruto
                rstBuscarIcaro.Close
                Set rstBuscarIcaro = Nothing
                rstListadoIcaro.MoveNext
                .Rows = .Rows + 1
            Wend
            '.Rows = .Rows - 1
            'Completamos la última fila de TOTAL
            intFila = intFila + 1
            Set rstBuscarIcaro = New ADODB.Recordset
            SQL = "Select Sum(RETENCIONES.Importe) AS TotalRetenciones " _
            & "From CARGA Inner Join RETENCIONES " _
            & "On CARGA.Comprobante = RETENCIONES.Comprobante " _
            & "Where " & SQLWhere
            rstBuscarIcaro.Open SQL, dbIcaro, adOpenDynamic, adLockOptimistic
            .TextMatrix(intFila, 0) = "Total Período"
            .TextMatrix(intFila, 1) = rstBuscarIcaro!TotalRetenciones
            rstBuscarIcaro.Close
            'Completamos los subtotales por codigo de retencion
            SQL = "Select RETENCIONES.Codigo, Sum(RETENCIONES.Importe) As TotalRetenciones " _
            & "From CARGA Inner Join RETENCIONES " _
            & "On CARGA.Comprobante = RETENCIONES.Comprobante " _
            & "Where " & SQLWhere _
            & "Group by RETENCIONES.Codigo "
            rstBuscarIcaro.Open SQL, dbIcaro, adOpenDynamic, adLockOptimistic
            rstBuscarIcaro.MoveFirst
            intColumna = 1
            While rstBuscarIcaro.EOF = False
                intColumna = intColumna + 1
                .TextMatrix(intFila, intColumna) = rstBuscarIcaro!TotalRetenciones
                rstBuscarIcaro.MoveNext
            Wend
            rstBuscarIcaro.Close
            Set rstBuscarIcaro = Nothing
        Else
            'Procedemos a limpiar la tabla
            .Clear
            .Rows = 2
            RetencionesPorCtaCte.dgRetencionesCtaCteSeleccionada.Clear
            RetencionesPorCtaCte.dgRetencionesCtaCteSeleccionada.Rows = 2
        End If
    End With
    
    rstListadoIcaro.Close
    Set rstListadoIcaro = Nothing
    
    With RetencionesPorCtaCte.dgRetencionesPorCtaCte
        If .TextMatrix(1, 0) <> "" Then
            Call CargardgRetencionesCtaCteSeleccionada(FechaInicio, FechaFin, _
            .TextMatrix(1, 0))
            RetencionesPorCtaCte.frmCtaCteSeleccionada.Caption = "Datos desagregados " _
            & "de la Cuenta Corriente " & .TextMatrix(1, 0)
        End If
    End With
    
End Sub

Public Sub CargardgRetencionesCtaCteSeleccionada(FechaInicio As String, _
FechaFin As String, CuentaCorriente As String)

    Dim intFila As Integer
    Dim intColumna As Integer
    Dim SQL As String
    Dim SQLRefCruzada As String
    Dim SQLWhere As String
    Dim SQLPivot As String
    Dim intNCampos As Integer
    Dim intNRegistros As Integer
    Dim datFechaInicio As Date
    Dim datFechaFin As Date
    
    Set rstListadoIcaro = New ADODB.Recordset
    datFechaInicio = DateTime.DateSerial(Right(FechaInicio, 4), Mid(FechaInicio, 4, 2), Left(FechaInicio, 2))
    datFechaFin = DateTime.DateSerial(Right(FechaFin, 4), Mid(FechaFin, 4, 2), Left(FechaFin, 2))
    If datFechaInicio > datFechaFin Then
        MsgBox "Debe ingresar una Fecha Inicio inferior a la Fecha Fin", vbCritical + vbOKOnly, "FECHAS INCORRECTAS"
        RetencionesPorCtaCte.txtFechaDesde.SetFocus
        Exit Sub
    Else
        If Left(CuentaCorriente, 5) = "Total" Then
            SQLWhere = "CARGA.Fecha >= #" & Format(datFechaInicio, "MM/DD/YYYY") & "# " _
            & "AND CARGA.Fecha <= #" & Format(datFechaFin, "MM/DD/YYYY") & "# "
        Else
            SQLWhere = "CARGA.Fecha >= #" & Format(datFechaInicio, "MM/DD/YYYY") & "# " _
            & "AND CARGA.Fecha <= #" & Format(datFechaFin, "MM/DD/YYYY") & "# " _
            & "AND CARGA.Cuenta = '" & CuentaCorriente & "' "
        End If
        SQLPivot = "CODIGOSRETENCIONES.Codigo"
        With RetencionesPorCtaCte.dgRetencionesPorCtaCte
            SQLPivot = SQLPivot & " In ("
            For intColumna = 2 To (.Cols - 2)
                SQLPivot = SQLPivot & "'" & .TextMatrix(0, intColumna) & "', "
            Next intColumna
            SQLPivot = Left(SQLPivot, Len(SQLPivot) - 2)
            SQLPivot = SQLPivot & ")"
        End With
    End If
    
    SQLRefCruzada = "TRANSFORM Sum(RETENCIONES.Importe) AS SumaDeImporte " _
    & "SELECT CARGA.Comprobante, Sum(RETENCIONES.Importe) AS [Total Retenciones] " _
    & "FROM CODIGOSRETENCIONES INNER JOIN (CARGA INNER JOIN RETENCIONES ON CARGA.Comprobante = RETENCIONES.Comprobante) ON CODIGOSRETENCIONES.Codigo = RETENCIONES.Codigo " _
    & "Where " & SQLWhere _
    & "GROUP BY CARGA.Comprobante " _
    & "ORDER BY CARGA.Comprobante Asc " _
    & "PIVOT " & SQLPivot
    rstListadoIcaro.Open SQLRefCruzada, dbIcaro, adOpenDynamic, adLockOptimistic
    
'    Debug.Print "N° Campos: " & rstListadoIcaro.Fields.Count
'    Debug.Print "N° Registros: " & rstListadoIcaro.RecordCount
    
    intNRegistros = rstListadoIcaro.RecordCount
    
    With RetencionesPorCtaCte.dgRetencionesCtaCteSeleccionada
        'Verificamos que la consulta tenga registros
        If intNRegistros > 0 Then
            'Procedemos a dar formato a la tabla
            intNCampos = rstListadoIcaro.Fields.Count
            .Clear
            .Cols = intNCampos + 4
            .Rows = 2
            For intColumna = 0 To (intNCampos - 1)
                .TextMatrix(0, intColumna) = rstListadoIcaro.Fields.Item(intColumna).Name
                .ColWidth(intColumna) = 1400
                .ColAlignment(intColumna) = 7
            Next
            .ColAlignment(0) = 1
            .TextMatrix(0, intNCampos) = "Total Bruto"
            .ColAlignment(intNCampos) = 7
            .ColWidth(intNCampos) = 1400
            .TextMatrix(0, intNCampos + 1) = "Total Neto"
            .ColAlignment(intNCampos + 1) = 7
            .ColWidth(intNCampos + 1) = 1400
            .TextMatrix(0, intNCampos + 2) = "Proveedor/Contratista"
            .ColAlignment(intNCampos + 2) = 1
            .ColWidth(intNCampos + 2) = 2800
            .TextMatrix(0, intNCampos + 3) = "Fecha"
            .ColAlignment(intNCampos + 3) = 1
            .ColWidth(intNCampos + 3) = 1400
            .FixedCols = 0
            .FocusRect = flexFocusHeavy
            .HighLight = flexHighlightWithFocus
            .SelectionMode = flexSelectionByRow
            .AllowUserResizing = flexResizeColumns
            'Procedemos a rellenar la tabla
            intFila = 0
            rstListadoIcaro.MoveFirst
            While rstListadoIcaro.EOF = False
                intFila = intFila + 1
                .RowHeight(intFila) = 300
                For intColumna = 0 To (intNCampos - 1)
                    If Len(rstListadoIcaro.Fields.Item(intColumna)) <> "0" Then
                        .TextMatrix(intFila, intColumna) = rstListadoIcaro.Fields.Item(intColumna)
                    Else
                        .TextMatrix(intFila, intColumna) = "0"
                    End If
                Next
                Set rstBuscarIcaro = New ADODB.Recordset
                SQL = "Select CARGA.Importe, CUIT.Nombre, CARGA.Fecha " _
                & "From CARGA INNER JOIN CUIT " _
                & "ON CARGA.CUIT = CUIT.CUIT " _
                & "Where CARGA.Comprobante = '" & .TextMatrix(intFila, 0) & "'"
                rstBuscarIcaro.Open SQL, dbIcaro, adOpenDynamic, adLockOptimistic
                .TextMatrix(intFila, intNCampos) = rstBuscarIcaro!Importe
                .TextMatrix(intFila, intNCampos + 1) = rstBuscarIcaro!Importe - Val(.TextMatrix(intFila, 1))
                .TextMatrix(intFila, intNCampos + 2) = rstBuscarIcaro!Nombre
                .TextMatrix(intFila, intNCampos + 3) = rstBuscarIcaro!Fecha
                rstBuscarIcaro.Close
                Set rstBuscarIcaro = Nothing
                rstListadoIcaro.MoveNext
                .Rows = .Rows + 1
            Wend
            .Rows = .Rows - 1
        Else
            'Procedemos a limpiar la tabla
            .Clear
            .Rows = 2
        End If
        'Cargamos los Pagos a Cuenta (Sin Retenciones) - ¿Qué pasa si no hay Pagos con retenciones en el mes?
        rstListadoIcaro.Close
        SQLRefCruzada = "SELECT RETENCIONES.Comprobante " _
            & "FROM RETENCIONES"
        SQL = "SELECT CARGA.Comprobante, CARGA.Importe, Carga.Fecha, CUIT.Nombre " _
            & "FROM CARGA INNER JOIN CUIT ON CARGA.CUIT = CUIT.CUIT " _
            & "WHERE Carga.Comprobante Not In (" & SQLRefCruzada & ") AND " & SQLWhere _
            & "ORDER BY CARGA.Comprobante Asc"
        rstListadoIcaro.Open SQL, dbIcaro, adOpenDynamic, adLockOptimistic
        If rstListadoIcaro.BOF = False Then
            intNCampos = RetencionesPorCtaCte.dgRetencionesPorCtaCte.Cols
            rstListadoIcaro.MoveFirst
            .Rows = .Rows + 1
            While rstListadoIcaro.EOF = False
                intFila = intFila + 1
                .RowHeight(intFila) = 300
                .TextMatrix(intFila, 0) = rstListadoIcaro!Comprobante
                For intColumna = 1 To (intNCampos - 2)
                    .TextMatrix(intFila, intColumna) = 0
                Next intColumna
                intColumna = (intNCampos - 1)
                .TextMatrix(intFila, intColumna) = rstListadoIcaro!Importe
                intColumna = intColumna + 1
                .TextMatrix(intFila, intColumna) = rstListadoIcaro!Importe
                intColumna = intColumna + 1
                .TextMatrix(intFila, intColumna) = rstListadoIcaro!Nombre
                intColumna = intColumna + 1
                .TextMatrix(intFila, intColumna) = rstListadoIcaro!Fecha
                rstListadoIcaro.MoveNext
                .Rows = .Rows + 1
            Wend
            .Rows = .Rows - 1
        End If
    End With
    
    rstListadoIcaro.Close
    Set rstListadoIcaro = Nothing
    
End Sub


Public Sub ConfigurardgEjecucionPorImputacion()
    
    With EjecucionPorImputacion.dgEjecucionPorImputacion
        .Clear
        .Cols = 5
        .Rows = 2
        .TextMatrix(0, 0) = "Código"
        .TextMatrix(0, 1) = "Imputación"
        .TextMatrix(0, 2) = "Monto Neto"
        .TextMatrix(0, 3) = "Retenciones"
        .TextMatrix(0, 4) = "Monto Bruto"
        .ColWidth(0) = 1400
        .ColWidth(1) = 4200
        .ColWidth(2) = 1400
        .ColWidth(3) = 1400
        .ColWidth(4) = 1400
        .FixedCols = 0
        .FocusRect = flexFocusHeavy
        .HighLight = flexHighlightWithFocus
        .SelectionMode = flexSelectionByRow
        .AllowUserResizing = flexResizeColumns
        .ColAlignment(0) = 1
        .ColAlignment(1) = 1
        .ColAlignment(2) = 7
        .ColAlignment(3) = 7
        .ColAlignment(4) = 7
    End With
    
End Sub


Public Sub CargardgEjecucionPorImputacion(FechaInicio As String, FechaFin As String)

    Dim intFila As Integer
    Dim SQL As String
    Dim SQLDate As String
    Dim SQLImputacion As String
    Dim dblImporteBruto As Double
    Dim dblRetenciones As Double
    Dim datFechaInicio As Date
    Dim datFechaFin As Date
    
    
    datFechaInicio = DateTime.DateSerial(Right(FechaInicio, 4), Mid(FechaInicio, 4, 2), Left(FechaInicio, 2))
    datFechaFin = DateTime.DateSerial(Right(FechaFin, 4), Mid(FechaFin, 4, 2), Left(FechaFin, 2))
    If datFechaInicio > datFechaFin Then
        MsgBox "Debe ingresar una Fecha Inicio inferior a la Fecha Fin", vbCritical + vbOKOnly, "FECHAS INCORRECTAS"
        EjecucionPorImputacion.txtFechaDesde.SetFocus
        Exit Sub
    Else
        Set rstListadoIcaro = New ADODB.Recordset
        SQLDate = "CARGA.Fecha >= #" & Format(datFechaInicio, "MM/DD/YYYY") & "# " _
        & "AND CARGA.Fecha <= #" & Format(datFechaFin, "MM/DD/YYYY") & "# "
    End If
    
    
    With EjecucionPorImputacion.dgEjecucionPorImputacion
        'Procedemos a rellenar la tabla renglón por renglón sin bucle (AGREGAR CONTROL DE NULL)
        intFila = 0
        'Primer Registro
        intFila = intFila + 1
        .RowHeight(intFila) = 300
        .TextMatrix(intFila, 0) = "CILP"
        .TextMatrix(intFila, 1) = "CREDITOS INDIVIDUALES - LOTE PROPIO"
        SQLImputacion = "(" & SQLDate & "And IMPUTACION.Imputacion = '" & "11-00-02-79" & "') " _
        & "Or (" & SQLDate & "And IMPUTACION.Imputacion = '" & "11-00-43-52" & "')"
        SQL = "Select Sum(CARGA.Importe) AS TotalImporte " _
        & "From CARGA Inner Join IMPUTACION " _
        & "On CARGA.Comprobante = IMPUTACION.Comprobante " _
        & "Where " & SQLImputacion
        rstListadoIcaro.Open SQL, dbIcaro, adOpenDynamic, adLockOptimistic
        If Len(rstListadoIcaro!TotalImporte) <> "0" Then
            dblImporteBruto = rstListadoIcaro!TotalImporte
        Else
            dblImporteBruto = 0
        End If
        rstListadoIcaro.Close
        .TextMatrix(intFila, 4) = FormatNumber(dblImporteBruto, 2)
        SQL = "Select Sum(RETENCIONES.Importe) AS TotalImporte " _
        & "From (CARGA Inner Join IMPUTACION ON CARGA.Comprobante = IMPUTACION.Comprobante) " _
        & "Inner Join RETENCIONES ON CARGA.Comprobante = RETENCIONES.Comprobante " _
        & "Where " & SQLImputacion
        rstListadoIcaro.Open SQL, dbIcaro, adOpenDynamic, adLockOptimistic
        If Len(rstListadoIcaro!TotalImporte) <> "0" Then
            dblRetenciones = rstListadoIcaro!TotalImporte
        Else
            dblRetenciones = 0
        End If
        rstListadoIcaro.Close
        .TextMatrix(intFila, 3) = FormatNumber(dblRetenciones, 2)
        .TextMatrix(intFila, 2) = FormatNumber(dblImporteBruto - dblRetenciones, 2)
        .Rows = .Rows + 1
        
        'Segundo Registro
        intFila = intFila + 1
        .RowHeight(intFila) = 300
        .TextMatrix(intFila, 0) = "LP"
        .TextMatrix(intFila, 1) = "LOTE PORÁ"
        SQLImputacion = "(" & SQLDate & "And IMPUTACION.Imputacion = '" & "12-00-51-80" & "') " _
        & "Or (" & SQLDate & "And IMPUTACION.Imputacion = '" & "12-00-43-51" & "')"
        SQL = "Select Sum(CARGA.Importe) AS TotalImporte " _
        & "From CARGA Inner Join IMPUTACION " _
        & "On CARGA.Comprobante = IMPUTACION.Comprobante " _
        & "Where " & SQLImputacion
        rstListadoIcaro.Open SQL, dbIcaro, adOpenDynamic, adLockOptimistic
        If Len(rstListadoIcaro!TotalImporte) <> "0" Then
            dblImporteBruto = rstListadoIcaro!TotalImporte
        Else
            dblImporteBruto = 0
        End If
        rstListadoIcaro.Close
        .TextMatrix(intFila, 4) = FormatNumber(dblImporteBruto, 2)
        SQL = "Select Sum(RETENCIONES.Importe) AS TotalImporte " _
        & "From (CARGA Inner Join IMPUTACION ON CARGA.Comprobante = IMPUTACION.Comprobante) " _
        & "Inner Join RETENCIONES ON CARGA.Comprobante = RETENCIONES.Comprobante " _
        & "Where " & SQLImputacion
        rstListadoIcaro.Open SQL, dbIcaro, adOpenDynamic, adLockOptimistic
        If Len(rstListadoIcaro!TotalImporte) <> "0" Then
            dblRetenciones = rstListadoIcaro!TotalImporte
        Else
            dblRetenciones = 0
        End If
        rstListadoIcaro.Close
        .TextMatrix(intFila, 3) = FormatNumber(dblRetenciones, 2)
        .TextMatrix(intFila, 2) = FormatNumber(dblImporteBruto - dblRetenciones, 2)
        .Rows = .Rows + 1
        
        'Tercer Registro
        intFila = intFila + 1
        .RowHeight(intFila) = 300
        .TextMatrix(intFila, 0) = "FFF"
        .TextMatrix(intFila, 1) = "FONDO FEDERAL SOLIDARIO"
        SQLImputacion = SQLDate & " And Left(IMPUTACION.Imputacion, 2) = '" & "21" & "'"
        SQL = "Select Sum(CARGA.Importe) AS TotalImporte " _
        & "From CARGA Inner Join IMPUTACION " _
        & "On CARGA.Comprobante = IMPUTACION.Comprobante " _
        & "Where " & SQLImputacion
        rstListadoIcaro.Open SQL, dbIcaro, adOpenDynamic, adLockOptimistic
        If Len(rstListadoIcaro!TotalImporte) <> "0" Then
            dblImporteBruto = rstListadoIcaro!TotalImporte
        Else
            dblImporteBruto = 0
        End If
        rstListadoIcaro.Close
        .TextMatrix(intFila, 4) = FormatNumber(dblImporteBruto, 2)
        SQL = "Select Sum(RETENCIONES.Importe) AS TotalImporte " _
        & "From (CARGA Inner Join IMPUTACION ON CARGA.Comprobante = IMPUTACION.Comprobante) " _
        & "Inner Join RETENCIONES ON CARGA.Comprobante = RETENCIONES.Comprobante " _
        & "Where " & SQLImputacion
        rstListadoIcaro.Open SQL, dbIcaro, adOpenDynamic, adLockOptimistic
        If Len(rstListadoIcaro!TotalImporte) <> "0" Then
            dblRetenciones = rstListadoIcaro!TotalImporte
        Else
            dblRetenciones = 0
        End If
        rstListadoIcaro.Close
        .TextMatrix(intFila, 3) = FormatNumber(dblRetenciones, 2)
        .TextMatrix(intFila, 2) = FormatNumber(dblImporteBruto - dblRetenciones, 2)

    End With

    Set rstListadoIcaro = Nothing
    
    With EjecucionPorImputacion.dgEjecucionPorImputacion
        If .TextMatrix(1, 0) <> "" Then
            Call CargardgEjecucionPorImputacionSeleccionada(FechaInicio, FechaFin, _
            .TextMatrix(1, 0))
'            RetencionesPorCtaCte.frmCtaCteSeleccionada.Caption = "Datos desagregados " _
'            & "de la Cuenta Corriente " & .TextMatrix(1, 0)
        End If
    End With
    
End Sub


Public Sub ConfigurardgEjecucionPorImputacionSeleccionada()
    
    With EjecucionPorImputacion.dgEjecucionPorImputacionSeleccionada
        .Clear
        .Cols = 9
        .Rows = 2
        .TextMatrix(0, 0) = "Nro. CyO"
        .TextMatrix(0, 1) = "Contratista"
        .TextMatrix(0, 2) = "Monto Neto"
        .TextMatrix(0, 3) = "Retenciones"
        .TextMatrix(0, 4) = "Monto Bruto"
        .TextMatrix(0, 5) = "Fecha"
        .TextMatrix(0, 6) = "Cta. Cte."
        .TextMatrix(0, 7) = "Estructura"
        .TextMatrix(0, 8) = "Obra"
        .ColWidth(0) = 1400
        .ColWidth(1) = 4200
        .ColWidth(2) = 1400
        .ColWidth(3) = 1400
        .ColWidth(4) = 1400
        .ColWidth(5) = 1400
        .ColWidth(6) = 1400
        .ColWidth(7) = 1400
        .ColWidth(8) = 5600
        .FixedCols = 0
        .FocusRect = flexFocusHeavy
        .HighLight = flexHighlightWithFocus
        .SelectionMode = flexSelectionByRow
        .AllowUserResizing = flexResizeColumns
        .ColAlignment(0) = 1
        .ColAlignment(1) = 1
        .ColAlignment(2) = 7
        .ColAlignment(3) = 7
        .ColAlignment(4) = 7
        .ColAlignment(5) = 7
        .ColAlignment(6) = 7
        .ColAlignment(7) = 7
        .ColAlignment(8) = 1
    End With
    
End Sub


Public Sub CargardgEjecucionPorImputacionSeleccionada(FechaInicio As String, _
FechaFin As String, Codigo As String)

    Dim intFila As Integer
    Dim SQL As String
    Dim SQLDate As String
    Dim SQLImputacion As String
    'Dim SQLImputacionSinRet As String
    Dim intNRegistros As Integer
    Dim datFechaInicio As Date
    Dim datFechaFin As Date
    
    Set rstListadoIcaro = New ADODB.Recordset
    datFechaInicio = DateTime.DateSerial(Right(FechaInicio, 4), Mid(FechaInicio, 4, 2), Left(FechaInicio, 2))
    datFechaFin = DateTime.DateSerial(Right(FechaFin, 4), Mid(FechaFin, 4, 2), Left(FechaFin, 2))
    If datFechaInicio > datFechaFin Then
        MsgBox "Debe ingresar una Fecha Inicio inferior a la Fecha Fin", vbCritical + vbOKOnly, "FECHAS INCORRECTAS"
        EjecucionPorImputacion.txtFechaDesde.SetFocus
        Exit Sub
    Else
        Set rstListadoIcaro = New ADODB.Recordset
        SQLDate = "CARGA.Fecha >= #" & Format(datFechaInicio, "MM/DD/YYYY") & "# " _
        & "AND CARGA.Fecha <= #" & Format(datFechaFin, "MM/DD/YYYY") & "# "
    End If
    
    Select Case Codigo
    Case Is = "CILP"
        SQLImputacion = SQLDate & " And (IMPUTACION.Imputacion = '" & "11-00-02-79" & "' " _
        & "Or IMPUTACION.Imputacion = '" & "11-00-43-52" & "')"
    Case Is = "LP"
        SQLImputacion = SQLDate & " And (IMPUTACION.Imputacion = '" & "12-00-51-80" & "' " _
        & "Or IMPUTACION.Imputacion = '" & "12-00-43-51" & "')"
    Case Is = "FFF"
        SQLImputacion = SQLDate & " And Left(IMPUTACION.Imputacion, 2) = '" & "21" & "'"
    End Select
        
    SQL = "Select CARGA.Comprobante, CUIT.Nombre, SUM(RETENCIONES.Importe) As TotalRetenciones, CARGA.Importe, " _
    & "CARGA.Fecha, CARGA.Cuenta, IMPUTACION.Imputacion, IMPUTACION.Partida, CARGA.Obra " _
    & "FROM CUIT INNER JOIN ((CARGA INNER JOIN IMPUTACION ON CARGA.Comprobante = IMPUTACION.Comprobante) " _
    & "INNER JOIN RETENCIONES ON CARGA.Comprobante = RETENCIONES.Comprobante) " _
    & "ON CUIT.CUIT = CARGA.CUIT " _
    & "Where " & SQLImputacion & " " _
    & "Group by CARGA.Comprobante, CUIT.Nombre, CARGA.Importe, CARGA.Fecha, CARGA.Cuenta, IMPUTACION.Imputacion, IMPUTACION.Partida, CARGA.Obra " _
    & "Order by CARGA.Comprobante Asc"
    
    rstListadoIcaro.Open SQL, dbIcaro, adOpenDynamic, adLockOptimistic
    
'    Debug.Print "N° Campos: " & rstListadoIcaro.Fields.Count
'    Debug.Print "N° Registros: " & rstListadoIcaro.RecordCount
    
    intNRegistros = rstListadoIcaro.RecordCount
    
    With EjecucionPorImputacion.dgEjecucionPorImputacionSeleccionada
        'Verificamos que la consulta tenga registros
        If intNRegistros > 0 Then
            intFila = 0
            rstListadoIcaro.MoveFirst
            While rstListadoIcaro.EOF = False
                intFila = intFila + 1
                .RowHeight(intFila) = 300
                .TextMatrix(intFila, 0) = rstListadoIcaro!Comprobante
                .TextMatrix(intFila, 1) = rstListadoIcaro!Nombre
                .TextMatrix(intFila, 2) = FormatNumber(rstListadoIcaro!Importe - rstListadoIcaro!TotalRetenciones, 2)
                .TextMatrix(intFila, 3) = FormatNumber(rstListadoIcaro!TotalRetenciones, 2)
                .TextMatrix(intFila, 4) = FormatNumber(rstListadoIcaro!Importe, 2)
                .TextMatrix(intFila, 5) = rstListadoIcaro!Fecha
                .TextMatrix(intFila, 6) = rstListadoIcaro!Cuenta
                .TextMatrix(intFila, 7) = rstListadoIcaro!Imputacion & "-" & rstListadoIcaro!Partida
                .TextMatrix(intFila, 8) = rstListadoIcaro!Obra
                rstListadoIcaro.MoveNext
                .Rows = .Rows + 1
            Wend
            .Rows = .Rows - 1
        Else
            'Procedemos a limpiar la tabla
            Call ConfigurardgEjecucionPorImputacionSeleccionada
        End If
        'Cargamos los Pagos a Cuenta (Sin Retenciones) - ¿Qué pasa si no hay Pagos con retenciones en el mes?
        rstListadoIcaro.Close
        SQL = "Select CARGA.Comprobante, CUIT.Nombre, CARGA.Importe, " _
        & "CARGA.Fecha, CARGA.Cuenta, IMPUTACION.Imputacion, IMPUTACION.Partida, CARGA.Obra " _
        & "FROM CUIT INNER JOIN (CARGA INNER JOIN IMPUTACION ON CARGA.Comprobante = IMPUTACION.Comprobante) " _
        & "ON CUIT.CUIT = CARGA.CUIT " _
        & "Where " & SQLImputacion & " And Carga.Comprobante Not In (SELECT RETENCIONES.Comprobante From RETENCIONES) " _
        & "Order by CARGA.Comprobante Asc"
        rstListadoIcaro.Open SQL, dbIcaro, adOpenDynamic, adLockOptimistic
        If rstListadoIcaro.BOF = False Then
            rstListadoIcaro.MoveFirst
            .Rows = .Rows + 1
            While rstListadoIcaro.EOF = False
                intFila = intFila + 1
                .RowHeight(intFila) = 300
                .TextMatrix(intFila, 0) = rstListadoIcaro!Comprobante
                .TextMatrix(intFila, 1) = rstListadoIcaro!Nombre
                .TextMatrix(intFila, 2) = FormatNumber(rstListadoIcaro!Importe, 2)
                .TextMatrix(intFila, 3) = FormatNumber(0, 2)
                .TextMatrix(intFila, 4) = FormatNumber(rstListadoIcaro!Importe, 2)
                .TextMatrix(intFila, 5) = rstListadoIcaro!Fecha
                .TextMatrix(intFila, 6) = rstListadoIcaro!Cuenta
                .TextMatrix(intFila, 7) = rstListadoIcaro!Imputacion & "-" & rstListadoIcaro!Partida
                .TextMatrix(intFila, 8) = rstListadoIcaro!Obra
                rstListadoIcaro.MoveNext
                .Rows = .Rows + 1
            Wend
            .Rows = .Rows - 1
        End If
    End With
    
    rstListadoIcaro.Close
    Set rstListadoIcaro = Nothing
    
End Sub

