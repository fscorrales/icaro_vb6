Attribute VB_Name = "EditarRegistros"
Public strEditandoComprobante As String
Public strEditandoProveedor As String
Public strEditandoCuentaBancaria As String
Public strEditandoPartida As String
Public strEditandoFuente As String
Public strEditandoCodigoRetencion As String
Public strEditandoLocalidad As String
Public strEditandoObra As String

Public Sub EditarComprobante()
    
    Dim i As String
    Dim SQL As String
    
    With CargaGasto
        i = ListadoImputacion.dgListado.Row
        If i <> 0 Then
            If ListadoImputacion.dgImputacion.Row = 0 Then
                strEditandoComprobante = ListadoImputacion.dgListado.TextMatrix(i, 0)
                Set rstRegistroIcaro = New ADODB.Recordset
                SQL = "Select * From CARGA Where COMPROBANTE = '" & strEditandoComprobante & "'"
                rstRegistroIcaro.Open SQL, dbIcaro, adOpenForwardOnly, adLockOptimistic
                .Show
                .txtComprobante.Text = Left(rstRegistroIcaro!Comprobante, 5)
                .txtFecha.Text = Format(rstRegistroIcaro!Fecha, "dd/mm/yyyy")
                .txtCUIT.Text = rstRegistroIcaro!CUIT
                .txtObra.Text = rstRegistroIcaro!Obra
                .txtImporte.Text = rstRegistroIcaro!Importe
                .txtFR.Text = rstRegistroIcaro!FondoDeReparo
                .txtCertificado.Text = rstRegistroIcaro!Certificado
                .txtAvance.Text = rstRegistroIcaro!Avance * 100
                .txtFuente.Text = rstRegistroIcaro!Fuente
                .txtCuenta.Text = rstRegistroIcaro!Cuenta
                rstRegistroIcaro.Close
                SQL = "Select * From CUIT Where CUIT = '" & .txtCUIT.Text & "'"
                rstRegistroIcaro.Open SQL, dbIcaro, adOpenForwardOnly, adLockOptimistic
                .txtDescripcionCUIT.Text = rstRegistroIcaro!Nombre
                rstRegistroIcaro.Close
                SQL = "Select * From FUENTES Where FUENTE = '" & .txtFuente.Text & "'"
                rstRegistroIcaro.Open SQL, dbIcaro, adOpenForwardOnly, adLockOptimistic
                .txtDescripcionFuente.Text = rstRegistroIcaro!Descripcion
                rstRegistroIcaro.Close
                SQL = "Select * From CUENTAS Where CUENTA = '" & .txtCuenta.Text & "'"
                rstRegistroIcaro.Open SQL, dbIcaro, adOpenForwardOnly, adLockOptimistic
                .txtDescripcionCuenta.Text = rstRegistroIcaro!Descripcion
                rstRegistroIcaro.Close
                Set rstRegistroIcaro = Nothing
                Unload ListadoImputacion
            Else
                MsgBox "Debe eliminar la Imputacion antes de poder " _
                & "modificar el comprobante", vbOKOnly + vbInformation, "IMPOSIBLE MODIFICAR"
                ListadoImputacion.dgListado.SetFocus
            End If
        End If
        i = ""
        SQL = ""
    End With
        
End Sub

Public Sub EditarProveedor()
    
    Dim i As String
    
    With ListadoProveedores.dgProveedores
        i = .Row
        If i <> 0 Then
            strEditandoProveedor = .TextMatrix(i, 0)
            CargaCUIT.Show
            CargaCUIT.txtCUIT.Text = .TextMatrix(i, 0)
            CargaCUIT.txtDescripcion.Text = .TextMatrix(i, 1)
            Unload ListadoProveedores
        End If
        i = ""
    End With
        
End Sub

Public Sub EditarCuentaBancaria()
    
    Dim i As String
    
    With ListadoCuentasBancarias.dgCuentasBancarias
        i = .Row
        If i <> 0 Then
            'strEditandoCuentaBancaria = (Left(.TextMatrix(i, 0), 2) & Mid(.TextMatrix(i, 0), 4, 1) & Mid(.TextMatrix(i, 0), 6, 5) & Right(.TextMatrix(i, 0), 1))
            strEditandoCuentaBancaria = .TextMatrix(i, 0)
            CargaCuentaBancaria.Show
            CargaCuentaBancaria.txtCuentaBancaria.Text = .TextMatrix(i, 0)
            CargaCuentaBancaria.txtDescripcion.Text = .TextMatrix(i, 1)
            Unload ListadoCuentasBancarias
        End If
        i = ""
    End With
        
End Sub

Public Sub EditarPartida()
    
    Dim i As String
    
    With ListadoPartidas.dgPartidas
        i = .Row
        If i <> 0 Then
            strEditandoPartida = .TextMatrix(i, 0)
            CargaPartida.Show
            CargaPartida.txtPartida.Text = .TextMatrix(i, 0)
            CargaPartida.txtDescripcion.Text = .TextMatrix(i, 1)
            Unload ListadoPartidas
        End If
        i = ""
    End With
        
End Sub

Public Sub EditarFuente()
    
    Dim i As String
    
    With ListadoFuentes.dgFuentes
        i = .Row
        If i <> 0 Then
            strEditandoFuente = .TextMatrix(i, 0)
            CargaFuente.Show
            CargaFuente.txtFuente.Text = .TextMatrix(i, 0)
            CargaFuente.txtDescripcion.Text = .TextMatrix(i, 1)
            Unload ListadoFuentes
        End If
        i = ""
    End With
        
End Sub

Public Sub EditarCodigoRetencion()
    
    Dim i As String
    
    With ListadoRetenciones.dgRetenciones
        i = .Row
        If i <> 0 Then
            strEditandoCodigoRetencion = .TextMatrix(i, 0)
            CodigoRetencion.Show
            CodigoRetencion.txtRetencion.Text = .TextMatrix(i, 0)
            CodigoRetencion.txtDescripcion.Text = .TextMatrix(i, 1)
            Unload ListadoRetenciones
        End If
        i = ""
    End With
        
End Sub

Public Sub EditarLocalidad()
        
    Dim i As String
    
    With ListadoLocalidades.dgLocalidades
        i = .Row
        If i <> 0 Then
            strEditandoLocalidad = .TextMatrix(i, 0)
            CargaLocalidad.Show
            CargaLocalidad.txtLocalidad.Text = .TextMatrix(i, 0)
            Unload ListadoLocalidades
        End If
        i = ""
    End With
        
End Sub

Public Sub EditarObra(DescripcionObra As String)
    
    'AGREGAR FILTRO PARA EL CASO DE REGISTROS NULOS O VACIOS.
    
    Dim SQL As String
    
    With CargaObra
        strEditandoObra = DescripcionObra
        CargaObra.Show
        Set rstRegistroIcaro = New ADODB.Recordset
        Set rstBuscarIcaro = New ADODB.Recordset
        SQL = "Select * from OBRAS Where DESCRIPCION = " & "'" & strEditandoObra & "'"
        rstRegistroIcaro.Open SQL, dbIcaro, adOpenForwardOnly, adLockOptimistic
        .txtCUIT.Text = rstRegistroIcaro!CUIT
        SQL = "Select * from CUIT Where CUIT = " & "'" & rstRegistroIcaro!CUIT & "'"
        rstBuscarIcaro.Open SQL, dbIcaro, adOpenForwardOnly, adLockOptimistic
        .txtDescripcionCUIT.Text = rstBuscarIcaro!Nombre
        rstBuscarIcaro.Close
        .txtPrograma.Text = Left(rstRegistroIcaro!Imputacion, 2)
        SQL = "Select * from PROGRAMAS Where PROGRAMA = " & "'" & Left(rstRegistroIcaro!Imputacion, 2) & "'"
        rstBuscarIcaro.Open SQL, dbIcaro, adOpenForwardOnly, adLockOptimistic
        .txtDescripcionPrograma.Text = rstBuscarIcaro!Descripcion
        rstBuscarIcaro.Close
        .txtSubprograma.Text = Mid(rstRegistroIcaro!Imputacion, 4, 2)
        SQL = "Select * from SUBPROGRAMAS Where SUBPROGRAMA = " & "'" & Left(rstRegistroIcaro!Imputacion, 5) & "'"
        rstBuscarIcaro.Open SQL, dbIcaro, adOpenForwardOnly, adLockOptimistic
        If IsNull(rstBuscarIcaro!Descripcion) = False Then
            .txtDescripcionSubprograma.Text = rstBuscarIcaro!Descripcion
        Else
            .txtDescripcionSubprograma.Text = ""
        End If
        rstBuscarIcaro.Close
        .txtProyecto.Text = Mid(rstRegistroIcaro!Imputacion, 7, 2)
        SQL = "Select * from PROYECTOS Where PROYECTO = " & "'" & Left(rstRegistroIcaro!Imputacion, 8) & "'"
        rstBuscarIcaro.Open SQL, dbIcaro, adOpenForwardOnly, adLockOptimistic
        .txtDescripcionProyecto.Text = rstBuscarIcaro!Descripcion
        rstBuscarIcaro.Close
        .txtActividad.Text = Right(rstRegistroIcaro!Imputacion, 2)
        SQL = "Select * from ACTIVIDADES Where ACTIVIDAD = " & "'" & rstRegistroIcaro!Imputacion & "'"
        rstBuscarIcaro.Open SQL, dbIcaro, adOpenForwardOnly, adLockOptimistic
        .txtDescripcionActividad.Text = rstBuscarIcaro!Descripcion
        rstBuscarIcaro.Close
        .txtPartida.Text = rstRegistroIcaro!Partida
        SQL = "Select * from PARTIDAS Where PARTIDA = " & "'" & rstRegistroIcaro!Partida & "'"
        rstBuscarIcaro.Open SQL, dbIcaro, adOpenForwardOnly, adLockOptimistic
        .txtDescripcionPartida.Text = rstBuscarIcaro!Descripcion
        rstBuscarIcaro.Close
        .txtMontoContrato.Text = rstRegistroIcaro!MontoDeContrato
        .txtAdicional.Text = rstRegistroIcaro!Adicional
        .txtObra.Text = rstRegistroIcaro!Descripcion
        .txtCuenta.Text = rstRegistroIcaro!Cuenta
        SQL = "Select * from CUENTAS Where CUENTA = " & "'" & rstRegistroIcaro!Cuenta & "'"
        rstBuscarIcaro.Open SQL, dbIcaro, adOpenForwardOnly, adLockOptimistic
        .txtDescripcionCuenta.Text = rstBuscarIcaro!Descripcion
        rstBuscarIcaro.Close
        .txtFuente.Text = rstRegistroIcaro!Fuente
        SQL = "Select * from FUENTES Where FUENTE = " & "'" & rstRegistroIcaro!Fuente & "'"
        rstBuscarIcaro.Open SQL, dbIcaro, adOpenForwardOnly, adLockOptimistic
        .txtDescripcionFuente.Text = rstBuscarIcaro!Descripcion
        rstBuscarIcaro.Close
        .txtLocalidad.Text = rstRegistroIcaro!Localidad
        If Len(rstRegistroIcaro!NormaLegal) <> "0" Then
            .txtNormaLegal.Text = rstRegistroIcaro!NormaLegal
        Else
            .txtNormaLegal.Text = ""
        End If
        rstRegistroIcaro.Close
        Set rstRegistroIcaro = Nothing
        Set rstBuscarIcaro = Nothing
        CargaObra.txtCUIT.SetFocus
    End With
        
End Sub

Public Sub EditarTipoDeObra(Destino As String, Optional Origen As String)

    Dim i As String
    Dim SQL As String
    Set dbIcaro = New ADODB.Connection
    dbIcaro.CursorLocation = adUseClient
    dbIcaro.Mode = adModeUnknown
    dbIcaro.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\ICARO.mdb"
    dbIcaro.Open
    Set rstRegistroIcaro = New ADODB.Recordset
    
    Select Case Destino
    Case "Reparaciones"
        With CargaTipoDeObra.dgNoCategorizadas
            i = .Row
            If i <> 0 Then
                SQL = "Select * From Actividades Where Actividad= '" & .TextMatrix(i, 0) & "'"
                rstRegistroIcaro.Open SQL, dbIcaro, adOpenForwardOnly, adLockOptimistic
                rstRegistroIcaro!TipoDeObra = "R"
                rstRegistroIcaro.Update
            End If
        End With
    Case "Viviendas"
        With CargaTipoDeObra.dgNoCategorizadas
            i = .Row
            If i <> 0 Then
                SQL = "Select * From Actividades Where Actividad= '" & .TextMatrix(i, 0) & "'"
                rstRegistroIcaro.Open SQL, dbIcaro, adOpenForwardOnly, adLockOptimistic
                rstRegistroIcaro!TipoDeObra = "V"
                rstRegistroIcaro.Update
            End If
        End With
    Case "NoCategorizadas"
        If Origen = "Reparaciones" Then
            With CargaTipoDeObra.dgReparaciones
                i = .Row
                If i <> 0 Then
                    SQL = "Select * From Actividades Where Actividad= '" & .TextMatrix(i, 0) & "'"
                    rstRegistroIcaro.Open SQL, dbIcaro, adOpenForwardOnly, adLockOptimistic
                    rstRegistroIcaro!TipoDeObra = Null
                    rstRegistroIcaro.Update
                End If
            End With
        ElseIf Origen = "Viviendas" Then
            With CargaTipoDeObra.dgViviendas
                i = .Row
                If i <> 0 Then
                    SQL = "Select * From Actividades Where Actividad= '" & .TextMatrix(i, 0) & "'"
                    rstRegistroIcaro.Open SQL, dbIcaro, adOpenForwardOnly, adLockOptimistic
                    rstRegistroIcaro!TipoDeObra = Null
                    rstRegistroIcaro.Update
                End If
            End With
        End If
    End Select
    
    SQL = ""
    i = ""
    rstRegistroIcaro.Close
    Set rstRegistroIcaro = Nothing
    dbIcaro.Close
    Set dbIcaro = Nothing
    
End Sub

Public Sub EditarEPAMConsolidado()
    
    Dim i As String
    
    With ImputacionEPAM
        i = EPAMConsolidado.dgEPAMConsolidado.Row
        If i <> 0 Then
            Set rstRegistroIcaro = New ADODB.Recordset
            strEditandoEPAMConsolidado = EPAMConsolidado.dgEPAMConsolidado.TextMatrix(i, 0)
            SQL = "Select * From DATOSFIJOSEPAM Where Codigo='" & strEditandoEPAMConsolidado & "' "
            rstRegistroIcaro.Open SQL, dbIcaro, adOpenForwardOnly, adLockOptimistic
            .Show
            .txtObra.Text = rstRegistroIcaro!Codigo
            .txtExpediente.Text = rstRegistroIcaro!Expediente
            .txtResolucion.Text = rstRegistroIcaro!Resolucion
            .txtPresupuestado.Text = De_Num_a_Tx_01(rstRegistroIcaro!Presupuestado)
            .txtPagado.Text = De_Num_a_Tx_01(rstRegistroIcaro!Pagado)
            .txtImputado.Text = De_Num_a_Tx_01(rstRegistroIcaro!Imputado)
            rstRegistroIcaro.Close
            Set rstRegistroIcaro = Nothing
            Unload EPAMConsolidado
            .txtObra.SetFocus
        End If
        i = ""
    End With
        
End Sub

Public Sub EditarResolucionEPAM()
    
    Dim i As String
    Dim x As String
    Dim SQL As String
    
    x = EPAMConsolidado.dgEPAMConsolidado.Row
    With EPAMConsolidado.dgResolucionesEPAM
        i = .Row
        If i <> 0 Then
            Set rstRegistroIcaro = New ADODB.Recordset
            strEditandoResolucionEPAM = .TextMatrix(i, 2)
            ImputacionResolucionesEPAM.Show
            SQL = "Select * From RESOLUCIONESEPAM Where Codigo = '" & EPAMConsolidado.dgEPAMConsolidado.TextMatrix(x, 0) & "' " _
            & "And Resolucion = '" & strEditandoResolucionEPAM & "'"
            rstRegistroIcaro.Open SQL, dbIcaro, adOpenForwardOnly, adLockOptimistic
            ImputacionResolucionesEPAM.txtObra.Text = rstRegistroIcaro!Codigo
            ImputacionResolucionesEPAM.txtDescripcion.Text = rstRegistroIcaro!Descripcion
            ImputacionResolucionesEPAM.txtExpediente.Text = rstRegistroIcaro!Expediente
            ImputacionResolucionesEPAM.txtResolucion.Text = rstRegistroIcaro!Resolucion
            ImputacionResolucionesEPAM.txtPresupuestado.Text = De_Num_a_Tx_01(rstRegistroIcaro!Importe)
            ImputacionResolucionesEPAM.txtImputado.Text = De_Num_a_Tx_01(rstRegistroIcaro!Imputado)
            rstRegistroIcaro.Close
            Set rstRegistroIcaro = Nothing
            Unload EPAMConsolidado
            ImputacionResolucionesEPAM.txtDescripcion.SetFocus
        End If
        i = ""
    End With
        
End Sub

Public Sub SincronizarObra(ObraICARO As String, ObraSGF As String)

    Dim SQL As String
    Dim i As Integer

    'If ValidarObra = True Then
    Set rstRegistroIcaro = New ADODB.Recordset
    SQL = "Select * from OBRAS Where DESCRIPCION = " & "'" & ObraICARO & "'"
    rstRegistroIcaro.Open SQL, dbIcaro, adOpenForwardOnly, adLockOptimistic
    SQL = ""
    rstRegistroIcaro!Descripcion = ObraSGF
'    rstRegistroIcaro!Localidad = .txtLocalidad.Text
'    rstRegistroIcaro!CUIT = .txtCUIT.Text
'    rstRegistroIcaro!Imputacion = .txtPrograma.Text & "-" & .txtSubprograma.Text & "-" & .txtProyecto.Text & "-" & .txtActividad.Text
'    rstRegistroIcaro!Partida = .txtPartida.Text
'    rstRegistroIcaro!MontoDeContrato = De_Txt_a_Num_01(.txtMontoContrato.Text)
'    rstRegistroIcaro!Adicional = De_Txt_a_Num_01(.txtAdicional.Text)
'    rstRegistroIcaro!Cuenta = .txtCuenta.Text
'    rstRegistroIcaro!Fuente = .txtFuente.Text
'    rstRegistroIcaro!NormaLegal = .txtNormaLegal.Text
    rstRegistroIcaro.Update
    MsgBox "Los datos fueron Modificados"
    rstRegistroIcaro.Close
    Set rstRegistroIcaro = Nothing
    With SincronizacionDescripcionSGF
        ConfigurarSincronizaciondgObrasSGF
        CargarSincronizaciondgObrasSGF
        i = .dgObrasSGF.Row
        .txtModificacion.Text = .dgObrasSGF.TextMatrix(i, 1)
        ConfigurarSincronizaciondgObrasICARO
        CargarSincronizaciondgObrasICARO (Left(.txtModificacion.Text, 50))
        i = .dgObrasICARO.Row
        If i > 0 Then
            .txtPrevioModificacion.Text = .dgObrasICARO.TextMatrix(i, 1)
        Else
            .txtPrevioModificacion.Text = ""
        End If
    End With
    'End If
        
End Sub
