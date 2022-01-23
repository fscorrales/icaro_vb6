Attribute VB_Name = "AgregarRegistros"
Public Sub GenerarComprobante()

    Dim x As String

    If ValidarComprobante = True Then
        Set rstRegistroIcaro = New ADODB.Recordset
        Dim strComprobante As String
        Dim datFecha As Date
        With CargaGasto
            strComprobante = Format(.txtComprobante.Text, "00000") & "/" & Right(.txtFecha.Text, 2)
            datFecha = DateTime.DateSerial(Right(.txtFecha, 4), Mid(.txtFecha, 4, 2), Left(.txtFecha, 2))
            If strEditandoComprobante = "" Then
                rstRegistroIcaro.Open "CARGA", dbIcaro, adOpenForwardOnly, adLockOptimistic
                rstRegistroIcaro.AddNew
            Else
                SQL = "Select * from CARGA Where COMPROBANTE = " & "'" & strEditandoComprobante & "'"
                rstRegistroIcaro.Open SQL, dbIcaro, adOpenForwardOnly, adLockOptimistic
                SQL = ""
                strEditandoComprobante = ""
            End If
            rstRegistroIcaro!Comprobante = strComprobante
            rstRegistroIcaro!Fecha = datFecha
            rstRegistroIcaro!CUIT = .txtCUIT.Text
            rstRegistroIcaro!Obra = .txtObra.Text
            rstRegistroIcaro!Importe = De_Txt_a_Num_01(.txtImporte.Text)
            rstRegistroIcaro!FondoDeReparo = De_Txt_a_Num_01(.txtFR.Text)
            rstRegistroIcaro!Certificado = .txtCertificado.Text
            rstRegistroIcaro!Avance = De_Txt_a_Num_01(.txtAvance.Text) / 100
            rstRegistroIcaro!Cuenta = .txtCuenta.Text
            rstRegistroIcaro!Fuente = .txtFuente.Text
            rstRegistroIcaro.Update
        End With
        rstRegistroIcaro.Close
        rstRegistroIcaro.Open "IMPUTACION", dbIcaro, adOpenForwardOnly, adLockOptimistic
        rstRegistroIcaro.AddNew
        Set rstBuscarIcaro = New ADODB.Recordset
        SQL = "Select * from OBRAS Where DESCRIPCION = " & "'" & CargaGasto.txtObra.Text & "'"
        rstBuscarIcaro.Open SQL, dbIcaro, adOpenForwardOnly, adLockOptimistic
        With rstRegistroIcaro
            .Fields("Comprobante") = strComprobante
            .Fields("Imputacion") = rstBuscarIcaro.Fields("Imputacion")
            .Fields("Partida") = rstBuscarIcaro.Fields("Partida")
            .Fields("Importe") = De_Txt_a_Num_01(CargaGasto.txtImporte.Text)
            .Update
            .Close
        End With
        rstBuscarIcaro.Close
        Set rstBuscarIcaro = Nothing

        If bolAutocargaCertificado = True Then
            rstRegistroIcaro.Open "RETENCIONES", dbIcaro, adOpenForwardOnly, adLockOptimistic
            If Trim(curIB) <> "0" Then
                With rstRegistroIcaro
                    .AddNew
                    .Fields("Comprobante") = strComprobante
                    .Fields("Codigo") = "110"
                    .Fields("Importe") = curIB
                    .Update
                End With
            End If
            If Trim(curLP) <> "0" Then
                With rstRegistroIcaro
                    .AddNew
                    .Fields("Comprobante") = strComprobante
                    .Fields("Codigo") = "112"
                    .Fields("Importe") = curLP
                    .Update
                End With
            End If
            If Trim(curGanancias) <> "0" Then
                With rstRegistroIcaro
                    .AddNew
                    .Fields("Comprobante") = strComprobante
                    .Fields("Codigo") = "113"
                    .Fields("Importe") = curGanancias
                    .Update
                End With
            End If
            If Trim(curSUSS) <> "0" Then
                With rstRegistroIcaro
                    .AddNew
                    .Fields("Comprobante") = strComprobante
                    .Fields("Codigo") = "114"
                    .Fields("Importe") = curSUSS
                    .Update
                End With
            End If
            If Trim(curINVICO) <> "0" Then
                With rstRegistroIcaro
                    .AddNew
                    .Fields("Comprobante") = strComprobante
                    .Fields("Codigo") = "337"
                    .Fields("Importe") = curINVICO
                    .Update
                End With
            End If
            rstRegistroIcaro.Close
            SQL = "Select * from CERTIFICADOS where Proveedor = '" & strProveedor & "' " _
            & "And Obra = '" & strObra & "' And Certificado = '" & strCertificado & "'"
            rstRegistroIcaro.Open SQL, dbIcaro, adOpenForwardOnly, adLockOptimistic
            rstRegistroIcaro.Delete
            rstRegistroIcaro.Close
            VaciarVariablesAutocarga
        End If
        
        If bolAutocargaEPAM = True Then
            rstRegistroIcaro.Open "RETENCIONES", dbIcaro, adOpenForwardOnly, adLockOptimistic
            If Trim(curIB) <> "0" Then
                With rstRegistroIcaro
                    .AddNew
                    .Fields("Comprobante") = strComprobante
                    .Fields("Codigo") = "110"
                    .Fields("Importe") = curIB
                    .Update
                End With
            End If
            If Trim(curSellos) <> "0" Then
                With rstRegistroIcaro
                    .AddNew
                    .Fields("Comprobante") = strComprobante
                    .Fields("Codigo") = "111"
                    .Fields("Importe") = curSellos
                    .Update
                End With
            End If
            If Trim(curLP) <> "0" Then
                With rstRegistroIcaro
                    .AddNew
                    .Fields("Comprobante") = strComprobante
                    .Fields("Codigo") = "112"
                    .Fields("Importe") = curLP
                    .Update
                End With
            End If
            If Trim(curGanancias) <> "0" Then
                With rstRegistroIcaro
                    .AddNew
                    .Fields("Comprobante") = strComprobante
                    .Fields("Codigo") = "113"
                    .Fields("Importe") = curGanancias
                    .Update
                End With
            End If
            If Trim(curSUSS) <> "0" Then
                With rstRegistroIcaro
                    .AddNew
                    .Fields("Comprobante") = strComprobante
                    .Fields("Codigo") = "114"
                    .Fields("Importe") = curSUSS
                    .Update
                End With
            End If
            rstRegistroIcaro.Close
            SQL = "Select * from EPAM where COMPROBANTE = """" " _
            & "And Codigo = '" & strObra & "' And MontoBruto = " & curMontoBruto & ""
            rstRegistroIcaro.Open SQL, dbIcaro, adOpenForwardOnly, adLockOptimistic
            With rstRegistroIcaro
                .Fields("Comprobante") = strComprobante
                .Fields("Fecha") = datFecha
                .Update
            End With
            rstRegistroIcaro.Close
            VaciarVariablesAutocarga
        End If
        datFecha = 0
        MsgBox "Los datos fueron agregados"
        Set rstRegistroIcaro = Nothing
        ListadoImputacion.Show
        Unload CargaGasto
        x = strComprobante
        strComprobante = ""
        ListadoImputacion.txtFecha.Text = "20" & Right(x, 2)
        ConfigurardgListadoGasto
        Call CargardgListadoGasto(, "20" & Right(x, 2), x)
        ConfigurardgImputacion
        CargardgImputacion (x)
        ConfigurardgRetencion
        CargardgRetencion (x)
        x = 0
        ListadoImputacion.dgListado.SetFocus
    End If
    
End Sub



Public Sub GenerarProveedor()

    Dim SQL As String

    If ValidarProveedor = True Then
        Set rstRegistroIcaro = New ADODB.Recordset
        With CargaCUIT
            If strEditandoProveedor = "" Then
                rstRegistroIcaro.Open "CUIT", dbIcaro, adOpenForwardOnly, adLockOptimistic
                rstRegistroIcaro.AddNew
            Else
                SQL = "Select * from CUIT Where CUIT = " & "'" & strEditandoProveedor & "'"
                rstRegistroIcaro.Open SQL, dbIcaro, adOpenForwardOnly, adLockOptimistic
                SQL = ""
                strEditandoProveedor = ""
            End If
            rstRegistroIcaro!CUIT = .txtCUIT.Text
            rstRegistroIcaro!Nombre = .txtDescripcion.Text
            rstRegistroIcaro.Update
        End With
        MsgBox "Los datos fueron agregados"
        rstRegistroIcaro.Close
        Set rstRegistroIcaro = Nothing
        If bolCUITDesdeCarga = True Then
            bolCUITDesdeCarga = False
            Unload CargaCUIT
            CargaGasto.Show
            CargaGasto.txtCUIT.SetFocus
        ElseIf bolAutocargaCertificado = True Then
            bolAutocargaCertificado = False
            Unload CargaCUIT
            AgregarCertificado
        Else
            Unload CargaCUIT
            ListadoProveedores.Show
            ConfigurardgProveedores
            CargardgProveedores
        End If
    End If
    
End Sub

Public Sub GenerarCuentaBancaria()

    Dim SQL As String
    Dim CuentaConGuiones As String

    If ValidarCuentaBancaria = True Then
        Set rstRegistroIcaro = New ADODB.Recordset
        With CargaCuentaBancaria
            If strEditandoCuentaBancaria = "" Then
                rstRegistroIcaro.Open "CUENTAS", dbIcaro, adOpenForwardOnly, adLockOptimistic
                rstRegistroIcaro.AddNew
            Else
                'CuentaConGuiones = Left(strEditandoProveedor, 2) & "-" & Mid(strEditandoProveedor, 3, 1) & "-" & Mid(strEditandoProveedor, 4, 5) & "-" & Right(strEditandoProveedor, 1)
                SQL = "Select * from CUENTAS Where CUENTA = " & "'" & strEditandoCuentaBancaria & "'"
                rstRegistroIcaro.Open SQL, dbIcaro, adOpenForwardOnly, adLockOptimistic
                SQL = ""
                strEditandoCuentaBancaria = ""
            End If
            'CuentaConGuiones = Left(.txtCuentaBancaria.Text, 2) & "-" & Mid(.txtCuentaBancaria.Text, 3, 1) & "-" & Mid(.txtCuentaBancaria.Text, 4, 5) & "-" & Right(.txtCuentaBancaria.Text, 1)
            rstRegistroIcaro!Cuenta = .txtCuentaBancaria.Text
            rstRegistroIcaro!Descripcion = .txtDescripcion.Text
            rstRegistroIcaro.Update
        End With
        MsgBox "Los datos fueron agregados"
        rstRegistroIcaro.Close
        Set rstRegistroIcaro = Nothing
        Unload CargaCuentaBancaria
        ListadoCuentasBancarias.Show
        ConfigurardgCuentasBancarias
        CargardgCuentasBancarias
    End If
    
End Sub

Public Sub GenerarPartida()

    Dim SQL As String

    If ValidarPartida = True Then
        Set rstRegistroIcaro = New ADODB.Recordset
        With CargaPartida
            If strEditandoPartida = "" Then
                rstRegistroIcaro.Open "PARTIDAS", dbIcaro, adOpenForwardOnly, adLockOptimistic
                rstRegistroIcaro.AddNew
            Else
                SQL = "Select * from PARTIDAS Where PARTIDA = " & "'" & strEditandoPartida & "'"
                rstRegistroIcaro.Open SQL, dbIcaro, adOpenForwardOnly, adLockOptimistic
                SQL = ""
                strEditandoPartida = ""
            End If
            rstRegistroIcaro!Partida = .txtPartida.Text
            rstRegistroIcaro!Descripcion = .txtDescripcion.Text
            rstRegistroIcaro.Update
        End With
        MsgBox "Los datos fueron agregados"
        rstRegistroIcaro.Close
        Set rstRegistroIcaro = Nothing
        Unload CargaPartida
        ListadoPartidas.Show
        ConfigurardgPartidas
        CargardgPartidas
    End If
        
End Sub

Public Sub GenerarFuente()

    Dim SQL As String

    If ValidarFuente = True Then
        Set rstRegistroIcaro = New ADODB.Recordset
        With CargaFuente
            If strEditandoFuente = "" Then
                rstRegistroIcaro.Open "FUENTES", dbIcaro, adOpenForwardOnly, adLockOptimistic
                rstRegistroIcaro.AddNew
            Else
                SQL = "Select * from FUENTES Where FUENTE = " & "'" & strEditandoFuente & "'"
                rstRegistroIcaro.Open SQL, dbIcaro, adOpenForwardOnly, adLockOptimistic
                SQL = ""
                strEditandoFuente = ""
            End If
            rstRegistroIcaro!Fuente = .txtFuente.Text
            rstRegistroIcaro!Descripcion = .txtDescripcion.Text
            rstRegistroIcaro.Update
        End With
        MsgBox "Los datos fueron agregados"
        rstRegistroIcaro.Close
        Set rstRegistroIcaro = Nothing
        Unload CargaFuente
        ListadoFuentes.Show
        ConfigurardgFuentes
        CargardgFuentes
    End If
        
End Sub

Public Sub GenerarCodigoRetencion()

    Dim SQL As String

    If ValidarCodigoRetencion = True Then
        Set rstRegistroIcaro = New ADODB.Recordset
        With CodigoRetencion
            If strEditandoCodigoRetencion = "" Then
                rstRegistroIcaro.Open "CODIGOSRETENCIONES", dbIcaro, adOpenForwardOnly, adLockOptimistic
                rstRegistroIcaro.AddNew
            Else
                SQL = "Select * from CODIGOSRETENCIONES Where CODIGO = " & "'" & strEditandoCodigoRetencion & "'"
                rstRegistroIcaro.Open SQL, dbIcaro, adOpenForwardOnly, adLockOptimistic
                SQL = ""
                strEditandoCodigoRetencion = ""
            End If
            rstRegistroIcaro!Retencion = .txtRetencion.Text
            rstRegistroIcaro!Descripcion = .txtDescripcion.Text
            rstRegistroIcaro.Update
        End With
        MsgBox "Los datos fueron agregados"
        rstRegistroIcaro.Close
        Set rstRegistroIcaro = Nothing
        Unload CodigoRetencion
        ListadoRetenciones.Show
        ConfigurardgCodigoRetenciones
        CargardgCodigoRetenciones
    End If
        
End Sub

Public Sub GenerarLocalidad()

    Dim SQL As String

    If ValidarLocalidad = True Then
        Set rstRegistroIcaro = New ADODB.Recordset
        With CargaLocalidad
            If strEditandoLocalidad = "" Then
                rstRegistroIcaro.Open "LOCALIDADES", dbIcaro, adOpenForwardOnly, adLockOptimistic
                rstRegistroIcaro.AddNew
            Else
                SQL = "Select * from LOCALIDADES Where LOCALIDAD = " & "'" & strEditandoLocalidad & "'"
                rstRegistroIcaro.Open SQL, dbIcaro, adOpenForwardOnly, adLockOptimistic
                SQL = ""
                strEditandoLocalidad = ""
            End If
            rstRegistroIcaro!Localidad = .txtLocalidad.Text
            rstRegistroIcaro.Update
        End With
        MsgBox "Los datos fueron agregados"
        rstRegistroIcaro.Close
        Set rstRegistroIcaro = Nothing
        Unload CargaLocalidad
        ListadoLocalidades.Show
        ConfigurardgLocalidades
        CargardgLocalidades
    End If
        
End Sub

Public Sub GenerarObra()

    Dim SQL As String

    If ValidarObra = True Then
        Set rstRegistroIcaro = New ADODB.Recordset
        With CargaObra
            If strEditandoObra = "" Then
                rstRegistroIcaro.Open "OBRAS", dbIcaro, adOpenForwardOnly, adLockOptimistic
                rstRegistroIcaro.AddNew
            Else
                SQL = "Select * from OBRAS Where DESCRIPCION = " & "'" & strEditandoObra & "'"
                rstRegistroIcaro.Open SQL, dbIcaro, adOpenForwardOnly, adLockOptimistic
                SQL = "Select * from DATOSFIJOSEPAM Where CODIGO = " & "'" & strEditandoObra & "'"
                If SQLNoMatch(SQL) = False Then
                    strEditandoEPAMConsolidado = strEditandoObra
                Else
                    strEditandoEPAMConsolidado = ""
                End If
                SQL = ""
                strEditandoObra = ""
            End If
            rstRegistroIcaro!Descripcion = .txtObra.Text
            rstRegistroIcaro!Localidad = .txtLocalidad.Text
            rstRegistroIcaro!CUIT = .txtCUIT.Text
            rstRegistroIcaro!Imputacion = .txtPrograma.Text & "-" & .txtSubprograma.Text & "-" & .txtProyecto.Text & "-" & .txtActividad.Text
            rstRegistroIcaro!Partida = .txtPartida.Text
            rstRegistroIcaro!MontoDeContrato = De_Txt_a_Num_01(.txtMontoContrato.Text)
            rstRegistroIcaro!Adicional = De_Txt_a_Num_01(.txtAdicional.Text)
            rstRegistroIcaro!Cuenta = .txtCuenta.Text
            rstRegistroIcaro!Fuente = .txtFuente.Text
            rstRegistroIcaro!NormaLegal = .txtNormaLegal.Text
            rstRegistroIcaro.Update
        End With
        rstRegistroIcaro.Close
        Set rstRegistroIcaro = Nothing
        If strEditandoEPAMConsolidado <> "" Then
            Call GenerarEPAMConsolidado(CargaObra.txtObra.Text)
            SQL = " y se procedió a actualizar la Obra Consolidada en EPAM"
        End If
        MsgBox "Los datos fueron agregados" & SQL
        SQL = ""
        If bolObraDesdeCarga = True Then
            Unload CargaObra
            CargaGasto.Show
            bolObraDesdeCarga = False
            CargaGasto.txtObra.SetFocus
        ElseIf bolAutocargaCertificado = True Then
            bolAutocargaCertificado = False
            Unload CargaObra
            AgregarCertificado
        ElseIf bolAutocargaEPAM = True Then
            bolAutocargaEPAM = False
            Unload CargaObra
            AgregarEPAM
        ElseIf bolObraDesdeEPAMConsolidado = True Then
            bolObraDesdeEPAMConsolidado = False
            Unload CargaObra
            ImputacionEPAM.Show
            ImputacionEPAM.txtObra.SetFocus
        Else
            Unload CargaObra
            ListadoObras.Show
            ConfigurardgObras
            CargardgObras
        End If
    End If
        
End Sub

Public Sub AgregarCertificado()

    Dim SQL As String

    If SQLNoMatch("Select * From CUIT Where Nombre = '" & strProveedor & "'") = True Then
        bolAutocargaCertificado = True
        CargaCUIT.Show
        CargaCUIT.txtDescripcion.Text = strProveedor
        CargaCUIT.txtDescripcion.Enabled = False
        Unload AutocargaCertificados
        Exit Sub
    End If
    If SQLNoMatch("Select * From OBRAS Where Descripcion = '" & strObra & "'") = True Then
        bolAutocargaCertificado = True
        CargaObra.Show
        CargaObra.txtDescripcionCUIT.Text = strProveedor
        CargaObra.txtDescripcionCUIT.Enabled = False
        Set rstRegistroIcaro = New ADODB.Recordset
        SQL = "Select * From CUIT Where Nombre = '" & strProveedor & "'"
        rstRegistroIcaro.Open SQL, dbIcaro, adOpenForwardOnly, adLockOptimistic
        CargaObra.txtCUIT.Text = rstRegistroIcaro!CUIT
        CargaObra.txtCUIT.Enabled = False
        CargaObra.txtObra.Text = strObra
        CargaObra.txtObra.Enabled = False
        CargaObra.txtPrograma.SetFocus
        rstRegistroIcaro.Close
        Set rstRegistroIcaro = Nothing
        Unload AutocargaCertificados
        Exit Sub
    End If
    With CargaGasto
        bolAutocargaCertificado = True
        .Show
        Set rstRegistroIcaro = New ADODB.Recordset
        SQL = "Select * From CUIT Where Nombre = '" & strProveedor & "'"
        rstRegistroIcaro.Open SQL, dbIcaro, adOpenForwardOnly, adLockOptimistic
        .txtFecha.Text = Format(Day(Date), "00") & "/" _
        & Format(Month(Date), "00") & "/" & Format(Year(Date), "0000")
        .txtCUIT.Text = rstRegistroIcaro!CUIT
        .txtCUIT.Enabled = False
        rstRegistroIcaro.Close
        .txtDescripcionCUIT.Text = strProveedor
        .txtDescripcionCUIT.Enabled = False
        .txtObra.Text = strObra
        .txtObra.Enabled = False
        .txtImporte.Text = curMontoBruto
        .txtFR.Text = curFondoDeReparo
        If Len(strCertificado) <> "0" Then
            .txtCertificado.Text = strCertificado
        Else
            .txtCertificado.Text = "0"
        End If
        .txtAvance.Text = "0"
        SQL = "Select * From OBRAS Where DESCRIPCION = '" & strObra & "'"
        rstRegistroIcaro.Open SQL, dbIcaro, adOpenForwardOnly, adLockOptimistic
        .txtFuente.Text = rstRegistroIcaro!Fuente
        .txtFuente.Enabled = True
        .txtCuenta.Text = rstRegistroIcaro!Cuenta
        .txtCuenta.Enabled = True
        rstRegistroIcaro.Close
        SQL = "Select * From CUENTAS Where CUENTA = '" & .txtCuenta.Text & "'"
        rstRegistroIcaro.Open SQL, dbIcaro, adOpenForwardOnly, adLockOptimistic
        .txtDescripcionCuenta.Text = rstRegistroIcaro!Descripcion
        .txtDescripcionCuenta.Enabled = False
        rstRegistroIcaro.Close
        SQL = "Select * From FUENTES Where FUENTE = '" & .txtFuente.Text & "'"
        rstRegistroIcaro.Open SQL, dbIcaro, adOpenForwardOnly, adLockOptimistic
        .txtDescripcionFuente.Text = rstRegistroIcaro!Descripcion
        .txtDescripcionFuente.Enabled = False
        .txtComprobante.SetFocus
        rstRegistroIcaro.Close
        Set rstRegistroIcaro = Nothing
    End With
    Unload AutocargaCertificados
    
End Sub

Public Sub AgregarEPAM()

    Dim SQL As String
    
    If SQLNoMatch("Select * From OBRAS Where Descripcion = '" & strObra & "'") = True Then
        bolAutocargaEPAM = True
        With CargaObra
            .Show
            .txtCUIT.Text = "30632351514"
            .txtCUIT.Enabled = False
            Set rstRegistroIcaro = New ADODB.Recordset
            SQL = "Select * From CUIT Where CUIT = '" & .txtCUIT.Text & "'"
            rstRegistroIcaro.Open SQL, dbIcaro, adOpenForwardOnly, adLockOptimistic
            .txtDescripcionCUIT.Text = rstRegistroIcaro!Nombre
            .txtDescripcionCUIT.Enabled = False
            .txtObra.Text = strObra
            .txtObra.Enabled = False
            .txtPrograma.SetFocus
        End With
        rstRegistroIcaro.Close
        Set rstRegistroIcaro = Nothing
        Unload AutocargaEPAM
        Exit Sub
    End If
    
    With CargaGasto
        bolAutocargaEPAM = True
        .Show
        .txtFecha.Text = Format(Day(Date), "00") & "/" _
        & Format(Month(Date), "00") & "/" & Format(Year(Date), "0000")
        .txtCUIT.Text = "30632351514"
        .txtCUIT.Enabled = False
        Set rstRegistroIcaro = New ADODB.Recordset
        SQL = "Select * From CUIT Where CUIT = '" & .txtCUIT.Text & "'"
        rstRegistroIcaro.Open SQL, dbIcaro, adOpenForwardOnly, adLockOptimistic
        .txtDescripcionCUIT.Text = rstRegistroIcaro!Nombre
        .txtDescripcionCUIT.Enabled = False
        rstRegistroIcaro.Close
        .txtObra.Text = strObra
        .txtObra.Enabled = False
        .txtImporte.Text = curMontoBruto
        .txtFR.Text = "0"
        .txtCertificado.Text = "0"
        .txtAvance.Text = "0"
        SQL = "Select * From OBRAS Where DESCRIPCION = '" & strObra & "'"
        rstRegistroIcaro.Open SQL, dbIcaro, adOpenForwardOnly, adLockOptimistic
        .txtFuente.Text = rstRegistroIcaro!Fuente
        .txtFuente.Enabled = True
        .txtCuenta.Text = rstRegistroIcaro!Cuenta
        .txtCuenta.Enabled = True
        rstRegistroIcaro.Close
        SQL = "Select * From CUENTAS Where CUENTA = '" & .txtCuenta.Text & "'"
        rstRegistroIcaro.Open SQL, dbIcaro, adOpenForwardOnly, adLockOptimistic
        .txtDescripcionCuenta.Text = rstRegistroIcaro!Descripcion
        .txtDescripcionCuenta.Enabled = False
        rstRegistroIcaro.Close
        SQL = "Select * From FUENTES Where FUENTE = '" & .txtFuente.Text & "'"
        rstRegistroIcaro.Open SQL, dbIcaro, adOpenForwardOnly, adLockOptimistic
        .txtDescripcionFuente.Text = rstRegistroIcaro!Descripcion
        .txtDescripcionFuente.Enabled = False
        .txtComprobante.SetFocus
        rstRegistroIcaro.Close
        Set rstRegistroIcaro = Nothing
    End With
    Unload AutocargaEPAM

End Sub

'Public Sub GenerarEPAMConsolidado()
'
'    Dim SQL As String
'
'    If ValidarEPAMConsolidado = True Then
'        Set rstRegistroIcaro = New ADODB.Recordset
'        With ImputacionEPAM
'            If strEditandoEPAMConsolidado = "" Then
'                rstRegistroIcaro.Open "DATOSFIJOSEPAM", dbIcaro, adOpenForwardOnly, adLockOptimistic
'                rstRegistroIcaro.AddNew
'            Else
'                SQL = "Select * from DATOSFIJOSEPAM Where CODIGO = " & "'" & strEditandoEPAMConsolidado & "'"
'                rstRegistroIcaro.Open SQL, dbIcaro, adOpenForwardOnly, adLockOptimistic
'                SQL = ""
'            End If
'            rstRegistroIcaro!Expediente = .txtExpediente.Text
'            rstRegistroIcaro!Resolucion = .txtResolucion.Text
'            rstRegistroIcaro!Codigo = .txtObra.Text
'            rstRegistroIcaro!Presupuestado = De_Txt_a_Num_01(.txtPresupuestado.Text)
'            rstRegistroIcaro!Pagado = De_Txt_a_Num_01(.txtPagado.Text)
'            rstRegistroIcaro!Imputado = De_Txt_a_Num_01(.txtImputado.Text)
'            rstRegistroIcaro.Update
'            'En caso de modificar el Código EPAM
'            If strEditandoEPAMConsolidado <> .txtObra.Text Then
'                rstRegistroIcaro.Close
'                SQL = "Select * from EPAM " _
'                & "Where CODIGO = " & "'" & strEditandoEPAMConsolidado & "'"
'                rstRegistroIcaro.Open SQL, dbIcaro, adOpenForwardOnly, adLockOptimistic
'                SQL = ""
'                If rstRegistroIcaro.BOF = False Then
'                    rstRegistroIcaro.MoveFirst
'                    While rstRegistroIcaro.EOF = False
'                        rstRegistroIcaro!Codigo = .txtObra.Text
'                        rstRegistroIcaro.Update
'                        rstRegistroIcaro.MoveNext
'                    Wend
'                End If
'            End If
'            strEditandoEPAMConsolidado = ""
'        End With
'        MsgBox "Los datos fueron agregados"
'        rstRegistroIcaro.Close
'        Set rstRegistroIcaro = Nothing
'        Unload ImputacionEPAM
'        EPAMConsolidado.Show
'        ConfigurardgEPAMConsolidado
'        CargardgEPAMConsolidado
'        Dim i As String
'        i = EPAMConsolidado.dgEPAMConsolidado.Row
'        ConfigurardgResolucionesEPAM
'        CargardgResolucionesEPAM (EPAMConsolidado.dgEPAMConsolidado.TextMatrix(i, 0))
'        i = ""
'    End If
'
'End Sub

Public Sub GenerarEPAMConsolidado(Optional Obra As Variant = Null, _
Optional Expediente As Variant = Null, Optional Resolucion As Variant = Null, _
Optional Presupuestado As Variant = Null, Optional Pagado As Variant = Null, _
Optional Imputado As Variant = Null)

    Dim SQL As String

    Set rstRegistroIcaro = New ADODB.Recordset
    If strEditandoEPAMConsolidado = "" Then
        rstRegistroIcaro.Open "DATOSFIJOSEPAM", dbIcaro, adOpenForwardOnly, adLockOptimistic
        rstRegistroIcaro.AddNew
    Else
        SQL = "Select * from DATOSFIJOSEPAM Where CODIGO = " & "'" & strEditandoEPAMConsolidado & "'"
        rstRegistroIcaro.Open SQL, dbIcaro, adOpenForwardOnly, adLockOptimistic
        SQL = ""
    End If
    If IsNull(Expediente) = False Then
        rstRegistroIcaro!Expediente = Expediente
    End If
    If IsNull(Resolucion) = False Then
        rstRegistroIcaro!Resolucion = Resolucion
    End If
    If IsNull(Obra) = False Then
        rstRegistroIcaro!Codigo = Obra
    End If
    If IsNull(Presupuestado) = False Then
        rstRegistroIcaro!Presupuestado = De_Txt_a_Num_01(Presupuestado)
    End If
    If IsNull(Pagado) = False Then
        rstRegistroIcaro!Pagado = De_Txt_a_Num_01(Pagado)
    End If
    If IsNull(Imputado) = False Then
        rstRegistroIcaro!Imputado = De_Txt_a_Num_01(Imputado)
    End If
    rstRegistroIcaro.Update
    'En caso de modificar el Código EPAM
    If IsNull(Obra) = False Then
        If strEditandoEPAMConsolidado <> Obra Then
            rstRegistroIcaro.Close
            SQL = "Select * from EPAM " _
            & "Where CODIGO = " & "'" & strEditandoEPAMConsolidado & "'"
            rstRegistroIcaro.Open SQL, dbIcaro, adOpenForwardOnly, adLockOptimistic
            SQL = ""
            If rstRegistroIcaro.BOF = False Then
                rstRegistroIcaro.MoveFirst
                While rstRegistroIcaro.EOF = False
                    rstRegistroIcaro!Codigo = Obra
                    rstRegistroIcaro.Update
                    rstRegistroIcaro.MoveNext
                Wend
            End If
        End If
    End If
    strEditandoEPAMConsolidado = ""
    rstRegistroIcaro.Close
    Set rstRegistroIcaro = Nothing
    
End Sub

Public Sub GenerarResolucionEPAM()

    Dim SQL As String

    If ValidarResolucionEPAM = True Then
        Set rstRegistroIcaro = New ADODB.Recordset
        With ImputacionResolucionesEPAM
            If strEditandoResolucionEPAM = "" Then
                rstRegistroIcaro.Open "RESOLUCIONESEPAM", dbIcaro, adOpenForwardOnly, adLockOptimistic
                rstRegistroIcaro.AddNew
            Else
                SQL = "Select * from RESOLUCIONESEPAM Where CODIGO = " & "'" & .txtObra & "' " _
                & "And RESOLUCION = " & "'" & strEditandoResolucionEPAM & "'"
                rstRegistroIcaro.Open SQL, dbIcaro, adOpenForwardOnly, adLockOptimistic
                SQL = ""
                strEditandoResolucionEPAM = ""
            End If
            rstRegistroIcaro!Codigo = .txtObra.Text
            rstRegistroIcaro!Descripcion = .txtDescripcion.Text
            rstRegistroIcaro!Expediente = .txtExpediente.Text
            rstRegistroIcaro!Resolucion = .txtResolucion.Text
            rstRegistroIcaro!Importe = De_Txt_a_Num_01(.txtPresupuestado.Text)
            rstRegistroIcaro!Imputado = De_Txt_a_Num_01(.txtImputado.Text)
            rstRegistroIcaro.Update
        End With
        MsgBox "Los datos fueron agregados"
        rstRegistroIcaro.Close
        Set rstRegistroIcaro = Nothing
        Unload ImputacionResolucionesEPAM
        EPAMConsolidado.Show
        ConfigurardgEPAMConsolidado
        CargardgEPAMConsolidado
        Dim i As String
        i = EPAMConsolidado.dgEPAMConsolidado.Row
        ConfigurardgResolucionesEPAM
        CargardgResolucionesEPAM (EPAMConsolidado.dgEPAMConsolidado.TextMatrix(i, 0))
        i = ""
    End If
        
End Sub

Public Sub GenerarRetencion()

    Dim x As String

    If ValidarRetencion = True Then
        Set rstRegistroIcaro = New ADODB.Recordset
        With CargaRetencion
            rstRegistroIcaro.Open "Retenciones", dbIcaro, adOpenForwardOnly, adLockOptimistic
            rstRegistroIcaro.AddNew
            x = Format(.txtComprobante.Text, "00000") & "/" & Right(.txtEjercicio.Text, 2)
            rstRegistroIcaro!Comprobante = x
            rstRegistroIcaro!Codigo = .txtRetencion.Text
            rstRegistroIcaro!Importe = De_Txt_a_Num_01(.txtMonto.Text)
            rstRegistroIcaro.Update
        End With
        MsgBox "Los datos fueron agregados"
        rstRegistroIcaro.Close
        Set rstRegistroIcaro = Nothing
        Unload CargaRetencion
        ListadoImputacion.Show
        ConfigurardgListadoGasto
        Call CargardgListadoGasto(, "20" & Right(x, 2), x)
        ConfigurardgImputacion
        CargardgImputacion (x)
        ConfigurardgRetencion
        CargardgRetencion (x)
        ListadoImputacion.txtFecha.Text = "20" & Right(x, 2)
        x = 0
    End If
    
End Sub

Public Sub GenerarEstructuraProgramatica()

    Dim strImputacion As String

    If ValidarEstructuraProgramatica = True Then
        Set rstRegistroIcaro = New ADODB.Recordset
        
        If Trim(CargaEstructuraProgramatica.txtSubprograma.Text) = "" Then
            With CargaEstructuraProgramatica
                rstRegistroIcaro.Open "PROGRAMAS", dbIcaro, adOpenForwardOnly, adLockOptimistic
                rstRegistroIcaro.AddNew
                strImputacion = Format(.txtPrograma.Text, "00")
                rstRegistroIcaro!Programa = strImputacion
                If Trim(.txtDescripcionPrograma.Text) <> "" Then
                    rstRegistroIcaro!Descripcion = .txtDescripcionPrograma.Text
                Else
                    rstRegistroIcaro!Descripcion = ""
                End If
                rstRegistroIcaro.Update
            End With
            GoTo Finalizar
        End If
        
        If Trim(CargaEstructuraProgramatica.txtProyecto.Text) = "" Then
            With CargaEstructuraProgramatica
                rstRegistroIcaro.Open "SUBPROGRAMAS", dbIcaro, adOpenForwardOnly, adLockOptimistic
                rstRegistroIcaro.AddNew
                strImputacion = .txtPrograma.Text & "-" & Format(.txtSubprograma.Text, "00")
                rstRegistroIcaro!Programa = .txtPrograma.Text
                rstRegistroIcaro!Subprograma = strImputacion
                If Trim(.txtDescripcionSubprograma.Text) <> "" Then
                    rstRegistroIcaro!Descripcion = .txtDescripcionSubprograma.Text
                Else
                    rstRegistroIcaro!Descripcion = ""
                End If
                rstRegistroIcaro.Update
            End With
            GoTo Finalizar
        End If
        
        If Trim(CargaEstructuraProgramatica.txtActividad.Text) = "" Then
            With CargaEstructuraProgramatica
                rstRegistroIcaro.Open "PROYECTOS", dbIcaro, adOpenForwardOnly, adLockOptimistic
                rstRegistroIcaro.AddNew
                strImputacion = .txtPrograma.Text & "-" & .txtSubprograma.Text
                rstRegistroIcaro!Subprograma = strImputacion
                strImputacion = .txtPrograma.Text & "-" & .txtSubprograma.Text & "-" & Format(.txtProyecto.Text, "00")
                rstRegistroIcaro!Proyecto = strImputacion
                If Trim(.txtDescripcionProyecto.Text) <> "" Then
                    rstRegistroIcaro!Descripcion = .txtDescripcionProyecto.Text
                Else
                    rstRegistroIcaro!Descripcion = ""
                End If
                rstRegistroIcaro.Update
            End With
            GoTo Finalizar
        End If
        
        If Trim(CargaEstructuraProgramatica.txtActividad.Text) = "" Then
            With CargaEstructuraProgramatica
                rstRegistroIcaro.Open "PROYECTOS", dbIcaro, adOpenForwardOnly, adLockOptimistic
                rstRegistroIcaro.AddNew
                strImputacion = .txtPrograma.Text & "-" & .txtSubprograma.Text
                rstRegistroIcaro!Subprograma = strImputacion
                strImputacion = .txtPrograma.Text & "-" & .txtSubprograma.Text & "-" & Format(.txtProyecto.Text, "00")
                rstRegistroIcaro!Proyecto = strImputacion
                If Trim(.txtDescripcionProyecto.Text) <> "" Then
                    rstRegistroIcaro!Descripcion = .txtDescripcionProyecto.Text
                Else
                    rstRegistroIcaro!Descripcion = ""
                End If
                rstRegistroIcaro.Update
            End With
            GoTo Finalizar
        End If

        With CargaEstructuraProgramatica
            rstRegistroIcaro.Open "ACTIVIDADES", dbIcaro, adOpenForwardOnly, adLockOptimistic
            rstRegistroIcaro.AddNew
            strImputacion = .txtPrograma.Text & "-" & .txtSubprograma.Text & "-" & .txtProyecto.Text
            rstRegistroIcaro!Proyecto = strImputacion
            strImputacion = .txtPrograma.Text & "-" & .txtSubprograma.Text & "-" & .txtProyecto.Text & "-" & Format(.txtActividad.Text, "00")
            rstRegistroIcaro!Actividad = strImputacion
            If Trim(.txtDescripcionActividad.Text) <> "" Then
                rstRegistroIcaro!Descripcion = .txtDescripcionActividad.Text
            Else
                rstRegistroIcaro!Descripcion = ""
            End If
            rstRegistroIcaro.Update
        End With
        GoTo Finalizar
        
    End If
    
    Exit Sub

Finalizar:
    MsgBox "Los datos fueron agregados"
    rstRegistroIcaro.Close
    Set rstRegistroIcaro = Nothing
    strImputacion = ""
    Unload CargaEstructuraProgramatica
    EstructuraProgramatica.Show
    ConfigurardgEstructura
    Call CargardgEstructura
    
End Sub
