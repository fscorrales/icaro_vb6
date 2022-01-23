Attribute VB_Name = "Validacion"
Function ValidarComprobante() As Boolean

    Dim strComprobante As String

    With CargaGasto
    
        If Trim(.txtComprobante.Text) = "" Or IsNumeric(.txtComprobante.Text) = False Or Len(.txtComprobante.Text) > "5" Then
            MsgBox "Debe ingresar un Nro de Comprobante de hasta 5 cifras", vbCritical + vbOKOnly, "NRO COMPROBANTE INCORRECTO"
            .txtComprobante.SetFocus
            ValidarComprobante = False
            Exit Function
        End If
        
        If Left(.txtFecha.Text, 2) = "" Or IsDate(.txtFecha.Text) = False Or Len(.txtFecha.Text) <> "10" Then
            MsgBox "Debe ingresar una fecha valida", vbCritical + vbOKOnly, "FECHA INCORRECTA"
            .txtFecha.SetFocus
            ValidarComprobante = False
            Exit Function
        End If

        strComprobante = Format(.txtComprobante.Text, "00000") & "/" & Right(.txtFecha.Text, 2)
        If strEditandoComprobante <> "" And strComprobante <> strEditandoComprobante Then
            If SQLNoMatch("Select * from CARGA Where COMPROBANTE= '" & strComprobante & "'") = False Then
                MsgBox "El Nro de Comprobante ya existe, verifique el Nro", vbCritical + vbOKOnly, "COMPROBANTE DUPLICADO"
                .txtComprobante.SetFocus
                ValidarComprobante = False
                Exit Function
            End If
        ElseIf strEditandoComprobante = "" Then
            If SQLNoMatch("Select * from CARGA Where COMPROBANTE= '" & strComprobante & "'") = False Then
                MsgBox "El Nro de Comprobante ya existe, verifique el Nro", vbCritical + vbOKOnly, "COMPROBANTE DUPLICADO"
                .txtComprobante.SetFocus
                ValidarComprobante = False
                Exit Function
            End If
        End If
        strComprobante = ""
        
        If Trim(.txtCUIT.Text) = "" Or IsNumeric(.txtCUIT.Text) = False Then
            MsgBox "Debe ingresar una numero de CUIT valido", vbCritical + vbOKOnly, "DNI/CUIT INCORRECTO"
            .txtCUIT.SetFocus
            ValidarComprobante = False
            Exit Function
        End If

        If Trim(.txtImporte.Text) = "" Or IsNumeric(.txtImporte.Text) = False Then
            MsgBox "Debe ingresar un Importe del comprobante valido", vbCritical + vbOKOnly, "IMPORTE INCORRECTO"
            .txtImporte.SetFocus
            ValidarComprobante = False
            Exit Function
        End If
        
        If Trim(.txtFR.Text) = "" Or IsNumeric(.txtFR.Text) = False Then
            MsgBox "No puede estar vacio el campo, debe ingresar un número", vbCritical + vbOKOnly, "FONDO DE REPARO INCORRECTO"
            .txtFR.SetFocus
            ValidarComprobante = False
            Exit Function
        End If
        
        If Trim(.txtCertificado.Text) = "" Then
            MsgBox "No puede estar vacio el campo, debe ingresar un número", vbCritical + vbOKOnly, "CERTIFICADO INCORRECTO"
            .txtCertificado.SetFocus
            ValidarComprobante = False
            Exit Function
        End If
        
        If Trim(.txtAvance.Text) = "" Or IsNumeric(.txtAvance.Text) = False Then
            MsgBox "No puede estar vacio el campo, debe ingresar un número", vbCritical + vbOKOnly, "AVANCE DE OBRA INCORRECTO"
            .txtAvance.SetFocus
            ValidarComprobante = False
            Exit Function
        End If
        If De_Txt_a_Num_01(.txtAvance.Text) > 100 Or De_Txt_a_Num_01(.txtAvance.Text) < 0 Then
            MsgBox "Debe ingresar un numero entre 0 y 100", vbCritical + vbOKOnly, "AVANCE DE OBRA INCORRECTO"
            .txtAvance.SetFocus
            ValidarComprobante = False
            Exit Function
        End If
        
        If Trim(.txtCuenta.Text) = "" Then
            MsgBox "No puede estar vacio el campo, debe ingresar un número", vbCritical + vbOKOnly, "CUENTA INCORRECTA"
            .txtCuenta.SetFocus
            ValidarComprobante = False
            Exit Function
        End If
        
        If Trim(.txtFuente.Text) = "" Or IsNumeric(.txtFuente.Text) = False Then
            MsgBox "No puede estar vacio el campo, debe ingresar un número", vbCritical + vbOKOnly, "FUENTE INCORRECTA"
            .txtFuente.SetFocus
            ValidarComprobante = False
            Exit Function
        End If

    End With
    
    ValidarComprobante = True
    
End Function

Public Function ValidarProveedor() As Boolean
    
    With CargaCUIT
    
        If Trim(.txtCUIT.Text) = "" Or IsNumeric(.txtCUIT.Text) = False Then
            MsgBox "Debe ingresar una numero de CUIT valido", vbCritical + vbOKOnly, "DNI/CUIT INCORRECTO"
            .txtCUIT.SetFocus
            ValidarProveedor = False
            Exit Function
        End If
        If strEditandoProveedor <> "" And .txtCUIT.Text <> strEditandoProveedor Then
            If SQLNoMatch("Select * from CUIT Where CUIT= '" & .txtCUIT.Text & "'") = False Then
                MsgBox "Debe ingresar un Nro de CUIT ÚNICO", vbCritical + vbOKOnly, "NRO CUIT DUPLICADO"
                .txtCUIT.SetFocus
                ValidarProveedor = False
                Exit Function
            End If
        ElseIf strEditandoProveedor = "" Then
            If SQLNoMatch("Select * from CUIT Where CUIT= '" & .txtCUIT.Text & "'") = False Then
                MsgBox "Debe ingresar un Nro de CUIT ÚNICO", vbCritical + vbOKOnly, "NRO CUIT DUPLICADO"
                .txtCUIT.SetFocus
                ValidarProveedor = False
                Exit Function
            End If
        End If

        If Trim(.txtDescripcion.Text) = "" Or IsNumeric(.txtDescripcion.Text) = True Or Len(Trim(.txtDescripcion.Text)) > 50 Then
            MsgBox "Debe ingresar el Nombre del Proveedor en forma correcta con menos de 50 caracteres", vbCritical + vbOKOnly, "DESCRIPCIÓN INCORRECTA"
            .txtDescripcion.SetFocus
            ValidarProveedor = False
            Exit Function
        End If
        If strEditandoProveedor <> "" Then
            Set rstRegistroIcaro = New ADODB.Recordset
            SQL = "Select * from CUIT Where CUIT = " & "'" & strEditandoProveedor & "'"
            rstRegistroIcaro.Open SQL, dbIcaro, adOpenForwardOnly, adLockReadOnly
            SQL = ""
            If .txtDescripcion.Text <> rstRegistroIcaro!Nombre Then
                If SQLNoMatch("Select * from CUIT Where NOMBRE= '" & .txtDescripcion.Text & "'") = False Then
                    MsgBox "Debe ingresar un Nombre de PROVEEDOR ÚNICO", vbCritical + vbOKOnly, "NOMBRE DEL PROVEEDOR DUPLICADO"
                    .txtDescripcion.SetFocus
                    rstRegistroIcaro.Close
                    Set rstRegistroIcaro = Nothing
                    ValidarProveedor = False
                    Exit Function
                End If
            End If
            rstRegistroIcaro.Close
            Set rstRegistroIcaro = Nothing
        ElseIf strEditandoProveedor = "" Then
            If SQLNoMatch("Select * from CUIT Where NOMBRE = '" & .txtDescripcion.Text & "'") = False Then
                MsgBox "Debe ingresar un Nombre de PROVEEDOR ÚNICO", vbCritical + vbOKOnly, "NOMBRE DEL PROVEEDOR DUPLICADO"
                .txtDescripcion.SetFocus
                ValidarProveedor = False
                Exit Function
            End If
        End If
        
    End With
    ValidarProveedor = True
    
End Function

'Public Function ValidarCuentaBancaria() As Boolean
'
'    Dim CuentaConGuiones As String
'
'    With CargaCuentaBancaria
'
'        If Len(.txtCuentaBancaria.Text) <> "9" Or IsNumeric(.txtCuentaBancaria.Text) = False Then
'            MsgBox "Debe ingresar una número de Cuenta Bancaria válido de 9 cifras y SIN GUIONES", vbCritical + vbOKOnly, "NÚMERO CUENTA BANCARIA INCORRECTO"
'            .txtCuentaBancaria.SetFocus
'            ValidarCuentaBancaria = False
'            Exit Function
'        End If
'        CuentaConGuiones = Left(.txtCuentaBancaria.Text, 2) & "-" & Mid(.txtCuentaBancaria.Text, 3, 1) & "-" & Mid(.txtCuentaBancaria.Text, 4, 5) & "-" & Right(.txtCuentaBancaria.Text, 1)
'        If strEditandoCuentaBancaria <> "" And .txtCuentaBancaria.Text <> strEditandoCuentaBancaria Then
'            If SQLNoMatch("Select * from CUENTAS Where CUENTA = '" & CuentaConGuiones & "'") = False Then
'                MsgBox "La CUENTA BANCARIA ya existe, verifique el Nro", vbCritical + vbOKOnly, "CUENTA BANCARIA DUPLICADA"
'                .txtCuentaBancaria.SetFocus
'                ValidarCuentaBancaria = False
'                Exit Function
'            End If
'        ElseIf strEditandoCuentaBancaria = "" Then
'            If SQLNoMatch("Select * from CUENTAS Where CUENTA = '" & CuentaConGuiones & "'") = False Then
'                MsgBox "La CUENTA BANCARIA ya existe, verifique el Nro", vbCritical + vbOKOnly, "CUENTA BANCARIA DUPLICADA"
'                .txtCuentaBancaria.SetFocus
'                ValidarCuentaBancaria = False
'                Exit Function
'            End If
'        End If
'
'        If Trim(.txtDescripcion.Text) = "" Or IsNumeric(.txtDescripcion.Text) = True Then
'            MsgBox "Debe ingresar la descripción de la Cuenta Bancaria en forma correcta", vbCritical + vbOKOnly, "DESCRIPCIÓN INCORRECTA"
'            .txtDescripcion.SetFocus
'            ValidarCuentaBancaria = False
'            Exit Function
'        End If
'
'    End With
'    ValidarCuentaBancaria = True
'
'End Function

Public Function ValidarCuentaBancaria() As Boolean
    
    Dim CuentaSinGuiones As String
    
    With CargaCuentaBancaria
           
        CuentaSinGuiones = Replace(.txtCuentaBancaria.Text, "-", "")
    
        If Len(.txtCuentaBancaria.Text) > "12" Or IsNumeric(CuentaSinGuiones) = False Then
            MsgBox "Debe ingresar una número de Cuenta Bancaria válido de hasta 12 cifras", vbCritical + vbOKOnly, "NÚMERO CUENTA BANCARIA INCORRECTO"
            .txtCuentaBancaria.SetFocus
            ValidarCuentaBancaria = False
            Exit Function
        End If
        If strEditandoCuentaBancaria <> "" And .txtCuentaBancaria.Text <> strEditandoCuentaBancaria Then
            If SQLNoMatch("Select * from CUENTAS Where CUENTA = '" & .txtCuentaBancaria.Text & "'") = False Then
                MsgBox "La CUENTA BANCARIA ya existe, verifique el Nro", vbCritical + vbOKOnly, "CUENTA BANCARIA DUPLICADA"
                .txtCuentaBancaria.SetFocus
                ValidarCuentaBancaria = False
                Exit Function
            End If
        ElseIf strEditandoCuentaBancaria = "" Then
            If SQLNoMatch("Select * from CUENTAS Where CUENTA = '" & .txtCuentaBancaria.Text & "'") = False Then
                MsgBox "La CUENTA BANCARIA ya existe, verifique el Nro", vbCritical + vbOKOnly, "CUENTA BANCARIA DUPLICADA"
                .txtCuentaBancaria.SetFocus
                ValidarCuentaBancaria = False
                Exit Function
            End If
        End If

        If Trim(.txtDescripcion.Text) = "" Or IsNumeric(.txtDescripcion.Text) = True Then
            MsgBox "Debe ingresar la descripción de la Cuenta Bancaria en forma correcta", vbCritical + vbOKOnly, "DESCRIPCIÓN INCORRECTA"
            .txtDescripcion.SetFocus
            ValidarCuentaBancaria = False
            Exit Function
        End If
        
    End With
    ValidarCuentaBancaria = True
    
End Function

Public Function ValidarPartida() As Boolean
    
    With CargaPartida
    
        If Len(.txtPartida.Text) <> "3" Or IsNumeric(.txtPartida.Text) = False Then
            MsgBox "Debe ingresar una numero de PARTIDA válido de 3 cifras", vbCritical + vbOKOnly, "CÓDIGO PARTIDA INCORRECTO"
            .txtPartida.SetFocus
            ValidarPartida = False
            Exit Function
        End If
        If strEditandoPartida <> "" And .txtPartida.Text <> strEditandoPartida Then
            If SQLNoMatch("Select * from PARTIDAS Where PARTIDA= '" & .txtPartida.Text & "'") = False Then
                MsgBox "La PARTIDA ya existe, verifique el Nro", vbCritical + vbOKOnly, "PARTIDA DUPLICADA"
                .txtPartida.SetFocus
                ValidarPartida = False
                Exit Function
            End If
        ElseIf strEditandoPartida = "" Then
            If SQLNoMatch("Select * from PARTIDAS Where PARTIDA= '" & .txtPartida.Text & "'") = False Then
                MsgBox "La PARTIDA ya existe, verifique el Nro", vbCritical + vbOKOnly, "PARTIDA DUPLICADA"
                .txtPartida.SetFocus
                ValidarPartida = False
                Exit Function
            End If
        End If

        If Trim(.txtDescripcion.Text) = "" Or IsNumeric(.txtDescripcion.Text) = True Then
            MsgBox "Debe ingresar la descripción de la Partida en forma correcta", vbCritical + vbOKOnly, "DESCRIPCIÓN INCORRECTA"
            .txtDescripcion.SetFocus
            ValidarPartida = False
            Exit Function
        End If
        
    End With
    ValidarPartida = True
    
End Function

Public Function ValidarFuente() As Boolean
    
    With CargaFuente
    
        If Len(.txtFuente.Text) <> "2" Or IsNumeric(.txtFuente.Text) = False Then
            MsgBox "Debe ingresar una numero de FUENTE válido de 2 cifras", vbCritical + vbOKOnly, "CÓDIGO FUENTE INCORRECTO"
            .txtFuente.SetFocus
            ValidarFuente = False
            Exit Function
        End If
        If strEditandoFuente <> "" And .txtFuente.Text <> strEditandoFuente Then
            If SQLNoMatch("Select * from FUENTES Where FUENTE= '" & .txtFuente.Text & "'") = False Then
                MsgBox "La FUENTE ya existe, verifique el Nro", vbCritical + vbOKOnly, "FUENTE DUPLICADA"
                .txtFuente.SetFocus
                ValidarFuente = False
                Exit Function
            End If
        ElseIf strEditandoFuente = "" Then
            If SQLNoMatch("Select * from FUENTES Where FUENTE= '" & .txtFuente.Text & "'") = False Then
                MsgBox "La FUENTE ya existe, verifique el Nro", vbCritical + vbOKOnly, "FUENTE DUPLICADA"
                .txtFuente.SetFocus
                ValidarFuente = False
                Exit Function
            End If
        End If

        If Trim(.txtDescripcion.Text) = "" Or IsNumeric(.txtDescripcion.Text) = True Then
            MsgBox "Debe ingresar la descripción de la Fuente en forma correcta", vbCritical + vbOKOnly, "DESCRIPCIÓN INCORRECTA"
            .txtDescripcion.SetFocus
            ValidarFuente = False
            Exit Function
        End If
        
    End With
    ValidarFuente = True
    
End Function

Public Function ValidarCodigoRetencion() As Boolean
    
    With CodigoRetencion
    
        If Len(.txtRetencion.Text) <> "3" Or IsNumeric(.txtRetencion.Text) = False Then
            MsgBox "Debe ingresar una numero de RETENCIÓN válido de 3 cifras", vbCritical + vbOKOnly, "CÓDIGO RETENCIÓN INCORRECTO"
            .txtRetencion.SetFocus
            ValidarCodigoRetencion = False
            Exit Function
        End If
        If strEditandoCodigoRetencion <> "" And .txtRetencion.Text <> strEditandoCodigoRetencion Then
            If SQLNoMatch("Select * from CODIGOSRETENCIONES Where CODIGO= '" & .txtRetencion.Text & "'") = False Then
                MsgBox "La RETENCION ya existe, verifique el Nro", vbCritical + vbOKOnly, "RETENCION DUPLICADA"
                .txtRetencion.SetFocus
                ValidarCodigoRetencion = False
                Exit Function
            End If
        ElseIf strEditandoCodigoRetencion = "" Then
            If SQLNoMatch("Select * from CODIGOSRETENCIONES Where CODIGO= '" & .txtRetencion.Text & "'") = False Then
                MsgBox "La RETENCION ya existe, verifique el Nro", vbCritical + vbOKOnly, "RETENCION DUPLICADA"
                .txtRetencion.SetFocus
                ValidarCodigoRetencion = False
                Exit Function
            End If
        End If

        If Trim(.txtDescripcion.Text) = "" Or IsNumeric(.txtDescripcion.Text) = True Then
            MsgBox "Debe ingresar la descripción de la RETENCION en forma correcta", vbCritical + vbOKOnly, "DESCRIPCIÓN INCORRECTA"
            .txtDescripcion.SetFocus
            ValidarCodigoRetencion = False
            Exit Function
        End If
        If strEditandoCodigoRetencion <> "" Then
            Set rstRegistroIcaro = New ADODB.Recordset
            SQL = "Select * from CODIGOSRETENCIONES Where CODIGO = " & "'" & strEditandoCodigoRetencion & "'"
            rstRegistroIcaro.Open SQL, dbIcaro, adOpenForwardOnly, adLockReadOnly
            SQL = ""
            If .txtDescripcion.Text <> rstRegistroIcaro!Descripcion Then
                If SQLNoMatch("Select * from CODIGOSRETENCIONES Where DESCRIPCION = '" & .txtDescripcion.Text & "'") = False Then
                    MsgBox "La RETENCION ya existe, verifique el Nombre / Descripción de la misma", vbCritical + vbOKOnly, "RETENCIÒN DUPLICADA"
                    .txtDescripcion.SetFocus
                    rstRegistroIcaro.Close
                    Set rstRegistroIcaro = Nothing
                    ValidarCodigoRetencion = False
                    Exit Function
                End If
            End If
            rstRegistroIcaro.Close
            Set rstRegistroIcaro = Nothing
        ElseIf strEditandoCodigoRetencion = "" Then
            If SQLNoMatch("Select * from CODIGOSRETENCIONES Where DESCRIPCION = '" & .txtDescripcion.Text & "'") = False Then
                MsgBox "La RETENCION ya existe, verifique el Nombre / Descripción de la misma", vbCritical + vbOKOnly, "RETENCIÒN DUPLICADA"
                .txtDescripcion.SetFocus
                ValidarCodigoRetencion = False
                Exit Function
            End If
        End If
        
    End With
    ValidarCodigoRetencion = True
   
End Function

Public Function ValidarLocalidad() As Boolean
    
    With CargaLocalidad
    
        If Trim(.txtLocalidad.Text) = "" Or IsNumeric(.txtLocalidad.Text) = True Then
            MsgBox "Debe ingresar la descripción de la Localidad en forma correcta", vbCritical + vbOKOnly, "LOCALIDAD INCORRECTA"
            .txtLocalidad.SetFocus
            ValidarLocalidad = False
            Exit Function
        End If
        If strEditandoLocalidad <> "" And .txtLocalidad.Text <> strEditandoLocalidad Then
            If SQLNoMatch("Select * from LOCALIDADES Where LOCALIDAD= '" & .txtLocalidad.Text & "'") = False Then
                MsgBox "La Localidad ya existe, verifique el Nombre / Descripción de la misma", vbCritical + vbOKOnly, "LOCALIDAD DUPLICADA"
                .txtLocalidad.SetFocus
                ValidarLocalidad = False
                Exit Function
            End If
        ElseIf strEditandoLocalidad = "" Then
            If SQLNoMatch("Select * from LOCALIDADES Where LOCALIDAD= '" & .txtLocalidad.Text & "'") = False Then
                MsgBox "La Localidad ya existe, verifique el Nombre / Descripción de la misma", vbCritical + vbOKOnly, "LOCALIDAD DUPLICADA"
                .txtLocalidad.SetFocus
                ValidarLocalidad = False
                Exit Function
            End If
        End If

    End With
    ValidarLocalidad = True
    
End Function

Public Function ValidarCUITCargaGasto(CUIT As String) As Boolean

    If CUIT = "" Then
        MsgBox "Debe ingresar un Nro de CUIT", vbCritical + vbOKOnly, "CUIT NULO"
        CargaGasto.txtCUIT.SetFocus
        ValidarCUITCargaGasto = False
        Exit Function
    Else
        If IsNumeric(CUIT) = False Then
            MsgBox "Debe Ingresar un número sin guiones", vbCritical + vbOKOnly, "CUIT NO NUMÉRICO"
            CargaGasto.txtCUIT.SetFocus
            ValidarCUITCargaGasto = False
            Exit Function
        Else
            If SQLNoMatch("Select * from CUIT Where CUIT= '" & CUIT & "'") = True Then
                Agregar = MsgBox("El NRO DE CUIT ingresado no existe en la Base de Datos. Desea Agregar el CUIT " & CUIT & " ?", vbQuestion + vbYesNo, "ADICIONAR CUIT")
                If Agregar = 6 Then
                    bolCUITDesdeCarga = True
                    CargaCUIT.Show
                    CargaCUIT.txtCUIT.Text = CUIT
                    CargaCUIT.txtCUIT.Enabled = False
                    CargaCUIT.txtDescripcion.SetFocus
                    CargaGasto.Hide
                Else
                    CargaGasto.txtCUIT.SetFocus
                End If
                ValidarCUITCargaGasto = False
                Exit Function
            End If
            'If SQLNoMatch("Select * from CUIT Where CUIT= '" & CUIT & "'") = True Then
            '    MsgBox "El CUIT ingresado no existe en la Base de Datos, deberá incorporar el Proveedor antes de seguir adelante", vbOKOnly, "CUIT INEXISTENTE"
            '    CargaGasto.txtCUIT.SetFocus
            '    ValidarCUITCargaGasto = False
            '    Exit Function
            'End If
        End If
    End If
    
    ValidarCUITCargaGasto = True

End Function

Public Function ValidarObraCargaGasto(DescripcionObra As String) As Boolean

    If ValidarCUITCargaGasto(CargaGasto.txtCUIT.Text) = True Then
        If DescripcionObra = "" Then
            MsgBox "Debe ingresar una Descripción de Obra", vbCritical + vbOKOnly, "DESCRIPCION OBRA INEXISTENTE"
            CargaGasto.txtObra.SetFocus
            ValidarObraCargaGasto = False
            Exit Function
        Else
            If SQLNoMatch("Select * from OBRAS Where DESCRIPCION= '" & DescripcionObra & "'") = True Then
                Agregar = MsgBox("La Descripcion de Obra ingresada no existe en la Base de Datos. Desea Agregar la Obra " & DescripcionObra & " ?", vbQuestion + vbYesNo, "ADICIONAR OBRA")
                If Agregar = 6 Then
                    bolObraDesdeCarga = True
                    CargaObra.Show
                    CargaObra.txtCUIT.Text = CargaGasto.txtCUIT.Text
                    CargaObra.txtCUIT.Enabled = False
                    CargaObra.txtDescripcionCUIT.Text = CargaGasto.txtDescripcionCUIT.Text
                    CargaObra.txtDescripcionCUIT.Enabled = False
                    CargaObra.txtObra.Text = CargaGasto.txtObra.Text
                    CargaObra.txtObra.Enabled = False
                    CargaObra.txtPrograma.SetFocus
                    CargaGasto.Hide
                Else
                    CargaGasto.txtCUIT.SetFocus
                End If
                ValidarObraCargaGasto = False
                Exit Function
            End If
            'If SQLNoMatch("Select * from OBRAS Where DESCRIPCION = '" & DescripcionObra & "'") = True Then
            '    MsgBox "La Descripcion de Obra ingresada no existe en la Base de Datos, deberá incorporar la misma para de seguir adelante", vbOKOnly, "DESCRIPCION OBRA INEXISTENTE"
            '    CargaGasto.txtObra.SetFocus
            '    ValidarObraCargaGasto = False
            '    Exit Function
            'End If
        End If
    Else
        ValidarObraCargaGasto = False
        Exit Function
    End If
    
    ValidarObraCargaGasto = True

End Function

Public Function ValidarCuentaCargaGasto(Cuenta As String) As Boolean

    If Cuenta = "" Then
        MsgBox "Debe ingresar un Nro de Cuenta", vbCritical + vbOKOnly, "CUENTA NULA"
        CargaGasto.txtCuenta.SetFocus
        ValidarCuentaCargaGasto = False
        Exit Function
    Else
        If SQLNoMatch("Select * from CUENTAS Where CUENTA= '" & Cuenta & "'") = True Then
            MsgBox "La Cuenta ingresadada no existe en la Base de Datos, deberá incorporar la Cuenta antes de seguir adelante", vbOKOnly, "CUENTA INEXISTENTE"
            CargaGasto.txtCuenta.SetFocus
            ValidarCuentaCargaGasto = False
            Exit Function
        End If
    End If
    
    ValidarCuentaCargaGasto = True

End Function

Public Function ValidarFuenteCargaGasto(Fuente As String) As Boolean

    If Fuente = "" Then
        MsgBox "Debe ingresar un Nro de Fuente", vbCritical + vbOKOnly, "FUENTE NULA"
        CargaGasto.txtFuente.SetFocus
        ValidarFuenteCargaGasto = False
        Exit Function
    Else
        If IsNumeric(Fuente) = False Then
            MsgBox "Debe Ingresar un valor numérico", vbCritical + vbOKOnly, "FUENTE NO NUMÉRICA"
            CargaGasto.txtFuente.SetFocus
            ValidarFuenteCargaGasto = False
            Exit Function
        Else
            If SQLNoMatch("Select * from FUENTES Where FUENTE= '" & Fuente & "'") = True Then
                MsgBox "La Fuente ingresada no existe en la Base de Datos, deberá incorporarla antes de seguir adelante", vbOKOnly, "FUENTE INEXISTENTE"
                CargaGasto.txtFuente.SetFocus
                ValidarFuenteCargaGasto = False
                Exit Function
            End If
        End If
    End If
    
    ValidarFuenteCargaGasto = True

End Function


Public Function ValidarCargaGasto() As Boolean

    If Trim(txtComprobante.Text) = "" Or IsNumeric(txtComprobante.Text) = False Or Len(txtComprobante.Text) > "4" Then
        MsgBox "Debe ingresar un Nro de Comprobante de hasta 4 cifras", vbCritical + vbOKOnly, "NRO COMPROBANTE INCORRECTO"
        txtComprobante.SetFocus
        ValidarCargaGasto = False
        Exit Function
    End If
    If Trim(txtEjercicio.Text) = "" Or IsNumeric(txtEjercicio.Text) = False Or Len(txtEjercicio.Text) <> "4" Then
        MsgBox "Debe ingresar el Ejercicio al cual corresponde el comprobante con 4 cifras", vbCritical + vbOKOnly, "NRO EJERCICIO INCORRECTO"
        txtEjercicio.SetFocus
        ValidarCargaGasto = False
        Exit Function
    ElseIf Format(Date, "yyyy") <> txtEjercicio.Text Then
        Dim Ejercicio As Integer
        Ejercicio = MsgBox("Esta seguro que desea imputar el comprobante al ejercicio " & txtEjercicio.Text, vbYesNo + vbQuestion, "EJERCICIO DE IMPUTACION")
        If Ejercicio = vbNo Then
            txtEjercicio.SetFocus
            ValidarCargaGasto = False
            Exit Function
        End If
    End If
    If bolEditandoComprobante = False Then
        rstCargaComprobante.FindFirst "Comprobante=" & "'" & Format(txtComprobante.Text, "0000") & "/" & Right(txtEjercicio, 2) & "'"
        If rstCargaComprobante.NoMatch = False Then
            MsgBox "El comprobante ya existe, verifique el Nro", vbCritical + vbOKOnly, "COMPROBANTE DUPLICADO"
            txtComprobante.SetFocus
            ValidarCargaGasto = False
            Exit Function
        End If
    End If
    If Trim(txtFecha.Text) = "" Or IsDate(txtFecha.Text) = False Then
        MsgBox "Debe ingresar una fecha valida", vbCritical + vbOKOnly, "FECHA INCORRECTA"
        txtFecha.SetFocus
        ValidarCargaGasto = False
        Exit Function
    End If
    If Trim(txtCUIT.Text) = "" Or IsNumeric(txtCUIT.Text) = False Then
        MsgBox "Debe ingresar una numero de CUIT valido", vbCritical + vbOKOnly, "DNI/CUIT INCORRECTO"
        txtCUIT.SetFocus
        ValidarCargaGasto = False
        Exit Function
    End If
    With rstCUIT
        .Index = "pkCUIT"
        .Seek "=", txtCUIT.Text
    End With
    If Trim(txtImporte.Text) = "" Or IsNumeric(txtImporte.Text) = False Then
        MsgBox "Debe ingresar un Importe del comprobante valido", vbCritical + vbOKOnly, "IMPORTE INCORRECTO"
        txtImporte.SetFocus
        ValidarCargaGasto = False
        Exit Function
    End If
    If Trim(txtFR.Text) = "" Or IsNumeric(txtFR.Text) = False Then
        MsgBox "No puede estar vacio el campo, debe ingresar un número", vbCritical + vbOKOnly, "FONDO DE REPARO INCORRECTO"
        txtFR.SetFocus
        ValidarCargaGasto = False
        Exit Function
    End If
    If Trim(txtAvance.Text) = "" Or IsNumeric(txtAvance.Text) = False Then
        MsgBox "No puede estar vacio el campo, debe ingresar un número", vbCritical + vbOKOnly, "AVANCE DE OBRA INCORRECTO"
        txtFR.SetFocus
        ValidarCargaGasto = False
        Exit Function
    End If

    ValidarCargaGasto = True
End Function

Public Function ValidarCUITCargaObra(CUIT As String) As Boolean

    If CUIT = "" Then
        MsgBox "Debe ingresar un Nro de CUIT", vbCritical + vbOKOnly, "CUIT NULO"
        CargaObra.txtDescripcionCUIT.Text = ""
        CargaObra.txtCUIT.SetFocus
        ValidarCUITCargaObra = False
        Exit Function
    Else
        If IsNumeric(CUIT) = False Then
            MsgBox "Debe Ingresar un número sin guiones", vbCritical + vbOKOnly, "CUIT NO NUMÉRICO"
            CargaObra.txtDescripcionCUIT.Text = ""
            CargaObra.txtCUIT.SetFocus
            ValidarCUITCargaObra = False
            Exit Function
        Else
            If SQLNoMatch("Select * from CUIT Where CUIT= '" & CUIT & "'") = True Then
                MsgBox "El CUIT ingresado no existe en la Base de Datos, deberá incorporar el Proveedor antes de seguir adelante", vbOKOnly, "CUIT INEXISTENTE"
                CargaObra.txtDescripcionCUIT.Text = ""
                CargaObra.txtCUIT.SetFocus
                ValidarCUITCargaObra = False
                Exit Function
            End If
        End If
    End If
    
    ValidarCUITCargaObra = True

End Function

Public Function ValidarProgramaCargaObra(Programa As String) As Boolean

    Dim Agregar As Integer
    
    If Programa = "" Then
        MsgBox "Debe ingresar un Nro de PROGRAMA", vbCritical + vbOKOnly, "PROGRAMA NULO"
        CargaObra.txtPrograma.SetFocus
        ValidarProgramaCargaObra = False
        Exit Function
    Else
        If IsNumeric(Programa) = False Or Len(Trim(Programa)) <> "2" Then
            MsgBox "Debe Ingresar un número de 2 cifras", vbCritical + vbOKOnly, "NRO PROGRAMA ERRONEO"
            CargaObra.txtPrograma.SetFocus
            ValidarProgramaCargaObra = False
            Exit Function
        Else
            If SQLNoMatch("Select * from PROGRAMAS Where PROGRAMA= '" & Programa & "'") = True Then
                Agregar = MsgBox("El NRO DE PROGRAMA ingresado no existe en la Base de Datos. Desea Agregar el PROGRAMA " & Programa & " ?", vbQuestion + vbYesNo, "ACTUALIZAR ESTRUCTURA")
                If Agregar = 6 Then
                    bolActualizarEstructura = True
                    CargaObra.txtDescripcionPrograma.Enabled = True
                    CargaObra.txtDescripcionPrograma.Text = "ESCRIBA AQUÍ EL NOMBRE DEL PROGRAMA"
                    CargaObra.txtDescripcionPrograma.SetFocus
                Else
                    CargaObra.txtPrograma.SetFocus
                End If
                ValidarProgramaCargaObra = False
                Exit Function
            End If
        End If
    End If
    
    ValidarProgramaCargaObra = True

End Function

Public Function ValidarSubprogramaCargaObra(Subprograma As String) As Boolean

    Dim Agregar As Integer
    
    If Subprograma = "" Then
        MsgBox "Debe ingresar un Nro de SUBPROGRAMA", vbCritical + vbOKOnly, "SUBPROGRAMA NULO"
        CargaObra.txtSubprograma.SetFocus
        ValidarSubprogramaCargaObra = False
        Exit Function
    Else
        If IsNumeric(Subprograma) = False Or Len(Trim(Subprograma)) <> "2" Then
            MsgBox "Debe Ingresar un número de 2 cifras", vbCritical + vbOKOnly, "NRO SUBPROGRAMA ERRONEO"
            CargaObra.txtSubprograma.SetFocus
            ValidarSubprogramaCargaObra = False
            Exit Function
        Else
            Subprograma = CargaObra.txtPrograma.Text & "-" & Subprograma
            If SQLNoMatch("Select * from SUBPROGRAMAS Where SUBPROGRAMA= '" & Subprograma & "'") = True Then
                Agregar = MsgBox("El NRO DE SUBPROGRAMA ingresado no existe en la Base de Datos. Desea Agregar el SUBPROGRAMA " & Subprograma & " ?", vbQuestion + vbYesNo, "ACTUALIZAR ESTRUCTURA")
                If Agregar = 6 Then
                    bolActualizarEstructura = True
                    CargaObra.txtDescripcionSubprograma.Enabled = True
                    CargaObra.txtDescripcionSubprograma.Text = "ESCRIBA AQUÍ EL NOMBRE DEL SUBPROGRAMA"
                    CargaObra.txtDescripcionSubprograma.SetFocus
                Else
                    CargaObra.txtSubprograma.SetFocus
                End If
                ValidarSubprogramaCargaObra = False
                Exit Function
            End If
        End If
    End If
    
    ValidarSubprogramaCargaObra = True

End Function

Public Function ValidarProyectoCargaObra(Proyecto As String) As Boolean

    Dim Agregar As Integer
    
    If Proyecto = "" Then
        MsgBox "Debe ingresar un Nro de PROYECTO", vbCritical + vbOKOnly, "PROYECTO NULO"
        CargaObra.txtProyecto.SetFocus
        ValidarProyectoCargaObra = False
        Exit Function
    Else
        If IsNumeric(Proyecto) = False Or Len(Trim(Proyecto)) <> "2" Then
            MsgBox "Debe Ingresar un número de 2 cifras", vbCritical + vbOKOnly, "NRO PROYECTO ERRONEO"
            CargaObra.txtProyecto.SetFocus
            ValidarProyectoCargaObra = False
            Exit Function
        Else
            Proyecto = CargaObra.txtPrograma.Text & "-" & CargaObra.txtSubprograma.Text & "-" & Proyecto
            If SQLNoMatch("Select * from PROYECTOS Where PROYECTO= '" & Proyecto & "'") = True Then
                Agregar = MsgBox("El NRO DE PROYECTO ingresado no existe en la Base de Datos. Desea Agregar el PROYECTO " & Proyecto & " ?", vbQuestion + vbYesNo, "ACTUALIZAR ESTRUCTURA")
                If Agregar = 6 Then
                    bolActualizarEstructura = True
                    CargaObra.txtDescripcionProyecto.Enabled = True
                    CargaObra.txtDescripcionProyecto.Text = "ESCRIBA AQUÍ EL NOMBRE DEL PROYECTO"
                    CargaObra.txtDescripcionProyecto.SetFocus
                Else
                    CargaObra.txtProyecto.SetFocus
                End If
                ValidarProyectoCargaObra = False
                Exit Function
            End If
        End If
    End If
    
    ValidarProyectoCargaObra = True

End Function

Public Function ValidarActividadCargaObra(Actividad As String) As Boolean

    Dim Agregar As Integer
    
    If Actividad = "" Then
        MsgBox "Debe ingresar un Nro de ACTIVIDAD", vbCritical + vbOKOnly, "ACTIVIDAD NULO"
        CargaObra.txtActividad.SetFocus
        ValidarActividadCargaObra = False
        Exit Function
    Else
        If IsNumeric(Actividad) = False Or Len(Trim(Actividad)) <> "2" Then
            MsgBox "Debe Ingresar un número de 2 cifras", vbCritical + vbOKOnly, "NRO ACTIVIDAD ERRONEO"
            CargaObra.txtActividad.SetFocus
            ValidarActividadCargaObra = False
            Exit Function
        Else
            Actividad = CargaObra.txtPrograma.Text & "-" & CargaObra.txtSubprograma.Text & "-" & CargaObra.txtProyecto.Text & "-" & Actividad
            If SQLNoMatch("Select * from ACTIVIDADES Where ACTIVIDAD= '" & Actividad & "'") = True Then
                Agregar = MsgBox("El NRO DE ACTIVIDAD ingresado no existe en la Base de Datos. Desea Agregar la ACTIVIDAD " & Actividad & " ?", vbQuestion + vbYesNo, "ACTUALIZAR ESTRUCTURA")
                If Agregar = 6 Then
                    bolActualizarEstructura = True
                    CargaObra.txtDescripcionActividad.Enabled = True
                    CargaObra.txtDescripcionActividad.Text = "ESCRIBA AQUÍ EL NOMBRE DE LA ACTIVIDAD"
                    CargaObra.txtDescripcionActividad.SetFocus
                Else
                    CargaObra.txtActividad.SetFocus
                End If
                ValidarActividadCargaObra = False
                Exit Function
            End If
        End If
    End If
    
    ValidarActividadCargaObra = True

End Function

Public Function ValidarPartidaCargaObra(Partida As String) As Boolean

    Dim Agregar As Integer
    
    If Partida = "" Then
        MsgBox "Debe ingresar un Nro de PARTIDA", vbCritical + vbOKOnly, "PARTIDA NULA"
        CargaObra.txtPartida.SetFocus
        ValidarPartidaCargaObra = False
        Exit Function
    Else
        If IsNumeric(Partida) = False Or Len(Trim(Partida)) <> "3" Then
            MsgBox "Debe Ingresar un número de 3 cifras", vbCritical + vbOKOnly, "NRO PARTUDA ERRONEA"
            CargaObra.txtPartida.SetFocus
            ValidarPartidaCargaObra = False
            Exit Function
        Else
            If SQLNoMatch("Select * from PARTIDAS Where PARTIDA= '" & Partida & "'") = True Then
                MsgBox "La PARTIDA ingresada no existe en la Base de Datos, deberá incorporarla antes de seguir adelante", vbOKOnly, "PARTIDA INEXISTENTE"
                CargaObra.txtPartida.SetFocus
                ValidarPartidaCargaObra = False
                Exit Function
            End If
        End If
    End If
    
    ValidarPartidaCargaObra = True

End Function

Public Function ValidarCuentaCargaObra(Cuenta As String) As Boolean

    If Cuenta = "" Then
        MsgBox "Debe ingresar un Nro de CUENTA", vbCritical + vbOKOnly, "CUENTA NULA"
        CargaObra.txtCuenta.SetFocus
        ValidarCuentaCargaObra = False
        Exit Function
    Else
        If SQLNoMatch("Select * from CUENTAS Where CUENTA= '" & Cuenta & "'") = True Then
            MsgBox "La CUENTA ingresada no existe en la Base de Datos, deberá incorporar la CUENTA antes de seguir adelante", vbOKOnly, "CUENTA INEXISTENTE"
            CargaObra.txtCuenta.SetFocus
            ValidarCuentaCargaObra = False
            Exit Function
        End If
    End If
    
    ValidarCuentaCargaObra = True

End Function

Public Function ValidarFuenteCargaObra(Fuente As String) As Boolean

    If Fuente = "" Then
        MsgBox "Debe ingresar un Nro de FUENTE", vbCritical + vbOKOnly, "FUENTE NULA"
        CargaObra.txtFuente.SetFocus
        ValidarFuenteCargaObra = False
        Exit Function
    Else
        If IsNumeric(Fuente) = False Then
            MsgBox "Debe Ingresar un número sin guiones", vbCritical + vbOKOnly, "FUENTE NO NUMÉRICO"
            CargaObra.txtFuente.SetFocus
            ValidarFuenteCargaObra = False
            Exit Function
        Else
            If SQLNoMatch("Select * from FUENTES Where FUENTE= '" & Fuente & "'") = True Then
                MsgBox "La FUENTE ingresada no existe en la Base de Datos, deberá incorporar la misma antes de seguir adelante", vbOKOnly, "FUENTE INEXISTENTE"
                CargaObra.txtFuente.SetFocus
                ValidarFuenteCargaObra = False
                Exit Function
            End If
        End If
    End If
    
    ValidarFuenteCargaObra = True

End Function

Public Function ValidarObra() As Boolean
    
    Dim strEstructura As String
    
    With CargaObra
    
        If Trim(.txtCUIT.Text) = "" Or IsNumeric(.txtCUIT.Text) = False Then
            MsgBox "Debe ingresar un Nro de CUIT/DNI valido", vbCritical + vbOKOnly, "DNI/CUIT INCORRECTO"
            .txtCUIT.SetFocus
            ValidarObra = False
            Exit Function
        End If
        If SQLNoMatch("Select * from CUIT Where CUIT= '" & .txtCUIT.Text & "'") = True Then
            MsgBox "El CUIT ingresado no existe en la Base de Datos, deberá incorporar el Proveedor antes de seguir adelante", vbCritical + vbOKOnly, "NRO CUIT INEXISTENTE"
            .txtCUIT.SetFocus
            ValidarObra = False
            Exit Function
        End If
                
        If Trim(.txtPrograma.Text) = "" Or IsNumeric(.txtPrograma.Text) = False Or Len(Trim(.txtPrograma.Text)) <> "2" Then
            MsgBox "Debe ingresar un Nro de PROGRAMA de 2 cifras", vbCritical + vbOKOnly, "NRO PROGRAMA INCORRECTO"
            .txtPrograma.SetFocus
            ValidarObra = False
            Exit Function
        End If
        strEstructura = .txtPrograma.Text
        If SQLNoMatch("Select * from PROGRAMAS Where PROGRAMA= '" & strEstructura & "'") = True Then
            MsgBox "El PROGRAMA ingresado no existe en la Base de Datos, deberá incorporar el PROGRAMA antes de seguir adelante", vbCritical + vbOKOnly, "PROGRAMA INEXISTENTE"
            .txtPrograma.SetFocus
            ValidarObra = False
            Exit Function
        End If
        
        If Trim(.txtSubprograma.Text) = "" Or IsNumeric(.txtSubprograma.Text) = False Or Len(Trim(.txtSubprograma.Text)) <> "2" Then
            MsgBox "Debe ingresar un Nro de SUBPROGRAMA de 2 cifras", vbCritical + vbOKOnly, "NRO SUBPROGRAMA INCORRECTO"
            .txtSubprograma.SetFocus
            ValidarObra = False
            Exit Function
        End If
        strEstructura = strEstructura & "-" & .txtSubprograma.Text
        If SQLNoMatch("Select * from SUBPROGRAMAS Where SUBPROGRAMA= '" & strEstructura & "'") = True Then
            MsgBox "El SUBPROGRAMA ingresado no existe en la Base de Datos, deberá incorporar el SUBPROGRAMA antes de seguir adelante", vbCritical + vbOKOnly, "SUBPROGRAMA INEXISTENTE"
            .txtSubprograma.SetFocus
            ValidarObra = False
            Exit Function
        End If
            
        If Trim(.txtProyecto.Text) = "" Or IsNumeric(.txtProyecto.Text) = False Or Len(Trim(.txtProyecto.Text)) <> "2" Then
            MsgBox "Debe ingresar un Nro de PROYECTO de 2 cifras", vbCritical + vbOKOnly, "NRO PROYECTO INCORRECTO"
            .txtProyecto.SetFocus
            ValidarObra = False
            Exit Function
        End If
        strEstructura = strEstructura & "-" & .txtProyecto.Text
        If SQLNoMatch("Select * from PROYECTOS Where PROYECTO= '" & strEstructura & "'") = True Then
            MsgBox "El PROYECTO ingresado no existe en la Base de Datos, deberá incorporar el PROYECTO antes de seguir adelante", vbCritical + vbOKOnly, "PROYECTO INEXISTENTE"
            .txtProyecto.SetFocus
            ValidarObra = False
            Exit Function
        End If
            
        If Trim(.txtActividad.Text) = "" Or IsNumeric(.txtActividad.Text) = False Or Len(Trim(.txtActividad.Text)) <> "2" Then
            MsgBox "Debe ingresar un Nro de ACTIVIDAD de 2 cifras", vbCritical + vbOKOnly, "NRO ACTIVIDAD INCORRECTO"
            .txtActividad.SetFocus
            ValidarObra = False
            Exit Function
        End If
        strEstructura = strEstructura & "-" & .txtActividad.Text
        If SQLNoMatch("Select * from ACTIVIDADES Where ACTIVIDAD= '" & strEstructura & "'") = True Then
            MsgBox "La ACTIVIDAD ingresado no existe en la Base de Datos, deberá incorporar la ACTIVIDAD antes de seguir adelante", vbCritical + vbOKOnly, "ACTIVIDAD INEXISTENTE"
            .txtActividad.SetFocus
            ValidarObra = False
            Exit Function
        End If
        
        If Trim(.txtPartida.Text) = "" Or IsNumeric(.txtPartida.Text) = False Or Len(Trim(.txtPartida.Text)) <> "3" Then
            MsgBox "Debe ingresar un Nro de PARTIDA de 3 cifras", vbCritical + vbOKOnly, "NRO PARTIDA INCORRECTO"
            .txtPartida.SetFocus
            ValidarObra = False
            Exit Function
        End If
        If SQLNoMatch("Select * from PARTIDAS Where PARTIDA= '" & .txtPartida.Text & "'") = True Then
            MsgBox "La PARTIDA ingresado no existe en la Base de Datos, deberá incorporar la PARTIDA antes de seguir adelante", vbCritical + vbOKOnly, "PARTIDA INEXISTENTE"
            .txtPartida.SetFocus
            ValidarObra = False
            Exit Function
        End If
        
        If Trim(.txtMontoContrato.Text) = "" Or IsNumeric(.txtMontoContrato.Text) = False Then
            MsgBox "Debe ingresar un MONTO DE CONTRATO", vbCritical + vbOKOnly, "MONTO DE CONTRATO INCORRECTO"
            .txtMontoContrato.SetFocus
            ValidarObra = False
            Exit Function
        End If
        
        If Trim(.txtAdicional.Text) = "" Or IsNumeric(.txtAdicional.Text) = False Then
            .txtAdicional.Text = "0"
        End If
    
        If Trim(.txtCuenta.Text) = "" Then
            MsgBox "Debe ingresar un Nro de CUENTA", vbCritical + vbOKOnly, "NRO CUENTA INCORRECTO"
            .txtCuenta.SetFocus
            ValidarObra = False
            Exit Function
        End If
        If SQLNoMatch("Select * from CUENTAS Where CUENTA= '" & .txtCuenta.Text & "'") = True Then
            MsgBox "Debe ingresar un Nro de CUENTA", vbCritical + vbOKOnly, "NRO CUENTA INEXISTENTE"
            .txtCuenta.SetFocus
            ValidarObra = False
            Exit Function
        End If
        
        If Trim(.txtFuente.Text) = "" Or IsNumeric(.txtFuente.Text) = False Then
            MsgBox "Debe ingresar un Nro de FUENTE", vbCritical + vbOKOnly, "NRO FUENTE INCORRECTO"
            .txtFuente.SetFocus
            ValidarObra = False
            Exit Function
        End If
        If SQLNoMatch("Select * from FUENTES Where FUENTE= '" & .txtFuente.Text & "'") = True Then
            MsgBox "Debe ingresar un Nro de FUENTE", vbCritical + vbOKOnly, "NRO FUENTE INEXISTENTE"
            .txtFuente.SetFocus
            ValidarObra = False
            Exit Function
        End If
    
        If Trim(.txtLocalidad.Text) = "" Or IsNumeric(.txtLocalidad.Text) = True Then
            MsgBox "Debe ingresar la LOCALIDAD", vbCritical + vbOKOnly, "NRO FUENTE LOCALIDAD"
            .txtLocalidad.SetFocus
            ValidarObra = False
            Exit Function
        End If
        If SQLNoMatch("Select * from LOCALIDADES Where LOCALIDAD= '" & .txtLocalidad.Text & "'") = True Then
            MsgBox "Debe ingresar la LOCALIDAD", vbCritical + vbOKOnly, "LOCALIDAD INEXISTENTE"
            .txtLocalidad.SetFocus
            ValidarObra = False
            Exit Function
        End If
    
        If Trim(.txtObra.Text) = "" Or IsNumeric(.txtObra.Text) = True Or Len(Trim(.txtObra.Text)) > 110 Then
            MsgBox "Debe ingresar una Descripción correcta de la Obra de no más de 110 caracteres", vbCritical + vbOKOnly, "DESCRIPCIÓN INCORRECTA"
            .txtObra.SetFocus
            ValidarObra = False
            Exit Function
        End If
        If strEditandoObra <> "" Then
            If .txtObra.Text <> strEditandoObra Then
                If SQLNoMatch("Select * from OBRAS Where DESCRIPCION = '" & .txtObra.Text & "'") = False Then
                    MsgBox "La Obra ya existe, verifique la Descripción de la misma", vbCritical + vbOKOnly, "OBRA DUPLICADA"
                    .txtObra.SetFocus
                    ValidarObra = False
                    Exit Function
                End If
            End If
        ElseIf strEditandoObra = "" Then
            If SQLNoMatch("Select * from OBRAS Where DESCRIPCION = '" & .txtObra.Text & "'") = False Then
                MsgBox "La Obra ya existe, verifique la Descripción de la misma", vbCritical + vbOKOnly, "OBRA DUPLICADA"
                .txtObra.SetFocus
                ValidarObra = False
                Exit Function
            End If
        End If
        
    End With
    ValidarObra = True
    
End Function

Public Function ValidarObraImputacionEPAM(DescripcionObra As String) As Boolean

        If DescripcionObra = "" Then
            MsgBox "Debe ingresar una Descripción de Obra", vbCritical + vbOKOnly, "DESCRIPCION OBRA INEXISTENTE"
            ImputacionEPAM.txtObra.SetFocus
            ValidarObraImputacionEPAM = False
            Exit Function
        Else
            If SQLNoMatch("Select * from OBRAS Where DESCRIPCION= '" & DescripcionObra & "'") = True Then
                Agregar = MsgBox("La Descripcion de Obra ingresada no existe en la Base de Datos. Desea Agregar la Obra " & DescripcionObra & " ?", vbQuestion + vbYesNo, "ADICIONAR OBRA")
                If Agregar = 6 Then
                    bolObraDesdeEPAMConsolidado = True
                    CargaObra.Show
                    CargaObra.txtCUIT.Text = "30632351514"
                    CargaObra.txtCUIT.Enabled = False
                    Set rstBuscarIcaro = New ADODB.Recordset
                    SQL = "Select * From CUIT Where CUIT = '" & CargaObra.txtCUIT.Text & "'"
                    rstBuscarIcaro.Open SQL, dbIcaro, adOpenDynamic, adLockOptimistic
                    CargaObra.txtDescripcionCUIT.Text = rstBuscarIcaro!Nombre
                    CargaObra.txtDescripcionCUIT.Enabled = False
                    CargaObra.txtObra.Text = DescripcionObra
                    CargaObra.txtObra.Enabled = False
                    CargaObra.txtPrograma.SetFocus
                    ImputacionEPAM.Hide
                Else
                    CargaObra.txtObra.SetFocus
                End If
                ValidarObraImputacionEPAM = False
                Exit Function
            End If
        End If

    ValidarObraImputacionEPAM = True

End Function

Public Function ValidarEPAMConsolidado() As Boolean

    With ImputacionEPAM
    
        If Trim(.txtObra.Text) = "" Or IsNumeric(.txtObra.Text) = True Then
            MsgBox "Debe ingresar un Nombre de OBRA valido", vbCritical + vbOKOnly, "NOMBRE DE OBRA INCORRECTO"
            .txtObra.SetFocus
            ValidarEPAMConsolidado = False
            Exit Function
        End If
        If strEditandoEPAMConsolidado <> "" And .txtObra.Text <> strEditandoEPAMConsolidado Then
            If SQLNoMatch("Select * from DATOSFIJOSEPAM Where CODIGO= '" & .txtObra.Text & "'") = False Then
                MsgBox "La OBRA ya se encuentra CONSOLIDADA, verifique el Nombre de la misma", vbCritical + vbOKOnly, "CÓDIGO DE OBRA DUPLICADA"
                .txtObra.SetFocus
                ValidarEPAMConsolidado = False
                Exit Function
            End If
        ElseIf strEditandoEPAMConsolidado = "" Then
            If SQLNoMatch("Select * from DATOSFIJOSEPAM Where CODIGO= '" & .txtObra.Text & "'") = False Then
                MsgBox "La OBRA ya se encuentra CONSOLIDADA, verifique el Nombre de la misma", vbCritical + vbOKOnly, "CÓDIGO DE OBRA DUPLICADA"
                .txtObra.SetFocus
                ValidarEPAMConsolidado = False
                Exit Function
            End If
        End If
        If Trim(.txtExpediente.Text) = "" Then
            MsgBox "Debe ingresar el Numero de Expediente en forma correcta", vbCritical + vbOKOnly, "EXPEDIENTE INCORRECTO"
            .txtExpediente.SetFocus
            ValidarEPAMConsolidado = False
            Exit Function
        End If
        If Trim(.txtResolucion.Text) = "" Then
            MsgBox "Debe ingresar el Numero de Resolución en forma correcta", vbCritical + vbOKOnly, "RESOLUCION INCORRECTA"
            .txtResolucion.SetFocus
            ValidarEPAMConsolidado = False
            Exit Function
        End If
        If Trim(.txtPresupuestado.Text) = "" Or IsNumeric(.txtPresupuestado.Text) = False Then
            MsgBox "Debe ingresar el Monto Presupuestado en forma correcta", vbCritical + vbOKOnly, "MONTO PRESUPUESTADO INCORRECTO"
            .txtPresupuestado.SetFocus
            ValidarEPAMConsolidado = False
            Exit Function
        End If
        If Trim(.txtImputado.Text) = "" Or IsNumeric(.txtImputado.Text) = False Then
            MsgBox "Debe ingresar el Monto Imputado en forma correcta", vbCritical + vbOKOnly, "MONTO IMPUTADO INCORRECTO"
            .txtImputado.SetFocus
            ValidarEPAMConsolidado = False
            Exit Function
        End If
        If Trim(.txtPagado.Text) = "" Or IsNumeric(.txtPagado.Text) = False Then
            MsgBox "Debe ingresar el Monto Pagado en forma correcta", vbCritical + vbOKOnly, "MONTO PAGADO INCORRECTO"
            .txtPagado.SetFocus
            ValidarEPAMConsolidado = False
            Exit Function
        End If
        
    End With
    ValidarEPAMConsolidado = True
    
End Function

Public Function ValidarResolucionEPAM() As Boolean

    With ImputacionResolucionesEPAM
    
        If Trim(.txtObra.Text) = "" Or IsNumeric(.txtObra.Text) = True Then
            MsgBox "Debe ingresar un Nombre de OBRA valido", vbCritical + vbOKOnly, "NOMBRE DE OBRA INCORRECTO"
            .txtObra.SetFocus
            ValidarResolucionEPAM = False
            Exit Function
        End If
        If strEditandoResolucionEPAM <> "" And .txtResolucion.Text <> strEditandoResolucionEPAM Then
            If SQLNoMatch("Select * from RESOLUCIONESEPAM Where CODIGO= '" & .txtObra.Text & "' And RESOLUCION= '" & .txtResolucion.Text & "'") = False Then
                MsgBox "El número de Resolución ya existe.", vbCritical + vbOKOnly, "RESOLUCIÓN DUPLICADA"
                .txtResolucion.SetFocus
                ValidarResolucionEPAM = False
                Exit Function
            End If
        ElseIf strEditandoResolucionEPAM = "" Then
            If SQLNoMatch("Select * from RESOLUCIONESEPAM Where CODIGO= '" & .txtObra.Text & "' And RESOLUCION= '" & .txtResolucion.Text & "'") = False Then
                MsgBox "El número de Resolución ya existe.", vbCritical + vbOKOnly, "RESOLUCIÓN DUPLICADA"
                .txtResolucion.SetFocus
                ValidarResolucionEPAM = False
                Exit Function
            End If
        End If
        If Trim(.txtDescripcion.Text) = "" Then
            MsgBox "Debe ingresar una DESCRIPCION en forma correcta", vbCritical + vbOKOnly, "DESCRIPCION INCORRECTA"
            .txtDescripcion.SetFocus
            ValidarResolucionEPAM = False
            Exit Function
        End If
        If Trim(.txtExpediente.Text) = "" Then
            MsgBox "Debe ingresar el Numero de Expediente en forma correcta", vbCritical + vbOKOnly, "EXPEDIENTE INCORRECTO"
            .txtExpediente.SetFocus
            ValidarResolucionEPAM = False
            Exit Function
        End If
        If Trim(.txtResolucion.Text) = "" Then
            MsgBox "Debe ingresar el Numero de Resolución en forma correcta", vbCritical + vbOKOnly, "RESOLUCION INCORRECTA"
            .txtResolucion.SetFocus
            ValidarResolucionEPAM = False
            Exit Function
        End If
        If Trim(.txtPresupuestado.Text) = "" Or IsNumeric(.txtPresupuestado.Text) = False Then
            MsgBox "Debe ingresar el Monto Presupuestado en forma correcta", vbCritical + vbOKOnly, "MONTO PRESUPUESTADO INCORRECTO"
            .txtPresupuestado.SetFocus
            ValidarResolucionEPAM = False
            Exit Function
        End If
        If Trim(.txtImputado.Text) = "" Or IsNumeric(.txtImputado.Text) = False Then
            MsgBox "Debe ingresar el Monto Imputado en forma correcta", vbCritical + vbOKOnly, "MONTO IMPUTADO INCORRECTO"
            .txtImputado.SetFocus
            ValidarResolucionEPAM = False
            Exit Function
        End If
        
    End With
    ValidarResolucionEPAM = True
    
End Function

Public Function ValidarRetencionCargaRetencion(Retencion As String) As Boolean

    If Retencion = "" Then
        MsgBox "Debe ingresar un Código de Retención", vbCritical + vbOKOnly, "CÓDIGO RETENCIÓN NULO"
        CargaRetencion.txtRetencion.SetFocus
        ValidarRetencionCargaRetencion = False
        Exit Function
    Else
        If SQLNoMatch("Select * from CODIGOSRETENCIONES Where CODIGO= '" & Retencion & "'") = True Then
            MsgBox "El Código de Retención ingresado no existe en la Base de Datos, " _
            & "deberá incorporar el CÓDIGO antes de seguir adelante", vbOKOnly, "CÓDIGO RETENCIÓN INEXISTENTE"
            CargaRetencion.txtRetencion.SetFocus
            ValidarRetencionCargaRetencion = False
            Exit Function
        End If
    End If
    
    ValidarRetencionCargaRetencion = True

End Function

Public Function ValidarRetencion() As Boolean
    
    With CargaRetencion
    
        If Len(.txtRetencion.Text) <> "3" Or IsNumeric(.txtRetencion.Text) = False Then
            MsgBox "Debe ingresar una numero de RETENCIÓN válido de 3 cifras", vbCritical + vbOKOnly, "CÓDIGO RETENCIÓN INCORRECTO"
            .txtRetencion.SetFocus
            ValidarRetencion = False
            Exit Function
        End If
        If SQLNoMatch("Select * from CODIGOSRETENCIONES Where CODIGO= '" & .txtRetencion.Text & "'") = True Then
            MsgBox "La RETENCION no existe en la Base de Datos. Verifique el Nro", vbCritical + vbOKOnly, "RETENCION INEXISTENTE"
            .txtRetencion.SetFocus
            ValidarRetencion = False
            Exit Function
        End If
        
        If SQLNoMatch("Select * From RETENCIONES Where Comprobante = '" & Format(.txtComprobante.Text, "00000") & "/" _
        & Right(.txtEjercicio.Text, 2) & "' And CODIGO= '" & .txtRetencion.Text & "'") = False Then
            MsgBox "La RETENCION ya fue ingresada en este comprobante. Verifique el Código de la misma", vbCritical + vbOKOnly, "RETENCIÓN YA INGRESADA EN EL COMPROBANTE"
            .txtRetencion.SetFocus
            ValidarRetencion = False
            Exit Function
        End If

        If Trim(.txtMonto.Text) = "" Or IsNumeric(.txtMonto.Text) = False Then
            MsgBox "Debe ingresar un Importe de la Retencion valido", vbCritical + vbOKOnly, "IMPORTE INCORRECTO"
            .txtMonto.SetFocus
            ValidarRetencion = False
            Exit Function
        End If
        
    End With
    ValidarRetencion = True

    
End Function

Public Function ValidarEstructuraProgramatica() As Boolean
    
    With CargaEstructuraProgramatica
    
        If Trim(.txtSubprograma.Text) = "" Then
            If Trim(.txtPrograma.Text) = "" Or IsNumeric(.txtPrograma.Text) = False Or Len(.txtPrograma.Text) > 2 Then
                MsgBox "Debe ingresar un Nro de Programa válido de 2 cifras", vbCritical + vbOKOnly, "PROGRAMA INCORRECTO"
                .txtPrograma.SetFocus
                ValidarEstructuraProgramatica = False
                Exit Function
            End If
            If SQLNoMatch("Select * from PROGRAMAS Where PROGRAMA= '" & .txtPrograma.Text & "'") = False Then
                MsgBox "El Programa ya existe, verifique el código del mismo", vbCritical + vbOKOnly, "PROGRAMA DUPLICADO"
                .txtPrograma.SetFocus
                ValidarEstructuraProgramatica = False
                Exit Function
            End If
            If Len(.txtDescripcionPrograma.Text) > 60 Then
                MsgBox "Debe ingresar una Descripción del Programa de no más de 60 caracteres", vbCritical + vbOKOnly, "DESCRIPCIÓN PROGRAMA INCORRECTA"
                .txtDescripcionPrograma.SetFocus
                ValidarEstructuraProgramatica = False
                Exit Function
            End If
            ValidarEstructuraProgramatica = True
            Exit Function
        End If
        
        If Trim(.txtProyecto.Text) = "" Then
            If Trim(.txtSubprograma.Text) = "" Or IsNumeric(.txtSubprograma.Text) = False Or Len(.txtSubprograma.Text) > 2 Then
                MsgBox "Debe ingresar un Nro de Subprograma válido de 2 cifras", vbCritical + vbOKOnly, "SUBPROGRAMA INCORRECTO"
                .txtSubprograma.SetFocus
                ValidarEstructuraProgramatica = False
                Exit Function
            End If
            If SQLNoMatch("Select * from SUBPROGRAMAS Where SUBPROGRAMA= '" & .txtPrograma.Text & "-" & .txtSubprograma.Text & "'") = False Then
                MsgBox "El Subprograma ya existe, verifique el código del mismo", vbCritical + vbOKOnly, "SUBPROGRAMA DUPLICADO"
                .txtSubprograma.SetFocus
                ValidarEstructuraProgramatica = False
                Exit Function
            End If
            If Len(.txtDescripcionPrograma.Text) > 60 Then
                MsgBox "Debe ingresar una Descripción del Subprograma de no más de 60 caracteres", vbCritical + vbOKOnly, "DESCRIPCIÓN SUBPROGRAMA INCORRECTA"
                .txtDescripcionSubprograma.SetFocus
                ValidarEstructuraProgramatica = False
                Exit Function
            End If
            ValidarEstructuraProgramatica = True
            Exit Function
        End If

        If Trim(.txtActividad.Text) = "" Then
            If Trim(.txtProyecto.Text) = "" Or IsNumeric(.txtProyecto.Text) = False Or Len(.txtProyecto.Text) > 2 Then
                MsgBox "Debe ingresar un Nro de Proyecto válido de 2 cifras", vbCritical + vbOKOnly, "PROYECTO INCORRECTO"
                .txtProyecto.SetFocus
                ValidarEstructuraProgramatica = False
                Exit Function
            End If
            If SQLNoMatch("Select * from PROYECTOS Where PROYECTO= '" & .txtPrograma.Text & "-" & .txtSubprograma.Text & "-" & .txtProyecto.Text & "'") = False Then
                MsgBox "El Proyecto ya existe, verifique el código del mismo", vbCritical + vbOKOnly, "PROYECTO DUPLICADO"
                .txtProyecto.SetFocus
                ValidarEstructuraProgramatica = False
                Exit Function
            End If
            If Len(.txtDescripcionProyecto.Text) > 60 Then
                MsgBox "Debe ingresar una Descripción del Proyecto de no más de 60 caracteres", vbCritical + vbOKOnly, "DESCRIPCIÓN PROYECTO INCORRECTA"
                .txtDescripcionProyecto.SetFocus
                ValidarEstructuraProgramatica = False
                Exit Function
            End If
            ValidarEstructuraProgramatica = True
            Exit Function
        End If

        If Trim(.txtActividad.Text) = "" Or IsNumeric(.txtActividad.Text) = False Or Len(.txtActividad.Text) > 2 Then
            MsgBox "Debe ingresar un Nro de Actividad válido de 2 cifras", vbCritical + vbOKOnly, "ACTIVIDAD INCORRECTA"
            .txtActividad.SetFocus
            ValidarEstructuraProgramatica = False
            Exit Function
        End If
        If SQLNoMatch("Select * from ACTIVIDADES Where ACTIVIDAD= '" & .txtPrograma.Text & "-" & .txtSubprograma.Text & "-" & .txtProyecto.Text & "-" & .txtActividad.Text & "'") = False Then
            MsgBox "La Actividad ya existe, verifique el código de la misma", vbCritical + vbOKOnly, "ACTIVIDAD DUPLICADA"
            .txtActividad.SetFocus
            ValidarEstructuraProgramatica = False
            Exit Function
        End If
        If Len(.txtDescripcionActividad.Text) > 60 Then
            MsgBox "Debe ingresar una Descripción de la Actividad de no más de 60 caracteres", vbCritical + vbOKOnly, "DESCRIPCIÓN ACTIVIDAD INCORRECTA"
            .txtDescripcionActividad.SetFocus
            ValidarEstructuraProgramatica = False
            Exit Function
        End If

    End With
    ValidarEstructuraProgramatica = True
    
End Function

Function ValidarActualizarRetencionesPorCtaCte() As Boolean

    Dim SQL As String
    
    With RetencionesPorCtaCte
        
        If Left(.txtFechaDesde.Text, 2) = "" Or IsDate(.txtFechaDesde.Text) = False Or Len(.txtFechaDesde.Text) <> "10" Then
            MsgBox "Debe ingresar una fecha valida", vbCritical + vbOKOnly, "FECHA DESDE INCORRECTA"
            .txtFechaDesde.SetFocus
            ValidarActualizarRetencionesPorCtaCte = False
            Exit Function
        End If
        
        If Left(.txtFechaHasta.Text, 2) = "" Or IsDate(.txtFechaHasta.Text) = False Or Len(.txtFechaHasta.Text) <> "10" Then
            MsgBox "Debe ingresar una fecha valida", vbCritical + vbOKOnly, "FECHA HASTA INCORRECTA"
            .txtFechaHasta.SetFocus
            ValidarActualizarRetencionesPorCtaCte = False
            Exit Function
        End If
        
'        If Trim(.cmbAno.Text) = "" Or IsNumeric(.cmbAno.Text) = False Or Len(.cmbAno.Text) <> "4" Then
'            MsgBox "Debe ingresar un Año de Ejercicio del Listado", vbCritical + vbOKOnly, "AÑO INCORRECTO"
'            .cmbAno.SetFocus
'            ValidarActualizarRetencionesPorCtaCte = False
'            Exit Function
'        End If
'
'        SQL = "Select Year(Fecha) As Ejercicio From CARGA " _
'        & "Where Year(Fecha) = '" & .cmbAno.Text & "'" _
'        & "Group by Year(Fecha) " _
'        & "Order By Year(Fecha) Desc"
'        If SQLNoMatch(SQL) Then
'            MsgBox "Debe ingresar un Año de Ejercicio del Listado", vbCritical + vbOKOnly, "AÑO INCORRECTO"
'            .SetFocus
'            ValidarActualizarRetencionesPorCtaCte = False
'            Exit Function
'        End If
                
    End With
    ValidarActualizarRetencionesPorCtaCte = True

End Function

Function ValidarActualizarEjecucionPorImputacion() As Boolean

    Dim SQL As String
    
    With EjecucionPorImputacion
        
        If Left(.txtFechaDesde.Text, 2) = "" Or IsDate(.txtFechaDesde.Text) = False Or Len(.txtFechaDesde.Text) <> "10" Then
            MsgBox "Debe ingresar una fecha valida", vbCritical + vbOKOnly, "FECHA DESDE INCORRECTA"
            .txtFechaDesde.SetFocus
            ValidarActualizarEjecucionPorImputacion = False
            Exit Function
        End If
        
        If Left(.txtFechaHasta.Text, 2) = "" Or IsDate(.txtFechaHasta.Text) = False Or Len(.txtFechaHasta.Text) <> "10" Then
            MsgBox "Debe ingresar una fecha valida", vbCritical + vbOKOnly, "FECHA HASTA INCORRECTA"
            .txtFechaHasta.SetFocus
            ValidarActualizarEjecucionPorImputacion = False
            Exit Function
        End If
        
'        If Trim(.cmbAno.Text) = "" Or IsNumeric(.cmbAno.Text) = False Or Len(.cmbAno.Text) <> "4" Then
'            MsgBox "Debe ingresar un Año de Ejercicio del Listado", vbCritical + vbOKOnly, "AÑO INCORRECTO"
'            .cmbAno.SetFocus
'            ValidarActualizarRetencionesPorCtaCte = False
'            Exit Function
'        End If
'
'        SQL = "Select Year(Fecha) As Ejercicio From CARGA " _
'        & "Where Year(Fecha) = '" & .cmbAno.Text & "'" _
'        & "Group by Year(Fecha) " _
'        & "Order By Year(Fecha) Desc"
'        If SQLNoMatch(SQL) Then
'            MsgBox "Debe ingresar un Año de Ejercicio del Listado", vbCritical + vbOKOnly, "AÑO INCORRECTO"
'            .SetFocus
'            ValidarActualizarRetencionesPorCtaCte = False
'            Exit Function
'        End If
                
    End With
    ValidarActualizarEjecucionPorImputacion = True

End Function


