Attribute VB_Name = "EliminarRegistros"
Public Sub EliminarComprobante()

    Dim i As String
    Dim Borrar As Integer
    Dim SQL As String
    
    With ListadoImputacion.dgListado
        i = .Row
        Borrar = MsgBox("Desea Borrar DEFINITIVAMENTE el Comprobante: " & .TextMatrix(i, 0) & "?" _
        & vbCrLf & "Tenga en cuenta que todas las Imputaciones y Retenciones Asociados al Comprobante " _
        & "también serán ELIMINADOS", vbQuestion + vbYesNo, "BORRANDO COMPROBANTE")
        If Borrar = 6 Then
            Set rstRegistroIcaro = New ADODB.Recordset
            If SQLNoMatch("Select * From EPAM Where COMPROBANTE= '" & .TextMatrix(i, 0) & "'") Then
                With rstRegistroIcaro
                    Set rstBuscarIcaro = New ADODB.Recordset
                    SQL = "Select * From CARGA Where COMPROBANTE= '" _
                    & ListadoImputacion.dgListado.TextMatrix(i, 0) & "'"
                    rstBuscarIcaro.Open SQL, dbIcaro, adOpenForwardOnly, adLockOptimistic
                    .Open "Certificados", dbIcaro, adOpenForwardOnly, adLockOptimistic
                    .AddNew
                    .Fields("Comprobante") = ""
                    '.Fields("Fecha") = rstBuscarIcaro.Fields("Fecha")
                    .Fields("Obra") = rstBuscarIcaro.Fields("Obra")
                    .Fields("Certificado") = rstBuscarIcaro.Fields("Certificado")
                    If Trim(.Fields("Certificado")) <> "FR" Then
                        .Fields("FondoDeReparo") = rstBuscarIcaro.Fields("FondoDeReparo")
                    Else
                        .Fields("FondoDeReparo") = "0"
                    End If
                    .Fields("MontoBruto") = rstBuscarIcaro.Fields("Importe")
                    .Fields("Periodo") = Format(rstBuscarIcaro.Fields("Fecha"), "mmmm" & "/" & "yy")
                    SQL = "Select * From CUIT Where CUIT= '" & rstBuscarIcaro.Fields("CUIT") & "'"
                    rstBuscarIcaro.Close
                    rstBuscarIcaro.Open SQL, dbIcaro, adOpenForwardOnly, adLockOptimistic
                    .Fields("Proveedor") = rstBuscarIcaro.Fields("Nombre")
                    rstBuscarIcaro.Close
                    SQL = "Select * From RETENCIONES Where Comprobante= '" _
                    & ListadoImputacion.dgListado.TextMatrix(i, 0) & "'"
                    .Fields("IB") = 0
                    .Fields("LP") = 0
                    .Fields("SUSS") = 0
                    .Fields("Ganancias") = 0
                    .Fields("INVICO") = 0
                    rstBuscarIcaro.Open SQL, dbIcaro, adOpenForwardOnly, adLockOptimistic
                    If rstBuscarIcaro.BOF = False Then
                        rstBuscarIcaro.MoveFirst
                        While rstBuscarIcaro.EOF = False
                            Select Case rstBuscarIcaro.Fields("Codigo")
                            Case Is = 101, 110
                                .Fields("IB") = rstBuscarIcaro.Fields("Importe")
                            Case Is = 104, 112
                                .Fields("LP") = rstBuscarIcaro.Fields("Importe")
                            Case Is = 114, 117
                                .Fields("SUSS") = rstBuscarIcaro.Fields("Importe")
                            Case Is = 106, 113
                                .Fields("Ganancias") = rstBuscarIcaro.Fields("Importe")
                            Case Is = 337
                                .Fields("INVICO") = rstBuscarIcaro.Fields("Importe")
                            End Select
                            rstBuscarIcaro.MoveNext
                        Wend
                    End If
                    .Update
                    rstRegistroIcaro.Close
                    rstBuscarIcaro.Close
                    Set rstBuscarIcaro = Nothing
                End With
            Else
                SQL = "Select * From EPAM Where COMPROBANTE= '" & .TextMatrix(i, 0) & "'"
                rstRegistroIcaro.Open SQL, dbIcaro, adOpenForwardOnly, adLockOptimistic
                rstRegistroIcaro.Fields("Comprobante") = ""
                rstRegistroIcaro.Update
                rstRegistroIcaro.Close
            End If
            SQL = "Select * from CARGA Where COMPROBANTE = " & "'" & .TextMatrix(i, 0) & "'"
            rstRegistroIcaro.Open SQL, dbIcaro, adOpenForwardOnly, adLockOptimistic
            rstRegistroIcaro.Delete
            rstRegistroIcaro.Close
            Set rstRegistroIcaro = Nothing
            ConfigurardgListadoGasto
            Call CargardgListadoGasto(, Year(Now()))
            ListadoImputacion.txtFecha.Text = Year(Now())
            i = .Row
            ConfigurardgImputacion
            CargardgImputacion (.TextMatrix(i, 0))
            ConfigurardgRetencion
            CargardgRetencion (.TextMatrix(i, 0))
        End If
    End With
    Borrar = 0
    i = ""
    SQL = ""

End Sub

Public Sub EliminarRetencion()

    Dim i As String
    Dim x As String
    Dim Borrar As Integer
    Dim SQL As String
    
    With ListadoImputacion
        i = .dgListado.Row
        x = .dgRetencion.Row
        If x > 0 And i > 0 Then
            Borrar = MsgBox("Desea Eliminar la Retención " & .dgRetencion.TextMatrix(x, 0) & " del Comprobante número " _
            & .dgListado.TextMatrix(i, 0) & "?", vbQuestion + vbYesNo, "BORRAR RETENCIÓN")
            If Borrar = 6 Then
                Set rstRegistroIcaro = New ADODB.Recordset
                SQL = "Select * From RETENCIONES Where Comprobante = '" _
                & .dgListado.TextMatrix(i, 0) & "' And Codigo = '" & .dgRetencion.TextMatrix(x, 0) & "'"
                rstRegistroIcaro.Open SQL, dbIcaro, adOpenForwardOnly, adLockOptimistic
                rstRegistroIcaro.Delete
                rstRegistroIcaro.Close
                Set rstRegistroIcaro = Nothing
                i = .dgListado.Row
                ConfigurardgRetencion
                CargardgRetencion (.dgListado.TextMatrix(i, 0))
            End If
        End If
    End With
    Borrar = 0
    i = ""
    x = ""
    SQL = ""
        
End Sub

Public Sub EliminarImputacion()

    Dim i As String
    Dim x As String
    Dim Borrar As Integer
    Dim SQL As String
    
    With ListadoImputacion
        i = .dgListado.Row
        x = .dgImputacion.Row
        If x > 0 And i > 0 Then
            Borrar = MsgBox("Desea Eliminar la Imputacion " & .dgImputacion.TextMatrix(x, 0) & "-" _
            & .dgImputacion.TextMatrix(x, 1) & " del Comprobante número " _
            & .dgListado.TextMatrix(i, 0) & "?", vbQuestion + vbYesNo, "BORRAR IMPUTACION")
            If Borrar = 6 Then
                Set rstRegistroIcaro = New ADODB.Recordset
                SQL = "Select * From Imputacion Where Comprobante = '" _
                & .dgListado.TextMatrix(i, 0) & "' And Imputacion = '" _
                & .dgImputacion.TextMatrix(x, 0) & "' And Partida = '" & .dgImputacion.TextMatrix(x, 1) & "'"
                rstRegistroIcaro.Open SQL, dbIcaro, adOpenForwardOnly, adLockOptimistic
                rstRegistroIcaro.Delete
                rstRegistroIcaro.Close
                Set rstRegistroIcaro = Nothing
                i = .dgListado.Row
                ConfigurardgImputacion
                CargardgImputacion (.dgListado.TextMatrix(i, 0))
            End If
        End If
    End With
    Borrar = 0
    i = ""
    x = ""
    SQL = ""
        
End Sub

Public Sub EliminarProveedor()

    Dim i As String
    Dim Borrar As Integer
    Dim SQL As String
    
    With ListadoProveedores.dgProveedores
        i = .Row
        Borrar = MsgBox("Desea Borrar DEFINITIVAMENTE el Proveedor: " & .TextMatrix(i, 1) & "?" & vbCrLf & "Tenga en cuenta que todas las Obras y Comprobantes Asociados al Proveedor también serán ELIMINADOS", vbQuestion + vbYesNo, "BORRANDO PROVEEDOR")
        If Borrar = 6 Then
            Set rstRegistroIcaro = New ADODB.Recordset
            SQL = "Select * from CUIT Where CUIT = " & "'" & .TextMatrix(i, 0) & "'"
            rstRegistroIcaro.Open SQL, dbIcaro, adOpenForwardOnly, adLockOptimistic
            rstRegistroIcaro.Delete
            rstRegistroIcaro.Close
            Set rstRegistroIcaro = Nothing
            ConfigurardgProveedores
            CargardgProveedores
        End If
    End With
    Borrar = 0
    i = ""
    SQL = ""
    
End Sub

Public Sub EliminarCuentaBancaria()

    Dim i As String
    Dim Borrar As Integer
    Dim SQL As String
    
    With ListadoCuentasBancarias.dgCuentasBancarias
        i = .Row
        Borrar = MsgBox("Desea Borrar DEFINITIVAMENTE la CUENTA BANCARIA: " & .TextMatrix(i, 0) & "?" & vbCrLf & "Tenga en cuenta que las Imputaciones Asociadas a la Cuenta Bancaria también serán ELIMINADAS", vbQuestion + vbYesNo, "BORRANDO CUENTA BANCARIA")
        If Borrar = 6 Then
            Set rstRegistroIcaro = New ADODB.Recordset
            SQL = "Select * from CUENTAS Where CUENTA = " & "'" & .TextMatrix(i, 0) & "'"
            rstRegistroIcaro.Open SQL, dbIcaro, adOpenForwardOnly, adLockOptimistic
            rstRegistroIcaro.Delete
            rstRegistroIcaro.Close
            Set rstRegistroIcaro = Nothing
            ConfigurardgCuentasBancarias
            CargardgCuentasBancarias
        End If
    End With
    Borrar = 0
    i = ""
    SQL = ""

End Sub

Public Sub EliminarPartida()

    Dim i As String
    Dim Borrar As Integer
    Dim SQL As String
    
    With ListadoPartidas.dgPartidas
        i = .Row
        Borrar = MsgBox("Desea Borrar DEFINITIVAMENTE la PARTIDA: " & .TextMatrix(i, 0) & "?" & vbCrLf & "Tenga en cuenta que las Imputaciones Asociadas a la Partida también serán ELIMINADAS", vbQuestion + vbYesNo, "BORRANDO PARTIDA")
        If Borrar = 6 Then
            Set rstRegistroIcaro = New ADODB.Recordset
            SQL = "Select * from PARTIDAS Where PARTIDA = " & "'" & .TextMatrix(i, 0) & "'"
            rstRegistroIcaro.Open SQL, dbIcaro, adOpenForwardOnly, adLockOptimistic
            rstRegistroIcaro.Delete
            rstRegistroIcaro.Close
            Set rstRegistroIcaro = Nothing
            ConfigurardgPartidas
            CargardgPartidas
        End If
    End With
    Borrar = 0
    i = ""
    SQL = ""

End Sub

Public Sub EliminarFuente()

    Dim i As String
    Dim Borrar As Integer
    Dim SQL As String
    
    With ListadoFuentes.dgFuentes
        i = .Row
        Borrar = MsgBox("Desea Borrar DEFINITIVAMENTE la FUENTE: " & .TextMatrix(i, 0) & "?" & vbCrLf & "Tenga en cuenta que las Imputaciones Asociadas a la Fuente también serán ELIMINADAS", vbQuestion + vbYesNo, "BORRANDO FUENTE")
        If Borrar = 6 Then
            Set rstRegistroIcaro = New ADODB.Recordset
            SQL = "Select * from FUENTES Where FUENTE = " & "'" & .TextMatrix(i, 0) & "'"
            rstRegistroIcaro.Open SQL, dbIcaro, adOpenForwardOnly, adLockOptimistic
            rstRegistroIcaro.Delete
            rstRegistroIcaro.Close
            Set rstRegistroIcaro = Nothing
            ConfigurardgFuentes
            CargardgFuentes
        End If
    End With
    Borrar = 0
    i = ""
    SQL = ""

End Sub

Public Sub EliminarCodigoRetencion()

    Dim i As String
    Dim Borrar As Integer
    Dim SQL As String
    
    With ListadoRetenciones.dgRetenciones
        i = .Row
        Borrar = MsgBox("Desea Borrar DEFINITIVAMENTE la RETENCION: " & .TextMatrix(i, 0) & "?" & vbCrLf & "Tenga en cuenta que las Deducciones Efectuadas Asociadas a la Partida también serán ELIMINADAS", vbQuestion + vbYesNo, "BORRANDO RETENCION")
        If Borrar = 6 Then
            Set rstRegistroIcaro = New ADODB.Recordset
            SQL = "Select * from CODIGOSRETENCIONES Where CODIGO = " & "'" & .TextMatrix(i, 0) & "'"
            rstRegistroIcaro.Open SQL, dbIcaro, adOpenForwardOnly, adLockOptimistic
            rstRegistroIcaro.Delete
            rstRegistroIcaro.Close
            Set rstRegistroIcaro = Nothing
            ConfigurardgCodigoRetenciones
            CargardgCodigoRetenciones
        End If
    End With
    Borrar = 0
    i = ""
    SQL = ""

End Sub

Public Sub EliminarLocalidad()

    Dim i As String
    Dim Borrar As Integer
    Dim SQL As String
    
    With ListadoLocalidades.dgLocalidades
        i = .Row
        Borrar = MsgBox("Desea Borrar DEFINITIVAMENTE la Localidad: " & .TextMatrix(i, 0) & "?" & vbCrLf & "Tenga en cuenta que las Imputaciones Asociadas a la Localidad también serán ELIMINADAS", vbQuestion + vbYesNo, "BORRANDO LOCALIDAD")
        If Borrar = 6 Then
            Set rstRegistroIcaro = New ADODB.Recordset
            SQL = "Select * from LOCALIDADES Where LOCALIDAD = " & "'" & .TextMatrix(i, 0) & "'"
            rstRegistroIcaro.Open SQL, dbIcaro, adOpenForwardOnly, adLockOptimistic
            rstRegistroIcaro.Delete
            rstRegistroIcaro.Close
            Set rstRegistroIcaro = Nothing
            ConfigurardgLocalidades
            CargardgLocalidades
        End If
    End With
    Borrar = 0
    i = ""
    SQL = ""

End Sub

Public Sub EliminarObra()

    Dim i As String
    Dim Borrar As Integer
    Dim SQL As String
    
    With ListadoObras.dgObras
        i = .Row
        Borrar = MsgBox("Desea Borrar DEFINITIVAMENTE la Obra: " & .TextMatrix(i, 0) & "?" & vbCrLf & "Tenga en cuenta que todas los Comprobantes Asociados a la Obra también serán ELIMINADOS", vbQuestion + vbYesNo, "BORRANDO OBRA")
        If Borrar = 6 Then
            Set rstRegistroIcaro = New ADODB.Recordset
            SQL = "Select * from OBRAS Where DESCRIPCION = " & "'" & .TextMatrix(i, 0) & "'"
            rstRegistroIcaro.Open SQL, dbIcaro, adOpenForwardOnly, adLockOptimistic
            rstRegistroIcaro.Delete
            rstRegistroIcaro.Close
            Set rstRegistroIcaro = Nothing
            ConfigurardgObras
            CargardgObras
        End If
    End With
    Borrar = 0
    i = ""
    SQL = ""

End Sub

Public Sub EliminarCertificado()

    Dim i As String
    Dim Borrar As Integer
    Dim SQL As String
    
    With AutocargaCertificados.dgAutocargaCertificado
        i = .Row
        Borrar = MsgBox("Desea Borrar el Certificado?", vbQuestion + vbYesNo, "Borrar")
        If Borrar = 6 Then
            Set rstRegistroIcaro = New ADODB.Recordset
            SQL = "Select * From CERTIFICADOS Where Comprobante = """"" _
            & " And Proveedor = " & "'" & .TextMatrix(i, 0) & "'" _
            & " And MontoBruto = Format((" & De_Num_a_Tx_01(.TextMatrix(i, 3), , 2) & "), '#.00') And Obra = " & " '" & .TextMatrix(i, 1) & "'" _
            & " And Certificado = " & "'" & .TextMatrix(i, 2) & "' Order by Periodo Desc, Obra Asc"
            rstRegistroIcaro.Open SQL, dbIcaro, adOpenForwardOnly, adLockOptimistic
            rstRegistroIcaro.Delete
            rstRegistroIcaro.Close
            Set rstRegistroIcaro = Nothing
            ConfigurardgAutocargaCertificados
            CargardgAutocargaCertificados
        End If
    End With
    Borrar = 0
    i = ""
    SQL = ""

End Sub

Public Sub EliminarTodoCertificados()

    Dim i As String
    Dim Borrar As Integer
    Dim SQL As String
    
    With AutocargaCertificados.dgAutocargaCertificado
        i = .Row
        Borrar = MsgBox("Desea Borrar TODOS los Certificados?", vbQuestion + vbYesNo, "ELIMINACIÓN COMPLETA")
        If Borrar = 6 Then
            SQL = "Delete * From CERTIFICADOS"
            dbIcaro.BeginTrans
            dbIcaro.Execute SQL
            dbIcaro.CommitTrans
            ConfigurardgAutocargaCertificados
            CargardgAutocargaCertificados
        End If
    End With
    Borrar = 0
    i = ""
    SQL = ""

End Sub

Public Sub EliminarEPAM()

    Dim i As String
    Dim Borrar As Integer
    Dim SQL As String
    
    With AutocargaEPAM.dgAutocargaEPAM
        i = .Row
        Borrar = MsgBox("Desea Borrar el Comprobante?", vbQuestion + vbYesNo, "Borrar")
        If Borrar = 6 Then
            Set rstRegistroIcaro = New ADODB.Recordset
            SQL = "Select * From EPAM Where Comprobante = """" And Codigo = '" & .TextMatrix(i, 0) & "'" _
            & " And MontoBruto = " & .TextMatrix(i, 1) & " And Inicio = '" & .TextMatrix(i, 3) & "'" _
            & " And Cierre = '" & .TextMatrix(i, 4) & "' Order by Inicio Desc, Codigo Asc"
            rstRegistroIcaro.Open SQL, dbIcaro, adOpenForwardOnly, adLockOptimistic
            rstRegistroIcaro.Delete
            rstRegistroIcaro.Close
            Set rstRegistroIcaro = Nothing
            ConfigurardgAutocargaEPAM
            CargardgAutocargaEPAM
        End If
    End With
    Borrar = 0
    i = ""
    SQL = ""

End Sub

Public Sub PagarEPAM()

    Dim i As String
    Dim Pagar As Integer
    Dim SQL As String
    
    With AutocargaEPAM.dgAutocargaEPAM
        i = .Row
        Pagar = MsgBox("Desea Pagar el Comprobante?", vbQuestion + vbYesNo, "PAGAR COMPROBANTE")
        If Pagar = 6 Then
            Set rstRegistroIcaro = New ADODB.Recordset
            SQL = "Select * From EPAM Where Comprobante = """" And Codigo = '" & .TextMatrix(i, 0) & "'" _
            & " And MontoBruto = " & .TextMatrix(i, 2) & " And Tipo = '" & .TextMatrix(i, 1) & "'" _
            & " And Inicio = '" & .TextMatrix(i, 3) & "' And Cierre = '" & .TextMatrix(i, 4) & "' Order by Inicio Desc, Codigo Asc"
            rstRegistroIcaro.Open SQL, dbIcaro, adOpenForwardOnly, adLockOptimistic
            rstRegistroIcaro!Comprobante = "PAGADO"
            rstRegistroIcaro!Fecha = .TextMatrix(i, 3)
            rstRegistroIcaro.Update
            rstRegistroIcaro.Close
            Set rstRegistroIcaro = Nothing
            ConfigurardgAutocargaEPAM
            CargardgAutocargaEPAM
        End If
    End With
    Pagar = 0
    i = ""
    SQL = ""
    
End Sub

Public Sub EliminarEPAMConsolidado()

    Dim i As String
    Dim Borrar As Integer
    Dim SQL As String
    
    With EPAMConsolidado.dgEPAMConsolidado
        i = .Row
        Borrar = MsgBox("Desea Borrar DEFINITIVAMENTE la CONSOLIDACION de la siguiente OBRA: " & .TextMatrix(i, 0) & "?", vbQuestion + vbYesNo, "BORRANDO OBRA EPAM CONSOLIDADA")
        If Borrar = 6 Then
            Set rstRegistroIcaro = New ADODB.Recordset
            SQL = "Select * From DATOSFIJOSEPAM Where Codigo = '" & .TextMatrix(i, 0) & "'"
            rstRegistroIcaro.Open SQL, dbIcaro, adOpenForwardOnly, adLockOptimistic
            rstRegistroIcaro.Delete
            rstRegistroIcaro.Close
            Set rstRegistroIcaro = Nothing
            ConfigurardgEPAMConsolidado
            CargardgEPAMConsolidado
            i = .Row
            ConfigurardgResolucionesEPAM
            CargardgResolucionesEPAM (.TextMatrix(i, 0))
        End If
    End With
    Borrar = 0
    i = ""
    SQL = ""
    
End Sub

Public Sub EliminarResolucionEPAM()

    Dim i As String
    Dim x As String
    Dim Borrar As Integer
    Dim SQL As String
    
    x = EPAMConsolidado.dgEPAMConsolidado.Row
    With EPAMConsolidado.dgResolucionesEPAM
        i = .Row
        If x >= 2 Then
            Borrar = MsgBox("Desea Borrar DEFINITIVAMENTE la REDETERMINACIÓN/ACTUALIZACIÓN de la siguiente OBRA: " _
            & EPAMConsolidado.dgEPAMConsolidado.TextMatrix(x, 0) & "-" & .TextMatrix(i, 0) & "?" _
            , vbQuestion + vbYesNo, "BORRANDO REDETERMINACIÓN/ACTUALIZACIÓN")
            If Borrar = 6 Then
                Set rstRegistroIcaro = New ADODB.Recordset
                SQL = "Select * From RESOLUCIONESEPAM Where Codigo = '" & EPAMConsolidado.dgEPAMConsolidado.TextMatrix(x, 0) & "' " _
                & "And Resolucion = '" & .TextMatrix(i, 2) & "'"
                rstRegistroIcaro.Open SQL, dbIcaro, adOpenForwardOnly, adLockOptimistic
                rstRegistroIcaro.Delete
                rstRegistroIcaro.Close
                Set rstRegistroIcaro = Nothing
                ConfigurardgResolucionesEPAM
                CargardgResolucionesEPAM (EPAMConsolidado.dgEPAMConsolidado.TextMatrix(x, 0))
            End If
        End If
    End With
    Borrar = 0
    i = ""
    x = ""
    SQL = ""
    
End Sub

Public Sub EliminarEstructuraProgramatica()

    Dim i As String
    Dim Borrar As Integer
    Dim SQL As String
    
    With EstructuraProgramatica.dgEstructura
        i = .Row
        Borrar = MsgBox("Desea Borrar DEFINITIVAMENTE la Estructura: " & .TextMatrix(i, 0) & "?" _
        & vbCrLf & "Tenga en cuenta que todas las Imputaciones Comprendidas en dicha estructura " _
        & "también serán ELIMINADAS", vbQuestion + vbYesNo, "BORRANDO ESTRUCTURA")
        If Borrar = 6 Then
            Set rstRegistroIcaro = New ADODB.Recordset
            Select Case Len(.TextMatrix(i, 0))
            Case Is = "2"
                SQL = "Select * From PROGRAMAS Where PROGRAMA = '" & .TextMatrix(i, 0) & "'"
                rstRegistroIcaro.Open SQL, dbIcaro, adOpenForwardOnly, adLockOptimistic
                rstRegistroIcaro.Delete
                rstRegistroIcaro.Close
                Set rstRegistroIcaro = Nothing
                ConfigurardgEstructura
                Call CargardgEstructura
            Case Is = "5"
                SQL = "Select * From SUBPROGRAMAS Where SUBPROGRAMA = '" & .TextMatrix(i, 0) & "'"
                rstRegistroIcaro.Open SQL, dbIcaro, adOpenForwardOnly, adLockOptimistic
                rstRegistroIcaro.Delete
                rstRegistroIcaro.Close
                Set rstRegistroIcaro = Nothing
                SQL = left(.TextMatrix(i, 0), 2)
                ConfigurardgEstructura
                Call CargardgEstructura(SQL)
            Case Is = "8"
                SQL = "Select * From PROYECTOS Where PROYECTO = '" & .TextMatrix(i, 0) & "'"
                rstRegistroIcaro.Open SQL, dbIcaro, adOpenForwardOnly, adLockOptimistic
                rstRegistroIcaro.Delete
                rstRegistroIcaro.Close
                Set rstRegistroIcaro = Nothing
                SQL = left(.TextMatrix(i, 0), 5)
                ConfigurardgEstructura
                Call CargardgEstructura(SQL)
            Case Is = "11"
                SQL = "Select * From ACTIVIDADES Where ACTIVIDAD = '" & .TextMatrix(i, 0) & "'"
                rstRegistroIcaro.Open SQL, dbIcaro, adOpenForwardOnly, adLockOptimistic
                rstRegistroIcaro.Delete
                rstRegistroIcaro.Close
                Set rstRegistroIcaro = Nothing
                SQL = left(.TextMatrix(i, 0), 8)
                ConfigurardgEstructura
                Call CargardgEstructura(SQL)
            Case Is = "15"
                SQL = "Select * From OBRAS Where DESCRIPCION = '" & .TextMatrix(i, 1) & "'"
                rstRegistroIcaro.Open SQL, dbIcaro, adOpenForwardOnly, adLockOptimistic
                rstRegistroIcaro.Delete
                rstRegistroIcaro.Close
                Set rstRegistroIcaro = Nothing
                SQL = left(.TextMatrix(i, 0), 11)
                ConfigurardgEstructura
                Call CargardgEstructura(SQL)
            End Select
        End If
    End With
    Borrar = 0
    i = ""
    SQL = ""

End Sub

