Attribute VB_Name = "HabilitarTextBox"
Public Sub BloquearTextBoxObra()
    
    With CargaObra
        .txtCUIT.Enabled = True
        .txtDescripcionCUIT.Enabled = False
        .txtPrograma.Enabled = True
        .txtDescripcionPrograma.Enabled = False
        .txtSubprograma.Enabled = False
        .txtDescripcionSubprograma.Enabled = False
        .txtProyecto.Enabled = False
        .txtDescripcionProyecto.Enabled = False
        .txtActividad.Enabled = False
        .txtDescripcionActividad.Enabled = False
        .txtPartida.Enabled = False
        .txtDescripcionPartida.Enabled = False
        .txtMontoContrato.Enabled = True
        .txtAdicional.Enabled = True
        .txtObra.Enabled = True
        .txtCuenta.Enabled = True
        .txtDescripcionCuenta.Enabled = False
        .txtFuente.Enabled = False
        .txtDescripcionFuente.Enabled = False
        .txtLocalidad.Enabled = True
        .txtNormaLegal.Enabled = True
    End With
    
End Sub

Public Sub BloquearTextBoxGasto()
    
    With CargaGasto
        .txtComprobante.Enabled = True
        .txtFecha.Enabled = True
        .txtCUIT.Enabled = True
        .txtObra.Enabled = False
        .txtImporte.Enabled = True
        .txtFR.Enabled = True
        .txtCertificado.Enabled = True
        .txtAvance.Enabled = True
        .txtCuenta.Enabled = False
        .txtFuente.Enabled = False
        .txtDescripcionCUIT.Enabled = False
        .txtDescripcionCuenta.Enabled = False
        .txtDescripcionFuente.Enabled = False
    End With

End Sub

Public Sub BloquearTextBoxCargaEstructuraProgramatica(Imputacion As String)

    Dim SQL As String
    
    Set rstRegistroIcaro = New ADODB.Recordset
    
    With CargaEstructuraProgramatica
        Select Case Len(Imputacion)
        Case Is = 2
            .Show
            .txtSubprograma.Enabled = False
            .txtDescripcionSubprograma.Enabled = False
            .txtProyecto.Enabled = False
            .txtDescripcionProyecto.Enabled = False
            .txtActividad.Enabled = False
            .txtDescripcionActividad.Enabled = False
        Case Is = 5
            .Show
            SQL = "Select * From PROGRAMAS Where Programa = '" & Left(Imputacion, 2) & "'"
            rstRegistroIcaro.Open SQL, dbIcaro, adOpenForwardOnly, adLockOptimistic
            .txtPrograma.Text = Left(Imputacion, 2)
            If Len(rstRegistroIcaro.Fields("Descripcion")) <> "0" Then
                .txtDescripcionPrograma.Text = rstRegistroIcaro.Fields("Descripcion")
            Else
                .txtDescripcionPrograma.Text = ""
            End If
            rstRegistroIcaro.Close
            .txtPrograma.Enabled = False
            .txtDescripcionPrograma.Enabled = False
            .txtSubprograma.Enabled = True
            .txtDescripcionSubprograma.Enabled = True
            .txtProyecto.Enabled = False
            .txtDescripcionProyecto.Enabled = False
            .txtActividad.Enabled = False
            .txtDescripcionActividad.Enabled = False
        Case Is = 8
            .Show
            SQL = "Select * From PROGRAMAS Where Programa = '" & Left(Imputacion, 2) & "'"
            rstRegistroIcaro.Open SQL, dbIcaro, adOpenForwardOnly, adLockOptimistic
            .txtPrograma.Text = Left(Imputacion, 2)
            If Len(rstRegistroIcaro.Fields("Descripcion")) <> "0" Then
                .txtDescripcionPrograma.Text = rstRegistroIcaro.Fields("Descripcion")
            Else
                .txtDescripcionPrograma.Text = ""
            End If
            rstRegistroIcaro.Close
            SQL = "Select * From SUBPROGRAMAS Where Subprograma = '" & Left(Imputacion, 5) & "'"
            rstRegistroIcaro.Open SQL, dbIcaro, adOpenForwardOnly, adLockOptimistic
            .txtSubprograma.Text = Mid(Imputacion, 4, 2)
            If Len(rstRegistroIcaro.Fields("Descripcion")) <> "0" Then
                .txtDescripcionSubprograma.Text = rstRegistroIcaro.Fields("Descripcion")
            Else
                .txtDescripcionSubprograma.Text = ""
            End If
            rstRegistroIcaro.Close
            .txtPrograma.Enabled = False
            .txtDescripcionPrograma.Enabled = False
            .txtSubprograma.Enabled = False
            .txtDescripcionSubprograma.Enabled = False
            .txtProyecto.Enabled = True
            .txtDescripcionProyecto.Enabled = True
            .txtActividad.Enabled = False
            .txtDescripcionActividad.Enabled = False
        Case Is = 11
            .Show
            SQL = "Select * From PROGRAMAS Where Programa = '" & Left(Imputacion, 2) & "'"
            rstRegistroIcaro.Open SQL, dbIcaro, adOpenForwardOnly, adLockOptimistic
            .txtPrograma.Text = Left(Imputacion, 2)
            If Len(rstRegistroIcaro.Fields("Descripcion")) <> "0" Then
                .txtDescripcionPrograma.Text = rstRegistroIcaro.Fields("Descripcion")
            Else
                .txtDescripcionPrograma.Text = ""
            End If
            rstRegistroIcaro.Close
            SQL = "Select * From SUBPROGRAMAS Where Subprograma = '" & Left(Imputacion, 5) & "'"
            rstRegistroIcaro.Open SQL, dbIcaro, adOpenForwardOnly, adLockOptimistic
            .txtSubprograma.Text = Mid(Imputacion, 4, 2)
            If Len(rstRegistroIcaro.Fields("Descripcion")) <> "0" Then
                .txtDescripcionSubprograma.Text = rstRegistroIcaro.Fields("Descripcion")
            Else
                .txtDescripcionSubprograma.Text = ""
            End If
            rstRegistroIcaro.Close
            SQL = "Select * From PROYECTOS Where PROYECTO = '" & Left(Imputacion, 8) & "'"
            rstRegistroIcaro.Open SQL, dbIcaro, adOpenForwardOnly, adLockOptimistic
            .txtProyecto.Text = Mid(Imputacion, 6, 2)
            If Len(rstRegistroIcaro.Fields("Descripcion")) <> "0" Then
                .txtDescripcionProyecto.Text = rstRegistroIcaro.Fields("Descripcion")
            Else
                .txtDescripcionProyecto.Text = ""
            End If
            rstRegistroIcaro.Close
            .txtPrograma.Enabled = False
            .txtDescripcionPrograma.Enabled = False
            .txtSubprograma.Enabled = False
            .txtDescripcionSubprograma.Enabled = False
            .txtProyecto.Enabled = False
            .txtDescripcionProyecto.Enabled = False
            .txtActividad.Enabled = True
            .txtDescripcionActividad.Enabled = True
        End Select
    End With
    
    Set rstRegistroIcaro = Nothing
    SQL = ""

End Sub
