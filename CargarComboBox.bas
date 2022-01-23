Attribute VB_Name = "CargarComboBox"
Public Sub CargarcmbProveedor()
    
    Set rstRegistroIcaro = New ADODB.Recordset
    rstRegistroIcaro.Open "Select * From CUIT Order By Nombre", dbIcaro, adOpenForwardOnly, adLockOptimistic
 
    If rstRegistroIcaro.BOF = False Then
        With rstRegistroIcaro
            .MoveFirst
            While .EOF = False
                ListadoControlRetenciones.cmbProveedor.AddItem rstRegistroIcaro!Nombre
                .MoveNext
            Wend
        End With
    End If
    
    ListadoControlRetenciones.cmbProveedor.AddItem "Todos los Proveedores"
    ListadoControlRetenciones.cmbProveedor.Text = "Todos los Proveedores"
    
    rstRegistroIcaro.Close
    Set rstRegistroIcaro = Nothing
    dbIcaro.Close
    Set dbIcaro = Nothing

End Sub

Public Sub CargarComboImputacionEPAM()

    Dim SQL As String

    Set rstListadoIcaro = New ADODB.Recordset
    SQL = "SELECT DISTINCT Codigo From EPAM WHERE Codigo <> ALL (SELECT Codigo FROM DATOSFIJOSEPAM)"
    rstListadoIcaro.Open SQL, dbIcaro, adOpenForwardOnly, adLockOptimistic
    If rstListadoIcaro.BOF = False Then
        With rstListadoIcaro
            .MoveFirst
            While .EOF = False
                ImputacionEPAM.txtObra.AddItem .Fields("Codigo")
            .MoveNext
            Wend
        End With
    End If
    rstListadoIcaro.Close
    Set rstListadoIcaro = Nothing
    
    SQL = ""

End Sub

Public Sub CargarComboRetenciones()

    Dim SQL As String

    Set rstListadoIcaro = New ADODB.Recordset
    SQL = "Select Codigo From CODIGOSRETENCIONES Order by Codigo"
    rstListadoIcaro.Open SQL, dbIcaro, adOpenForwardOnly, adLockOptimistic
    If rstListadoIcaro.BOF = False Then
        With rstListadoIcaro
            .MoveFirst
            While .EOF = False
                CargaRetencion.txtRetencion.AddItem .Fields("Codigo")
            .MoveNext
            Wend
        End With
    End If
    rstListadoIcaro.Close
    Set rstListadoIcaro = Nothing
    
    SQL = ""

End Sub

Public Sub CargarCombosGasto()

    Set rstListadoIcaro = New ADODB.Recordset
    
    rstListadoIcaro.Open "Select * From CUIT Order By CUIT", dbIcaro, adOpenDynamic, adLockOptimistic
    With CargaGasto.txtCUIT
        If rstListadoIcaro.BOF = False Then
            rstListadoIcaro.MoveFirst
            While rstListadoIcaro.EOF = False
                .AddItem rstListadoIcaro!CUIT
                rstListadoIcaro.MoveNext
            Wend
        End If
    End With
    rstListadoIcaro.Close

    rstListadoIcaro.Open "Select * From CUENTAS Order By CUENTA", dbIcaro, adOpenDynamic, adLockOptimistic
    With CargaGasto.txtCuenta
        If rstListadoIcaro.BOF = False Then
            rstListadoIcaro.MoveFirst
            While rstListadoIcaro.EOF = False
                .AddItem rstListadoIcaro!Cuenta
                rstListadoIcaro.MoveNext
            Wend
        End If
    End With
    rstListadoIcaro.Close
    
    rstListadoIcaro.Open "Select * From FUENTES Order By FUENTE", dbIcaro, adOpenDynamic, adLockOptimistic
    With CargaGasto.txtFuente
        If rstListadoIcaro.BOF = False Then
            rstListadoIcaro.MoveFirst
            While rstListadoIcaro.EOF = False
                .AddItem rstListadoIcaro!Fuente
                rstListadoIcaro.MoveNext
            Wend
        End If
    End With
    rstListadoIcaro.Close
    Set rstListadoIcaro = Nothing

End Sub

Public Sub CargarComboObraDeGasto(CUIT As String)
    
    Dim SQL As String
    
    Set rstListadoIcaro = New ADODB.Recordset
    SQL = "Select * From OBRAS Where CUIT = '" & CUIT & "' Order By DESCRIPCION"
    rstListadoIcaro.Open SQL, dbIcaro, adOpenDynamic, adLockOptimistic
    With CargaGasto.txtObra
        If rstListadoIcaro.BOF = False Then
            rstListadoIcaro.MoveFirst
            While rstListadoIcaro.EOF = False
                .AddItem rstListadoIcaro!Descripcion
                rstListadoIcaro.MoveNext
            Wend
        End If
    End With
    rstListadoIcaro.Close
    Set rstListadoIcaro = Nothing

End Sub

Public Sub CargarCombosObra()
        
    Set rstListadoIcaro = New ADODB.Recordset
    
    rstListadoIcaro.Open "Select * From CUIT Order By CUIT", dbIcaro, adOpenDynamic, adLockOptimistic
    With CargaObra.txtCUIT
        If rstListadoIcaro.BOF = False Then
            rstListadoIcaro.MoveFirst
            While rstListadoIcaro.EOF = False
                .AddItem rstListadoIcaro!CUIT
                rstListadoIcaro.MoveNext
            Wend
        End If
    End With
    rstListadoIcaro.Close

    rstListadoIcaro.Open "Select * From CUENTAS Order By CUENTA", dbIcaro, adOpenDynamic, adLockOptimistic
    With CargaObra.txtCuenta
        If rstListadoIcaro.BOF = False Then
            rstListadoIcaro.MoveFirst
            While rstListadoIcaro.EOF = False
                .AddItem rstListadoIcaro!Cuenta
                rstListadoIcaro.MoveNext
            Wend
        End If
    End With
    rstListadoIcaro.Close
    
    rstListadoIcaro.Open "Select * From FUENTES Order By FUENTE", dbIcaro, adOpenDynamic, adLockOptimistic
    With CargaObra.txtFuente
        If rstListadoIcaro.BOF = False Then
            rstListadoIcaro.MoveFirst
            While rstListadoIcaro.EOF = False
                .AddItem rstListadoIcaro!Fuente
                rstListadoIcaro.MoveNext
            Wend
        End If
    End With
    rstListadoIcaro.Close
    
    rstListadoIcaro.Open "Select * From LOCALIDADES Order By LOCALIDAD", dbIcaro, adOpenDynamic, adLockOptimistic
    With CargaObra.txtLocalidad
        If rstListadoIcaro.BOF = False Then
            rstListadoIcaro.MoveFirst
            While rstListadoIcaro.EOF = False
                .AddItem rstListadoIcaro!Localidad
                rstListadoIcaro.MoveNext
            Wend
        End If
    End With
    rstListadoIcaro.Close
    Set rstListadoIcaro = Nothing

End Sub

'Public Sub CargarCmbEjercicioRetencionesPorCtaCte()
'
'    Dim SQL As String
'    Dim i As Integer
'
'    Set rstListadoIcaro = New ADODB.Recordset
'
'    SQL = "Select Year(Fecha) As Ejercicio From CARGA " _
'    & "Group by Year(Fecha) " _
'    & "Order By Year(Fecha) Desc"
'
'    rstListadoIcaro.Open SQL, dbIcaro, adOpenDynamic, adLockOptimistic
'    With RetencionesPorCtaCte.cmbAno
'        If rstListadoIcaro.BOF = False Then
'            rstListadoIcaro.MoveFirst
'            While rstListadoIcaro.EOF = False
'                .AddItem rstListadoIcaro!Ejercicio
'                rstListadoIcaro.MoveNext
'            Wend
'        End If
'    End With
'    rstListadoIcaro.Close
'    Set rstListadoIcaro = Nothing
'
'    With RetencionesPorCtaCte.cmbMes
'        For i = 1 To 12
'            .AddItem i
'        Next i
'        .AddItem "1er Semestre"
'        .AddItem "2do Semestre"
'        .AddItem "Año Completo"
'    End With
'
'End Sub
