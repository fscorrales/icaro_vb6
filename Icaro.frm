VERSION 5.00
Begin VB.MDIForm Icaro 
   BackColor       =   &H8000000C&
   Caption         =   "ICARO"
   ClientHeight    =   4875
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   7380
   Icon            =   "Icaro.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu MnuArchivo 
      Caption         =   "&Archivo"
      Begin VB.Menu MnuEstructuraProgramatica 
         Caption         =   "Estructura Programática"
         Begin VB.Menu MnuListadoEstructura 
            Caption         =   "Listado Estructura"
         End
         Begin VB.Menu LineArchivo1 
            Caption         =   "-"
         End
         Begin VB.Menu MnuCargarPartida 
            Caption         =   "Cargar Partida"
         End
         Begin VB.Menu MnuListadoPartidas 
            Caption         =   "Listado de Partidas"
         End
         Begin VB.Menu LineArchivo2 
            Caption         =   "-"
         End
         Begin VB.Menu MnuCargarFuente 
            Caption         =   "Cargar Fuente"
         End
         Begin VB.Menu MnuListadoFuentes 
            Caption         =   "Listado de Fuentes"
         End
         Begin VB.Menu LineArchivo3 
            Caption         =   "-"
         End
         Begin VB.Menu MnuCargarRetencion 
            Caption         =   "Cargar Retención"
         End
         Begin VB.Menu MnuListadoRetenciones 
            Caption         =   "Listado de Retenciones"
         End
      End
      Begin VB.Menu LineArchivo4 
         Caption         =   "-"
      End
      Begin VB.Menu MnuCargaCuit 
         Caption         =   "Carga Cuit"
         Begin VB.Menu MnuAgregarProveedor 
            Caption         =   "Agregar Proveedor"
         End
         Begin VB.Menu MnuListadoProveedores 
            Caption         =   "Listado de Proveedores"
         End
      End
      Begin VB.Menu MnuCargaCuenta 
         Caption         =   "Cargar Cuenta Bancaria"
         Begin VB.Menu MnuCargarCuentaBancaria 
            Caption         =   "Agregar Cuenta Bancaria"
         End
         Begin VB.Menu MnuListadoCuentasBancarias 
            Caption         =   "Listado Cuentas Bancarias"
         End
      End
      Begin VB.Menu MnuCargaLocalidad 
         Caption         =   "Carga Localidad"
         Begin VB.Menu MnuAgregarLocalidad 
            Caption         =   "Agregar Localidad"
         End
         Begin VB.Menu MnuListadoLocalidades 
            Caption         =   "Listado de Localidades"
         End
      End
      Begin VB.Menu MnuCargaObra 
         Caption         =   "Carga Obra"
         Begin VB.Menu MnuAgregarObra 
            Caption         =   "Agregar Obra"
         End
         Begin VB.Menu MnuListadoObras 
            Caption         =   "Listado de Obras"
         End
         Begin VB.Menu LineArchivo5 
            Caption         =   "-"
         End
         Begin VB.Menu MnuSincronizacionConSGF 
            Caption         =   "Sinc. Desc. con SGF"
         End
         Begin VB.Menu LineArchivo6 
            Caption         =   "-"
         End
         Begin VB.Menu TipoDeObra 
            Caption         =   "Tipo de Obra"
         End
      End
      Begin VB.Menu LineArchivo7 
         Caption         =   "-"
      End
      Begin VB.Menu MnuImportar 
         Caption         =   "Importar"
         Begin VB.Menu MnuEjecucionObras 
            Caption         =   "Ejecucion Obras"
         End
         Begin VB.Menu MnuFondoDeReparo 
            Caption         =   "Fondo de Reparo"
         End
         Begin VB.Menu LineArchivo8 
            Caption         =   "-"
         End
         Begin VB.Menu MnuCondicionTributariaAFIP 
            Caption         =   "Condición Tributaria AFIP"
         End
      End
      Begin VB.Menu LineArchivo9 
         Caption         =   "-"
      End
      Begin VB.Menu MnuSalir 
         Caption         =   "Salir"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu MnuCargar 
      Caption         =   "&Cargar"
      Begin VB.Menu MnuEjecuciondeGasto 
         Caption         =   "Ejecucion de Gasto"
         Begin VB.Menu MnuCargaManualCya 
            Caption         =   "Carga Manual de CyO"
         End
         Begin VB.Menu MnuListadoDeComprobantes 
            Caption         =   "Listado de Comprobantes"
         End
         Begin VB.Menu Line9 
            Caption         =   "-"
         End
         Begin VB.Menu MnuAsistenteDeCargaCertificados 
            Caption         =   "Asistente de Carga Certificados"
         End
         Begin VB.Menu MnuAsistenteDeCargaEPAM 
            Caption         =   "Asistente de Carga EPAM"
         End
      End
      Begin VB.Menu MnuEjecucionDeFondos 
         Caption         =   "Ejecucion de Fondos"
      End
   End
   Begin VB.Menu MnuInformes 
      Caption         =   "&Informes"
      Begin VB.Menu MnuControlObras 
         Caption         =   "Control Obras"
         Begin VB.Menu MnuEjecucionMensual 
            Caption         =   "Ejecución Mensual"
         End
         Begin VB.Menu MnuControlFuente10y13 
            Caption         =   "Control Fte.10 y 13"
         End
      End
      Begin VB.Menu MnuRetencionesPorCtaCte 
         Caption         =   "Retenciones por Cta. Cte."
      End
      Begin VB.Menu MnuEjecucionPorImputacion 
         Caption         =   "Ejecución por Imputacion (Sist. Banco)"
      End
      Begin VB.Menu MnuEpamConsolidado 
         Caption         =   "EPAM Consolidado"
      End
      Begin VB.Menu MnuObrasTerminadas 
         Caption         =   "Obras Terminadas"
         Begin VB.Menu MnuInformeEstructuraPresupuestaria 
            Caption         =   "Informe de Estructura Presupuestaria"
         End
         Begin VB.Menu MnuCertificados100 
            Caption         =   "Detalle Certificados 100%"
         End
      End
   End
   Begin VB.Menu MnuControles 
      Caption         =   "Controles"
      Begin VB.Menu MnuControlEjecucionAcumulada 
         Caption         =   "Ejecucion Acumulada"
         Begin VB.Menu MnuControEjecucionAcumuladaMes 
            Caption         =   "Mes Determinado (rf636m)"
         End
         Begin VB.Menu MnuControEjecucionAcumuladaAño 
            Caption         =   "Ejecicio Completo (rf602)"
         End
      End
   End
End
Attribute VB_Name = "Icaro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MDIForm_Load()

    Dim datFechaLimite As Date
     
    datFechaLimite = #3/15/2019#
     
    If Date < datFechaLimite Then
        Conectar
    End If

End Sub

Private Sub MDIForm_Unload(Cancel As Integer)

    Desconectar

End Sub

Private Sub MnuAgregarObra_Click()

    CargaObra.Show
    
End Sub

Private Sub MnuAgregarProveedor_Click()

    CargaCUIT.Show
    
End Sub

Private Sub MnuAsistenteDeCargaCertificados_Click()

    AutocargaCertificados.Show
    ConfigurardgAutocargaCertificados
    CargardgAutocargaCertificados

End Sub

Private Sub MnuAsistenteDeCargaEPAM_Click()

    AutocargaEPAM.Show
    ConfigurardgAutocargaEPAM
    CargardgAutocargaEPAM

End Sub


Private Sub MnuCargaManualCya_Click()

    CargaGasto.Show
    
End Sub

Private Sub MnuCargarCuentaBancaria_Click()

    CargaCuentaBancaria.Show
    
End Sub

Private Sub MnuCargarFuente_Click()

    CargaFuente.Show
    
End Sub

Private Sub MnuAgregarLocalidad_Click()

    CargaLocalidad.Show
    
End Sub

Private Sub MnuCargarPartida_Click()

    CargaPartida.Show
    
End Sub

Private Sub MnuCargarRetencion_Click()

    CodigoRetencion.Show
    
End Sub

Private Sub MnuCertificados100_Click()
    
    strEjercicio = InputBox("Escriba el Año que desea consultar", "Obras Terminadas", Format(Date, "yyyy"))
    If strEjercicio <> "" Then
        Certificados100.Show
    End If
    
End Sub

Private Sub MnuCondicionTributariaAFIP_Click()
    'ImportacionAFIP.Show
End Sub

Private Sub MnuControEjecucionAcumuladaAño_Click()
    
    Dim intRespuesta As Integer
    
    strControlCruzadoSIIF = "rf602"
    intRespuesta = MsgBox("¿Tiene acceso a generar archivos .txt directamente del SIIF? " _
    & vbCrLf & "De ser así, se recomienda que conteste en forma afirmativa y prosiga con los pasos " _
    & "que a continuación se detallan", vbQuestion + vbYesNo + vbDefaultButton2, "TIPO DE ARCHIVO A EXPORTAR SIIF")
    If intRespuesta = 6 Then
        MsgBox "Antes de iniciar el control cruzado, tenga en cuenta lo siguiente:" _
        & vbCrLf & "1) GENERAR el reporte, de tipo .txt, " & strControlCruzadoSIIF & " del Modulo Gasto del SIIF." _
        & vbCrLf & "2) GUARDAR el archivo generado en una ubicación que recuerde." _
        & vbCrLf & "3) En la siguiente ventana, COLOCAR la FECHA límite hasta donde ICARO acumulará la ejecución." _
        & vbCrLf & "4) Luego de pulsar el boton IMPORTAR, busque el archivo guardado en el punto 2).", , "CONTROL CRUZADO ICARO VS SIIF (opción .txt)"
        strControlCruzadoSIIF = "rf602txt"
    Else
        MsgBox "Antes de iniciar el control cruzado, tenga en cuenta lo siguiente:" _
        & vbCrLf & "1) Se recomienda utilizar PDF XChange Viewer como visor predeterminado de .pdf" _
        & vbCrLf & "2) GENERAR el reporte, de tipo .pdf, " & strControlCruzadoSIIF & " del Modulo Gasto del SIIF." _
        & vbCrLf & "3) Seleccionar TODOS los datos del reporte generado (.pdf) y COPIAR los mismos." _
        & vbCrLf & "4) PEGAR los datos copiados en un BLOC DE NOTAS nuevo." _
        & vbCrLf & "5) Utilizar la opción EDICIÓN/REEMPLAZAR, del Bloc de Notas, para sustituir TODOS los '.' por ' ' (espacio vacio)." _
        & vbCrLf & "6) GUARDAR el BLOC DE NOTAS en una ubicación que recuerde." _
        & vbCrLf & "7) En la siguiente ventana, COLOCAR la FECHA límite hasta donde ICARO acumulará la ejecución." _
        & vbCrLf & "8) Luego de pulsar el boton IMPORTAR, busque el archivo guardado en el punto 6).", , "CONTROL CRUZADO ICARO VS SIIF (opción .pdf)"
        strControlCruzadoSIIF = "rf602pdf"
    End If
    ControlCruzadoSIIF.Show
    ControlCruzadoSIIF.txtFecha.Text = Format(Now(), "dd/mm/yyyy")

End Sub

Private Sub MnuControEjecucionAcumuladaMes_Click()

    strControlCruzadoSIIF = "rf636m"
    MsgBox "Antes de iniciar el control cruzado, tenga en cuenta lo siguiente:" _
    & vbCrLf & "1) Se recomienda utilizar PDF XChange Viewer como visor predeterminado de .pdf" _
    & vbCrLf & "2) GENERAR el reporte, de tipo .pdf, " & strControlCruzadoSIIF & " del Modulo Gasto del SIIF." _
    & vbCrLf & "3) Seleccionar TODOS los datos del reporte generado (.pdf) y COPIAR los mismos." _
    & vbCrLf & "4) PEGAR los datos copiados en un BLOC DE NOTAS nuevo." _
    & vbCrLf & "5) Utilizar la opción EDICIÓN/REEMPLAZAR, del Bloc de Notas, para sustituir TODOS los '.' por ' ' (espacio vacio)." _
    & vbCrLf & "6) GUARDAR el BLOC DE NOTAS en una ubicación que recuerde." _
    & vbCrLf & "7) En la siguiente ventana, COLOCAR la FECHA límite hasta donde ICARO acumulará la ejecución." _
    & vbCrLf & "8) Luego de pulsar el boton IMPORTAR, busque el archivo guardado en el punto 6).", , "CONTROL CRUZADO ICARO VS SIIF (opción .pdf)"
    ControlCruzadoSIIF.Show
    ControlCruzadoSIIF.txtFecha.Text = Format(Now(), "dd/mm/yyyy")

End Sub

Private Sub MnuControlFuente10y13_Click()
    
    ListadoControlFuente10y13.Show
    ConfigurardgEjecucionFuente10y13
    
End Sub

Private Sub MnuControlRetenciones_Click()
    'ListadoControlRetenciones.Show
    'ConfigurardgControlRetenciones
End Sub

Private Sub MnuEjecucionMensual_Click()

    strEjercicio = InputBox("Escriba el Año que desea consultar", "Ejercicio Presupuestario", Format(Date, "yyyy"))
    If strEjercicio <> "" Then
        EjecucionMensual.Show
    End If
    
End Sub

Private Sub MnuEjecucionObras_Click()

    ImportacionObras.Show
    ImportacionObras.cmdFR.Visible = False
    
End Sub

Private Sub MnuEjecucionPorImputacion_Click()

    EjecucionPorImputacion.Show
    EjecucionPorImputacion.txtFechaDesde.Text = "01/01/" & Format(Now(), "yyyy")
    EjecucionPorImputacion.txtFechaHasta.Text = Format(Now(), "dd/mm/yyyy")
    Call ConfigurardgEjecucionPorImputacion
    Call ConfigurardgEjecucionPorImputacionSeleccionada
    
End Sub

Private Sub MnuEpamConsolidado_Click()
    
    Dim i As Integer

    With EPAMConsolidado
        i = .dgEPAMConsolidado.Row
        .Show
        ConfigurardgEPAMConsolidado
        CargardgEPAMConsolidado
        ConfigurardgResolucionesEPAM
        CargardgResolucionesEPAM (.dgEPAMConsolidado.TextMatrix(i, 0))
    End With

End Sub

Private Sub MnuFondoDeReparo_Click()

    ImportacionObras.Show
    ImportacionObras.cmdDemandaLibre.Visible = False
    ImportacionObras.cmdAutogestivo.Visible = False
    
End Sub

Private Sub MnuInformeEstructuraPresupuestaria_Click()
    
    strEjercicio = InputBox("Escriba el Año que desea consultar", "Obras Terminadas", Format(Date, "yyyy"))
    If strEjercicio <> "" Then
        ObrasTerminadasPorEstructura.Show
    End If
    
End Sub

Private Sub MnuListadoCuentasBancarias_Click()
    
    ListadoCuentasBancarias.Show
    ConfigurardgCuentasBancarias
    CargardgCuentasBancarias
    
End Sub

Private Sub MnuListadoDeComprobantes_Click()
    
    Dim x As Integer
    
    With ListadoImputacion
        .Show
        ConfigurardgListadoGasto
        Call CargardgListadoGasto(, Year(Now()))
        .txtFecha = Year(Now())
        x = .dgListado.Row
        ConfigurardgImputacion
        CargardgImputacion (.dgListado.TextMatrix(x, 0))
        ConfigurardgRetencion
        CargardgRetencion (.dgListado.TextMatrix(x, 0))
        x = 0
    End With

End Sub

Private Sub MnuListadoEstructura_Click()

    EstructuraProgramatica.Show
    ConfigurardgEstructura
    Call CargardgEstructura
    
End Sub

Private Sub MnuListadoFuentes_Click()

    ListadoFuentes.Show
    ConfigurardgFuentes
    CargardgFuentes
    
End Sub

Private Sub MnuListadoLocalidades_Click()
    
    ListadoLocalidades.Show
    ConfigurardgLocalidades
    CargardgLocalidades

End Sub

Private Sub MnuListadoObras_Click()
    
    ListadoObras.Show
    ConfigurardgObras
    CargardgObras
    
End Sub

Private Sub MnuListadoPartidas_Click()
    
    ListadoPartidas.Show
    ConfigurardgPartidas
    CargardgPartidas
    
End Sub

Private Sub MnuListadoProveedores_Click()
    
    ListadoProveedores.Show
    ConfigurardgProveedores
    CargardgProveedores
    
End Sub

Private Sub MnuListadoRetenciones_Click()

    ListadoRetenciones.Show
    ConfigurardgCodigoRetenciones
    CargardgCodigoRetenciones
    
End Sub

Private Sub MnuRetencionesPorCtaCte_Click()
        
    RetencionesPorCtaCte.Show
    RetencionesPorCtaCte.txtFechaDesde.Text = "01/01/" & Format(Now(), "yyyy")
    RetencionesPorCtaCte.txtFechaHasta.Text = Format(Now(), "dd/mm/yyyy")
        
End Sub

Private Sub MnuSalir_Click()

    DesconectarRecordset
    dbBase.Close
    End
    
End Sub

Private Sub MnuSincronizacionConSGF_Click()

    Dim i As Integer

    With SincronizacionDescripcionSGF
        .Show
        ConfigurarSincronizaciondgObrasSGF
        CargarSincronizaciondgObrasSGF
        i = .dgObrasSGF.Row
        If i > 0 Then
            .txtModificacion.Text = .dgObrasSGF.TextMatrix(i, 1)
            ConfigurarSincronizaciondgObrasICARO
            CargarSincronizaciondgObrasICARO (Left(.txtModificacion.Text, 50))
        Else
            .txtModificacion.Text = ""
            ConfigurarSincronizaciondgObrasICARO
        End If
        i = .dgObrasICARO.Row
        If i > 0 Then
            .txtPrevioModificacion.Text = .dgObrasICARO.TextMatrix(i, 1)
        Else
            .txtPrevioModificacion.Text = ""
        End If
    End With

End Sub

Private Sub TipoDeObra_Click()
    
    CargaTipoDeObra.Show
    ConfigurardgTipoDeObra
    CargardgTipoDeObra
    
End Sub
