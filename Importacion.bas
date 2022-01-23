Attribute VB_Name = "Importacion"
Option Explicit

Public strControlCruzadoSIIF As String

Public Registros As Collection
Public Registro As New CCertificados
Public Fields As Collection
Public Field As New CEPAM
Dim Certificado As CCertificados
Dim EPAMCertificado As CEPAM
Dim EPAMProveedor As CEPAM
Dim intCuenta As Integer
Dim strProveedor As String
Dim Periodo As Date
'Variable de tipo Aplicación de Excel
Dim objExcel As Object
'Una variable de tipo Libro de Excel
Dim xLibro As Object
Dim Col As Integer, Fila As Integer
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long


' Importante : Agregar la referencia a Micorosft Excel xx object library
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Sub ImportarDemandaLibre(Direccion As String)
    If Direccion = "" Then
        Unload ImportacionObras
        Exit Sub
    End If
    On Error GoTo HayError
    Set Registros = New Collection
    'creamos un nuevo objeto excel
    Set objExcel = CreateObject("Excel.Application")
    'Usamos el método open para abrir el archivo que está _
     en el directorio del programa llamado archivo.xls
    objExcel.Workbooks.Open FileName:=(Direccion)
    'Hacemos el Excel NO sea Visible
    Set xLibro = objExcel
    objExcel.Visible = False

    With xLibro.Sheets(1)
        If ImportacionObras.cmdDemandaLibre.Visible = True Then
            If .Cells(2, 9) <> "Resumen de Certificaciones: " Then
                Call InfoGeneral("ARCHIVO INCORRECTO", ImportacionObras)
                DesconectarDemandaLibre
                Exit Sub
            End If
        Else
            If .Cells(2, 9) <> "Devoluc. de Fondo de Reparo: " Then
                Call InfoGeneral("ARCHIVO INCORRECTO", ImportacionObras)
                DesconectarDemandaLibre
                Exit Sub
            End If
        End If
        Fila = 11
        Periodo = .Cells(2, 10)
        While Trim(.Cells(Fila, 4)) = ""
            Fila = Fila + 1
            If Trim(.Cells(Fila, 5)) = "" Then
                If Trim(.Cells(Fila, 1)) <> "" Then
                    strProveedor = .Cells(Fila, 1)
                End If
            Else
                Set Certificado = New CCertificados
                If Trim(.Cells(Fila, 6)) = "" Then
                    Certificado.Proveedor = strProveedor
                    Certificado.Obra = .Cells(Fila, 1)
                    Certificado.Certificado = .Cells(Fila, 5)
                    Certificado.FondoDeReparo = .Cells(Fila, 8)
                    If .Cells(Fila, 5) = "FR" Then
                        Certificado.FondoDeReparo = .Cells(Fila, 10)
                    End If
                    Certificado.MontoBruto = .Cells(Fila, 10)
                    Certificado.IB = .Cells(Fila, 11)
                    Certificado.LP = .Cells(Fila, 12)
                    Certificado.SUSS = .Cells(Fila, 13)
                    Certificado.Ganancias = .Cells(Fila, 14)
                    Certificado.INVICO = .Cells(Fila, 15)
                    Certificado.Periodo = Periodo
                    Registros.Add Certificado
                Else
                    Certificado.Proveedor = strProveedor
                    Certificado.Obra = .Cells(Fila, 1)
                    Certificado.Certificado = .Cells(Fila, 5)
                    Certificado.FondoDeReparo = .Cells(Fila, 7)
                    If .Cells(Fila, 5) = "FR" Then
                        Certificado.FondoDeReparo = .Cells(Fila, 9)
                    End If
                    Certificado.MontoBruto = .Cells(Fila, 9)
                    Certificado.IB = .Cells(Fila, 10)
                    Certificado.LP = .Cells(Fila, 11)
                    Certificado.SUSS = .Cells(Fila, 12)
                    Certificado.Ganancias = .Cells(Fila, 13)
                    Certificado.INVICO = .Cells(Fila, 14)
                    Certificado.Periodo = Periodo
                    Registros.Add Certificado
                End If
            End If
        Wend
    End With
      
    Call InfoGeneral("", ImportacionObras)
    Call InfoGeneral("REGISTROS CAPTURADOS = " & Registros.Count, ImportacionObras)
    Call InfoGeneral("", ImportacionObras)

    InsertarDemandaLibre
    
    Exit Sub
    
HayError:
    MsgBox Err.Number & " - " & Err.Description, , "Error!!!"
    DesconectarDemandaLibre
    RemoverDemandaLibre
    
End Sub

Private Sub InsertarDemandaLibre()
    
    Dim SQL As String
    Dim intCuenta As Integer
    intCuenta = Registros.Count
    On Error GoTo HayError
    Set rstRegistroIcaro = New ADODB.Recordset
    rstRegistroIcaro.Open "CERTIFICADOS", dbIcaro, adOpenForwardOnly, adLockOptimistic
    For Each Registro In Registros
        With Registro
            SQL = "Select Obra, Certificado From CERTIFICADOS Where Obra = '" & .Obra & "' And Certificado = '" & .Certificado & "' " _
            & "Union Select Obra, Certificado From CARGA Where Obra = '" & .Obra & "' And Certificado = '" & .Certificado & "'"
            If SQLNoMatch(SQL) = False Then
                Call InfoGeneral("ERROR AL INGRESAR CERTIFICADO - " & .Obra & " - " & .Certificado, ImportacionObras)
                intCuenta = intCuenta - 1
            Else
                rstRegistroIcaro.AddNew
                rstRegistroIcaro.Fields("Comprobante") = ""
                rstRegistroIcaro.Fields("Proveedor") = .Proveedor
                rstRegistroIcaro.Fields("Obra") = .Obra
                rstRegistroIcaro.Fields("Certificado") = .Certificado
                rstRegistroIcaro.Fields("FondoDeReparo") = .FondoDeReparo
                rstRegistroIcaro.Fields("MontoBruto") = .MontoBruto
                rstRegistroIcaro.Fields("IB") = .IB
                rstRegistroIcaro.Fields("LP") = .LP
                rstRegistroIcaro.Fields("SUSS") = .SUSS
                rstRegistroIcaro.Fields("Ganancias") = .Ganancias
                rstRegistroIcaro.Fields("INVICO") = .INVICO
                rstRegistroIcaro.Fields("Periodo") = .Periodo
                rstRegistroIcaro.Update
            End If
        End With
    Next Registro
    
    Call InfoGeneral("", ImportacionObras)
    Call InfoGeneral("REGISTROS AGREGADOS = " & intCuenta, ImportacionObras)
    Call InfoGeneral("", ImportacionObras)
    
    DesconectarDemandaLibre
    RemoverDemandaLibre
    
    Exit Sub
    
HayError:
    MsgBox Err.Number & " - " & Err.Description, , "Error!!!"
    DesconectarDemandaLibre
    RemoverDemandaLibre

End Sub

Private Sub DesconectarDemandaLibre()
    
    objExcel.ActiveWorkbook.Close False
    objExcel.Quit
    Set Certificado = Nothing
    Set objExcel = Nothing
    Set xLibro = Nothing
    'rstRegistroIcaro.Close
    Set rstRegistroIcaro = Nothing

End Sub

Private Sub RemoverDemandaLibre()
    Dim i As Long
    
    For i = Registros.Count To 1 Step -1
        Registros.Remove (i)
    Next i
    
    Debug.Print Registros.Count
End Sub

Public Sub ImportarEPAM(Direccion As String)
    If Direccion = "" Then
        Unload ImportacionObras
        Exit Sub
    End If
    On Error GoTo HayError
    Set Fields = New Collection
    'creamos un nuevo objeto excel
    Set objExcel = CreateObject("Excel.Application")
    'Usamos el método open para abrir el archivo que está _
     en el directorio del programa llamado archivo.xls
    objExcel.Workbooks.Open FileName:=(Direccion)
    'Hacemos el Excel NO sea Visible
    Set xLibro = objExcel
    objExcel.Visible = False

    Dim strCodigoObra As String

    Dim FechaInicio As String
    Dim FechaCierre As String
        
    With xLibro.Sheets(1)
        If .Cells(5, 8) <> "Resumen de Rendiciones (por Obras)" Then
            Call InfoGeneral("ARCHIVO INCORRECTO", ImportacionObras)
            DesconectarEPAM
            Exit Sub
        End If
        Fila = 17
        FechaInicio = Trim(.Cells(8, 7))
        FechaCierre = Trim(.Cells(8, 9))
        While .Cells(Fila, 1) <> "TOTALES : "
            Fila = Fila + 1
            If Trim(.Cells(Fila, 5)) = "" Then
                If Trim(.Cells(Fila, 1)) <> "" Then
                    strCodigoObra = .Cells(Fila, 1)
                    Set EPAMCertificado = New CEPAM
                    'Set EPAMProveedor = New CEPAM
                End If
            Else
                While Trim(.Cells(Fila, 1)) <> ""
                    'If Left(Trim(.Cells(Fila, 5)), 4) = "CERT" Then
                        EPAMCertificado.CodigoObra = strCodigoObra
                        EPAMCertificado.Tipo = "CERTIFICADO"
                        EPAMCertificado.MontoBruto = EPAMCertificado.MontoBruto + .Cells(Fila, 17)
                        If Trim(.Cells(Fila, 9)) = "" Then
                            EPAMCertificado.Ganancias = EPAMCertificado.Ganancias + 0
                        Else
                            EPAMCertificado.Ganancias = EPAMCertificado.Ganancias + .Cells(Fila, 9)
                        End If
                        If Trim(.Cells(Fila, 10)) = "" Then
                            EPAMCertificado.Sellos = EPAMCertificado.Sellos + 0
                        Else
                            EPAMCertificado.Sellos = EPAMCertificado.Sellos + .Cells(Fila, 10)
                        End If
                        If Trim(.Cells(Fila, 11)) = "" Then
                            EPAMCertificado.LP = EPAMCertificado.LP + 0
                        Else
                            EPAMCertificado.LP = EPAMCertificado.LP + .Cells(Fila, 11)
                        End If
                        If Trim(.Cells(Fila, 12)) = "" Then
                            EPAMCertificado.IB = EPAMCertificado.IB + 0
                        Else
                            EPAMCertificado.IB = EPAMCertificado.IB + .Cells(Fila, 12)
                        End If
                        If Trim(.Cells(Fila, 13)) = "" Then
                            EPAMCertificado.SUSS = EPAMCertificado.SUSS + 0
                        Else
                            EPAMCertificado.SUSS = EPAMCertificado.SUSS + .Cells(Fila, 13)
                        End If
                        EPAMCertificado.FechaInicio = FechaInicio
                        EPAMCertificado.FechaCierre = FechaCierre
                    'Else
                    '    EPAMProveedor.CodigoObra = strCodigoObra
                    '    EPAMProveedor.Tipo = "PROVEEDOR"
                    '    EPAMProveedor.MontoBruto = EPAMProveedor.MontoBruto + .Cells(Fila, 17)
                    '    If Trim(.Cells(Fila, 9)) = "" Then
                    '        EPAMProveedor.Ganancias = EPAMProveedor.Ganancias + 0
                    '    Else
                    '        EPAMProveedor.Ganancias = EPAMProveedor.Ganancias + .Cells(Fila, 9)
                    '    End If
                    '    If Trim(.Cells(Fila, 10)) = "" Then
                    '        EPAMProveedor.Sellos = EPAMProveedor.Sellos + 0
                    '    Else
                    '        EPAMProveedor.Sellos = EPAMProveedor.Sellos + .Cells(Fila, 10)
                    '    End If
                    '    If Trim(.Cells(Fila, 11)) = "" Then
                    '        EPAMProveedor.LP = EPAMProveedor.LP + 0
                    '    Else
                    '        EPAMProveedor.LP = EPAMProveedor.LP + .Cells(Fila, 11)
                    '    End If
                    '    If Trim(.Cells(Fila, 12)) = "" Then
                    '        EPAMProveedor.IB = EPAMProveedor.IB + 0
                    '    Else
                    '        EPAMProveedor.IB = EPAMProveedor.IB + .Cells(Fila, 12)
                    '    End If
                    '    If Trim(.Cells(Fila, 13)) = "" Then
                    '        EPAMProveedor.SUSS = EPAMProveedor.SUSS + 0
                    '    Else
                    '        EPAMProveedor.SUSS = EPAMProveedor.SUSS + .Cells(Fila, 13)
                    '    End If
                    '    EPAMProveedor.FechaInicio = FechaInicio
                    '    EPAMProveedor.FechaCierre = FechaCierre
                    'End If
                    Fila = Fila + 1
                Wend
                If Not Len(EPAMCertificado.CodigoObra) = 0 Then
                    Fields.Add EPAMCertificado
                    Set EPAMCertificado = New CEPAM
                End If
                'If Not Len(EPAMProveedor.CodigoObra) = 0 Then
                '    Fields.Add EPAMProveedor
                '    Set EPAMProveedor = New CEPAM
                'End If
            End If
        Wend
    End With
    
    For Each Field In Fields
        With Field
            Debug.Print .CodigoObra, .Tipo, .MontoBruto, .Ganancias, .Sellos, .LP, .IB, .SUSS, .FechaInicio, .FechaCierre
        End With
    Next
    Debug.Print Fields.Count
    
    InsertarEPAM
    
    Exit Sub
    
HayError:
    MsgBox Err.Number & " - " & Err.Description, , "Error!!!"
    DesconectarEPAM
    RemoverEPAM

End Sub

Private Sub InsertarEPAM()
    
    Dim intCuenta As Integer
    Dim SQL As String
    intCuenta = Fields.Count
    On Error GoTo HayError
    Set rstRegistroIcaro = New ADODB.Recordset
    rstRegistroIcaro.Open "EPAM", dbIcaro, adOpenForwardOnly, adLockOptimistic
    For Each Field In Fields
        With Field
            SQL = "Select * From EPAM Where Codigo =" & "'" & .CodigoObra & "'" & "AND MontoBruto =" & .MontoBruto & "AND Tipo =" & "'" & .Tipo & "'" & "AND Inicio =" & "'" & .FechaInicio & "'" & "AND Cierre =" & "'" & .FechaCierre & "'"
            If SQLNoMatch(SQL) = False Then
                Call InfoGeneral("ERROR AL INGRESAR COMPROBANTE - " & .CodigoObra & " - " & .Tipo & " - " & .MontoBruto, ImportacionObras)
                intCuenta = intCuenta - 1
            Else
                rstRegistroIcaro.AddNew
                rstRegistroIcaro.Fields("Comprobante") = ""
                rstRegistroIcaro.Fields("Codigo") = .CodigoObra
                rstRegistroIcaro.Fields("Tipo") = .Tipo
                rstRegistroIcaro.Fields("MontoBruto") = .MontoBruto
                rstRegistroIcaro.Fields("Ganancias") = .Ganancias
                rstRegistroIcaro.Fields("Sellos") = .Sellos
                rstRegistroIcaro.Fields("IB") = .IB
                rstRegistroIcaro.Fields("LP") = .LP
                rstRegistroIcaro.Fields("SUSS") = .SUSS
                rstRegistroIcaro.Fields("Inicio") = .FechaInicio
                rstRegistroIcaro.Fields("Cierre") = .FechaCierre
                rstRegistroIcaro.Update
            End If
        End With
    Next Field
    
    Call InfoGeneral("", ImportacionObras)
    Call InfoGeneral("REGISTROS AGREGADOS = " & intCuenta, ImportacionObras)
    Call InfoGeneral("", ImportacionObras)
    
    DesconectarEPAM
    RemoverEPAM
    
    Exit Sub
    
HayError:
    MsgBox Err.Number & " - " & Err.Description, , "Error!!!"
    DesconectarEPAM
    RemoverEPAM

End Sub

Private Sub DesconectarEPAM()
    objExcel.ActiveWorkbook.Close False
    objExcel.Quit
    Set EPAMCertificado = Nothing
    Set EPAMProveedor = Nothing
    Set objExcel = Nothing
    Set xLibro = Nothing
    'rstRegistroIcaro.Close
    Set rstRegistroIcaro = Nothing

End Sub

Private Sub RemoverEPAM()
    Dim i As Long
    
    For i = Fields.Count To 1 Step -1
        Fields.Remove (i)
    Next i
    
    Debug.Print Fields.Count
End Sub

Public Sub ImportarAFIP()

    
    Dim strDireccion As String
    Dim intNumeroArchivo As Integer
    Dim strLineaImportada As String
    
    ImportacionAFIP.txtInforme.Text = ""
    Call InfoGeneral("- INICIANDO IMPORTACIÓN BASE DE DATOS AFIP -", ImportacionAFIP)
    Call InfoGeneral("", ImportacionAFIP)
    ImportacionAFIP.dlgMultifuncion.FileName = ""
    ImportacionAFIP.dlgMultifuncion.ShowOpen
    strDireccion = ImportacionAFIP.dlgMultifuncion.FileName
    If strDireccion = "" Then 'En caso de apretar cancelar
        Call InfoGeneral("IMPORTACION CANCELADA", ImportacionAFIP)
        Exit Sub
    Else
        
        'Conectando a Base de Datos AFIP
        Set dbAFIP = New ADODB.Connection
        dbAFIP.CursorLocation = adUseClient
        dbAFIP.Mode = adModeUnknown
        dbAFIP.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\AFIP.mdb"
        dbAFIP.Open
    
        Call InfoGeneral("Datos importados desde = " & strDireccion, ImportacionAFIP)
        Call InfoGeneral("", ImportacionAFIP)
        Call InfoGeneral("El proceso puede demorar varios minutos", ImportacionAFIP)
        Call InfoGeneral("", ImportacionAFIP)
        intNumeroArchivo = FreeFile
        Open strDireccion For Input As #intNumeroArchivo
        strDireccion = ""
        
        'Creando Recordset
        Set rstRegistroAFIP = New ADODB.Recordset
        rstRegistroAFIP.Open "CONDICIONTRIBUTARIA", dbAFIP, adOpenForwardOnly, adLockOptimistic

        Do Until EOF(intNumeroArchivo)
            Line Input #intNumeroArchivo, strLineaImportada
            With rstRegistroAFIP
                .AddNew
                .Fields("CUIT") = Left$(strLineaImportada, 11)
                .Fields("Denominacion") = Mid$(strLineaImportada, 12, 30)
                .Fields("Ganancias") = Mid$(strLineaImportada, 42, 2)
                .Fields("IVA") = Mid$(strLineaImportada, 44, 2)
                .Fields("Monotributo") = Mid$(strLineaImportada, 46, 2)
                .Fields("IntegranteSociedad") = Mid$(strLineaImportada, 48, 1)
                .Fields("Empleador") = Mid$(strLineaImportada, 49, 1)
                .Fields("ActividadMonotributo") = Mid$(strLineaImportada, 50, 2)
                .Update
                strLineaImportada = ""
            End With
        Loop
        Call InfoGeneral("PROCESO FINALIZADO", ImportacionAFIP)
        Call InfoGeneral("", ImportacionAFIP)
        'Call InfoGeneral("Cantidad de registros importados " & i, ImportacionAFIP)

        rstRegistroAFIP.Close
        Set rstRegistroAFIP = Nothing
        dbAFIP.Close
        Set dbAFIP = Nothing
        strLineaImportada = ""
        Close #intNumeroArchivo
    End If

End Sub

Public Sub ControlarEjecucionAcumuladaSIIF(FechaTope As String, ReporteImportado As String)

    Dim datFecha As Date
    Dim strDireccion As String
    Dim strNombreArchivo As String
    Dim intNumeroArchivo As Integer
    Dim SQL As String
    Dim lngRows As Long
    Dim lngRowsAcum As Long

    If Trim(FechaTope) = "" Or IsDate(FechaTope) = False Then
        MsgBox "Debe ingresar una Fecha Tope adecuada", vbCritical + vbOKOnly, "FECHA INCORRECTA"
        ControlCruzadoSIIF.txtFecha.SetFocus
        Exit Sub
    Else
        datFecha = DateTime.DateSerial(Right(FechaTope, 4), Mid(FechaTope, 4, 2), Left(FechaTope, 2))
    End If
    
On Error GoTo HayError
    
    With ControlCruzadoSIIF
        .txtInforme.Text = ""
        Call InfoGeneral("- INICIANDO CONTROL EJECUCIÓN ACUMULADA ICARO VS SIIF (" & ReporteImportado & ") -", ControlCruzadoSIIF)
        Call InfoGeneral("", ControlCruzadoSIIF)
        .dlgMultifuncion.Filter = "Todos los archivos de Texto(*.txt)|*.txt|"
        .dlgMultifuncion.FileName = ""
        .dlgMultifuncion.ShowOpen
        'Cargamos en el nombre del archivo y la ubicación del mismo
        strNombreArchivo = .dlgMultifuncion.FileTitle
        strDireccion = .dlgMultifuncion.FileName
        strDireccion = Replace(.dlgMultifuncion.FileName, "\" & strNombreArchivo, "")
    End With
    If strDireccion = "" Then 'En caso de apretar cancelar
        Call InfoGeneral("CONTROL CANCELADO", ControlCruzadoSIIF)
        Exit Sub
    Else
        'Generamos el archivo Schema.ini que utilizaremos para importar el txt
        intNumeroArchivo = FreeFile
        'Open strDireccion For Input As #intNumeroArchivo
        Open strDireccion & "\schema.ini" For Output As #intNumeroArchivo
        Print #intNumeroArchivo, "[" & strNombreArchivo & "]"
        Select Case ReporteImportado
        Case "rf602txt"
            Print #intNumeroArchivo, "Format=FixedLength"
            Print #intNumeroArchivo, "TextDelimiter=none"
        Case Else
            Print #intNumeroArchivo, "Format=Delimited( )"
            Print #intNumeroArchivo, "NumberLeadingZeros = True"
        End Select
        Print #intNumeroArchivo, "DecimalSymbol = ,"
        Print #intNumeroArchivo, "ColNameHeader = False"
        Print #intNumeroArchivo, "MaxScanRows = 0"
        Print #intNumeroArchivo, "CharacterSet = ANSI"
        Print #intNumeroArchivo, "CurrencyThousandSymbol = ."
'        Print #intNumeroArchivo, "CurrencyDecimalSymbol = ,"
        Select Case ReporteImportado
        Case "rf602txt"
            Print #intNumeroArchivo, "Col1=Vacio1 Text Width 8"
            Print #intNumeroArchivo, "Col2=Programa Text Width 2"
            Print #intNumeroArchivo, "Col3=Vacio2 Text Width 3"
            Print #intNumeroArchivo, "Col4=Subprograma Text Width 2"
            Print #intNumeroArchivo, "Col5=Vacio3 Text Width 3"
            Print #intNumeroArchivo, "Col6=Proyecto Text Width 2"
            Print #intNumeroArchivo, "Col7=Vacio4 Text Width 3"
            Print #intNumeroArchivo, "Col8=Actividad Text Width 2"
            Print #intNumeroArchivo, "Col9=Vacio5 Text Width 3"
            Print #intNumeroArchivo, "Col10=Partida Text Width 3"
            Print #intNumeroArchivo, "Col11=Vacio6 Text Width 3"
            Print #intNumeroArchivo, "Col12=Fuente Text Width 2"
            Print #intNumeroArchivo, "Col13=Vacio7 Text Width 3"
            Print #intNumeroArchivo, "Col14=OrganismoFinanciador Text Width 3"
            Print #intNumeroArchivo, "Col15=CreditoOriginaltxt Text Width 21"
            Print #intNumeroArchivo, "Col16=CreditoVigentetxt Text Width 20"
            Print #intNumeroArchivo, "Col17=Comprometidotxt Text Width 20"
            Print #intNumeroArchivo, "Col18=Ordenado Double Width 20"
            Print #intNumeroArchivo, "Col19=SaldoPresupuestariotxt Text Width 21"
            Print #intNumeroArchivo, "Col20=Pendientetxt Text Width 20"
            Print #intNumeroArchivo, "Col21=Vacio8 Text Width 4"
        Case "rf602pdf"
            Print #intNumeroArchivo, "Col1=Programa Text"
            Print #intNumeroArchivo, "Col2=Subprograma Text"
            Print #intNumeroArchivo, "Col3=Proyecto Text"
            Print #intNumeroArchivo, "Col4=Actividad Text"
            Print #intNumeroArchivo, "Col5=Partida Text"
            Print #intNumeroArchivo, "Col6=CreditoOriginal Double"
            Print #intNumeroArchivo, "Col7=CreditoVigente Double"
            Print #intNumeroArchivo, "Col8=Comprometido Double"
            Print #intNumeroArchivo, "Col9=Ordenado Double"
            Print #intNumeroArchivo, "Col10=SaldoPresupuestario Double"
            Print #intNumeroArchivo, "Col11=Fuente Text"
            Print #intNumeroArchivo, "Col12=OrganismoFinanciador Text"
            Print #intNumeroArchivo, "Col13=Pendiente Double"
        Case "rf636m"
            Print #intNumeroArchivo, "Col1=Partida Text"
            Print #intNumeroArchivo, "Col2=CreditoVigente Double"
            Print #intNumeroArchivo, "Col3=Comprometido Double"
            Print #intNumeroArchivo, "Col4=Ordenado Double"
            Print #intNumeroArchivo, "Col5=SaldoPresupuestario Double"
            Print #intNumeroArchivo, "Col6=CreditoOriginal Double"
            Print #intNumeroArchivo, "Col7=Programa Text"
            Print #intNumeroArchivo, "Col8=Subprograma Text"
            Print #intNumeroArchivo, "Col9=Proyecto Text"
            Print #intNumeroArchivo, "Col10=Actividad Text"
            Print #intNumeroArchivo, "Col11=Fuente Text"
            Print #intNumeroArchivo, "Col12=OrganismoFinanciador Text"
        End Select
        Close #intNumeroArchivo
        
        Call InfoGeneral("1) Archivo Schema.ini creado en: " & strDireccion, ControlCruzadoSIIF)
        Call InfoGeneral("", ControlCruzadoSIIF)
        
        'Creamos una tabla auxiliar (utilizando Schema.ini) para trabajar en ella
        SQL = "SELECT * INTO AUXILIARIMPORTACION " _
        & "FROM [" & strNombreArchivo & "] IN '" & strDireccion & "' 'TEXT;'"
        dbIcaro.BeginTrans
        dbIcaro.Execute SQL, lngRows, adCmdText Or adExecuteNoRecords
        dbIcaro.CommitTrans
        
        lngRowsAcum = lngRows

        Call InfoGeneral("2) Datos importados en tabla AUXILIARIMPORTACION con " & lngRows & " registros", ControlCruzadoSIIF)
        Call InfoGeneral("", ControlCruzadoSIIF)
        
        'Procedemos a depurar la tabla AUXILIARIMPORTACION
        Call InfoGeneral("3) Depurando la tabla AUXILIARIMPORTACION", ControlCruzadoSIIF)

        'Eliminamos los registros que, en el campo Partida, se hayan vacios
        SQL = "Delete From AUXILIARIMPORTACION " _
        & "Where IsNull(AUXILIARIMPORTACION.Partida) = TRUE"
        dbIcaro.BeginTrans
        dbIcaro.Execute SQL, lngRows
        dbIcaro.CommitTrans

        lngRowsAcum = lngRowsAcum - lngRows
        Call InfoGeneral("- " & lngRows & " registros eliminados por estar vacios", ControlCruzadoSIIF)
        
        'Eliminamos los registros que no son Partidas presupuestarias según el listado de Partidas de ICARO
        SQL = "Delete From AUXILIARIMPORTACION " _
        & "Where (AUXILIARIMPORTACION.Partida) Not In " _
        & "(Select Partida From Partidas)"
        dbIcaro.BeginTrans
        dbIcaro.Execute SQL, lngRows
        dbIcaro.CommitTrans

        lngRowsAcum = lngRowsAcum - lngRows
        Call InfoGeneral("- " & lngRows & " registros eliminados por no corresponder con el listado de PARTIDAS de ICARO", ControlCruzadoSIIF)
        
        'Eliminamos los registros que no corresponden a las partida 421 y 422 del Campo Partida
        SQL = "Delete From AUXILIARIMPORTACION " _
        & "Where Trim(AUXILIARIMPORTACION.Partida) <> '421' " _
        & "And Trim(AUXILIARIMPORTACION.Partida) <> '422'"
        dbIcaro.BeginTrans
        dbIcaro.Execute SQL, lngRows
        dbIcaro.CommitTrans

        lngRowsAcum = lngRowsAcum - lngRows
        Call InfoGeneral("- " & lngRows & " registros eliminados por no ser 421 y 422", ControlCruzadoSIIF)
        
        'Agregamos una columna a la tabla AUXILIARIMPORTACION
        SQL = "Alter Table AUXILIARIMPORTACION " _
        & "Add ImputacionSIIF Text(18)"
        dbIcaro.BeginTrans
        dbIcaro.Execute SQL
        dbIcaro.CommitTrans
        
        'Formateamos la reciente columna añadida
        SQL = "Update AUXILIARIMPORTACION " _
        & "Set ImputacionSIIF = Format(AUXILIARIMPORTACION.Programa, '00') & '-' & Format(AUXILIARIMPORTACION.Subprograma, '00') & '-' & " _
        & "Format(AUXILIARIMPORTACION.Proyecto, '00') & '-' & Format(AUXILIARIMPORTACION.Actividad, '00') & '-' & " _
        & "AUXILIARIMPORTACION.Partida & '-' & AUXILIARIMPORTACION.Fuente"
        dbIcaro.BeginTrans
        dbIcaro.Execute SQL
        dbIcaro.CommitTrans

'        If ReporteImportado = "rf602txt" Then
'            'Agregamos una columna a la tabla AUXILIARIMPORTACION
'            SQL = "Alter Table AUXILIARIMPORTACION " _
'            & "Add Ordenado Double"
'            dbIcaro.BeginTrans
'            dbIcaro.Execute SQL
'            dbIcaro.CommitTrans
'            'Formateamos el campo Ordenado
'            SQL = "Update AUXILIARIMPORTACION " _
'            & "Set Ordenado = CStr(Val(Left(Ordenado,Len(Ordenadotxt)-3))) & '.' & Right(Ordenadotxt,2)"
'            dbIcaro.BeginTrans
'            dbIcaro.Execute SQL
'            dbIcaro.CommitTrans
'        End If
        
        Call InfoGeneral("- Columna ImputacionSIIF añadida a la tabla AUXILIARIMPUTACION", ControlCruzadoSIIF)
        Call InfoGeneral("", ControlCruzadoSIIF)
        
        'Creamos la consulta SQL que compara SIIF con ICARO
        Select Case ReporteImportado
        Case "rf602pdf", "rf602txt"
'            SQL = "SELECT ImputacionPresupuestaria, Sum([EjecucionAcumulada]) As [DiferenciaEjecucion] " _
'                & "FROM (SELECT (OBRAS.Imputacion & '-' & OBRAS.Partida & '-' & OBRAS.Fuente) As ImputacionPresupuestaria, Sum(CARGA.Importe) AS EjecucionAcumulada " _
'                & "FROM OBRAS INNER JOIN CARGA ON OBRAS.Descripcion = CARGA.Obra " _
'                & "WHERE Year([CARGA].[Fecha]) = " & Year(datFecha) & " " _
'                & "AND CARGA.Fecha <= #" & Format(datFecha, "MM/DD/YYYY") & "# " _
'                & "AND Val(OBRAS.Partida) Between 421 AND 422 " _
'                & "GROUP BY OBRAS.Imputacion, OBRAS.Partida, OBRAS.Fuente " _
'                & "UNION ALL " _
'                & "SELECT AUXILIARIMPORTACION.ImputacionSIIF, (AUXILIARIMPORTACION.Ordenado * -1) " _
'                & "FROM AUXILIARIMPORTACION) " _
'                & "GROUP BY ImputacionPresupuestaria " _
'                & "HAVING ABS(Sum([EjecucionAcumulada])) > 0.001"
            SQL = "SELECT ImputacionPresupuestaria, Sum([EjecucionAcumulada]) As [DiferenciaEjecucion] " _
                & "FROM (SELECT (OBRAS.Imputacion & '-' & OBRAS.Partida & '-' & CARGA.Fuente) As ImputacionPresupuestaria, Sum(CARGA.Importe) AS EjecucionAcumulada " _
                & "FROM OBRAS INNER JOIN CARGA ON OBRAS.Descripcion = CARGA.Obra " _
                & "WHERE Year([CARGA].[Fecha]) = " & Year(datFecha) & " " _
                & "AND CARGA.Fecha <= #" & Format(datFecha, "MM/DD/YYYY") & "# " _
                & "AND Val(OBRAS.Partida) Between 421 AND 422 " _
                & "GROUP BY OBRAS.Imputacion, OBRAS.Partida, CARGA.Fuente " _
                & "UNION ALL " _
                & "SELECT AUXILIARIMPORTACION.ImputacionSIIF, (AUXILIARIMPORTACION.Ordenado * -1) " _
                & "FROM AUXILIARIMPORTACION) " _
                & "GROUP BY ImputacionPresupuestaria " _
                & "HAVING ABS(Sum([EjecucionAcumulada])) > 0.001"
        Case "rf636m"
            SQL = "SELECT ImputacionPresupuestaria, Sum([EjecucionAcumulada]) As [DiferenciaEjecucion] " _
                & "FROM (SELECT (OBRAS.Imputacion & '-' & OBRAS.Partida & '-' & CARGA.Fuente) As ImputacionPresupuestaria, Sum(CARGA.Importe) AS EjecucionAcumulada " _
                & "FROM OBRAS INNER JOIN CARGA ON OBRAS.Descripcion = CARGA.Obra " _
                & "WHERE Year([CARGA].[Fecha]) = " & Year(datFecha) & " " _
                & "AND Month([CARGA].[Fecha]) = " & Month(datFecha) & " " _
                & "AND CARGA.Fecha <= #" & Format(datFecha, "MM/DD/YYYY") & "# " _
                & "AND Val(OBRAS.Partida) Between 421 AND 422 " _
                & "GROUP BY OBRAS.Imputacion, OBRAS.Partida, CARGA.Fuente " _
                & "UNION ALL " _
                & "SELECT AUXILIARIMPORTACION.ImputacionSIIF, (AUXILIARIMPORTACION.Ordenado * -1) " _
                & "FROM AUXILIARIMPORTACION) " _
                & "GROUP BY ImputacionPresupuestaria " _
                & "HAVING ABS(Sum([EjecucionAcumulada])) > 0.001"
        End Select
            
        Set rstRegistroIcaro = New ADODB.Recordset
        rstRegistroIcaro.Open SQL, dbIcaro, adOpenForwardOnly, adLockOptimistic

        Call InfoGeneral("4) Cruzando información de datos importados con datos cargados en Icaro al: ", ControlCruzadoSIIF)
        Call InfoGeneral("", ControlCruzadoSIIF)
        
        Call InfoGeneral("5) Desvíos detectados ($ Icaro menos $ SIIF):", ControlCruzadoSIIF)
        
        If rstRegistroIcaro.BOF = False Then
            rstRegistroIcaro.MoveFirst
            While rstRegistroIcaro.EOF = False
                Call InfoGeneral(" +" & rstRegistroIcaro!ImputacionPresupuestaria & " = $ " & FormatNumber(rstRegistroIcaro!DiferenciaEjecucion, 2), ControlCruzadoSIIF)
                rstRegistroIcaro.MoveNext
            Wend
        End If
        rstRegistroIcaro.Close
        Set rstRegistroIcaro = Nothing
        Call InfoGeneral("", ControlCruzadoSIIF)

        'Borramos schema.ini
        Kill strDireccion & "\schema.ini"

        'Borramos tabla creada
        SQL = "Drop Table AUXILIARIMPORTACION"
        dbIcaro.BeginTrans
        dbIcaro.Execute SQL
        dbIcaro.CommitTrans
        Call InfoGeneral("6) Tabla AUXILIARIMPORTACION y archivo Schema.ini eliminados", ControlCruzadoSIIF)
        Call InfoGeneral("", ControlCruzadoSIIF)
        
        SQL = "SELECT CARGA.Comprobante, CARGA.Fecha FROM CARGA " _
            & "WHERE Year([CARGA].[Fecha]) = " & Year(datFecha) & " " _
            & "AND CARGA.Fecha <= #" & Format(datFecha, "MM/DD/YYYY") & "# " _
            & "AND CARGA.Comprobante NOT IN (SELECT IMPUTACION.Comprobante " _
            & "FROM IMPUTACION INNER JOIN CARGA " _
            & "ON IMPUTACION.Comprobante = CARGA.Comprobante " _
            & "WHERE Year([CARGA].[Fecha]) = " & Year(datFecha) & " " _
            & "AND CARGA.Fecha <= #" & Format(datFecha, "MM/DD/YYYY") & "#) " _
            & "ORDER BY CARGA.Fecha Asc"
        
        Set rstRegistroIcaro = New ADODB.Recordset
        rstRegistroIcaro.Open SQL, dbIcaro, adOpenForwardOnly, adLockOptimistic

        Call InfoGeneral("7) Cantidad de CYO sin imputación presupuestaria: " & rstRegistroIcaro.RecordCount, ControlCruzadoSIIF)
        
        If rstRegistroIcaro.RecordCount > 0 Then
            rstRegistroIcaro.MoveFirst
            While rstRegistroIcaro.EOF = False
                Call InfoGeneral(" +" & rstRegistroIcaro!Fecha & " -> Nro CYO: " & rstRegistroIcaro!Comprobante, ControlCruzadoSIIF)
                rstRegistroIcaro.MoveNext
            Wend
        End If
        
        rstRegistroIcaro.Close
        Set rstRegistroIcaro = Nothing

        Exit Sub
    End If
    
HayError:
    MsgBox Err.Number & " - " & Err.Description, , "Error!!!"

    
End Sub

Public Function InfoGeneral(StrAgregar As String, FormularioDestino As Form)
    
    FormularioDestino.txtInforme.Text = FormularioDestino.txtInforme.Text & vbCrLf & StrAgregar
    
End Function
