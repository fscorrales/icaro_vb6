Attribute VB_Name = "Exportacion"
Public Function Exportar_Excel(strOutputPath As String, FlexGrid As Object) As Boolean
    On Error GoTo Error_Handler
  
    Dim o_Excel     As Object
    Dim o_Libro     As Object
    Dim o_Hoja      As Object
    Dim Fila        As Long
    Dim Columna     As Long
       
    ' -- Crea el objeto Excel, el objeto workBook y el objeto sheet
    Set o_Excel = CreateObject("Excel.Application")
    Set o_Libro = o_Excel.Workbooks.Add
    Set o_Hoja = o_Libro.Worksheets.Add
       
    ' -- Bucle para Exportar los datos
    With FlexGrid
        For Fila = 0 To .Rows - 1
            For Columna = 0 To .Cols - 1
                o_Hoja.Cells(Fila + 1, Columna + 1).Value = .TextMatrix(Fila, Columna)
            Next
        Next
    End With
    o_Libro.Close True, strOutputPath
    ' -- Cerrar Excel
    o_Excel.Quit
    ' -- Terminar instancias
    Call ReleaseObjects(o_Excel, o_Libro, o_Hoja)
    Exportar_Excel = True
Exit Function
  
' -- Controlador de Errores
Error_Handler:
    ' -- Cierra la hoja y el la aplicación Excel
    If Not o_Libro Is Nothing Then: o_Libro.Close False
    If Not o_Excel Is Nothing Then: o_Excel.Quit
    Call ReleaseObjects(o_Excel, o_Libro, o_Hoja)
    If Err.Number <> 1004 Then MsgBox Err.Description, vbCritical
End Function
' -------------------------------------------------------------------
' \\ -- Eliminar objetos para liberar recursos
' -------------------------------------------------------------------
Private Sub ReleaseObjects(o_Excel As Object, o_Libro As Object, o_Hoja As Object)
    If Not o_Excel Is Nothing Then Set o_Excel = Nothing
    If Not o_Libro Is Nothing Then Set o_Libro = Nothing
    If Not o_Hoja Is Nothing Then Set o_Hoja = Nothing
End Sub

' \\ -- función para psar los datos hacia una hoja de un libro exisitente
' -------------------------------------------------------------------------------------------
Public Function Exportar_Excel2(sBookFileName As String, FlexGrid As Object, Optional sNameSheet As String = vbNullString) As Boolean
  
    On Error GoTo Error_Handler
  
    Dim o_Excel     As Object
    Dim o_Libro     As Object
    Dim o_Hoja      As Object
    Dim Fila        As Long
    Dim Columna     As Long
  
    ' -- Error en la ruta del libro
    If sBookFileName = vbNullString Or Len(Dir(sBookFileName)) = 0 Then
           
        MsgBox " Falta el Path del archivo de Excel o no se ha encontrado el libro en la ruta especificada ", vbCritical
        Exit Function
    End If
       
    ' -- Crea el objeto Excel, el objeto workBook y el objeto sheet
    Set o_Excel = CreateObject("Excel.Application")
    Set o_Libro = o_Excel.Workbooks.Open(sBookFileName)
    ' -- Comprobar si se abre la hoja por defecto, o la indicada en el parámetro de la función
    If Len(sNameSheet) = 0 Then
        Set o_Hoja = o_Libro.Worksheets(1)
    Else
        Set o_Hoja = o_Libro.Worksheets(sNameSheet)
    End If
    ' -- Bucle para Exportar los datos
    With FlexGrid
        For Fila = 1 To .Rows - 1
            For Columna = 0 To .Cols - 1
                o_Hoja.Cells(Fila, Columna + 1).Value = .TextMatrix(Fila, Columna)
            Next
        Next
    End With
    ' -- Cerrar libro y guardar los datos
    o_Libro.Close True
    ' -- Cerrar Excel
    o_Excel.Quit
    ' -- Terminar instancias
    Call ReleaseObjects(o_Excel, o_Libro, o_Hoja)
    Exportar_Excel2 = True
Exit Function
  
' -- Controlador de Errores
Error_Handler:
    ' -- Cierra la hoja y el la aplicación Excel
    If Not o_Libro Is Nothing Then: o_Libro.Close False
    If Not o_Excel Is Nothing Then: o_Excel.Quit
    Call ReleaseObjects(o_Excel, o_Libro, o_Hoja)
    MsgBox Err.Description, vbCritical
End Function


Public Function ExportarEPAMConsolidado(strOutputPath As String, FlexGridGeneral As Object, FlexGridDiscriminado As Object) As Boolean
    On Error GoTo Error_Handler
  
    Dim o_Excel             As Object
    Dim o_Libro             As Object
    Dim o_Hoja              As Object
    Dim FilaGeneral         As Long
    Dim FilaDiscriminada    As Long
    Dim ColumnaDiscriminada As Long
    Dim x                   As Long

       
    ' -- Crea el objeto Excel, el objeto workBook y el objeto sheet
    Set o_Excel = CreateObject("Excel.Application")
    Set o_Libro = o_Excel.Workbooks.Add
    Set o_Hoja = o_Libro.Worksheets.Add
       
    ' -- Bucle para Exportar los datos
    o_Hoja.Cells(1, 1).Value = "PRESUPUESTO DE OBRAS EPAM AL " & Date
    o_Hoja.Range("A1:H1").MergeCells = True
    o_Hoja.Range("A1:H1").Font.Size = 14
    o_Hoja.Range("A1:H1").Font.Bold = True
    o_Hoja.Range("A1:H1").HorizontalAlignment = xlVAlignCenter
    o_Hoja.Cells(3, 1).Value = "Descripción"
    o_Hoja.Cells(3, 3).Value = "Imputación Presupuestaria"
    o_Hoja.Cells(3, 4).Value = "Presupuesto"
    o_Hoja.Cells(3, 5).Value = "Pagado"
    o_Hoja.Cells(3, 6).Value = "Imputado"
    o_Hoja.Cells(3, 7).Value = "Decisión a Tomar"
    o_Hoja.Cells(3, 8).Value = "Límite"
    o_Hoja.Range("C3:H3").HorizontalAlignment = xlVAlignCenter
    o_Hoja.Range("A3:H3").Font.Bold = True
    o_Hoja.Range("A3:H3").Interior.ColorIndex = 48
    o_Hoja.Columns(1).ColumnWidth = 35
    o_Hoja.Columns(2).ColumnWidth = 12
    o_Hoja.Columns(3).ColumnWidth = 25
    o_Hoja.Columns(4).ColumnWidth = 12
    o_Hoja.Columns(5).ColumnWidth = 11
    o_Hoja.Columns(6).ColumnWidth = 11
    o_Hoja.Columns(7).ColumnWidth = 22
    o_Hoja.Columns(8).ColumnWidth = 11
    
    With o_Hoja.Range("A3:H3").Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With o_Hoja.Range("A3:H3").Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With o_Hoja.Range("A3:H3").Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With o_Hoja.Range("A3:H3").Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    
    With FlexGridGeneral
        x = 3
        For FilaGeneral = 1 To .Rows - 1
            ConfigurardgResolucionesEPAM
            CargardgResolucionesEPAM (FlexGridGeneral.TextMatrix(FilaGeneral, 0))
            x = x + 1
            o_Hoja.Cells(x, 1).Value = .TextMatrix(FilaGeneral, 0)
            o_Hoja.Range("A" & x & ":B" & x).MergeCells = True
            o_Hoja.Range("A" & x & ":B" & x).HorizontalAlignment = xlVAlignCenter
            o_Hoja.Cells(x, 3).Value = .TextMatrix(FilaGeneral, 1)
            o_Hoja.Cells(x, 5).Value = .TextMatrix(FilaGeneral, 2)
            o_Hoja.Cells(x, 6).Value = .TextMatrix(FilaGeneral, 3)
            o_Hoja.Cells(x, 7).Value = .TextMatrix(FilaGeneral, 4)
            With FlexGridDiscriminado
                For FilaDiscriminada = 1 To .Rows - 1
                    x = x + 1
                    For ColumnaDiscriminada = 0 To .Cols - 1
                        o_Hoja.Cells(x, ColumnaDiscriminada + 1).Value = .TextMatrix(FilaDiscriminada, ColumnaDiscriminada)
                    Next
                Next
            End With
        Next
    End With

'Configurando Bordes de la Grilla
    o_Hoja.Range("A4:H" & x).Select
    o_Excel.Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    o_Excel.Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With o_Excel.Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With o_Excel.Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With o_Excel.Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With o_Excel.Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With o_Excel.Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With o_Excel.Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
'Configurando Para Impresión
    o_Hoja.Rows("1:3").Select
    With o_Libro.ActiveSheet.PageSetup
        .PrintTitleRows = "$1:$3"
        .PrintTitleColumns = ""
    End With
    o_Libro.ActiveSheet.PageSetup.PrintArea = ""
    With o_Libro.ActiveSheet.PageSetup
        .LeftHeader = ""
        .CenterHeader = ""
        .RightHeader = ""
        .LeftFooter = ""
        .CenterFooter = ""
        .RightFooter = ""
        .LeftMargin = Application.InchesToPoints(0.393700787401575)
        .RightMargin = Application.InchesToPoints(0.393700787401575)
        .TopMargin = Application.InchesToPoints(0.984251968503937)
        .BottomMargin = Application.InchesToPoints(0.984251968503937)
        .HeaderMargin = Application.InchesToPoints(0)
        .FooterMargin = Application.InchesToPoints(0)
        .PrintHeadings = False
        .PrintGridlines = False
        .PrintComments = xlPrintNoComments
        .PrintQuality = 600
        .CenterHorizontally = True
        .CenterVertically = False
        .Orientation = xlLandscape
        .Draft = False
        .PaperSize = xlPaperLetter
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = 90
        .PrintErrors = xlPrintErrorsDisplayed
    End With
'Inmovilizar Paneles
    o_Hoja.Rows("4:4").Select
    o_Excel.ActiveWindow.FreezePanes = True

    o_Libro.Close True, strOutputPath
    ' -- Cerrar Excel
    o_Excel.Quit
    ' -- Terminar instancias
    Call ReleaseObjects(o_Excel, o_Libro, o_Hoja)
    ExportarEPAMConsolidado = True
Exit Function
  
' -- Controlador de Errores
Error_Handler:
    ' -- Cierra la hoja y el la aplicación Excel
    If Not o_Libro Is Nothing Then: o_Libro.Close False
    If Not o_Excel Is Nothing Then: o_Excel.Quit
    Call ReleaseObjects(o_Excel, o_Libro, o_Hoja)
    If Err.Number <> 1004 Then MsgBox Err.Description, vbCritical
End Function
