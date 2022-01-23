Attribute VB_Name = "EjemplosVarios"
'Option Explicit

' Importante : Agregar la referencia a Micorosft Excel xx object library
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'Private Sub ImportarExcel()


    'Variable de tipo Aplicación de Excel
    'Dim objExcel As Excel.Application

    'Una variable de tipo Libro de Excel
    'Dim xLibro As Excel.Workbook
    'Dim Col As Integer, Fila As Integer

    'creamos un nuevo objeto excel
    'Set objExcel = New Excel.Application

    'Usamos el método open para abrir el archivo que está _
     en el directorio del programa llamado archivo.xls
    'Set xLibro = objExcel.Workbooks.Open(App.Path + "\archivo.xls")

    'Hacemos el Excel Visible
    'objExcel.Visible = True
    
    


    'With xLibro

        ' Hacemos referencia a la Hoja
        'With .Sheets(1)

            'Recorremos la fila desde la 1 hasta la 7
            'For Fila = 1 To 7

                'Agregamos el valor de la fila que _
                 corresponde a la columna 2
                'Combo1.AddItem .Cells(Fila, 2)
            'Next

        'End With
    'End With

    'Eliminamos los objetos si ya no los usamos
    'Set objExcel = Nothing
    'Set xLibro = Nothing

'End Sub

 'Option Explicit

'***************************************************************************
'*  Name         : Ejemplo para pasar datos de Excel a Access con DAO

'*  Referencias  : Microsoft DAO Object Library
'*  Controles    : Command1

'***************************************************************************

' \\ Función
' --------------------------------------------------------------------------

'Private Function Excel_a_Access( _
    'Path_BD As String, _
    'Path_XLS As String, _
    'La_Tabla As String, _
    'Filas As Integer, _
    'Columnas As Integer) As Boolean

    ' -- Variables para acceder al libro de excel,
    ' -- para la base de datos y el recordset dao
    'Dim Obj_Excel       As Object
    'Dim Obj_Hoja        As Object
    'Dim bd              As DataBase
    'Dim rst             As Recordset
    
    ' -- Variables para la fila, la columna y el dato a copiar
    'Dim Fila_Actual     As Integer
    'Dim Columna_Actual  As Integer
    'Dim Dato            As Variant

    'Screen.MousePointer = vbHourglass

    ' -- Crea una Nueva instancia de Excel
    'Set Obj_Excel = CreateObject("Excel.Application")

    ' -- Abre el libro pasandole el path
    'Obj_Excel.Workbooks.Open FileName:=Path_XLS

    ' -- si es la versión de Excel 97, asigna la hoja activa ( ActiveSheet )
    'If Val(Obj_Excel.Application.Version) >= 8 Then
        'Set Obj_Hoja = Obj_Excel.ActiveSheet
    'Else
        'Set Obj_Hoja = Obj_Excel
    'End If
    
    ' -- Abrir la base de datos
    'Set bd = OpenDatabase(Path_BD)
        
    ' -- Llenar el recordset indicándole la tabla
    'Set rst = bd.OpenRecordset(La_Tabla, dbOpenTable)
    
    
    ' -- Recorrer las filas y columnas de la hoja
    'For Fila_Actual = 1 To Filas
        ' -- Agregar un nuevo registro
        'rst.AddNew
        
        ' -- REcorrer las columnas del libro
        'For Columna_Actual = 0 To Columnas - 1
        
            ' -- Almacena le dato de la celda actual
            'Dato = Trim$(Obj_Hoja.Cells(Fila_Actual, Columna_Actual + 1))
            
            ' -- Agrega los datos al campo indicado
            'rst(Columna_Actual).Value = Dato

        'Next
        ' -- Actualiza los datos en la tabla
        'rst.Update
    'Next
    
    'Excel_a_Access = True
    ' -- Descargar los objetos
    
    'Call Descargar_Objetos(rst, bd, Obj_Excel, Obj_Hoja)
    'Screen.MousePointer = vbDefault

'Exit Function

' -- Error
' --------------------------------------------------------------
'ErrSub:

' -- Descargar las referencias y cierra la base de datos
'Call Descargar_Objetos(rst, bd, Obj_Excel, Obj_Hoja)
'MsgBox Err.Description, vbCritical
'Screen.MousePointer = vbDefault
    
'End Function

'Descarga los objetos y los cierra
'Sub Descargar_Objetos( _
    'rst As Recordset, _
    'bd As DataBase, _
    'Obj_Excel As Object, _
    'Obj_Hoja As Object)

    
    'Set rst = Nothing
    'bd.Close
    'Set bd = Nothing
    'Obj_Excel.ActiveWorkbook.Close False
    'Obj_Excel.Quit
    'Set Obj_Hoja = Nothing
    'Set Obj_Excel = Nothing

'End Sub


' \\ Botón para exportar
' --------------------------------------------------------------------------

'Private Sub Command1_Click()
  
  'Dim ret As Boolean
  
' -- Parámetros:
' -- El nombre y path de la base de datos y del libro excel
' -- El nombre de la tabla
' -- La cantidad de filas y columnas de la hoja a leer
  
  ' -- En este caso se leen el rango desde A1 hasta C10
    'ret = Excel_a_Access("c:\bd1.mdb", "c:\Libro1.xls", "Tabla1", 10, 3)
    
    'If ret Then 'Ok
        'MsgBox " Datos copiados a Access ", vbInformation
    'End If

'End Sub

'Private Sub Form_Load()
    'Command1.Caption = " Excel a Access "
'End Sub


'Option Explicit

'Agregar la referencia de ADO
'----------------------------------------------------

'Private Sub ExcelaAccess(Path_BD As String, _
                           Path_XLS As String, _
                           La_Tabla As String, _
                           Filas As Integer, _
                           Columnas As Integer)


'Dim Obj_Excel As Object
'Dim Obj_Hoja As Object
'Dim cn_Ado As ADODB.Connection
'Dim rst_Ado As ADODB.Recordset
'Dim Fila_Actual As Integer
'Dim Columna_Actual As Integer
'Dim Dato As Variant

    'Screen.MousePointer = vbHourglass

    'Nueva instancia de Excel
    'Set Obj_Excel = CreateObject("Excel.Application")

    ' Abre el libro de Excel
    'Obj_Excel.Workbooks.Open FileName:=Path_XLS

    ' si es la versión de Excel 97, asigna la hoja activa ( ActiveSheet )
    'If Val(Obj_Excel.Application.Version) >= 8 Then
        'Set Obj_Hoja = Obj_Excel.ActiveSheet
    'Else
        'Set Obj_Hoja = Obj_Excel
    'End If
    
    'Abre una nueva conexión Ado
    'Set cn_Ado = New ADODB.Connection
    
    ' Cadena de conexión
    'cn_Ado.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                              "Data Source=" & Path_BD & ";" & _
                              "Persist Security Info=False"
    
    ' Abre la base de datos
    'cn_Ado.Open
    
    'Nuevo objeto recordset
    'Set rst_Ado = New ADODB.Recordset
    
    'Abre el recordset
    'rst_Ado.Open "Select * FROM " & La_Tabla, _
                  cn_Ado, adOpenStatic, adLockPessimistic
    
    'Se posiciona al final    If rst_Ado.RecordCount <> 0 Then rst_Ado.MoveLast
    ' Recorre las filas y columnas de la hoja
    'For Fila_Actual = 1 To Filas
        'Nuevo registro
        'rst_Ado.AddNew
        'For Columna_Actual = 0 To Columnas - 1
        
            ' Va leyendo los datos de la celda indicada
            'Dato = Trim$(Obj_Hoja.Cells(Fila_Actual, Columna_Actual + 1))
            
            'Agrega los datos al campo indicado
            'rst_Ado(Columna_Actual).Value = Dato

        'Next
    'Next
    
    'Call Descargar_Objetos(rst_Ado, cn_Ado, Obj_Excel, Obj_Hoja)
    'Screen.MousePointer = vbDefault
    'MsgBox " Datos copiados ", vbInformation

'Exit Sub

'Error
'ErrSub:

'Call Descargar_Objetos(rst_Ado, cn_Ado, Obj_Excel, Obj_Hoja)
'MsgBox Err.Description, vbCritical
'Screen.MousePointer = vbDefault
    
'End Sub

'Descarga los objetos y los cierra
'Sub DescargarObjetos(rst_Ado As ADODB.Recordset, cn_Ado As ADODB.Connection, _
                                      Obj_Excel As Object, Obj_Hoja As Object)

    
    'Set rst_Ado = Nothing
    'cn_Ado.Close
    'Set cn_Ado = Nothing
    'Obj_Excel.ActiveWorkbook.Close False
    'Obj_Excel.Quit
    'Set Obj_Hoja = Nothing
    'Set Obj_Excel = Nothing

'End Sub



'Private Sub Command1_Click()

' Pasar como parámetro el nombre y path de la _
  base de datos y del libro excel, el nombre de la tabla _
  y la cantidad de filas y columnas de la hoja a leer

'Call Excel_a_Access(App.Path & "\bd1.mdb", _
                    App.Path & "\Libro1.xls", _
                    "Tabla1", 10, 3)
                    
'End Sub

'Option Explicit
'Public Sub Importar_Excel( _
    'Libro As String, _
    'hoja As String, _
    'Optional rango As String = "")

    'Dim conexion As ADODB.Connection, rs As ADODB.Recordset

    'Set conexion = New ADODB.Connection
    
    'conexion.Open '"Provider=Microsoft.Jet.OLEDB.4.0;" & _
                  '"Data Source=" & Libro & _
                  '";Extended Properties=""Excel 8.0;HDR=Yes;"""

      
    ' Nuevo recordset
    'Set rs = New ADODB.Recordset
    
    'With rs
        '.CursorLocation = adUseClient
        '.CursorType = adOpenStatic
        '.LockType = adLockOptimistic
    'End With

    'If rango <> ":" Then
       'hoja = hoja & "$" & rango
    'End If
    
    'rs.Open "SELECT * FROM [" & hoja & "]", conexion, , , adCmdText
    
    ' Mostramos los datos en el datagrid
    'Set DataGrid1.DataSource = rs


'End Sub



'Private Sub Command1_Click()
    'Dim rango As String, hoja As String, ruta As String

    'ruta = Text1 'ruta del archivo excel
    'rango = Text2 & ":" & Text3 'Rango de datos (opcional)
    'hoja = Text4 'Nombre de la hoja

    'Importar_Excel ruta, hoja, rango

'End Sub

'---- ACTUALIZAR TODOS LOS REGISTROS DE UN CAMPO ----

'Lo que pretendes hacer creo que la mejor manera es meidante una consulta de actualización

'Dim TotalMiValor As Double
'Dim dba As DataBase
'Dim rsa As Recorset

'Set dba = currentDb
'Set rsa = dba.OpenRecordset("Select SUM(Valor) AS TotalValor FROM Tabla1")

'TotalMiValor = rsa!TotalValor

'rsa.Close

'dba.Execute "UPDATE TABLA1 SET TOTAL=VALOR/" & TotalMiValor

'dba.Close
