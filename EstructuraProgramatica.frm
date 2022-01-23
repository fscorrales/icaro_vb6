VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form EstructuraProgramatica 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Estructura Programática"
   ClientHeight    =   7320
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9015
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7320
   ScaleWidth      =   9015
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frmListado 
      Caption         =   "Listado de Programas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   4935
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   8775
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid dgEstructura 
         Height          =   4455
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   8500
         _ExtentX        =   15002
         _ExtentY        =   7858
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Acciones Posibles"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   5160
      Width           =   8775
      Begin VB.CommandButton cmdVolver 
         BackColor       =   &H008080FF&
         Caption         =   "Volver (F12)"
         Height          =   375
         Left            =   4560
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   960
         Width           =   4095
      End
      Begin VB.CommandButton cmdIngresar 
         BackColor       =   &H008080FF&
         Caption         =   "Ver Subprogramas"
         Height          =   375
         Left            =   4560
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   360
         Width           =   4095
      End
      Begin VB.CommandButton cmdEliminar 
         BackColor       =   &H008080FF&
         Caption         =   "Eliminar Programa"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1560
         Width           =   4095
      End
      Begin VB.CommandButton cmdEditar 
         BackColor       =   &H008080FF&
         Caption         =   "Editar Programa"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   960
         Width           =   4095
      End
      Begin VB.CommandButton cmdAgregar 
         BackColor       =   &H008080FF&
         Caption         =   "Agregar Programa"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   360
         Width           =   4095
      End
   End
End
Attribute VB_Name = "EstructuraProgramatica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Imputacion As String

Private Sub dgEstructura_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = vbKeyF4
            IngresarEstructura
        Case Is = vbKeyF12
            VolverEstructura
        Case Is = vbKeyF6
            AgregarEstructura
        Case Is = vbKeyF9
            'EditarEstructura
        Case Is = vbKeyF8
            EliminarEstructura
    End Select
End Sub

Private Sub cmdAgregar_Click()

    AgregarEstructura
    
End Sub

Private Sub cmdEditar_Click()

    EditarEstructura
    
End Sub

Private Sub cmdEliminar_Click()
    
    EliminarEstructura
    
End Sub

Private Sub cmdIngresar_Click()
    
    IngresarEstructura
    
End Sub

Private Sub cmdVolver_Click()
    
    VolverEstructura
    
End Sub

Private Sub Form_Load()
    
    Call CenterMe(EstructuraProgramatica, 9100, 7700)

End Sub


Private Sub Form_Unload(Cancel As Integer)
    
    Imputacion = ""
    
End Sub

Private Sub IngresarEstructura()
    
    Dim i As String

    i = dgEstructura.Row
    
    'Para qué sirve?
    If Len(dgEstructura.TextMatrix(i, 0)) = 0 Then
        i = ""
        Exit Sub
    End If

    Imputacion = Me.dgEstructura.TextMatrix(i, 0)
    Select Case Len(Imputacion)
    Case Is = 2
        ConfigurardgEstructura
        Call CargardgEstructura(Imputacion)
    Case Is = 5
        ConfigurardgEstructura
        Call CargardgEstructura(Imputacion)
    Case Is = 8
        ConfigurardgEstructura
        Call CargardgEstructura(Imputacion)
    Case Is = 11
        ConfigurardgEstructura
        Call CargardgEstructura(Imputacion)
    End Select

    i = ""
    
End Sub

Private Sub VolverEstructura()

    Dim i As String
    
    i = dgEstructura.Row
    
    If Len(dgEstructura.TextMatrix(i, 0)) <> 0 Then
        Imputacion = Me.dgEstructura.TextMatrix(i, 0)
        Select Case Len(Imputacion)
        Case Is = 15
            Imputacion = Left(Imputacion, Len(Imputacion) - 4)
        Case Else
            Imputacion = Left(Imputacion, Len(Imputacion) - 3)
        End Select
    Else
        If Len(Imputacion) = 0 Then
            i = 0
            Exit Sub
        End If
    End If
    
    Select Case Len(Imputacion)
    Case Is = 2
        ConfigurardgEstructura
        Call CargardgEstructura
    Case Is = 5
        ConfigurardgEstructura
        Call CargardgEstructura(Left(Imputacion, 2))
    Case Is = 8
        ConfigurardgEstructura
        Call CargardgEstructura(Left(Imputacion, 5))
    Case Is = 11
        ConfigurardgEstructura
        Call CargardgEstructura(Left(Imputacion, 8))
    End Select

    i = ""
    
End Sub

Private Sub AgregarEstructura()

    Dim SQL As String

    If Len(dgEstructura.TextMatrix(1, 0)) <> 0 Then
        Imputacion = Me.dgEstructura.TextMatrix(1, 0)
    Else
        Select Case Len(Imputacion)
        Case Is = 0
            Imputacion = "00"
        Case Is = 15
            Imputacion = Imputacion & "-000"
        Case Else
            Imputacion = Imputacion & "-00"
        End Select
    End If
    
    Select Case Len(Imputacion)
    Case Is = 2
        Call BloquearTextBoxCargaEstructuraProgramatica(Imputacion)
        Unload EstructuraProgramatica
    Case Is = 5
        Call BloquearTextBoxCargaEstructuraProgramatica(Imputacion)
        Unload EstructuraProgramatica
    Case Is = 8
        Call BloquearTextBoxCargaEstructuraProgramatica(Imputacion)
        Unload EstructuraProgramatica
    Case Is = 11
        Call BloquearTextBoxCargaEstructuraProgramatica(Imputacion)
        Unload EstructuraProgramatica
    Case Is = 15
        With CargaObra
            Set rstRegistroIcaro = New ADODB.Recordset
            .Show
            .txtPrograma.Text = Left(Imputacion, 2)
            SQL = "Select * from PROGRAMAS Where PROGRAMA = " & "'" & Left(Imputacion, 2) & "'"
            rstRegistroIcaro.Open SQL, dbIcaro, adOpenForwardOnly, adLockOptimistic
            If Len(rstRegistroIcaro!Descripcion) <> "0" Then
                .txtDescripcionPrograma.Text = rstRegistroIcaro!Descripcion
            Else
                .txtDescripcionPrograma.Text = ""
            End If
            rstRegistroIcaro.Close
            .txtSubprograma.Text = Mid(Imputacion, 4, 2)
            SQL = "Select * from SUBPROGRAMAS Where SUBPROGRAMA = " & "'" & Left(Imputacion, 5) & "'"
            rstRegistroIcaro.Open SQL, dbIcaro, adOpenForwardOnly, adLockOptimistic
            If Len(rstRegistroIcaro!Descripcion) <> "0" Then
                .txtDescripcionSubprograma.Text = rstRegistroIcaro!Descripcion
            Else
                .txtDescripcionSubprograma.Text = ""
            End If
            rstRegistroIcaro.Close
            .txtProyecto.Text = Mid(Imputacion, 7, 2)
            SQL = "Select * from PROYECTOS Where PROYECTO = " & "'" & Left(Imputacion, 8) & "'"
            rstRegistroIcaro.Open SQL, dbIcaro, adOpenForwardOnly, adLockOptimistic
            If Len(rstRegistroIcaro!Descripcion) <> "0" Then
                .txtDescripcionProyecto.Text = rstRegistroIcaro!Descripcion
            Else
                .txtDescripcionProyecto.Text = ""
            End If
            rstRegistroIcaro.Close
            .txtProyecto.Text = Mid(Imputacion, 10, 2)
            SQL = "Select * from ACTIVIDADES Where ACTIVIDAD = " & "'" & Left(Imputacion, 11) & "'"
            rstRegistroIcaro.Open SQL, dbIcaro, adOpenForwardOnly, adLockOptimistic
            If Len(rstRegistroIcaro!Descripcion) <> "0" Then
                .txtDescripcionActividad.Text = rstRegistroIcaro!Descripcion
            Else
                .txtDescripcionActividad.Text = ""
            End If
            rstRegistroIcaro.Close
            .txtPrograma.Enabled = False
            .txtDescripcionPrograma.Enabled = False
            .txtSubprograma.Enabled = False
            .txtDescripcionSubprograma.Enabled = False
            .txtProyecto.Enabled = False
            .txtDescripcionProyecto.Enabled = False
            .txtActividad.Enabled = False
            .txtDescripcionActividad.Enabled = False
            .txtPartida.Enabled = True
            .txtDescripcionPartida.Enabled = False
            .txtCUIT.SetFocus
            Unload EstructuraProgramatica
        End With
    End Select

End Sub

Private Sub EditarEstructura()

    Dim i As String
    i = Me.dgEstructura.Row
    If i <> 0 Then
        Select Case Len(Me.dgEstructura.TextMatrix(i, 0))
        'Case Is = 0
        '    i = ""
        '    Exit Sub
        'Case Is = 2
        '    .Show
        '    .txtPrograma.Text = dgEstructura.TextMatrix(i, 0)
        '    .txtDescripcionPrograma.Text = dgEstructura.TextMatrix(i, 1)
        '    strEditandoEstructura = dgEstructura.TextMatrix(i, 0)
        '    strCargaEstructura = "Programa"
        '    .txtSubprograma.Enabled = False
        '    .txtDescripcionSubprograma.Enabled = False
        '    .txtProyecto.Enabled = False
        '    .txtDescripcionProyecto.Enabled = False
        '    .txtActividad.Enabled = False
        '    .txtDescripcionActividad.Enabled = False
        'Case Is = 5
        '    .Show
        '    strCargaEstructura = "Subprograma"
        '    strEditandoEstructura = dgEstructura.TextMatrix(i, 0)
        '    Where = "Where SUBPROGRAMAS.Subprograma = " & "'" & dgEstructura.TextMatrix(i, 0) & "'"
        '    Call SetRecordset(rstEstructuraProgramatica, "SELECT PROGRAMAS.Programa, PROGRAMAS.Descripcion As DescripcionPrograma, SUBPROGRAMAS.Subprograma, SUBPROGRAMAS.Descripcion As DescripcionSubprograma FROM  PROGRAMAS INNER JOIN SUBPROGRAMAS ON PROGRAMAS.Programa = SUBPROGRAMAS.Programa " & Where)
        '    .txtPrograma.Text = rstEstructuraProgramatica.Fields("Programa")
        '    If Len(rstEstructuraProgramatica.Fields("DescripcionPrograma")) <> "0" Then
        '        .txtDescripcionPrograma.Text = rstEstructuraProgramatica.Fields("DescripcionPrograma")
        '    Else
        '        .txtDescripcionPrograma.Text = ""
        '    End If
        '    .txtSubprograma.Text = Right(rstEstructuraProgramatica.Fields("Subprograma"), 2)
        '    If Len(rstEstructuraProgramatica.Fields("DescripcionSubprograma")) <> "0" Then
        '        .txtDescripcionSubprograma.Text = rstEstructuraProgramatica.Fields("DescripcionSubprograma")
        '    Else
        '        .txtDescripcionSubprograma.Text = ""
        '    End If
        '    .txtPrograma.Enabled = False
        '    .txtDescripcionPrograma.Enabled = False
        '    .txtSubprograma.Enabled = True
        '    .txtDescripcionSubprograma.Enabled = True
        '    .txtProyecto.Enabled = False
        '    .txtDescripcionProyecto.Enabled = False
        '    .txtActividad.Enabled = False
        '    .txtDescripcionActividad.Enabled = False
        'Case Is = 8
        '    .Show
        '    strCargaEstructura = "Proyecto"
        '    strEditandoEstructura = dgEstructura.TextMatrix(i, 0)
        '    Where = "Where PROYECTOS.Proyecto = " & "'" & dgEstructura.TextMatrix(i, 0) & "'"
        '    Call SetRecordset(rstEstructuraProgramatica, "SELECT PROGRAMAS.Programa, PROGRAMAS.Descripcion As DescripcionPrograma, SUBPROGRAMAS.Subprograma, SUBPROGRAMAS.Descripcion As DescripcionSubprograma, PROYECTOS.Proyecto, PROYECTOS.Descripcion As DescripcionProyecto, ACTIVIDADES.Actividad, ACTIVIDADES.Descripcion As DescripcionActividad FROM   ((PROGRAMAS INNER JOIN SUBPROGRAMAS ON PROGRAMAS.Programa = SUBPROGRAMAS.Programa) INNER JOIN PROYECTOS ON SUBPROGRAMAS.Subprograma = PROYECTOS.Subprograma) INNER JOIN ACTIVIDADES ON PROYECTOS.Proyecto = ACTIVIDADES.Proyecto " & Where)
        '    .txtPrograma.Text = rstEstructuraProgramatica.Fields("Programa")
        '    If Len(rstEstructuraProgramatica.Fields("DescripcionPrograma")) <> "0" Then
        '        .txtDescripcionPrograma.Text = rstEstructuraProgramatica.Fields("DescripcionPrograma")
        '    Else
        '        .txtDescripcionPrograma.Text = ""
        '    End If
        '    .txtSubprograma.Text = Right(rstEstructuraProgramatica.Fields("Subprograma"), 2)
        '    If Len(rstEstructuraProgramatica.Fields("DescripcionSubprograma")) <> "0" Then
        '        .txtDescripcionSubprograma.Text = rstEstructuraProgramatica.Fields("DescripcionSubprograma")
        '    Else
        '        .txtDescripcionSubprograma.Text = ""
        '    End If
        '    .txtProyecto.Text = Right(rstEstructuraProgramatica.Fields("Proyecto"), 2)
        '    If Len(rstEstructuraProgramatica.Fields("DescripcionProyecto")) <> "0" Then
        '        .txtDescripcionProyecto.Text = rstEstructuraProgramatica.Fields("DescripcionProyecto")
        '    Else
        '        .txtDescripcionProyecto.Text = ""
        '    End If
        '    .txtPrograma.Enabled = False
        '    .txtDescripcionPrograma.Enabled = False
        '    .txtSubprograma.Enabled = False
        '    .txtDescripcionSubprograma.Enabled = False
        '    .txtProyecto.Enabled = True
        '    .txtDescripcionProyecto.Enabled = True
        '    .txtActividad.Enabled = False
        '    .txtDescripcionActividad.Enabled = False
        'Case Is = 11
        '    .Show
        '    strCargaEstructura = "Actividad"
        '    strEditandoEstructura = dgEstructura.TextMatrix(i, 0)
        '    Where = "Where ACTIVIDADES.Actividad = " & "'" & dgEstructura.TextMatrix(i, 0) & "'"
        '    Call SetRecordset(rstEstructuraProgramatica, "SELECT PROGRAMAS.Programa, PROGRAMAS.Descripcion As DescripcionPrograma, SUBPROGRAMAS.Subprograma, SUBPROGRAMAS.Descripcion As DescripcionSubprograma, PROYECTOS.Proyecto, PROYECTOS.Descripcion As DescripcionProyecto, ACTIVIDADES.Actividad, ACTIVIDADES.Descripcion As DescripcionActividad FROM ((PROGRAMAS INNER JOIN SUBPROGRAMAS ON PROGRAMAS.Programa = SUBPROGRAMAS.Programa) INNER JOIN PROYECTOS ON SUBPROGRAMAS.Subprograma = PROYECTOS.Subprograma) INNER JOIN ACTIVIDADES ON PROYECTOS.Proyecto = ACTIVIDADES.Proyecto " & Where)
        '    .txtPrograma.Text = rstEstructuraProgramatica.Fields("Programa")
        '    If Len(rstEstructuraProgramatica.Fields("DescripcionPrograma")) <> "0" Then
        '        .txtDescripcionPrograma.Text = rstEstructuraProgramatica.Fields("DescripcionPrograma")
        '    Else
        '        .txtDescripcionPrograma.Text = ""
        '    End If
        '    .txtSubprograma.Text = Right(rstEstructuraProgramatica.Fields("Subprograma"), 2)
        '    If Len(rstEstructuraProgramatica.Fields("DescripcionSubprograma")) <> "0" Then
        '        .txtDescripcionSubprograma.Text = rstEstructuraProgramatica.Fields("DescripcionSubprograma")
        '    Else
        '        .txtDescripcionSubprograma.Text = ""
        '    End If
        '    .txtProyecto.Text = Right(rstEstructuraProgramatica.Fields("Proyecto"), 2)
        '    If Len(rstEstructuraProgramatica.Fields("DescripcionProyecto")) <> "0" Then
        '        .txtDescripcionProyecto.Text = rstEstructuraProgramatica.Fields("DescripcionProyecto")
        '    Else
        '        .txtDescripcionProyecto.Text = ""
        '    End If
        '    .txtActividad.Text = Right(rstEstructuraProgramatica.Fields("Actividad"), 2)
        '    If Len(rstEstructuraProgramatica.Fields("DescripcionActividad")) <> "0" Then
        '        .txtDescripcionActividad.Text = rstEstructuraProgramatica.Fields("DescripcionActividad")
        '    Else
        '        .txtDescripcionActividad.Text = ""
        '    End If
        '    .txtPrograma.Enabled = False
        '    .txtDescripcionPrograma.Enabled = False
        '    .txtSubprograma.Enabled = False
        '    .txtDescripcionSubprograma.Enabled = False
        '    .txtProyecto.Enabled = False
        '    .txtDescripcionProyecto.Enabled = False
        '    .txtActividad.Enabled = True
        '    .txtDescripcionActividad.Enabled = True
        Case Is = 15
            Call EditarObra(Me.dgEstructura.TextMatrix(i, 1))
            Unload EstructuraProgramatica
        End Select
    End If
    
    i = ""

End Sub

Private Sub EliminarEstructura()

    EliminarEstructuraProgramatica
    
End Sub
