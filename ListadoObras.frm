VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form ListadoObras 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Listado de Obras"
   ClientHeight    =   7290
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9015
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7290
   ScaleWidth      =   9015
   ShowInTaskbar   =   0   'False
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
      TabIndex        =   2
      Top             =   5160
      Width           =   8775
      Begin VB.CommandButton cmdComprobantesImputados 
         BackColor       =   &H008080FF&
         Caption         =   "Ver Comprobantes Imputados"
         Height          =   375
         Left            =   4560
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   360
         Width           =   4095
      End
      Begin VB.CommandButton cmdAgregar 
         BackColor       =   &H008080FF&
         Caption         =   "Agregar Obra (F6)"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   360
         Width           =   4095
      End
      Begin VB.CommandButton cmdEditar 
         BackColor       =   &H008080FF&
         Caption         =   "Editar Obra (F9)"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   960
         Width           =   4095
      End
      Begin VB.CommandButton cmdEliminar 
         BackColor       =   &H008080FF&
         Caption         =   "Eliminar Obra (F8)"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1560
         Width           =   4095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Listado de Obras"
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
      TabIndex        =   0
      Top             =   120
      Width           =   8775
      Begin VB.TextBox txtDescripcionObra 
         Height          =   285
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   6500
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid dgObras 
         Height          =   4095
         Left            =   120
         TabIndex        =   1
         Top             =   720
         Width           =   8505
         _ExtentX        =   15002
         _ExtentY        =   7223
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
End
Attribute VB_Name = "ListadoObras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAgregar_Click()

    Unload ListadoObras
    CargaObra.Show
    
End Sub

Private Sub cmdComprobantesImputados_Click()

    Dim x As Integer
    
    x = Me.dgObras.Row
    
    With ListadoImputacion
        .Show
        ConfigurardgListadoGasto
        Call CargardgListadoGasto(, , , Me.dgObras.TextMatrix(x, 0))
        x = .dgListado.Row
        ConfigurardgImputacion
        CargardgImputacion (.dgListado.TextMatrix(x, 0))
        ConfigurardgRetencion
        CargardgRetencion (.dgListado.TextMatrix(x, 0))
        x = 0
    End With

End Sub

Private Sub cmdEditar_Click()
    
    Dim i As String
    
    i = Me.dgObras.Row
    If i <> 0 Then
        Call EditarObra(Me.dgObras.TextMatrix(i, 0))
        Unload ListadoObras
    End If
    
End Sub

Private Sub cmdEliminar_Click()

    EliminarObra
    
End Sub

Private Sub Form_Load()

    Call CenterMe(ListadoObras, 9100, 7700)

End Sub

Private Sub dgObras_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
        Case Is = vbKeyF6
            Unload ListadoObras
            CargaObra.Show
        Case Is = vbKeyF9
            Dim i As String
            i = Me.dgObras.Row
            If i <> 0 Then
                Call EditarObra(Me.dgObras.TextMatrix(i, 0))
                Unload ListadoObras
            End If
            i = ""
        Case Is = vbKeyF8
            EliminarObra
    End Select
    
End Sub

Private Sub txtDescripcionObra_KeyDown(KeyCode As Integer, Shift As Integer)

    Dim strCadena As String

    If KeyCode = 13 Then
        strCadena = Trim(Me.txtDescripcionObra.Text)
        If Len(strCadena) <= 100 Then
            ConfigurardgObras
            Call CargardgObras(strCadena)
        End If
    End If

End Sub
