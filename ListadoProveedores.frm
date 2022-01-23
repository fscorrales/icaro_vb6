VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form ListadoProveedores 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Proveedores"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7005
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   7005
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Listado de Proveedores"
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
      TabIndex        =   2
      Top             =   120
      Width           =   6735
      Begin VB.TextBox txtCUIT 
         Height          =   285
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   1200
      End
      Begin VB.TextBox txtNombre 
         Height          =   285
         Left            =   1320
         TabIndex        =   6
         Top             =   360
         Width           =   5300
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid dgProveedores 
         Height          =   4095
         Left            =   120
         TabIndex        =   1
         Top             =   720
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   7223
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
      Width           =   6735
      Begin VB.CommandButton cmdComprobantesImputados 
         BackColor       =   &H008080FF&
         Caption         =   "Ver Comprobantes Imputados"
         Height          =   375
         Left            =   3480
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   960
         Width           =   3095
      End
      Begin VB.CommandButton cmdObrasEjecutadas 
         BackColor       =   &H008080FF&
         Caption         =   "Ver Obras Ejecutadas"
         Height          =   375
         Left            =   3480
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   360
         Width           =   3095
      End
      Begin VB.CommandButton cmdEliminar 
         BackColor       =   &H008080FF&
         Caption         =   "Eliminar Proveedor (F8)"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1560
         Width           =   3095
      End
      Begin VB.CommandButton cmdEditar 
         BackColor       =   &H008080FF&
         Caption         =   "Editar Proveedor (F9)"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   960
         Width           =   3095
      End
      Begin VB.CommandButton cmdAgregar 
         BackColor       =   &H008080FF&
         Caption         =   "Agregar Proveedor (F6) "
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   360
         Width           =   3095
      End
   End
End
Attribute VB_Name = "ListadoProveedores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAgregar_Click()
    
    Unload ListadoProveedores
    CargaCUIT.Show
    
End Sub

Private Sub cmdComprobantesImputados_Click()

    Dim x As Integer
    
    x = Me.dgProveedores.Row
    
    With ListadoImputacion
        .Show
        ConfigurardgListadoGasto
        Call CargardgListadoGasto(, , , , Me.dgProveedores.TextMatrix(x, 0))
        x = .dgListado.Row
        ConfigurardgImputacion
        CargardgImputacion (.dgListado.TextMatrix(x, 0))
        ConfigurardgRetencion
        CargardgRetencion (.dgListado.TextMatrix(x, 0))
        x = 0
    End With

End Sub

Private Sub cmdEditar_Click()
    
    EditarProveedor

End Sub

Private Sub cmdEliminar_Click()
    
    EliminarProveedor

End Sub

Private Sub cmdObrasEjecutadas_Click()

    Dim x As Integer
    
    x = Me.dgProveedores.Row
    
    With ListadoObras
        .Show
        ConfigurardgObras
        Call CargardgObras(, Me.dgProveedores.TextMatrix(x, 0))
        x = 0
    End With

End Sub

Private Sub dgProveedores_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
        Case Is = vbKeyF6
            Unload ListadoProveedores
            CargaCUIT.Show
        Case Is = vbKeyF9
            EditarProveedor
        Case Is = vbKeyF8
            EliminarProveedor
    End Select
    
End Sub

Private Sub Form_Load()
    
    Call CenterMe(ListadoProveedores, 7100, 7700)

End Sub

Private Sub txtNombre_KeyDown(KeyCode As Integer, Shift As Integer)

    Dim strCadena As String

    If KeyCode = 13 Then
        strCadena = Trim(Me.txtNombre.Text)
        If Len(strCadena) <= 50 Then
            ConfigurardgProveedores
            Call CargardgProveedores(, strCadena)
        End If
    End If

End Sub

Private Sub txtCUIT_KeyDown(KeyCode As Integer, Shift As Integer)

    Dim strCadena As String

    If KeyCode = 13 Then
        strCadena = Trim(Me.txtCUIT.Text)
        If Len(strCadena) <= 11 Then
            ConfigurardgProveedores
            Call CargardgProveedores(strCadena)
        End If
    End If

End Sub
