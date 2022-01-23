VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form ListadoCuentasBancarias 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cuentas Bancarias"
   ClientHeight    =   7200
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5280
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7200
   ScaleWidth      =   5280
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Listado de Cuentas Bancarias"
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
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   5175
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid dgCuentasBancarias 
         Height          =   4455
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   4900
         _ExtentX        =   8652
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
      Left            =   360
      TabIndex        =   0
      Top             =   5040
      Width           =   4455
      Begin VB.CommandButton cmdAgregar 
         BackColor       =   &H008080FF&
         Caption         =   "Agregar Cuenta Bancaria (F6)"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   360
         Width           =   4215
      End
      Begin VB.CommandButton cmdEditar 
         BackColor       =   &H008080FF&
         Caption         =   "Editar Cuenta Bancarias (F9)"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   960
         Width           =   4215
      End
      Begin VB.CommandButton cmdEliminar 
         BackColor       =   &H008080FF&
         Caption         =   "Eliminar Cuenta Bancaria (F8)"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   1560
         Width           =   4215
      End
   End
End
Attribute VB_Name = "ListadoCuentasBancarias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAgregar_Click()

    Unload ListadoCuentasBancarias
    CargaCuentaBancaria.Show
            
End Sub

Private Sub cmdEditar_Click()

    EditarCuentaBancaria
    
End Sub

Private Sub cmdEliminar_Click()

    EliminarCuentaBancaria
    
End Sub

Private Sub Form_Load()
    
    Call CenterMe(ListadoCuentasBancarias, 5500, 7500)

End Sub

Private Sub dgCuentasBancarias_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
        Case Is = vbKeyF6
            Unload ListadoCuentasBancarias
            CargaCuentaBancaria.Show
        Case Is = vbKeyF9
            EditarCuentaBancaria
        Case Is = vbKeyF8
            EliminarCuentaBancaria
    End Select
    
End Sub

