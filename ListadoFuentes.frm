VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form ListadoFuentes 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Fuentes"
   ClientHeight    =   7155
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5235
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7155
   ScaleWidth      =   5235
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Listado de Fuentes"
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
      TabIndex        =   5
      Top             =   0
      Width           =   5175
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid dgFuentes 
         Height          =   4455
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   4905
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
         Caption         =   "Agregar Fuente Presupuestaria (F6)"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   360
         Width           =   4215
      End
      Begin VB.CommandButton cmdEditar 
         BackColor       =   &H008080FF&
         Caption         =   "Editar Fuente Presupuestaria (F9)"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   960
         Width           =   4215
      End
      Begin VB.CommandButton cmdEliminar 
         BackColor       =   &H008080FF&
         Caption         =   "Eliminar Fuente Presupuestaria (F8)"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1560
         Width           =   4215
      End
   End
End
Attribute VB_Name = "ListadoFuentes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAgregar_Click()

    Unload ListadoFuentes
    CargaFuente.Show
    
End Sub

Private Sub cmdEditar_Click()

    EditarFuente
    
End Sub

Private Sub cmdEliminar_Click()

    EliminarFuente
    
End Sub

Private Sub Form_Load()
    
    Call CenterMe(ListadoFuentes, 5350, 7500)

End Sub

Private Sub dgFuentes_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
        Case Is = vbKeyF6
            Unload ListadoFuentes
            CargaFuente.Show
        Case Is = vbKeyF9
            EditarFuente
        Case Is = vbKeyF8
            EliminarFuente
    End Select
    
End Sub
