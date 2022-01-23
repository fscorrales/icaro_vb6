VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form ListadoRetenciones 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Listado de Retenciones"
   ClientHeight    =   7170
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5415
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7170
   ScaleWidth      =   5415
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Listado de Partidas"
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
      TabIndex        =   4
      Top             =   0
      Width           =   5175
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid dgRetenciones 
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
      Left            =   480
      TabIndex        =   0
      Top             =   5040
      Width           =   4455
      Begin VB.CommandButton cmdAgregar 
         BackColor       =   &H008080FF&
         Caption         =   "Agregar Nueva Retención (F6)"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   360
         Width           =   4215
      End
      Begin VB.CommandButton cmdEditar 
         BackColor       =   &H008080FF&
         Caption         =   "Editar Retención Seleccionada (F9)"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   960
         Width           =   4215
      End
      Begin VB.CommandButton cmdEliminar 
         BackColor       =   &H008080FF&
         Caption         =   "Eliminar Eliminar Retención Seleccionada (F8)"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   1560
         Width           =   4215
      End
   End
End
Attribute VB_Name = "ListadoRetenciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAgregar_Click()

    Unload ListadoRetenciones
    CodigoRetencion.Show
    
End Sub

Private Sub cmdEditar_Click()

    EditarCodigoRetencion
    
End Sub

Private Sub cmdEliminar_Click()

    EliminarCodigoRetencion
    
End Sub

Private Sub dgRetenciones_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
        Case Is = vbKeyF6
            Unload ListadoRetenciones
            CodigoRetencion.Show
        Case Is = vbKeyF9
            EditarCodigoRetencion
        Case Is = vbKeyF8
            EliminarCodigoRetencion
    End Select
    
End Sub

Private Sub Form_Load()

    Call CenterMe(ListadoRetenciones, 5500, 7500)

End Sub
