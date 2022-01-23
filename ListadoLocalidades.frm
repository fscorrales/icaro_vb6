VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form ListadoLocalidades 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Localidades"
   ClientHeight    =   7095
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2985
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7095
   ScaleWidth      =   2985
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
      Left            =   240
      TabIndex        =   2
      Top             =   5040
      Width           =   2415
      Begin VB.CommandButton cmdEliminar 
         BackColor       =   &H008080FF&
         Caption         =   "Eliminar Localidad (F8)"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1560
         Width           =   2175
      End
      Begin VB.CommandButton cmdEditar 
         BackColor       =   &H008080FF&
         Caption         =   "Editar Localidad (F9)"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   960
         Width           =   2175
      End
      Begin VB.CommandButton cmdAgregar 
         BackColor       =   &H008080FF&
         Caption         =   "Agregar Localidad (F6)"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Listado de Localidades"
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
      TabIndex        =   0
      Top             =   0
      Width           =   2895
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid dgLocalidades 
         Height          =   4455
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   7858
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
End
Attribute VB_Name = "ListadoLocalidades"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAgregar_Click()

    Unload ListadoLocalidades
    CargaLocalidad.Show

End Sub

Private Sub cmdEditar_Click()

    EditarLocalidad
 
End Sub

Private Sub cmdEliminar_Click()
    
    EliminarLocalidad
    
End Sub

Private Sub Form_Load()

    Call CenterMe(ListadoLocalidades, 3100, 7500)

End Sub

Private Sub dgLocalidades_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
        Case Is = vbKeyF6
            Unload ListadoLocalidades
            CargaLocalidad.Show
        Case Is = vbKeyF9
            EditarLocalidad
        Case Is = vbKeyF8
            EliminarLocalidad
    End Select
    
End Sub
