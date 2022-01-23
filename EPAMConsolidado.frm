VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form EPAMConsolidado 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "EPAM Consolidado"
   ClientHeight    =   7905
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7905
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Obras EPAM Consolidadas"
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
      Height          =   5055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11655
      Begin VB.CommandButton cmdExportarTotal 
         BackColor       =   &H008080FF&
         Caption         =   "Exportar Todas las Obras Consolidadas"
         Height          =   495
         Left            =   9000
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   4440
         Width           =   1695
      End
      Begin VB.CommandButton cmdEliminarConsolidado 
         BackColor       =   &H008080FF&
         Caption         =   "Eliminar Obra Consolidada (F8)"
         Height          =   495
         Left            =   6360
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   4440
         Width           =   1695
      End
      Begin VB.CommandButton cmdEditarConsolidado 
         BackColor       =   &H008080FF&
         Caption         =   "Modificar Obra Consolidada (F9)"
         Height          =   495
         Left            =   3600
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   4440
         Width           =   1695
      End
      Begin VB.CommandButton cmdAgregarConsolidado 
         BackColor       =   &H008080FF&
         Caption         =   "Agregar Obra EPAM a Consolidar (F6)"
         Height          =   495
         Left            =   960
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   4440
         Width           =   1695
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid dgEPAMConsolidado 
         Height          =   3975
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   7011
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Descriminación de la Obra (Redeterminaciones y Actualizaciones)"
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
      Height          =   2535
      Left            =   120
      TabIndex        =   2
      Top             =   5280
      Width           =   11655
      Begin VB.CommandButton cmdEliminarRedeterminacion 
         BackColor       =   &H008080FF&
         Caption         =   "Eliminar Redet. o Actualización (F8)"
         Height          =   495
         Left            =   6360
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   1920
         Width           =   1695
      End
      Begin VB.CommandButton cmdExportarEspecifica 
         BackColor       =   &H008080FF&
         Caption         =   "Exportar Obra Específica"
         Height          =   495
         Left            =   9000
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   1920
         Width           =   1695
      End
      Begin VB.CommandButton cmdEditarRedeterminacion 
         BackColor       =   &H008080FF&
         Caption         =   "Modificar Redet. o Actualización (F9)"
         Height          =   495
         Left            =   3600
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1920
         Width           =   1695
      End
      Begin VB.CommandButton cmdAgregarRedeterminacion 
         BackColor       =   &H008080FF&
         Caption         =   "Agregar Redet. o Actualización (F6)"
         Height          =   495
         Left            =   960
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1920
         Width           =   1695
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid dgResolucionesEPAM 
         Height          =   1455
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   2566
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
End
Attribute VB_Name = "EPAMConsolidado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAgregarRedeterminacion_Click()
    
    Dim i As String
    
    ImputacionResolucionesEPAM.Show
    i = dgEPAMConsolidado.Row
    ImputacionResolucionesEPAM.txtObra.Text = dgEPAMConsolidado.TextMatrix(i, 0)
    Unload EPAMConsolidado
    i = ""
    
End Sub

Private Sub cmdEditarRedeterminacion_Click()

    EditarResolucionEPAM
    
End Sub

Private Sub cmdEliminarRedeterminacion_Click()

    EliminarResolucionEPAM
    
End Sub

Private Sub cmdExportarEspecifica_Click()

    If Exportar_Excel(App.Path & "\EPAMDiscriminado.xls", dgResolucionesEPAM) Then
        MsgBox " Datos exportados en " & App.Path, vbInformation
    End If
    
End Sub

Private Sub cmdExportarTotal_Click()

    If ExportarEPAMConsolidado(App.Path & "\EPAMConsolidado.xls", dgEPAMConsolidado, dgResolucionesEPAM) Then
        MsgBox " Datos exportados en " & App.Path, vbInformation
    End If
    
End Sub

Private Sub dgEPAMConsolidado_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
        Case Is = vbKeyF4
            dgResolucionesEPAM.SetFocus
        Case Is = vbKeyF6
            Unload EPAMConsolidado
            ImputacionEPAM.Show
        Case Is = vbKeyF9
            EditarEPAMConsolidado
        Case Is = vbKeyF8
            EliminarEPAMConsolidado
    End Select
    
End Sub

Private Sub dgResolucionesEPAM_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
        Case Is = vbKeyF6
            Dim i As String
            ImputacionResolucionesEPAM.Show
            i = dgEPAMConsolidado.Row
            ImputacionResolucionesEPAM.txtObra.Text = dgEPAMConsolidado.TextMatrix(i, 0)
            Unload EPAMConsolidado
            i = ""
        Case Is = vbKeyF9
            EditarResolucionEPAM
        Case Is = vbKeyF8
            EliminarResolucionEPAM
        Case Is = vbKeyF12
            dgEPAMConsolidado.SetFocus
    End Select
    
End Sub

Private Sub cmdExportar_Click()

    If Exportar_Excel(App.Path & "\EPAMConsolidado.xls", dgEPAMConsolidado) Then
        MsgBox " Datos exportados en " & App.Path, vbInformation
    End If
    
End Sub

Private Sub cmdAgregarConsolidado_Click()

    Unload EPAMConsolidado
    ImputacionEPAM.Show
    
End Sub

Private Sub cmdEditarConsolidado_Click()
    
    EditarEPAMConsolidado
    
End Sub

Private Sub cmdEliminarConsolidado_Click()
    
    EliminarEPAMConsolidado
    
End Sub

Private Sub Form_Load()
    
    Call CenterMe(EPAMConsolidado, 12000, 8250)
    
End Sub

Private Sub dgEPAMConsolidado_RowColChange()

    Dim i As Integer
    i = dgEPAMConsolidado.Row
    ConfigurardgResolucionesEPAM
    CargardgResolucionesEPAM (dgEPAMConsolidado.TextMatrix(i, 0))
    i = 0
    
End Sub

