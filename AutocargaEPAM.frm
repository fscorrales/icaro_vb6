VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form AutocargaEPAM 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Asistente de Carga"
   ClientHeight    =   6015
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9870
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   9870
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Caption         =   "Autogestivo"
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
      Height          =   5775
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9615
      Begin VB.CommandButton cmdPagarEPAM 
         BackColor       =   &H008080FF&
         Caption         =   "Pagar EPAM y No Imputar"
         Enabled         =   0   'False
         Height          =   495
         Left            =   6960
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   5160
         Width           =   1695
      End
      Begin VB.CommandButton cmdAgregarEPAM 
         BackColor       =   &H008080FF&
         Caption         =   "Cargar EPAM (F6)"
         Height          =   495
         Left            =   960
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   5160
         Width           =   1695
      End
      Begin VB.CommandButton cmdEliminarEPAM 
         BackColor       =   &H008080FF&
         Caption         =   "Eliminar EPAM (F9)"
         Height          =   495
         Left            =   3960
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   5160
         Width           =   1695
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid dgAutocargaEPAM 
         Height          =   4815
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   8493
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
End
Attribute VB_Name = "AutocargaEPAM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAgregarEPAM_Click()

    LlenarVariablesAutocargaEPAM
    AgregarEPAM
    
End Sub

Private Sub cmdEliminarEPAM_Click()

    EliminarEPAM
    
End Sub

Private Sub cmdPagarEPAM_Click()

    PagarEPAM
    
End Sub

Private Sub dgEPAM_KeyDown(KeyCode As Integer, Shift As Integer)

    Dim i As Integer
    If KeyCode = vbKeyF6 Then
        AgregarEPAM
    End If
    
    If KeyCode = vbKeyF10 Then
        PagarEPAM
    End If
    
    If KeyCode = vbKeyF9 Then
        EliminarEPAM
    End If
    
End Sub


Private Sub Form_Load()

    Call CenterMe(AutocargaEPAM, 10000, 6350)
    VaciarVariablesAutocarga
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'dgCertificado.Clear
    'dgEPAM.Clear
End Sub
