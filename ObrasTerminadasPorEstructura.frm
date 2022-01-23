VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form ObrasTerminadasPorEstructura 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Informe de Estructura Presupuestaria"
   ClientHeight    =   4920
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   14910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   14910
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdExportar 
      Caption         =   "Exportar Excel"
      Height          =   615
      Left            =   240
      TabIndex        =   2
      Top             =   4200
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Sistema de Información de Programación Presupuestaria"
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
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   14655
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid dgObrasTerminadas 
         Height          =   3615
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   14415
         _ExtentX        =   25426
         _ExtentY        =   6376
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
End
Attribute VB_Name = "ObrasTerminadasPorEstructura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExportar_Click()
    If Exportar_Excel(App.Path & "\ObrasTerminadasPorEstructura.xls", dgObrasTerminadas) Then
        MsgBox " Datos exportados en " & App.Path, vbInformation
    End If
End Sub

Private Sub Form_Load()
    
    Call CenterMe(ObrasTerminadasPorEstructura, 15000, 5250)
    ConfigurardgObrasTerminadas
    CargardgObrasTerminadas
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    strEjercicio = ""
    
End Sub

