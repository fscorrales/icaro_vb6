VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form ListadoControlFuente10y13 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Control Fuente 10 y 13"
   ClientHeight    =   6405
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   12885
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6405
   ScaleWidth      =   12885
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Ejecución Presupuestaria"
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
      TabIndex        =   2
      Top             =   120
      Width           =   12615
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid dgEjecucionFuente10y13 
         Height          =   3615
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   12375
         _ExtentX        =   21828
         _ExtentY        =   6376
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Comprobantes"
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
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   4200
      Width           =   12615
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid dgComprobantesPorDecretos 
         Height          =   1695
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   12375
         _ExtentX        =   21828
         _ExtentY        =   2990
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
End
Attribute VB_Name = "ListadoControlFuente10y13"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

    Call CenterMe(ListadoControlFuente10y13, 12950, 6750)
    
End Sub

