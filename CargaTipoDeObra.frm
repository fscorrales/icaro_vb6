VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form CargaTipoDeObra 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Tipo de Obra"
   ClientHeight    =   6690
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   15210
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6690
   ScaleWidth      =   15210
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Caption         =   "Listados de Actividades Categorizadas"
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
      Height          =   6495
      Left            =   0
      TabIndex        =   4
      Top             =   120
      Width           =   15135
      Begin VB.CommandButton cmdEliminarViviendas 
         BackColor       =   &H008080FF&
         Caption         =   "No Categorizado"
         Height          =   375
         Left            =   11640
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   6000
         Width           =   2000
      End
      Begin VB.CommandButton cmdEliminarReparaciones 
         BackColor       =   &H008080FF&
         Caption         =   "No Categorizado"
         Height          =   375
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   6000
         Width           =   2000
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid dgReparaciones 
         Height          =   5295
         Left            =   240
         TabIndex        =   7
         Top             =   600
         Width           =   4845
         _ExtentX        =   8546
         _ExtentY        =   9340
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.CommandButton cmdAgregarReparaciones 
         BackColor       =   &H008080FF&
         Caption         =   "Agregar a Reparaciones"
         Height          =   375
         Left            =   5160
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   6000
         Width           =   2000
      End
      Begin VB.CommandButton cmdAgregarViviendas 
         BackColor       =   &H008080FF&
         Caption         =   "Agregar a Viviendas"
         Height          =   375
         Left            =   8040
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   6000
         Width           =   2000
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid dgNoCategorizadas 
         Height          =   5295
         Left            =   5160
         TabIndex        =   8
         Top             =   600
         Width           =   4845
         _ExtentX        =   8546
         _ExtentY        =   9340
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid dgViviendas 
         Height          =   5295
         Left            =   10200
         TabIndex        =   9
         Top             =   600
         Width           =   4845
         _ExtentX        =   8546
         _ExtentY        =   9340
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Construcción de Viviendas (IVA 10,5%)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   10200
         TabIndex        =   10
         Top             =   360
         Width           =   4815
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "Reparaciones e Infra (IVA 21%)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   4815
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Actividades no Categorizadas"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   5160
         TabIndex        =   5
         Top             =   360
         Width           =   4815
      End
   End
End
Attribute VB_Name = "CargaTipoDeObra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAgregarReparaciones_Click()

    Call EditarTipoDeObra("Reparaciones")
    ConfigurardgTipoDeObra
    CargardgTipoDeObra

End Sub

Private Sub cmdAgregarViviendas_Click()

    Call EditarTipoDeObra("Viviendas")
    ConfigurardgTipoDeObra
    CargardgTipoDeObra

End Sub

Private Sub cmdEliminarReparaciones_Click()

    Call EditarTipoDeObra("NoCategorizadas", "Reparaciones")
    ConfigurardgTipoDeObra
    CargardgTipoDeObra

End Sub

Private Sub cmdEliminarViviendas_Click()

    Call EditarTipoDeObra("NoCategorizadas", "Viviendas")
    ConfigurardgTipoDeObra
    CargardgTipoDeObra

End Sub

Private Sub Form_Load()

    Call CenterMe(CargaTipoDeObra, 15300, 7000)
    
End Sub
