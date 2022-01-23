VERSION 5.00
Begin VB.Form ImputacionResolucionesEPAM 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Reterminaciones y Ampliaciones Obras EPAM Consolidadas"
   ClientHeight    =   3390
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6975
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   6975
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdAgregar 
      BackColor       =   &H008080FF&
      Caption         =   "Guardar Datos"
      Height          =   495
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   2880
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      Caption         =   "Imputación"
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
      Height          =   855
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   6735
      Begin VB.TextBox txtObra 
         Height          =   285
         Left            =   1560
         TabIndex        =   13
         Top             =   360
         Width           =   5055
      End
      Begin VB.Label Label3 
         Caption         =   "Obra"
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
         TabIndex        =   10
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Presupuesto Incial"
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
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   6735
      Begin VB.TextBox txtImputado 
         Height          =   285
         Left            =   5040
         TabIndex        =   6
         Top             =   1320
         Width           =   1575
      End
      Begin VB.TextBox txtDescripcion 
         Height          =   285
         Left            =   1560
         TabIndex        =   1
         Top             =   360
         Width           =   5055
      End
      Begin VB.TextBox txtExpediente 
         Height          =   285
         Left            =   1560
         TabIndex        =   2
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox txtResolucion 
         Height          =   285
         Left            =   5040
         TabIndex        =   3
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox txtPresupuestado 
         Height          =   285
         Left            =   1560
         TabIndex        =   5
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label6 
         Caption         =   "Imputado"
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
         Left            =   3600
         TabIndex        =   12
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "Descripción"
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
         TabIndex        =   11
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Expediente"
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
         TabIndex        =   8
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Resolución"
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
         Left            =   3600
         TabIndex        =   7
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Presupuestado"
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
         TabIndex        =   4
         Top             =   1320
         Width           =   1455
      End
   End
End
Attribute VB_Name = "ImputacionResolucionesEPAM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAgregar_Click()

    GenerarResolucionEPAM
    
End Sub

Private Sub Form_Load()

    Call CenterMe(ImputacionResolucionesEPAM, 7100, 3800)
    txtObra.Enabled = False
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    strEditandoResolucionEPAM = ""
    
End Sub
