VERSION 5.00
Begin VB.Form ImputacionEPAM 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Obras EPAM Consolidada"
   ClientHeight    =   4080
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6990
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4080
   ScaleWidth      =   6990
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdAgregar 
      BackColor       =   &H008080FF&
      Caption         =   "Guardar Datos"
      Height          =   495
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3480
      Width           =   1815
   End
   Begin VB.Frame Frame3 
      Caption         =   "Montos Acumulados al 30/04/2009"
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
      TabIndex        =   13
      Top             =   2520
      Width           =   6735
      Begin VB.TextBox txtPagado 
         Enabled         =   0   'False
         Height          =   285
         Left            =   5040
         TabIndex        =   5
         Text            =   "0"
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox txtImputado 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1560
         TabIndex        =   4
         Text            =   "0"
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label6 
         Caption         =   "Pagado"
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
         TabIndex        =   15
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label5 
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
         Left            =   120
         TabIndex        =   14
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
      Height          =   1335
      Left            =   120
      TabIndex        =   9
      Top             =   1080
      Width           =   6735
      Begin VB.TextBox txtPresupuestado 
         Height          =   285
         Left            =   1560
         TabIndex        =   3
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox txtResolucion 
         Height          =   285
         Left            =   5040
         TabIndex        =   2
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox txtExpediente 
         Height          =   285
         Left            =   1560
         TabIndex        =   1
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "Monto Presp."
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
         TabIndex        =   12
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
         TabIndex        =   10
         Top             =   360
         Width           =   1455
      End
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
      TabIndex        =   7
      Top             =   120
      Width           =   6735
      Begin VB.ComboBox txtObra 
         Height          =   315
         Left            =   1560
         TabIndex        =   0
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
         TabIndex        =   8
         Top             =   360
         Width           =   1455
      End
   End
End
Attribute VB_Name = "ImputacionEPAM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAgregar_Click()
    
    If ValidarEPAMConsolidado = True Then
        Call GenerarEPAMConsolidado(Me.txtObra.Text, Me.txtExpediente.Text, _
        Me.txtResolucion.Text, Me.txtPresupuestado.Text, Me.txtPagado.Text, _
        Me.txtImputado.Text)
        MsgBox "Los datos fueron agregados"
        EPAMConsolidado.Show
        ConfigurardgEPAMConsolidado
        CargardgEPAMConsolidado
        Dim i As String
        i = EPAMConsolidado.dgEPAMConsolidado.Row
        ConfigurardgResolucionesEPAM
        CargardgResolucionesEPAM (EPAMConsolidado.dgEPAMConsolidado.TextMatrix(i, 0))
        i = ""
        Unload ImputacionEPAM
    End If
    
End Sub

Private Sub Form_Load()

    Call CenterMe(ImputacionEPAM, 7100, 4400)
    CargarComboImputacionEPAM
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    strEditandoEPAMConsolidado = ""

End Sub

Private Sub txtObra_LostFocus()
    
    If ValidarObraImputacionEPAM(Me.txtObra.Text) Then
        Me.txtExpediente.SetFocus
    End If
        
End Sub
