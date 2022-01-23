VERSION 5.00
Begin VB.Form CargaRetencion 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Retenciones y Deducciones"
   ClientHeight    =   3570
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6855
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   6855
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox txtRetencion 
      Height          =   315
      Left            =   1560
      TabIndex        =   0
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Frame Frame3 
      Caption         =   "Datos de la Deducción"
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
      Height          =   2775
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   6615
      Begin VB.TextBox txtEjercicio 
         Height          =   285
         Left            =   4920
         TabIndex        =   9
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox txtComprobante 
         Height          =   285
         Left            =   1440
         TabIndex        =   8
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox txtMonto 
         Height          =   285
         Left            =   1440
         TabIndex        =   3
         Top             =   2280
         Width           =   1575
      End
      Begin VB.TextBox txtDescripcion 
         Height          =   285
         Left            =   1440
         TabIndex        =   2
         Top             =   1680
         Width           =   5055
      End
      Begin VB.Label Label11 
         Caption         =   "Ejercicio"
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
         Left            =   3480
         TabIndex        =   11
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Comrprobante"
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
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Monto"
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
         TabIndex        =   7
         Top             =   2280
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Código"
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
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label3 
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
         TabIndex        =   5
         Top             =   1680
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmdAgregar 
      BackColor       =   &H008080FF&
      Caption         =   "Guardar Datos"
      Height          =   495
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3000
      Width           =   1815
   End
End
Attribute VB_Name = "CargaRetencion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

    Call CenterMe(CargaRetencion, 6950, 3900)
    txtComprobante.Enabled = False
    txtEjercicio.Enabled = False
    txtDescripcion.Enabled = False
    CargarComboRetenciones
    
End Sub

Private Sub txtRetencion_LostFocus()

    Dim SQL As String

    If Me.txtRetencion.Text <> "" Then
        If ValidarRetencionCargaRetencion(Me.txtRetencion.Text) Then
            Set rstBuscarIcaro = New ADODB.Recordset
            SQL = "Select * From CODIGOSRETENCIONES Where CODIGO = '" & Me.txtRetencion.Text & "'"
            rstBuscarIcaro.Open SQL, dbIcaro, adOpenDynamic, adLockOptimistic
            Me.txtDescripcion.Text = rstBuscarIcaro!Descripcion
            rstBuscarIcaro.Close
            Set rstBuscarIcaro = Nothing
            Me.txtMonto.SetFocus
        End If
    End If
        
End Sub

Private Sub cmdAgregar_Click()

    GenerarRetencion

End Sub
