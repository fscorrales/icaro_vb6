VERSION 5.00
Begin VB.Form CargaCUIT 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "CUIT"
   ClientHeight    =   2400
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6900
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2400
   ScaleWidth      =   6900
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdAgregar 
      BackColor       =   &H008080FF&
      Caption         =   "Guardar Datos"
      Height          =   495
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Frame Frame3 
      Caption         =   "Datos del Proveedor"
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
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6615
      Begin VB.TextBox txtDescripcion 
         Height          =   285
         Left            =   1440
         TabIndex        =   2
         Top             =   1080
         Width           =   5055
      End
      Begin VB.TextBox txtCUIT 
         Height          =   285
         Left            =   1440
         TabIndex        =   1
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Descripcion"
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
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "CUIT / DNI"
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
         TabIndex        =   3
         Top             =   480
         Width           =   1455
      End
   End
End
Attribute VB_Name = "CargaCUIT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAgregar_Click()

    GenerarProveedor

End Sub

Private Sub Form_Load()

    Call CenterMe(CargaCUIT, 7000, 2700)

End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    strEditandoProveedor = ""
    
    If bolCUITDesdeCarga = True Then
        CargaGasto.Show
        bolCUITDesdeCarga = False
        CargaGasto.txtCUIT.SetFocus
    End If
    
    If bolAutocargaCertificado = True Then
        VaciarVariablesAutocarga
        AutocargaCertificados.Show
    End If
    
End Sub
