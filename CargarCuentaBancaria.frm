VERSION 5.00
Begin VB.Form CargaCuentaBancaria 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cuentas Bancarias"
   ClientHeight    =   2340
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5940
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2340
   ScaleWidth      =   5940
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdAgregar 
      BackColor       =   &H008080FF&
      Caption         =   "Guardar Datos"
      Height          =   495
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Frame Frame3 
      Caption         =   "Datos de la Cuenta Bancaria"
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
      Width           =   5655
      Begin VB.TextBox txtCuentaBancaria 
         Height          =   285
         Left            =   1440
         TabIndex        =   1
         Top             =   480
         Width           =   2000
      End
      Begin VB.TextBox txtDescripcion 
         Height          =   285
         Left            =   1440
         TabIndex        =   2
         Top             =   1080
         Width           =   4000
      End
      Begin VB.Label Label1 
         Caption         =   "Nro. Cuenta"
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
         Top             =   480
         Width           =   1455
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
         TabIndex        =   3
         Top             =   1080
         Width           =   1455
      End
   End
End
Attribute VB_Name = "CargaCuentaBancaria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAgregar_Click()

    GenerarCuentaBancaria

End Sub

Private Sub Form_Load()
    
    Call CenterMe(CargaCuentaBancaria, 6000, 2700)
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    strEditandoCuentaBancaria = ""
        
End Sub


