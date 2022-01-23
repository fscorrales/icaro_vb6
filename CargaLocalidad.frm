VERSION 5.00
Begin VB.Form CargaLocalidad 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Localidades"
   ClientHeight    =   1650
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4575
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1650
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdAgregar 
      BackColor       =   &H008080FF&
      Caption         =   "Guardar Datos"
      Height          =   495
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      Caption         =   "Ingreso de Localidad"
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
      TabIndex        =   0
      Top             =   120
      Width           =   4335
      Begin VB.TextBox txtLocalidad 
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   4095
      End
   End
End
Attribute VB_Name = "CargaLocalidad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAgregar_Click()

    GenerarLocalidad
    
End Sub

Private Sub Form_Load()

    Call CenterMe(CargaLocalidad, 4700, 2050)

End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    strEditandoLocalidad = ""

End Sub
