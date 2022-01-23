VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form ListadoControlRetenciones 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Control Retenciones"
   ClientHeight    =   5970
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   14805
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5970
   ScaleWidth      =   14805
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Datos a filtrar"
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
      TabIndex        =   2
      Top             =   0
      Width           =   14535
      Begin VB.CommandButton cmdExportar 
         BackColor       =   &H008080FF&
         Caption         =   "Exportar Excel"
         Height          =   495
         Left            =   12120
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox txtEjercicio 
         Height          =   375
         Left            =   7800
         TabIndex        =   7
         Top             =   360
         Width           =   1095
      End
      Begin VB.ComboBox cmbProveedor 
         Height          =   315
         Left            =   1560
         TabIndex        =   4
         Top             =   360
         Width           =   4965
      End
      Begin VB.CommandButton cmdActualizar 
         BackColor       =   &H008080FF&
         Caption         =   "Actualizar"
         Height          =   495
         Left            =   9480
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "Proveedor"
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
         Width           =   1455
      End
      Begin VB.Label Label1 
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
         Left            =   6840
         TabIndex        =   5
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame frmMultiuso 
      Caption         =   "Control de Retenciones Practicadas vs Rentenciones Calculadas"
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
      Height          =   4935
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   14535
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid dgControlRetenciones 
         Height          =   4455
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   14295
         _ExtentX        =   25215
         _ExtentY        =   7858
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
End
Attribute VB_Name = "ListadoControlRetenciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdActualizar_Click()

    Call CargardgControlRetenciones(Me.cmbProveedor.Text, Me.txtEjercicio.Text)

End Sub

Private Sub cmdExportar_Click()

    If Exportar_Excel(App.Path & "\ControlRetenciones.xls", dgControlRetenciones) Then
        MsgBox " Datos exportados en " & App.Path, vbInformation
    End If

End Sub

Private Sub Form_Load()

    Call CenterMe(ListadoControlRetenciones, 14900, 6300)
    CargarcmbProveedor
    Me.txtEjercicio = Year(Date)
    
End Sub
