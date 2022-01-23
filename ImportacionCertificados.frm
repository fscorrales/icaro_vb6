VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form ImportacionCertificados 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   5130
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3690
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5130
   ScaleWidth      =   3690
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtInforme 
      Height          =   4335
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   2
      Top             =   120
      Width           =   3375
   End
   Begin VB.CommandButton cmdDemandaLibre 
      Caption         =   "&Demanda Libre"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   4560
      Width           =   1335
   End
   Begin VB.CommandButton cmdAutogestivo 
      Caption         =   "&Autogestivo"
      Height          =   495
      Left            =   2160
      TabIndex        =   0
      Top             =   4560
      Width           =   1335
   End
   Begin MSComDlg.CommonDialog dlgCertificados 
      Left            =   1560
      Top             =   4560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "ImportacionCertificados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdDemandaLibre_Click()
    Call Info("- INICIANDO IMPORTACION DE DEMANDA LIBRE -")
    dlgCertificados.Filter = "Todos los Access(*.xls)|*.xls|"
    dlgCertificados.ShowOpen
    
    Call ImportarCertificado(dlgCertificados.FileName)
    
End Sub

Private Sub Form_Load()
    With ImportacionCertificados
        .Height = 5600
        .Width = 3700
    End With
    Call CenterMe(ImportacionCertificados)
End Sub

Public Function Info(StrAgregar As String)
    txtInforme.Text = txtInforme.Text & vbCrLf & StrAgregar
End Function
