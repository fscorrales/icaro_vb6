VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form ImportacionObras 
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
   Begin VB.CommandButton cmdFR 
      Caption         =   "FondoReparo"
      Height          =   495
      Left            =   1080
      TabIndex        =   3
      Top             =   4560
      Width           =   1455
   End
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
Attribute VB_Name = "ImportacionObras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAutogestivo_Click()

    Call InfoGeneral("", ImportacionObras)
    Call InfoGeneral("- IMPORTACION AUTOGESTIVO -", ImportacionObras)
    dlgCertificados.Filter = "Todos los Access(*.xls)|*.xls|"
    dlgCertificados.ShowOpen
    
    Call ImportarEPAM(dlgCertificados.FileName)

End Sub

Private Sub cmdDemandaLibre_Click()

    Call InfoGeneral("", ImportacionObras)
    Call InfoGeneral("- IMPORTACION DEMANDA LIBRE -", ImportacionObras)
    dlgCertificados.Filter = "Todos los Access(*.xls)|*.xls|"
    dlgCertificados.ShowOpen
    
    Call ImportarDemandaLibre(dlgCertificados.FileName)
    
End Sub

Private Sub cmdFR_Click()

    Call InfoGeneral("", ImportacionObras)
    Call InfoGeneral("- IMPORTACION FONDO DE REPARO -", ImportacionObras)
    dlgCertificados.Filter = "Todos los Access(*.xls)|*.xls|"
    dlgCertificados.ShowOpen
    
    Call ImportarDemandaLibre(dlgCertificados.FileName)
    
End Sub

Private Sub Form_Load()

    Call CenterMe(ImportacionObras, 3700, 5600)

End Sub
