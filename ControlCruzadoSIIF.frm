VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form ControlCruzadoSIIF 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Control Ejecución Presupuestaria Acumulada"
   ClientHeight    =   5520
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4815
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   4815
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdImportar 
      BackColor       =   &H008080FF&
      Caption         =   "Importar Reporte SIIF para Control Cruzado"
      Height          =   495
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4920
      Width           =   1935
   End
   Begin VB.Frame frmPeriodoHasta 
      Caption         =   "Identificación del Período Hasta"
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
      Height          =   1095
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   4575
      Begin VB.TextBox txtFecha 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2640
         TabIndex        =   5
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha de Control"
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
         Left            =   480
         TabIndex        =   4
         Top             =   480
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Resumen del Control"
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
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   4575
      Begin VB.TextBox txtInforme 
         Height          =   3015
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   2
         Top             =   360
         Width           =   4335
      End
   End
   Begin MSComDlg.CommonDialog dlgMultifuncion 
      Left            =   4080
      Top             =   4920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "ControlCruzadoSIIF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdImportar_Click()

'    Dim SQL As String
'
'    SQL = "Select * From CODIGOLIQUIDACIONES Where CODIGO= '" & Left(Me.cmbCodigoLiquidacion.Text, 4) & "'"
'
'    If SQLNoMatch(SQL) Then
'        MsgBox "Debe seleccionar un Código de Liquidación del Listado", vbCritical + vbOKOnly, "CÓDIGO LIQUIDACIÓN INEXISTENTE"
'        Me.cmbCodigoLiquidacion.SetFocus
'    Else
'        ImportarLiquidacionSueldosSchema (Left(Me.cmbCodigoLiquidacion.Text, 4))
'    End If

    Call ControlarEjecucionAcumuladaSIIF(Me.txtFecha.Text, strControlCruzadoSIIF)

End Sub

Private Sub Form_Load()

    Call CenterMe(ControlCruzadoSIIF, 4900, 5850)
    
    Select Case strControlCruzadoSIIF
    Case "rf602txt"
        Me.Caption = "Control Presupuestario Anual (rf602 opción .txt)"
'        Me.frRecuadro.Caption = "Datos del Parentesco SIRADIG"
    Case "rf602pdf"
        Me.Caption = "Control Presupuestario Anual (rf602 opción .pdf)"
'        Me.frRecuadro.Caption = "Datos del Parentesco SIRADIG"
    Case "rf636m"
        Me.Caption = "Control Presupuestario Mes Específico (rf636m)"
'        Me.frRecuadro.Caption = "Datos de las Deducciones SIRADIG"
    End Select

End Sub

Private Sub Form_Unload(Cancel As Integer)

    strControlCruzadoSIIF = ""

End Sub
