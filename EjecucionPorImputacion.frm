VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form EjecucionPorImputacion 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Control de Ejecución SIIF por Imputación (Sistema de Seguimiento de Cuentas - INVICO)"
   ClientHeight    =   8040
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   13425
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8040
   ScaleWidth      =   13425
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdExportarImputacionSeleccionada 
      BackColor       =   &H008080FF&
      Caption         =   "Exportar Imputación Seleccionada"
      Height          =   495
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   7440
      Width           =   2055
   End
   Begin VB.Frame frmImputacionSeleccionada 
      Caption         =   "Datos desagregados de la Imputación seleccionada"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   3135
      Left            =   120
      TabIndex        =   9
      Top             =   4200
      Width           =   13215
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid dgEjecucionPorImputacionSeleccionada 
         Height          =   2655
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   12975
         _ExtentX        =   22886
         _ExtentY        =   4683
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.CommandButton cmdExportarResumen 
      BackColor       =   &H008080FF&
      Caption         =   "Exportar Resumen de Ejecución por Imputación"
      Height          =   495
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7440
      Width           =   2055
   End
   Begin VB.Frame Frame2 
      Caption         =   "Identificación del Período de consulta (dd/mm/aaaa)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
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
      Width           =   13215
      Begin VB.TextBox txtFechaHasta 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4800
         TabIndex        =   7
         Top             =   360
         Width           =   1695
      End
      Begin VB.TextBox txtFechaDesde 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1440
         TabIndex        =   6
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton cmdActualizar 
         BackColor       =   &H008080FF&
         Caption         =   "Actualizar"
         Height          =   495
         Left            =   11280
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label7 
         Caption         =   "Fecha Desde:"
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
         Top             =   360
         Width           =   1465
      End
      Begin VB.Label Label8 
         Caption         =   "Fecha Hasta:"
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
         TabIndex        =   3
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Ejecución por Imputación según Período filtrado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   13215
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid dgEjecucionPorImputacion 
         Height          =   2775
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   12975
         _ExtentX        =   22886
         _ExtentY        =   4895
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
End
Attribute VB_Name = "EjecucionPorImputacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdActualizar_Click()
    
    If ValidarActualizarEjecucionPorImputacion = True Then
        Call ConfigurardgEjecucionPorImputacion
        Call CargardgEjecucionPorImputacion(Me.txtFechaDesde.Text, Me.txtFechaHasta.Text)
    End If

End Sub

Private Sub cmdExportarImputacionSeleccionada_Click()

    Dim strLink As String
    Dim strFechaDesde As String
    Dim strFechaHasta As String
    
    strFechaDesde = Replace(Me.txtFechaDesde, "/", "-")
    strFechaHasta = Replace(Me.txtFechaHasta, "/", "-")
    
    strLink = "\EjecuciónPorImputacionSeleccionada del " & strFechaDesde & " al " & strFechaHasta & ".xls"
        
    If Exportar_Excel(App.Path & strLink, Me.dgEjecucionPorImputacionSeleccionada) Then
        MsgBox " Datos exportados en " & App.Path, vbInformation
    End If

End Sub

Private Sub cmdExportarResumen_Click()

    Dim strLink As String
    Dim strFechaDesde As String
    Dim strFechaHasta As String
    
    strFechaDesde = Replace(Me.txtFechaDesde, "/", "-")
    strFechaHasta = Replace(Me.txtFechaHasta, "/", "-")
    
    strLink = "\EjecuciónPorImputación del " & strFechaDesde & " al " & strFechaHasta & ".xls"
        
    If Exportar_Excel(App.Path & strLink, Me.dgEjecucionPorImputacion) Then
        MsgBox " Datos exportados en " & App.Path, vbInformation
    End If

End Sub

Private Sub dgEjecucionPorImputacion_RowColChange()

    Dim i As Integer
    
    With Me.dgEjecucionPorImputacion
        i = .Row
        If i <> 0 Then
            Call ConfigurardgEjecucionPorImputacionSeleccionada
            Call CargardgEjecucionPorImputacionSeleccionada(Me.txtFechaDesde.Text, _
            Me.txtFechaHasta.Text, .TextMatrix(i, 0))
            Me.frmImputacionSeleccionada.Caption = "Datos desagregados " _
            & "de la Imputación " & .TextMatrix(i, 1)
        End If
    End With
    
    i = 0

End Sub

Private Sub Form_Load()

    Call CenterMe(EjecucionPorImputacion, 13500, 8500)
    Me.dgEjecucionPorImputacion.Clear
    Me.dgEjecucionPorImputacionSeleccionada.Clear

End Sub
