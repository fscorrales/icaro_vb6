VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form ListadoImputacion 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Listado de Imputación"
   ClientHeight    =   7065
   ClientLeft      =   3525
   ClientTop       =   285
   ClientWidth     =   7950
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7065
   ScaleWidth      =   7950
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdAgregar 
      BackColor       =   &H008080FF&
      Caption         =   "Agregar Comprobante (F6)"
      Height          =   495
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3360
      Width           =   1695
   End
   Begin VB.Frame Frame3 
      Caption         =   "Retenciones (F2)"
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
      Height          =   2895
      Left            =   4080
      TabIndex        =   2
      Top             =   4080
      Width           =   3735
      Begin VB.CommandButton cmdEliminarRetencion 
         BackColor       =   &H008080FF&
         Caption         =   "Eliminar Retención (F8)"
         Height          =   495
         Left            =   1920
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   2280
         Width           =   1695
      End
      Begin VB.CommandButton cmdAgregarRetencion 
         BackColor       =   &H008080FF&
         Caption         =   "Agregar Retención (F6)"
         Height          =   495
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   2280
         Width           =   1695
      End
      Begin VB.TextBox txtTotalRetencion 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   300
         Left            =   2220
         TabIndex        =   7
         Top             =   1920
         Width           =   1000
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid dgRetencion 
         Height          =   1455
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   2566
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Imputación (F4)"
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
      Height          =   2895
      Left            =   120
      TabIndex        =   1
      Top             =   4080
      Width           =   3735
      Begin VB.CommandButton cmdEliminarImputacion 
         BackColor       =   &H008080FF&
         Caption         =   "Eliminar Imputación (F8)"
         Height          =   495
         Left            =   1920
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   2280
         Width           =   1695
      End
      Begin VB.CommandButton cmdAgregarImputacion 
         BackColor       =   &H008080FF&
         Caption         =   "Agregar Imputación (F6)"
         Height          =   495
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   2280
         Width           =   1695
      End
      Begin VB.TextBox txtTotalImputacion 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   300
         Left            =   2220
         TabIndex        =   6
         Top             =   1920
         Width           =   1000
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid dgImputacion 
         Height          =   1455
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   2566
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Comprobantes (Asistente Carga Certif. F10 y Asistente EPAM F11)"
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
      Height          =   3855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7695
      Begin VB.TextBox txtImporte 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   6240
         TabIndex        =   17
         Top             =   360
         Width           =   1050
      End
      Begin VB.TextBox txtFecha 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1200
         TabIndex        =   16
         Top             =   360
         Width           =   1050
      End
      Begin VB.TextBox txtComprobante 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   1100
      End
      Begin VB.CommandButton cmdEliminar 
         BackColor       =   &H008080FF&
         Caption         =   "Eliminar Comprobante (F8)"
         Height          =   495
         Left            =   5880
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   3240
         Width           =   1695
      End
      Begin VB.CommandButton cmdModificar 
         BackColor       =   &H008080FF&
         Caption         =   "Modificar Comprobante (F9)"
         Height          =   495
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   3240
         Width           =   1695
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid dgListado 
         Height          =   2415
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   4260
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
End
Attribute VB_Name = "ListadoImputacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAgregarRetencion_Click()

    Dim i As Integer
    i = dgListado.Row
    With CargaRetencion
        .Show
        .txtComprobante.Text = Left(dgListado.TextMatrix(i, 0), 5)
        .txtEjercicio.Text = "20" & Right(dgListado.TextMatrix(i, 0), 2)
    End With
    Unload ListadoImputacion
    i = 0
    
End Sub

Private Sub cmdEliminar_Click()

    EliminarComprobante
    
End Sub

Private Sub cmdEliminarImputacion_Click()

    EliminarImputacion
    
End Sub

Private Sub cmdEliminarRetencion_Click()

    EliminarRetencion
    
End Sub

Private Sub cmdModificar_Click()

    EditarComprobante
    
End Sub

Private Sub dgListado_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
        Case Is = vbKeyF6
            CargaGasto.Show
            Unload ListadoImputacion
        Case Is = vbKeyF10
            AutocargaCertificados.Show
            ConfigurardgAutocargaCertificados
            CargardgAutocargaCertificados
            Unload ListadoImputacion
        Case Is = vbKeyF11
            AutocargaEPAM.Show
            ConfigurardgAutocargaEPAM
            CargardgAutocargaEPAM
            Unload ListadoImputacion
        Case Is = vbKeyF9
            EditarComprobante
        Case Is = vbKeyF8
            EliminarComprobante
        Case Is = vbKeyF4
            dgImputacion.SetFocus
        Case Is = vbKeyF2
            dgRetencion.SetFocus
    End Select
    
End Sub

Private Sub dgImputacion_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
        Case Is = vbKeyF12
            dgListado.SetFocus
        'Case Is = vbKeyF6
        '    AgregarImputacion
        Case Is = vbKeyF8
            EliminarImputacion
    End Select
    
End Sub

Private Sub dgRetencion_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
        Case Is = vbKeyF12
            dgListado.SetFocus
        Case Is = vbKeyF6
            Dim i As Integer
            i = dgListado.Row
            With CargaRetencion
                .Show
                .txtComprobante.Text = Left(dgListado.TextMatrix(i, 0), 5)
                .txtEjercicio.Text = "20" & Right(dgListado.TextMatrix(i, 0), 2)
            End With
            Unload ListadoImputacion
            i = 0
        Case Is = vbKeyF8
            EliminarRetencion
    End Select
    
End Sub

Private Sub dgListado_RowColChange()
    
    Dim x As Integer
    x = dgListado.Row
    ConfigurardgImputacion
    CargardgImputacion (dgListado.TextMatrix(x, 0))
    ConfigurardgRetencion
    CargardgRetencion (dgListado.TextMatrix(x, 0))
    x = 0
    
End Sub

Private Sub Form_Load()
    
    Call CenterMe(ListadoImputacion, 8000, 7400)

End Sub

Private Sub Form_Unload(Cancel As Integer)

    dgListado.Clear
    dgImputacion.Clear
    dgRetencion.Clear
    
End Sub

Private Sub cmdAgregar_Click()

    CargaGasto.Show
    Unload ListadoImputacion

End Sub

'Sub AgregarImputacion()
'
'    Dim i As Integer
'    i = dgListado.Row
'    If Val(dgListado.TextMatrix(i, 6)) = Val(txtTotalImputacion.Text) Then
'        MsgBox "El Comprobante en cuestión ya no posee Saldo Diponible para Ingresar otra Imputacion"
'    Else
'        With CargaImputacion
'            .Show
'            .txtComprobante.Text = Left(dgListado.TextMatrix(i, 0), 4)
'            .txtEjercicio.Text = "20" & Right(dgListado.TextMatrix(i, 0), 2)
'            .txtFecha.Text = dgListado.TextMatrix(i, 1)
'            .txtSaldoDisponible.Text = Val(txtTotalImputacion.Text) - Val(dgListado.TextMatrix(i, 6))
'        End With
'        Unload ListadoImputacion
'    End If
'    i = 0
'
'End Sub

Private Sub txtFecha_KeyDown(KeyCode As Integer, Shift As Integer)

    Dim x As String

    If KeyCode = 13 Then
        Select Case Len(Me.txtFecha.Text)
        Case Is = 4
            If IsNumeric(Me.txtFecha.Text) = True Then
                ConfigurardgListadoGasto
                Call CargardgListadoGasto(, Me.txtFecha.Text)
            End If
        Case Is = 7
            If IsNumeric(Left(Me.txtFecha.Text, 2)) = True _
            And IsNumeric(Right(Me.txtFecha.Text, 4)) = True Then
                ConfigurardgListadoGasto
                Call CargardgListadoGasto(, Me.txtFecha.Text)
            End If
        Case Is = 10
            If IsNumeric(Left(Me.txtFecha.Text, 2)) = True _
            And IsNumeric(Mid(Me.txtFecha.Text, 4, 2)) = True _
            And IsNumeric(Right(Me.txtFecha.Text, 4)) = True Then
                ConfigurardgListadoGasto
                Call CargardgListadoGasto(, Me.txtFecha.Text)
            End If
        End Select
        With ListadoImputacion
            x = .dgListado.Row
            ConfigurardgImputacion
            CargardgImputacion (.dgListado.TextMatrix(x, 0))
            ConfigurardgRetencion
            CargardgRetencion (.dgListado.TextMatrix(x, 0))
            x = 0
            .dgListado.SetFocus
        End With
    End If

End Sub

Private Sub txtComprobante_KeyDown(KeyCode As Integer, Shift As Integer)

    Dim x As String

    If KeyCode = 13 Then
        ConfigurardgListadoGasto
        Call CargardgListadoGasto(Me.txtComprobante.Text)
        With ListadoImputacion
            x = .dgListado.Row
            ConfigurardgImputacion
            CargardgImputacion (.dgListado.TextMatrix(x, 0))
            ConfigurardgRetencion
            CargardgRetencion (.dgListado.TextMatrix(x, 0))
            x = 0
            .dgListado.SetFocus
        End With
    End If

End Sub
