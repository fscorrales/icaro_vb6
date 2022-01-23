VERSION 5.00
Begin VB.Form CargaGasto 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Carga Gasto"
   ClientHeight    =   5985
   ClientLeft      =   4755
   ClientTop       =   285
   ClientWidth     =   6885
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5985
   ScaleWidth      =   6885
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdAgregar 
      BackColor       =   &H008080FF&
      Caption         =   "Guardar Datos"
      Height          =   495
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5400
      Width           =   1815
   End
   Begin VB.Frame Frame3 
      Caption         =   "Datos Numéricos del Comprobante"
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
      Height          =   2655
      Left            =   120
      TabIndex        =   20
      Top             =   2640
      Width           =   6615
      Begin VB.TextBox txtCertificado 
         Height          =   285
         Left            =   4920
         TabIndex        =   8
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox txtImporte 
         Height          =   285
         Left            =   1560
         TabIndex        =   6
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox txtFR 
         Height          =   285
         Left            =   1560
         TabIndex        =   7
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox txtAvance 
         Height          =   285
         Left            =   4920
         TabIndex        =   9
         Top             =   960
         Width           =   1575
      End
      Begin VB.ComboBox txtCuenta 
         Height          =   315
         Left            =   1560
         TabIndex        =   10
         Top             =   1560
         Width           =   1575
      End
      Begin VB.ComboBox txtFuente 
         Height          =   315
         Left            =   1560
         TabIndex        =   12
         Top             =   2160
         Width           =   1575
      End
      Begin VB.TextBox txtDescripcionCuenta 
         Height          =   285
         Left            =   3240
         TabIndex        =   11
         Top             =   1560
         Width           =   3255
      End
      Begin VB.TextBox txtDescripcionFuente 
         Height          =   285
         Left            =   3240
         TabIndex        =   13
         Top             =   2160
         Width           =   3255
      End
      Begin VB.Label Label5 
         Caption         =   "Importe"
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
         TabIndex        =   26
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label6 
         Caption         =   "Fondo Reparo"
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
         TabIndex        =   25
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label7 
         Caption         =   "Certificado"
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
         TabIndex        =   24
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label8 
         Caption         =   "Avance"
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
         TabIndex        =   23
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label9 
         Caption         =   "Cuenta"
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
         TabIndex        =   22
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label Label10 
         Caption         =   "Fuente"
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
         TabIndex        =   21
         Top             =   2160
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Proveedor y Obra Destino"
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
      Height          =   1335
      Left            =   120
      TabIndex        =   17
      Top             =   1200
      Width           =   6615
      Begin VB.ComboBox txtCUIT 
         Height          =   315
         Left            =   1560
         TabIndex        =   3
         Top             =   360
         Width           =   1575
      End
      Begin VB.ComboBox txtObra 
         Height          =   315
         Left            =   1560
         TabIndex        =   5
         Top             =   960
         Width           =   4935
      End
      Begin VB.TextBox txtDescripcionCUIT 
         Height          =   285
         Left            =   3240
         TabIndex        =   4
         Top             =   360
         Width           =   3255
      End
      Begin VB.Label Label2 
         Caption         =   "CUIT"
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
         TabIndex        =   19
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Obra"
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
         TabIndex        =   18
         Top             =   960
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Número de Comprobante y Ejercicio de Imputación"
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
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6615
      Begin VB.TextBox txtComprobante 
         Height          =   285
         Left            =   1560
         TabIndex        =   1
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Left            =   4920
         TabIndex        =   2
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Comprobante"
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
         TabIndex        =   16
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha (dd/mm/aaaa)"
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
         Height          =   495
         Left            =   3480
         TabIndex        =   15
         Top             =   360
         Width           =   1455
      End
   End
End
Attribute VB_Name = "CargaGasto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAgregar_Click()

    GenerarComprobante
    
End Sub

Private Sub Form_Load()

    Call CenterMe(CargaGasto, 7000, 6400)
    CargarCombosGasto
    BloquearTextBoxGasto
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    If bolAutocargaCertificado = True Then
        VaciarVariablesAutocarga
        AutocargaCertificados.Show
        ConfigurardgAutocargaCertificados
        CargardgAutocargaCertificados
    ElseIf bolAutocargaEPAM = True Then
        VaciarVariablesAutocarga
        AutocargaEPAM.Show
        ConfigurardgAutocargaEPAM
        CargardgAutocargaEPAM
    End If
    strEditandoComprobante = ""
    
End Sub

Private Sub txtCUIT_LostFocus()

    Dim SQL As String
    
    If ValidarCUITCargaGasto(txtCUIT.Text) Then
        Set rstBuscarIcaro = New ADODB.Recordset
        SQL = "Select * From CUIT Where CUIT = '" & txtCUIT.Text & "'"
        rstBuscarIcaro.Open SQL, dbIcaro, adOpenDynamic, adLockOptimistic
        txtDescripcionCUIT.Text = rstBuscarIcaro!Nombre
        rstBuscarIcaro.Close
        Set rstBuscarIcaro = Nothing
        txtObra.Clear
        txtFuente.Text = ""
        txtCuenta.Text = ""
        txtDescripcionFuente.Text = ""
        txtDescripcionCuenta.Text = ""
        txtObra.Enabled = True
        txtFuente.Enabled = False
        txtCuenta.Enabled = False
        Call CargarComboObraDeGasto(txtCUIT.Text)
        txtObra.SetFocus
    End If
    
End Sub

Private Sub txtObra_LostFocus()
    
    Dim SQL As String
    
    If ValidarObraCargaGasto(txtObra.Text) Then
        Set rstBuscarIcaro = New ADODB.Recordset
        SQL = "Select * From OBRAS Where DESCRIPCION = '" & txtObra.Text & "'"
        rstBuscarIcaro.Open SQL, dbIcaro, adOpenDynamic, adLockOptimistic
        txtCuenta.Text = rstBuscarIcaro!Cuenta
        txtFuente.Text = rstBuscarIcaro!Fuente
        rstBuscarIcaro.Close
        SQL = "Select * From CUENTAS Where CUENTA = '" & txtCuenta.Text & "'"
        rstBuscarIcaro.Open SQL, dbIcaro, adOpenDynamic, adLockOptimistic
        txtDescripcionCuenta.Text = rstBuscarIcaro!Descripcion
        rstBuscarIcaro.Close
        SQL = "Select * From FUENTES Where FUENTE = '" & txtFuente.Text & "'"
        rstBuscarIcaro.Open SQL, dbIcaro, adOpenDynamic, adLockOptimistic
        txtDescripcionFuente.Text = rstBuscarIcaro!Descripcion
        rstBuscarIcaro.Close
        Set rstBuscarIcaro = Nothing
        txtFuente.Enabled = True
        txtCuenta.Enabled = True
    End If
    
End Sub

Private Sub txtCuenta_LostFocus()
    
    Dim SQL As String
    
    If ValidarCuentaCargaGasto(txtCuenta.Text) Then
        Set rstBuscarIcaro = New ADODB.Recordset
        SQL = "Select * From CUENTAS Where CUENTA = '" & txtCuenta.Text & "'"
        rstBuscarIcaro.Open SQL, dbIcaro, adOpenDynamic, adLockOptimistic
        txtDescripcionCuenta.Text = rstBuscarIcaro!Descripcion
        rstBuscarIcaro.Close
        Set rstBuscarIcaro = Nothing
    End If
    
End Sub

Private Sub txtFuente_LostFocus()
    
    Dim SQL As String
    
    If ValidarFuenteCargaGasto(txtFuente.Text) Then
        Set rstBuscarIcaro = New ADODB.Recordset
        SQL = "Select * From FUENTES Where FUENTE = '" & txtFuente.Text & "'"
        rstBuscarIcaro.Open SQL, dbIcaro, adOpenDynamic, adLockOptimistic
        txtDescripcionFuente.Text = rstBuscarIcaro!Descripcion
        rstBuscarIcaro.Close
        Set rstBuscarIcaro = Nothing
    End If
    
End Sub

