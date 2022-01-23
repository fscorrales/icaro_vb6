VERSION 5.00
Begin VB.Form CargaObra 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Carga Obra"
   ClientHeight    =   7410
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7815
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7410
   ScaleWidth      =   7815
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdAgregar 
      BackColor       =   &H008080FF&
      Caption         =   "Guardar Obra"
      Height          =   495
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   6840
      Width           =   1695
   End
   Begin VB.Frame Frame3 
      Caption         =   "Contrato"
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
      Height          =   3015
      Left            =   120
      TabIndex        =   27
      Top             =   3720
      Width           =   7575
      Begin VB.TextBox txtNormaLegal 
         Height          =   285
         Left            =   4800
         TabIndex        =   35
         Top             =   2520
         Width           =   2655
      End
      Begin VB.ComboBox txtLocalidad 
         Height          =   315
         Left            =   1680
         TabIndex        =   34
         Top             =   2520
         Width           =   1575
      End
      Begin VB.TextBox txtMontoContrato 
         Height          =   285
         Left            =   1680
         TabIndex        =   12
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox txtAdicional 
         Height          =   285
         Left            =   4800
         TabIndex        =   13
         Top             =   360
         Width           =   2655
      End
      Begin VB.TextBox txtObra 
         Height          =   525
         Left            =   1680
         MultiLine       =   -1  'True
         TabIndex        =   14
         Top             =   840
         Width           =   5775
      End
      Begin VB.ComboBox txtCuenta 
         Height          =   315
         Left            =   1680
         TabIndex        =   15
         Top             =   1560
         Width           =   1575
      End
      Begin VB.ComboBox txtFuente 
         Height          =   315
         Left            =   1680
         TabIndex        =   17
         Top             =   2040
         Width           =   1575
      End
      Begin VB.TextBox txtDescripcionCuenta 
         Height          =   285
         Left            =   3360
         TabIndex        =   16
         Top             =   1560
         Width           =   4095
      End
      Begin VB.TextBox txtDescripcionFuente 
         Height          =   285
         Left            =   3360
         TabIndex        =   18
         Top             =   2040
         Width           =   4095
      End
      Begin VB.Label Label13 
         Caption         =   "Norma Legal"
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
         Left            =   3360
         TabIndex        =   37
         Top             =   2520
         Width           =   1455
      End
      Begin VB.Label Label12 
         Caption         =   "Localidad"
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
         Left            =   240
         TabIndex        =   33
         Top             =   2520
         Width           =   1455
      End
      Begin VB.Label Label11 
         Caption         =   "Adicional"
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
         Left            =   3360
         TabIndex        =   32
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label8 
         Caption         =   "Monto Contrato"
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
         Left            =   240
         TabIndex        =   31
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
         Left            =   240
         TabIndex        =   30
         Top             =   840
         Width           =   1455
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
         Left            =   240
         TabIndex        =   29
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
         Left            =   240
         TabIndex        =   28
         Top             =   2040
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Imputacion"
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
      TabIndex        =   21
      Top             =   960
      Width           =   7575
      Begin VB.TextBox txtDescripcionPartida 
         Height          =   285
         Left            =   2280
         TabIndex        =   11
         Top             =   2280
         Width           =   5175
      End
      Begin VB.TextBox txtDescripcionActividad 
         Height          =   285
         Left            =   2280
         TabIndex        =   9
         Top             =   1800
         Width           =   5175
      End
      Begin VB.TextBox txtDescripcionProyecto 
         Height          =   285
         Left            =   2280
         TabIndex        =   7
         Top             =   1320
         Width           =   5175
      End
      Begin VB.TextBox txtDescripcionSubprograma 
         Height          =   285
         Left            =   2280
         TabIndex        =   5
         Top             =   840
         Width           =   5175
      End
      Begin VB.TextBox txtDescripcionPrograma 
         Height          =   285
         Left            =   2280
         TabIndex        =   3
         Top             =   360
         Width           =   5175
      End
      Begin VB.TextBox txtPartida 
         Height          =   285
         Left            =   1680
         TabIndex        =   10
         Top             =   2280
         Width           =   495
      End
      Begin VB.TextBox txtActividad 
         Height          =   285
         Left            =   1680
         TabIndex        =   8
         Top             =   1800
         Width           =   495
      End
      Begin VB.TextBox txtProyecto 
         Height          =   285
         Left            =   1680
         TabIndex        =   6
         Top             =   1320
         Width           =   495
      End
      Begin VB.TextBox txtSubprograma 
         Height          =   285
         Left            =   1680
         TabIndex        =   4
         Top             =   840
         Width           =   495
      End
      Begin VB.TextBox txtPrograma 
         Height          =   285
         Left            =   1680
         TabIndex        =   2
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label7 
         Caption         =   "Partida"
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
         Left            =   240
         TabIndex        =   26
         Top             =   2280
         Width           =   1455
      End
      Begin VB.Label Label6 
         Caption         =   "Actividad"
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
         Left            =   240
         TabIndex        =   25
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "Proyecto"
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
         Left            =   240
         TabIndex        =   24
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Subprograma"
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
         Left            =   240
         TabIndex        =   23
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Programa"
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
         Left            =   240
         TabIndex        =   22
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Proveedor"
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
      Height          =   735
      Left            =   120
      TabIndex        =   19
      Top             =   120
      Width           =   7575
      Begin VB.ComboBox txtCUIT 
         Height          =   315
         Left            =   1680
         TabIndex        =   0
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox txtDescripcionCUIT 
         Height          =   285
         Left            =   3360
         TabIndex        =   1
         Top             =   240
         Width           =   4095
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
         Left            =   240
         TabIndex        =   20
         Top             =   240
         Width           =   1455
      End
   End
End
Attribute VB_Name = "CargaObra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAgregar_Click()
    
    GenerarObra
    
End Sub

Private Sub Form_Load()

    Call CenterMe(CargaObra, 7900, 7770)
    CargarCombosObra
    BloquearTextBoxObra
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    strEditandoObra = ""
    bolActualizarEstructura = False
    
    If bolObraDesdeCarga = True Then
        CargaGasto.Show
        bolObraDesdeCarga = False
        CargaGasto.txtObra.SetFocus
    End If
    
    If bolObraDesdeEPAMConsolidado = True Then
        bolObraDesdeEPAMConsolidado = False
        ImputacionEPAM.Show
        ImputacionEPAM.txtObra.SetFocus
    End If
    
    If bolAutocargaCertificado = True Then
        VaciarVariablesAutocarga
        AutocargaCertificados.Show
    End If
    
    If bolAutocargaEPAM = True Then
        VaciarVariablesAutocarga
        AutocargaEPAM.Show
    End If

End Sub

Private Sub txtPrograma_GotFocus()

    txtSubprograma.Enabled = False
    txtProyecto.Enabled = False
    txtActividad.Enabled = False
    txtPartida.Enabled = False
    txtDescripcionPrograma.Enabled = False
    txtDescripcionSubprograma.Enabled = False
    txtDescripcionProyecto.Enabled = False
    txtDescripcionActividad.Enabled = False
    txtDescripcionPartida.Enabled = False
    
End Sub

Private Sub txtPrograma_LostFocus()

    Dim SQL As String

    If Me.txtDescripcionCUIT.Text <> "" Then
        If ValidarProgramaCargaObra(Me.txtPrograma.Text) Then
            Set rstBuscarIcaro = New ADODB.Recordset
            SQL = "Select * From PROGRAMAS Where PROGRAMA = '" & Me.txtPrograma.Text & "'"
            rstBuscarIcaro.Open SQL, dbIcaro, adOpenDynamic, adLockOptimistic
            Me.txtDescripcionPrograma.Text = rstBuscarIcaro!Descripcion
            rstBuscarIcaro.Close
            Set rstBuscarIcaro = Nothing
            Me.txtSubprograma.Enabled = True
            Me.txtSubprograma.SetFocus
        End If
    End If
        
End Sub

Private Sub txtDescripcionPrograma_LostFocus()

    If bolActualizarEstructura = True Then
        Dim Agregar As Integer
        Agregar = MsgBox("Desea realmente agregar el PROGRAMA " & txtPrograma.Text & " con el siguiente nombre " & txtDescripcionPrograma.Text & "?", vbQuestion + vbYesNo, "ACTUALIZAR ESTRUCTURA")
        If Agregar = 6 Then
            Set rstRegistroIcaro = New ADODB.Recordset
            With rstRegistroIcaro
                .Open "PROGRAMAS", dbIcaro, adOpenForwardOnly, adLockOptimistic
                .AddNew
                .Fields("Programa") = Me.txtPrograma.Text
                If Len(Me.txtDescripcionPrograma.Text) <> 0 Then
                    .Fields("Descripcion") = Me.txtDescripcionPrograma.Text
                Else
                    .Fields("Descripcion") = ""
                End If
                .Update
            End With
            rstRegistroIcaro.Close
            Set rstRegistroIcaro = Nothing
            txtDescripcionPrograma.Enabled = False
            bolActualizarEstructura = False
            txtSubprograma.Enabled = True
            txtSubprograma.SetFocus
        Else
            txtDescripcionPrograma.Text = ""
            txtPrograma.SetFocus
        End If
    End If
    
End Sub

Private Sub txtSubprograma_GotFocus()
    
    txtProyecto.Enabled = False
    txtActividad.Enabled = False
    txtPartida.Enabled = False
    txtDescripcionPrograma.Enabled = False
    txtDescripcionSubprograma.Enabled = False
    txtDescripcionProyecto.Enabled = False
    txtDescripcionActividad.Enabled = False
    txtDescripcionPartida.Enabled = False
    
End Sub

Private Sub txtSubprograma_LostFocus()

    Dim SQL As String

    If ValidarSubprogramaCargaObra(Me.txtSubprograma.Text) Then
        Set rstBuscarIcaro = New ADODB.Recordset
        SQL = Me.txtPrograma.Text & "-" & Me.txtSubprograma.Text
        SQL = "Select * From SUBPROGRAMAS Where SUBPROGRAMA = '" & SQL & "'"
        rstBuscarIcaro.Open SQL, dbIcaro, adOpenDynamic, adLockOptimistic
        If IsNull(rstBuscarIcaro!Descripcion) = False Then
            Me.txtDescripcionSubprograma.Text = rstBuscarIcaro!Descripcion
        Else
            Me.txtDescripcionSubprograma.Text = ""
        End If
        rstBuscarIcaro.Close
        Set rstBuscarIcaro = Nothing
        Me.txtProyecto.Enabled = True
        Me.txtProyecto.SetFocus
    End If

End Sub

Private Sub txtDescripcionSubprograma_LostFocus()

    If bolActualizarEstructura = True Then
        Dim Agregar As Integer
        Agregar = MsgBox("Desea realmente agregar el SUBPROGRAMA " & txtSubprograma.Text & " con el siguiente nombre " & txtDescripcionSubprograma.Text & "?", vbQuestion + vbYesNo, "ACTUALIZAR ESTRUCTURA")
        If Agregar = 6 Then
            Set rstRegistroIcaro = New ADODB.Recordset
            With rstRegistroIcaro
                .Open "SUBPROGRAMAS", dbIcaro, adOpenForwardOnly, adLockOptimistic
                .AddNew
                .Fields("Programa") = txtPrograma.Text
                .Fields("Subprograma") = txtPrograma.Text & "-" & txtSubprograma.Text
                If Len(txtDescripcionSubprograma.Text) <> 0 Then
                    .Fields("Descripcion") = txtDescripcionSubprograma.Text
                Else
                    .Fields("Descripcion") = ""
                End If
                .Update
            End With
            txtDescripcionSubprograma.Enabled = False
            bolActualizarEstructura = False
            txtProyecto.Enabled = True
            txtProyecto.SetFocus
            rstRegistroIcaro.Close
            Set rstRegistroIcaro = Nothing
        Else
            txtDescripcionSubprograma.Text = ""
            txtSubprograma.SetFocus
        End If
    End If
    
End Sub

Private Sub txtProyecto_GotFocus()

    txtActividad.Enabled = False
    txtPartida.Enabled = False
    txtDescripcionPrograma.Enabled = False
    txtDescripcionSubprograma.Enabled = False
    txtDescripcionProyecto.Enabled = False
    txtDescripcionActividad.Enabled = False
    txtDescripcionPartida.Enabled = False
    
End Sub

Private Sub txtProyecto_LostFocus()
   
    Dim SQL As String
    
    If ValidarProyectoCargaObra(Me.txtProyecto.Text) Then
        Set rstBuscarIcaro = New ADODB.Recordset
        SQL = Me.txtPrograma.Text & "-" & Me.txtSubprograma.Text & "-" & Me.txtProyecto.Text
        SQL = "Select * From PROYECTOS Where PROYECTO = '" & SQL & "'"
        rstBuscarIcaro.Open SQL, dbIcaro, adOpenDynamic, adLockOptimistic
        If IsNull(rstBuscarIcaro!Descripcion) = False Then
            Me.txtDescripcionProyecto.Text = rstBuscarIcaro!Descripcion
        Else
            Me.txtDescripcionProyecto.Text = ""
        End If
        rstBuscarIcaro.Close
        Set rstBuscarIcaro = Nothing
        Me.txtActividad.Enabled = True
        Me.txtActividad.SetFocus
    End If
   
End Sub

Private Sub txtDescripcionProyecto_LostFocus()

    If bolActualizarEstructura = True Then
        Dim Agregar As Integer
        Agregar = MsgBox("Desea realmente agregar el PROYECTO " & txtProyecto.Text & " con el siguiente nombre " & txtDescripcionProyecto.Text & "?", vbQuestion + vbYesNo, "ACTUALIZAR ESTRUCTURA")
        If Agregar = 6 Then
            Set rstRegistroIcaro = New ADODB.Recordset
            With rstRegistroIcaro
                .Open "PROYECTOS", dbIcaro, adOpenForwardOnly, adLockOptimistic
                .AddNew
                .Fields("Subprograma") = txtPrograma.Text & "-" & txtSubprograma.Text
                .Fields("Proyecto") = txtPrograma.Text & "-" & txtSubprograma.Text & "-" & txtProyecto.Text
                If Len(txtDescripcionProyecto.Text) <> 0 Then
                    .Fields("Descripcion") = txtDescripcionProyecto.Text
                Else
                    .Fields("Descripcion") = ""
                End If
                .Update
            End With
            txtDescripcionProyecto.Enabled = False
            bolActualizarEstructura = False
            txtActividad.Enabled = True
            txtActividad.SetFocus
            rstRegistroIcaro.Close
            Set rstRegistroIcaro = Nothing
        Else
            txtDescripcionProyecto.Text = ""
            txtProyecto.SetFocus
        End If
    End If
    
End Sub

Private Sub txtActividad_GotFocus()

    txtPartida.Enabled = False
    txtDescripcionPrograma.Enabled = False
    txtDescripcionSubprograma.Enabled = False
    txtDescripcionProyecto.Enabled = False
    txtDescripcionActividad.Enabled = False
    txtDescripcionPartida.Enabled = False
    
End Sub

Private Sub txtActividad_LostFocus()
   
    Dim SQL As String
    
    If ValidarActividadCargaObra(Me.txtActividad.Text) Then
        Set rstBuscarIcaro = New ADODB.Recordset
        SQL = Me.txtPrograma.Text & "-" & Me.txtSubprograma.Text & "-" & Me.txtProyecto.Text & "-" & Me.txtActividad.Text
        SQL = "Select * From ACTIVIDADES Where ACTIVIDAD = '" & SQL & "'"
        rstBuscarIcaro.Open SQL, dbIcaro, adOpenDynamic, adLockOptimistic
        Me.txtDescripcionActividad.Text = rstBuscarIcaro!Descripcion
        rstBuscarIcaro.Close
        Set rstBuscarIcaro = Nothing
        txtPartida.Enabled = True
        txtPartida.SetFocus
    End If
   
End Sub

Private Sub txtDescripcionActividad_LostFocus()

    If bolActualizarEstructura = True Then
        Dim Agregar As Integer
        Agregar = MsgBox("Desea realmente agregar la ACTIVIDAD " & txtActividad.Text & " con el siguiente nombre " & txtDescripcionActividad.Text & "?", vbQuestion + vbYesNo, "ACTUALIZAR ESTRUCTURA")
            Set rstRegistroIcaro = New ADODB.Recordset
            With rstRegistroIcaro
                .Open "ACTIVIDADES", dbIcaro, adOpenForwardOnly, adLockOptimistic
                .AddNew
                .Fields("Proyecto") = txtPrograma.Text & "-" & txtSubprograma.Text & "-" & txtProyecto.Text
                .Fields("Actividad") = txtPrograma.Text & "-" & txtSubprograma.Text & "-" & txtProyecto.Text & "-" & txtActividad.Text
                If Len(txtDescripcionActividad.Text) <> 0 Then
                    .Fields("Descripcion") = txtDescripcionActividad.Text
                Else
                    .Fields("Descripcion") = ""
                End If
                .Update
            End With
            txtDescripcionActividad.Enabled = False
            bolActualizarEstructura = False
            txtPartida.Enabled = True
            txtPartida.SetFocus
            rstRegistroIcaro.Close
            Set rstRegistroIcaro = Nothing
    Else
        txtDescripcionActividad.Text = ""
        txtActividad.SetFocus
    End If
    
End Sub

Private Sub txtPartida_LostFocus()

    If ValidarPartidaCargaObra(Me.txtPartida.Text) Then
        Set rstBuscarIcaro = New ADODB.Recordset
        SQL = "Select * From PARTIDAS Where PARTIDA = '" & Me.txtPartida.Text & "'"
        rstBuscarIcaro.Open SQL, dbIcaro, adOpenDynamic, adLockOptimistic
        Me.txtDescripcionPartida.Text = rstBuscarIcaro!Descripcion
        rstBuscarIcaro.Close
        Set rstBuscarIcaro = Nothing
        txtMontoContrato.SetFocus
    End If
    
End Sub

Private Sub txtCUIT_LostFocus()
    
    Dim SQL As String

    If ValidarCUITCargaObra(Me.txtCUIT.Text) Then
        Set rstBuscarIcaro = New ADODB.Recordset
        SQL = "Select * From CUIT Where CUIT = '" & Me.txtCUIT.Text & "'"
        rstBuscarIcaro.Open SQL, dbIcaro, adOpenDynamic, adLockOptimistic
        Me.txtDescripcionCUIT.Text = rstBuscarIcaro!Nombre
        rstBuscarIcaro.Close
        Set rstBuscarIcaro = Nothing
    End If
    
End Sub

Private Sub txtCuenta_LostFocus()

    Dim SQL As String

    If ValidarCuentaCargaObra(Me.txtCuenta.Text) Then
        Set rstBuscarIcaro = New ADODB.Recordset
        SQL = "Select * From CUENTAS Where CUENTA = '" & Me.txtCuenta.Text & "'"
        rstBuscarIcaro.Open SQL, dbIcaro, adOpenDynamic, adLockOptimistic
        Me.txtDescripcionCuenta.Text = rstBuscarIcaro!Descripcion
        rstBuscarIcaro.Close
        Set rstBuscarIcaro = Nothing
        Me.txtFuente.Enabled = True
        Me.txtFuente.SetFocus
    End If
    
End Sub

Private Sub txtFuente_LostFocus()
    
    Dim SQL As String

    If ValidarFuenteCargaObra(Me.txtFuente.Text) Then
        Set rstBuscarIcaro = New ADODB.Recordset
        SQL = "Select * From FUENTES Where FUENTE = '" & Me.txtFuente.Text & "'"
        rstBuscarIcaro.Open SQL, dbIcaro, adOpenDynamic, adLockOptimistic
        Me.txtDescripcionFuente.Text = rstBuscarIcaro!Descripcion
        rstBuscarIcaro.Close
        Set rstBuscarIcaro = Nothing
    End If

End Sub

