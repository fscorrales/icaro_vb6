VERSION 5.00
Begin VB.Form CargaImputacion 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Carga Imputación"
   ClientHeight    =   5595
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7830
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5595
   ScaleWidth      =   7830
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdAgregar 
      BackColor       =   &H008080FF&
      Caption         =   "Guardar"
      Height          =   500
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5040
      Width           =   1700
   End
   Begin VB.Frame Frame3 
      Caption         =   "Datos del Comprobante"
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
      Height          =   1215
      Left            =   120
      TabIndex        =   22
      Top             =   120
      Width           =   7575
      Begin VB.TextBox txtFecha 
         Height          =   285
         Left            =   1680
         TabIndex        =   28
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox txtSaldoDisponible 
         Height          =   285
         Left            =   5880
         TabIndex        =   27
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox txtEjercicio 
         Height          =   285
         Left            =   1680
         TabIndex        =   24
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox txtComprobante 
         Height          =   285
         Left            =   5880
         TabIndex        =   23
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label11 
         Caption         =   "Fecha"
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
      Begin VB.Label Label10 
         Caption         =   "Saldo Disponible"
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
         Left            =   4440
         TabIndex        =   29
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label9 
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
         Left            =   240
         TabIndex        =   26
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label8 
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
         Left            =   4440
         TabIndex        =   25
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Importe a Consignar"
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
      TabIndex        =   18
      Top             =   4200
      Width           =   7575
      Begin VB.TextBox txtSaldoRestante 
         Height          =   285
         Left            =   5880
         TabIndex        =   20
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox txtMonto 
         Height          =   285
         Left            =   1680
         TabIndex        =   10
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "Saldo Restante"
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
         Left            =   4440
         TabIndex        =   21
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Monto"
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
         TabIndex        =   19
         Top             =   360
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
      TabIndex        =   12
      Top             =   1440
      Width           =   7575
      Begin VB.TextBox txtPrograma 
         Height          =   285
         Left            =   1680
         TabIndex        =   0
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox txtSubprograma 
         Height          =   285
         Left            =   1680
         TabIndex        =   2
         Top             =   840
         Width           =   495
      End
      Begin VB.TextBox txtProyecto 
         Height          =   285
         Left            =   1680
         TabIndex        =   4
         Top             =   1320
         Width           =   495
      End
      Begin VB.TextBox txtActividad 
         Height          =   285
         Left            =   1680
         TabIndex        =   6
         Top             =   1800
         Width           =   495
      End
      Begin VB.TextBox txtPartida 
         Height          =   285
         Left            =   1680
         TabIndex        =   8
         Top             =   2280
         Width           =   495
      End
      Begin VB.TextBox txtDescripcionPrograma 
         Height          =   285
         Left            =   2280
         TabIndex        =   1
         Top             =   360
         Width           =   5175
      End
      Begin VB.TextBox txtDescripcionSubprograma 
         Height          =   285
         Left            =   2280
         TabIndex        =   3
         Top             =   840
         Width           =   5175
      End
      Begin VB.TextBox txtDescripcionProyecto 
         Height          =   285
         Left            =   2280
         TabIndex        =   5
         Top             =   1320
         Width           =   5175
      End
      Begin VB.TextBox txtDescripcionActividad 
         Height          =   285
         Left            =   2280
         TabIndex        =   7
         Top             =   1800
         Width           =   5175
      End
      Begin VB.TextBox txtDescripcionPartida 
         Height          =   285
         Left            =   2280
         TabIndex        =   9
         Top             =   2280
         Width           =   5175
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
         TabIndex        =   17
         Top             =   360
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
         TabIndex        =   16
         Top             =   840
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
         TabIndex        =   15
         Top             =   1320
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
         TabIndex        =   14
         Top             =   1800
         Width           =   1455
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
         TabIndex        =   13
         Top             =   2280
         Width           =   1455
      End
   End
End
Attribute VB_Name = "CargaImputacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bolActualizarEstructura As Boolean

Private Sub Form_Load()
    Call CenterMe(CargaImputacion, 6000, 7950)

    txtComprobante.Enabled = False
    txtEjercicio.Enabled = False
    txtSaldoDisponible.Enabled = False
    txtFecha.Enabled = False
    txtSubprograma.Enabled = False
    txtProyecto.Enabled = False
    txtActividad.Enabled = False
    txtPartida.Enabled = False
    txtDescripcionPrograma.Enabled = False
    txtDescripcionSubprograma.Enabled = False
    txtDescripcionProyecto.Enabled = False
    txtDescripcionActividad.Enabled = False
    txtDescripcionPartida.Enabled = False
    txtSaldoRestante.Enabled = False
End Sub

Private Sub cmdAgregar_Click()
    If Validar = True Then
        'Call SetRecordset(rstCargaImputacion, "IMPUTACION")
        With rstCargaImputacion
            .AddNew
            .Fields("Comprobante") = Format(txtComprobante.Text, "0000") & "/" & Right(txtEjercicio.Text, 2)
            .Fields("Imputacion") = txtPrograma.Text & "-" & txtSubprograma.Text & "-" & txtProyecto.Text & "-" & txtActividad.Text
            .Fields("Partida") = txtPartida.Text
            .Fields("Importe") = txtMonto.Text
            .Update
        End With
        MsgBox "Los datos fueron agregados"
        Set rstCargaImputacion = Nothing
        strComprobante = Format(txtComprobante.Text, "0000") & "/" & Right(txtEjercicio.Text, 2)
        ListadoImputacion.Show
        Unload CargaImputacion
    End If
End Sub

Function Validar() As Boolean
    If Trim(txtMonto.Text) = "" Or IsNumeric(txtMonto.Text) = False Then
        MsgBox "Debe ingresar un Importe número superior a 0", vbCritical + vbOKOnly, "IMPORTE INCORRECTO"
        txtMonto.SetFocus
        Validar = False
        Exit Function
    End If
    If Val(txtSaldoRestante.Text) < "0" Then
        MsgBox "El Importe Ingresado supera el Saldo Disponible del Comprobante", vbCritical + vbOKOnly, "IMPORTE INCORRECTO"
        txtMonto.SetFocus
        Validar = False
        Exit Function
    End If
    
    'Call SetRecordset(rstEstructura, "SELECT PROGRAMAS.Programa, PROGRAMAS.Descripcion As DescripcionPrograma, SUBPROGRAMAS.Subprograma, SUBPROGRAMAS.Descripcion As DescripcionSubprograma, PROYECTOS.Proyecto, PROYECTOS.Descripcion As DescripcionProyecto, ACTIVIDADES.Actividad, ACTIVIDADES.Descripcion As DescripcionActividad FROM ((PROGRAMAS INNER JOIN SUBPROGRAMAS ON PROGRAMAS.Programa = SUBPROGRAMAS.Programa) INNER JOIN PROYECTOS ON SUBPROGRAMAS.Subprograma = PROYECTOS.Subprograma) INNER JOIN ACTIVIDADES ON PROYECTOS.Proyecto = ACTIVIDADES.Proyecto")
    If Trim(txtPrograma.Text) = "" Or IsNumeric(txtPrograma.Text) = False Then
        MsgBox "Debe ingresar un Nro de PROGRAMA", vbCritical + vbOKOnly, "NRO PROGRAMA INCORRECTO"
        txtPrograma.SetFocus
        Validar = False
        Exit Function
    End If
    With rstEstructura
        .FindFirst "Programa =" & "'" & txtPrograma.Text & "'"
        If .NoMatch = True Then
            MsgBox "Debe ingresar un Nro de PROGRAMA", vbCritical + vbOKOnly, "NRO PROGRAMA INEXISTENTE"
            txtPrograma.SetFocus
            Validar = False
            Exit Function
        End If
    End With
    If Trim(txtSubprograma.Text) = "" Or IsNumeric(txtSubprograma.Text) = False Then
        MsgBox "Debe ingresar un Nro de SUBPROGRAMA", vbCritical + vbOKOnly, "NRO SUBPROGRAMA INCORRECTO"
        txtSubprograma.SetFocus
        Validar = False
        Exit Function
    End If
    With rstEstructura
        .FindFirst "Subprograma =" & "'" & txtPrograma.Text & "-" & txtSubprograma.Text & "'"
        If .NoMatch = True Then
            MsgBox "Debe ingresar un Nro de SUBPROGRAMA", vbCritical + vbOKOnly, "NRO SUBPROGRAMA INEXISTENTE"
            txtSubprograma.SetFocus
            Validar = False
            Exit Function
        End If
    End With
    If Trim(txtProyecto.Text) = "" Or IsNumeric(txtProyecto.Text) = False Then
        MsgBox "Debe ingresar un Nro de PROYECTO", vbCritical + vbOKOnly, "NRO PROYECTO INCORRECTO"
        txtProyecto.SetFocus
        Validar = False
        Exit Function
    End If
    With rstEstructura
        .FindFirst "Proyecto =" & "'" & txtPrograma.Text & "-" & txtSubprograma.Text & "-" & txtProyecto.Text & "'"
        If .NoMatch = True Then
            MsgBox "Debe ingresar un Nro de PROYECTO", vbCritical + vbOKOnly, "NRO PROYECTO INEXISTENTE"
            txtProyecto.SetFocus
            Validar = False
            Exit Function
        End If
    End With
    If Trim(txtActividad.Text) = "" Or IsNumeric(txtActividad.Text) = False Then
        MsgBox "Debe ingresar un Nro de ACTIVIDAD", vbCritical + vbOKOnly, "NRO ACTIVIDAD INCORRECTO"
        txtActividad.SetFocus
        Validar = False
        Exit Function
    End If
    With rstEstructura
        .FindFirst "Actividad =" & "'" & txtPrograma.Text & "-" & txtSubprograma.Text & "-" & txtProyecto.Text & "-" & txtActividad.Text & "'"
        If .NoMatch = True Then
            MsgBox "Debe ingresar un Nro de ACTIVIDAD", vbCritical + vbOKOnly, "NRO ACTIVIDAD INEXISTENTE"
            txtActividad.SetFocus
            Validar = False
            Exit Function
        End If
    End With
    If Trim(txtPartida.Text) = "" Or IsNumeric(txtPartida.Text) = False Then
        MsgBox "Debe ingresar un Nro de PARTIDA", vbCritical + vbOKOnly, "NRO PARTIDA INCORRECTO"
        txtPartida.SetFocus
        Validar = False
        Exit Function
    End If
    'Call SetRecordset(rstPartida, "PARTIDAS")
    With rstPartida
            .Index = "pkPartida"
            .Seek "=", txtPartida
        If .NoMatch = True Then
            MsgBox "Debe ingresar un Nro de PARTIDA", vbCritical + vbOKOnly, "NRO PARTIDA INEXISTENTE"
            txtPartida.SetFocus
            Validar = False
            Set rstPartida = Nothing
            Exit Function
        End If
    End With
    Set rstPartida = Nothing
    
    Validar = True
End Function

Private Sub Form_Unload(Cancel As Integer)
    Set rstEstructura = Nothing
    Set rstPartida = Nothing
    Set rstPrograma = Nothing
    Set rstSubprograma = Nothing
    Set rstProyecto = Nothing
    Set rstActividad = Nothing
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
    If Len(Trim(txtPrograma.Text)) <> "2" Then
        MsgBox "Debe ingresar un código de 2 cifras para continuar"
        txtPrograma.SetFocus
    Else
        'Call Buscar(txtPrograma.Text, rstPrograma, "PROGRAMAS", "pkPrograma")
        With rstPrograma
            If .NoMatch = False Then
                    If Len(.Fields("Descripcion")) <> 0 Then
                        txtDescripcionPrograma.Text = .Fields("Descripcion")
                    Else
                        txtDescripcionPrograma.Text = ""
                    End If
                txtSubprograma.Enabled = True
                txtSubprograma.SetFocus
            Else
                Dim Agregar As Integer
                Agregar = MsgBox("Desea Agregar el PROGRAMA " & txtPrograma.Text & " ?", vbQuestion + vbYesNo, "ACTUALIZAR ESTRUCTURA")
                If Agregar = 6 Then
                    bolActualizarEstructura = True
                    txtDescripcionPrograma.Enabled = True
                    txtDescripcionPrograma.SetFocus
                    txtDescripcionPrograma.Text = "ESCRIBA AQUÍ EL NOMBRE DEL PROGRAMA"
                Else
                    txtPrograma.SetFocus
                End If
            End If
        End With
        Set rstPrograma = Nothing
    End If
End Sub

Private Sub txtDescripcionPrograma_LostFocus()
    If bolActualizarEstructura = True Then
        Dim Agregar As Integer
        Agregar = MsgBox("Desea realmente agregar el PROGRAMA " & txtPrograma.Text & " con el siguiente nombre " & txtDescripcionPrograma.Text & "?", vbQuestion + vbYesNo, "ACTUALIZAR ESTRUCTURA")
        If Agregar = 6 Then
            'Call SetRecordset(rstPrograma, "PROGRAMAS")
            With rstPrograma
                .AddNew
                .Fields("Programa") = txtPrograma.Text
                If Len(txtDescripcionPrograma.Text) <> 0 Then
                    .Fields("Descripcion") = txtDescripcionPrograma.Text
                Else
                    .Fields("Descripcion") = ""
                End If
                .Update
            End With
            txtDescripcionPrograma.Enabled = False
            bolActualizarEstructura = False
            txtSubprograma.Enabled = True
            txtSubprograma.SetFocus
            Set rstPrograma = Nothing
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
   If Trim(txtSubprograma.Text) <> "" Then
        If Len(Trim(txtSubprograma.Text)) = 2 Then
            'Call Buscar(txtPrograma.Text & "-" & txtSubprograma.Text, rstSubprograma, "SUBPROGRAMAS", "pkSubprograma")
            With rstSubprograma
                If .NoMatch = False Then
                    If Len(.Fields("Descripcion")) <> 0 Then
                        txtDescripcionSubprograma.Text = .Fields("Descripcion")
                    Else
                        txtDescripcionSubprograma.Text = ""
                    End If
                    txtProyecto.Enabled = True
                    txtProyecto.SetFocus
                Else
                    Dim Agregar As Integer
                    Agregar = MsgBox("Desea Agregar el SUBPROGRAMA " & txtSubprograma.Text & " ?", vbQuestion + vbYesNo, "ACTUALIZAR ESTRUCTURA")
                    If Agregar = 6 Then
                        bolActualizarEstructura = True
                        txtDescripcionSubprograma.Enabled = True
                        txtDescripcionSubprograma.SetFocus
                        txtDescripcionSubprograma.Text = "ESCRIBA AQUÍ EL NOMBRE DEL SUBPROGRAMA"
                    Else
                        txtSubprograma.SetFocus
                    End If
                End If
            End With
            Set rstSubprograma = Nothing
        Else
            MsgBox "Debe ingresar un código de 2 cifras para continuar"
            txtSubprograma.SetFocus
        End If
    End If
End Sub

Private Sub txtDescripcionSubprograma_LostFocus()
    If bolActualizarEstructura = True Then
        Dim Agregar As Integer
        Agregar = MsgBox("Desea realmente agregar el SUBPROGRAMA " & txtSubprograma.Text & " con el siguiente nombre " & txtDescripcionSubprograma.Text & "?", vbQuestion + vbYesNo, "ACTUALIZAR ESTRUCTURA")
        If Agregar = 6 Then
            'Call SetRecordset(rstSubprograma, "SUBPROGRAMAS")
            With rstSubprograma
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
            Set rstSubprograma = Nothing
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
   If Trim(txtProyecto.Text) <> "" Then
        If Len(Trim(txtProyecto.Text)) = 2 Then
            'Call Buscar(txtPrograma.Text & "-" & txtSubprograma.Text & "-" & txtProyecto.Text, rstProyecto, "PROYECTOS", "pkProyecto")
            With rstProyecto
                If .NoMatch = False Then
                    If Len(.Fields("Descripcion")) <> 0 Then
                        txtDescripcionProyecto.Text = .Fields("Descripcion")
                    Else
                        txtDescripcionProyecto.Text = ""
                    End If
                    txtActividad.Enabled = True
                    txtActividad.SetFocus
                Else
                    Dim Agregar As Integer
                    Agregar = MsgBox("Desea Agregar el PROYECTO " & txtProyecto.Text & " ?", vbQuestion + vbYesNo, "ACTUALIZAR ESTRUCTURA")
                    If Agregar = 6 Then
                        bolActualizarEstructura = True
                        txtDescripcionProyecto.Enabled = True
                        txtDescripcionProyecto.SetFocus
                        txtDescripcionProyecto.Text = "ESCRIBA AQUÍ EL NOMBRE DE LA ACTIVIDAD"
                    Else
                        txtProyecto.SetFocus
                    End If
                End If
            End With
            Set rstProyecto = Nothing
        Else
            MsgBox "Debe ingresar un código de 2 cifras para continuar"
            txtProyecto.SetFocus
        End If
    End If
End Sub

Private Sub txtDescripcionProyecto_LostFocus()
    If bolActualizarEstructura = True Then
        Dim Agregar As Integer
        Agregar = MsgBox("Desea realmente agregar el PROYECTO " & txtProyecto.Text & " con el siguiente nombre " & txtDescripcionProyecto.Text & "?", vbQuestion + vbYesNo, "ACTUALIZAR ESTRUCTURA")
        If Agregar = 6 Then
            'Call SetRecordset(rstProyecto, "PROYECTOS")
            With rstProyecto
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
            Set rstProyecto = Nothing
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
   If Trim(txtActividad.Text) <> "" Then
        If Len(Trim(txtActividad.Text)) = 2 Then
            'Call Buscar(txtPrograma.Text & "-" & txtSubprograma.Text & "-" & txtProyecto.Text & "-" & txtActividad.Text, rstActividad, "ACTIVIDADES", "pkActividad")
            With rstActividad
                If .NoMatch = False Then
                    If Len(.Fields("Descripcion")) <> 0 Then
                        txtDescripcionActividad.Text = .Fields("Descripcion")
                    Else
                        txtDescripcionActividad.Text = ""
                    End If
                    txtPartida.Enabled = True
                    txtPartida.SetFocus
                Else
                    Dim Agregar As Integer
                    Agregar = MsgBox("Desea Agregar la ACTIVIDAD " & txtActividad.Text & " ?", vbQuestion + vbYesNo, "ACTUALIZAR ESTRUCTURA")
                    If Agregar = 6 Then
                        bolActualizarEstructura = True
                        txtDescripcionActividad.Enabled = True
                        txtDescripcionActividad.SetFocus
                        txtDescripcionActividad.Text = "ESCRIBA AQUÍ EL NOMBRE DE LA ACTIVIDAD"
                    Else
                        txtActividad.SetFocus
                    End If
                End If
            End With
            Set rstActividad = Nothing
        Else
            MsgBox "Debe ingresar un código de 2 cifras para continuar"
            txtActividad.SetFocus
        End If
    End If
End Sub

Private Sub txtDescripcionActividad_LostFocus()
    If bolActualizarEstructura = True Then
        Dim Agregar As Integer
        Agregar = MsgBox("Desea realmente agregar la ACTIVIDAD " & txtActividad.Text & " con el siguiente nombre " & txtDescripcionActividad.Text & "?", vbQuestion + vbYesNo, "ACTUALIZAR ESTRUCTURA")
        If Agregar = 6 Then
            'Call SetRecordset(rstActividad, "ACTIVIDADES")
            With rstActividad
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
            Set rstActividad = Nothing
        Else
            txtDescripcionActividad.Text = ""
            txtActividad.SetFocus
        End If
    End If
End Sub

Private Sub txtPartida_LostFocus()
    If Trim(txtPartida.Text) <> "" Then
        'Call SetRecordset(rstPartida, "PARTIDAS")
        With rstPartida
            .Index = "pkPartida"
            .Seek "=", txtPartida
        If .NoMatch = False Then
            txtDescripcionPartida.Text = .Fields("Descripcion")
        Else
            txtDescripcionPartida.Text = ""
        End If
        End With
        Set rstPartida = Nothing
    Else
        txtDescripcionPartida.Text = ""
    End If
End Sub

Private Sub txtMonto_LostFocus()
    If Len(Trim(txtMonto.Text)) <> 0 And IsNumeric(txtMonto.Text) = True Then
        txtSaldoRestante.Text = Val(txtSaldoDisponible.Text) - Val(txtMonto.Text)
    End If
End Sub

