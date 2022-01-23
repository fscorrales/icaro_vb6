VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form SincronizacionDescripcionSGF 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "SINCRONIZACIÓN"
   ClientHeight    =   7665
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   14400
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7665
   ScaleWidth      =   14400
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Detalle de Modifcación a Realizar"
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
      Height          =   1455
      Left            =   120
      TabIndex        =   5
      Top             =   6120
      Width           =   14175
      Begin VB.CommandButton cmdModificar 
         BackColor       =   &H008080FF&
         Caption         =   "Modificar Descripción Obra ICARO"
         Height          =   855
         Left            =   12600
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   360
         Width           =   1395
      End
      Begin VB.TextBox txtModificacion 
         Height          =   285
         Left            =   1680
         TabIndex        =   8
         Top             =   960
         Width           =   10695
      End
      Begin VB.TextBox txtPrevioModificacion 
         Height          =   285
         Left            =   1680
         TabIndex        =   6
         Top             =   480
         Width           =   10695
      End
      Begin VB.Label Label3 
         Caption         =   "Modificación"
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
         TabIndex        =   9
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Previo a Modif."
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
         TabIndex        =   7
         Top             =   480
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Listado de Obras a Sincronizar"
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
      Height          =   5895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   14175
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid dgObrasSGF 
         Height          =   2415
         Left            =   240
         TabIndex        =   3
         Top             =   600
         Width           =   13725
         _ExtentX        =   24209
         _ExtentY        =   4260
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid dgObrasICARO 
         Height          =   2055
         Left            =   240
         TabIndex        =   4
         Top             =   3600
         Width           =   13725
         _ExtentX        =   24209
         _ExtentY        =   3625
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label Label5 
         Caption         =   "Obras SGF a Sincronizar con ICARO"
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
         TabIndex        =   2
         Top             =   360
         Width           =   4815
      End
      Begin VB.Label Label1 
         Caption         =   "Obras ICARO propuestas para modificar"
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
         TabIndex        =   1
         Top             =   3360
         Width           =   4815
      End
   End
End
Attribute VB_Name = "SincronizacionDescripcionSGF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Private Sub cmdAgregarReparaciones_Click()
'
'    Call EditarTipoDeObra("Reparaciones")
'    ConfigurardgTipoDeObra
'    CargardgTipoDeObra
'
'End Sub
'
'Private Sub cmdAgregarViviendas_Click()
'
'    Call EditarTipoDeObra("Viviendas")
'    ConfigurardgTipoDeObra
'    CargardgTipoDeObra
'
'End Sub
'
'Private Sub cmdEliminarReparaciones_Click()
'
'    Call EditarTipoDeObra("NoCategorizadas", "Reparaciones")
'    ConfigurardgTipoDeObra
'    CargardgTipoDeObra
'
'End Sub
'
'Private Sub cmdEliminarViviendas_Click()
'
'    Call EditarTipoDeObra("NoCategorizadas", "Viviendas")
'    ConfigurardgTipoDeObra
'    CargardgTipoDeObra
'
'End Sub

Private Sub cmdModificar_Click()

    Dim Modificar As Integer
    Dim strSGF As String
    Dim strIcaro As String

    strSGF = Me.txtModificacion.Text
    strIcaro = Me.txtPrevioModificacion.Text
    
    If strSGF <> "" And strIcaro <> "" Then
        Modificar = MsgBox("Desea MODIFICAR la descripción de la Obra registrada en ICARO como: " & strIcaro _
        & " por la siguiente: " & strSGF & " ?", vbQuestion + vbYesNo, "SINCRONIZANDO DESCRIPCION OBRA ICARO VS SGF")
        If Modificar = 6 Then
            Call SincronizarObra(strIcaro, strSGF)
        End If
    End If

End Sub

Private Sub Form_Load()

    Call CenterMe(SincronizacionDescripcionSGF, 14500, 8025)
    
End Sub

Private Sub dgObrasSGF_RowColChange()

    Dim i As Integer
    i = dgObrasSGF.Row
    Me.txtModificacion.Text = dgObrasSGF.TextMatrix(i, 1)
    ConfigurarSincronizaciondgObrasICARO
    CargarSincronizaciondgObrasICARO (Left(Me.txtModificacion.Text, 40))
    Me.txtPrevioModificacion.Text = ""
    i = Me.dgObrasICARO.Row
    If i > 0 Then
        Me.txtPrevioModificacion.Text = Me.dgObrasICARO.TextMatrix(i, 1)
    Else
        Me.txtPrevioModificacion.Text = ""
    End If

    i = 0
    
End Sub

Private Sub dgObrasICARO_RowColChange()

    Dim i As Integer
    i = Me.dgObrasICARO.Row
    If i > 0 Then
        Me.txtPrevioModificacion.Text = Me.dgObrasICARO.TextMatrix(i, 1)
    Else
        Me.txtPrevioModificacion.Text = ""
    End If
    i = 0
    
End Sub
