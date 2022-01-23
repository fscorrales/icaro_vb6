VERSION 5.00
Begin VB.Form CargaEstructuraProgramatica 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Carga Estructura Programatica"
   ClientHeight    =   2970
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7710
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2970
   ScaleWidth      =   7710
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdAgregar 
      BackColor       =   &H008080FF&
      Caption         =   "Guardar"
      Height          =   495
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2400
      Width           =   1695
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
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7455
      Begin VB.TextBox txtPrograma 
         Height          =   285
         Left            =   1680
         TabIndex        =   1
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox txtSubprograma 
         Height          =   285
         Left            =   1680
         TabIndex        =   3
         Top             =   840
         Width           =   495
      End
      Begin VB.TextBox txtProyecto 
         Height          =   285
         Left            =   1680
         TabIndex        =   5
         Top             =   1320
         Width           =   495
      End
      Begin VB.TextBox txtActividad 
         Height          =   285
         Left            =   1680
         TabIndex        =   7
         Top             =   1800
         Width           =   495
      End
      Begin VB.TextBox txtDescripcionPrograma 
         Height          =   285
         Left            =   2280
         TabIndex        =   2
         Top             =   360
         Width           =   5055
      End
      Begin VB.TextBox txtDescripcionSubprograma 
         Height          =   285
         Left            =   2280
         TabIndex        =   4
         Top             =   840
         Width           =   5055
      End
      Begin VB.TextBox txtDescripcionProyecto 
         Height          =   285
         Left            =   2280
         TabIndex        =   6
         Top             =   1320
         Width           =   5055
      End
      Begin VB.TextBox txtDescripcionActividad 
         Height          =   285
         Left            =   2280
         TabIndex        =   8
         Top             =   1800
         Width           =   5055
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
         TabIndex        =   12
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
         TabIndex        =   11
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
         TabIndex        =   10
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
         TabIndex        =   9
         Top             =   1800
         Width           =   1455
      End
   End
End
Attribute VB_Name = "CargaEstructuraProgramatica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    
    Call CenterMe(CargaEstructuraProgramatica, 7800, 3300)
    
End Sub

Private Sub cmdAgregar_Click()

    GenerarEstructuraProgramatica
    
End Sub

Private Sub EdicionCompleta()
    Select Case strCargaEstructura
    Case Is = "Programa"
        'Call SetRecordset(rstEditandoEstructura, "Select * From Subprogramas Where Programa = " & "'" & strEditandoEstructura & "'")
        With rstEditandoEstructura
            If .BOF = False Then
                .MoveFirst
                While .EOF = False
                    .Edit
                    .Fields("Subprograma") = strEditandoEstructura & Right(.Fields("Subprograma"), 3)
                    .Update
                    .MoveNext
                Wend
            End If
        End With
        'Call SetRecordset(rstEditandoEstructura, "Select * From Proyectos Where Left(Subprograma , 2) = " & "'" & strEditandoEstructura & "'")
        With rstEditandoEstructura
            If .BOF = False Then
                .MoveFirst
                While .EOF = False
                    .Edit
                    .Fields("Proyecto") = strEditandoEstructura & Right(.Fields("Proyecto"), 6)
                    .Update
                    .MoveNext
                Wend
            End If
        End With
        'Call SetRecordset(rstEditandoEstructura, "Select * From Actividades Where Left(Proyecto , 2) = " & "'" & strEditandoEstructura & "'")
        With rstEditandoEstructura
            If .BOF = False Then
                .MoveFirst
                While .EOF = False
                    .Edit
                    .Fields("Actividad") = strEditandoEstructura & Right(.Fields("Actividad"), 9)
                    .Update
                    .MoveNext
                Wend
            End If
        End With
    Case Is = "Subprograma"
        'Call SetRecordset(rstEditandoEstructura, "Select * From Proyectos Where Subprograma = " & "'" & strEditandoEstructura & "'")
        With rstEditandoEstructura
            If .BOF = False Then
                .MoveFirst
                While .EOF = False
                    .Edit
                    .Fields("Proyecto") = strEditandoEstructura & Right(.Fields("Proyecto"), 3)
                    .Update
                    .MoveNext
                Wend
            End If
        End With
        'Call SetRecordset(rstEditandoEstructura, "Select * From Actividades Where Left(Proyecto , 5) = " & "'" & strEditandoEstructura & "'")
        With rstEditandoEstructura
            If .BOF = False Then
                .MoveFirst
                While .EOF = False
                    .Edit
                    .Fields("Actividad") = strEditandoEstructura & Right(.Fields("Actividad"), 6)
                    .Update
                    .MoveNext
                Wend
            End If
        End With
    Case Is = "Proyecto"
        'Call SetRecordset(rstEditandoEstructura, "Select * From Actividades Where Proyecto = " & "'" & strEditandoEstructura & "'")
        With rstEditandoEstructura
            If .BOF = False Then
                .MoveFirst
                While .EOF = False
                    .Edit
                    .Fields("Actividad") = strEditandoEstructura & Right(.Fields("Actividad"), 3)
                    .Update
                    .MoveNext
                Wend
            End If
        End With
    End Select
End Sub
