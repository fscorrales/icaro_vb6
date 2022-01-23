VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form EjecucionMensual 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   6465
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   12840
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6465
   ScaleWidth      =   12840
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Caption         =   "Comprobantes"
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
      TabIndex        =   1
      Top             =   4200
      Width           =   12615
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid dgComprobantes 
         Height          =   1695
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   12375
         _ExtentX        =   21828
         _ExtentY        =   2990
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Ejecución Presupuestaria"
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
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   12615
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid dgEjecucion 
         Height          =   3615
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   12375
         _ExtentX        =   21828
         _ExtentY        =   6376
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
End
Attribute VB_Name = "EjecucionMensual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim intMeses As Integer

Private Sub Form_Load()

    Call CenterMe(EjecucionMensual, 13000, 7000)
    EstablecerMeses
    Call ConfigurardgEjecucionMensual(intMeses)
    Call CargardgEjecucionMensual(intMeses)
    intMeses = 0
    ConfigurardgComprobantesEjecucionMensual
    With Me.dgEjecucion
        If .TextMatrix(1, 0) <> "" Then
            ConfigurardgComprobantesEjecucionMensual
            CargardgComprobantesEjecucionMensual (.TextMatrix(1, 0) & "-" & .TextMatrix(1, 1))
        End If
    End With
    
End Sub

Sub EstablecerMeses()

    Dim SQL As String
    
    Set rstBuscarIcaro = New ADODB.Recordset
    SQL = "Select COMPROBANTE, FECHA From CARGA Where Right(COMPROBANTE,2) = " & Right(strEjercicio, 2) & " Order by Fecha Desc"
    rstBuscarIcaro.Open SQL, dbIcaro, adOpenForwardOnly, adLockOptimistic
    If rstBuscarIcaro.BOF = False Then
        rstBuscarIcaro.MoveFirst
        intMeses = Format(rstBuscarIcaro!Fecha, "mm")
    Else
        intMeses = 12
    End If
    
    Set rstBuscarIcaro = Nothing
    SQL = ""
    
End Sub

Private Sub dgEjecucion_RowColChange()

    Dim i As Integer
    
    With Me.dgEjecucion
        i = .Row
        If i <> 0 Then
            ConfigurardgComprobantesEjecucionMensual
            CargardgComprobantesEjecucionMensual (.TextMatrix(i, 0) & "-" & .TextMatrix(i, 1))
        End If
    End With
    
    i = 0
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    strEjercicio = ""
    intMeses = 0
    
End Sub

