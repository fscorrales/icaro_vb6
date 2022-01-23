VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CEPAM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_strCodigoObra As String
Private m_strTipo As String
Private m_curMontoBruto As Currency
Private m_curGanancias As Currency
Private m_curSellos As Currency
Private m_curLP As Currency
Private m_curIB As Currency
Private m_curSUSS As Currency
Private m_datFechaInicio As String
Private m_datFechaCierre As String

Public Property Get CodigoObra() As String
    CodigoObra = m_strCodigoObra
End Property

Public Property Let CodigoObra(NuevoCodigoObra As String)
    m_strCodigoObra = NuevoCodigoObra
End Property

Public Property Get Tipo() As String
    Tipo = m_strTipo
End Property

Public Property Let Tipo(NuevaTipo As String)
    m_strTipo = NuevaTipo
End Property

Public Property Get MontoBruto() As Currency
    MontoBruto = m_curMontoBruto
End Property

Public Property Let MontoBruto(NuevoMontoBruto As Currency)
    m_curMontoBruto = NuevoMontoBruto
End Property

Public Property Get Ganancias() As Currency
    Ganancias = m_curGanancias
End Property

Public Property Let Ganancias(NuevoGanancias As Currency)
    m_curGanancias = NuevoGanancias
End Property

Public Property Get Sellos() As Currency
    Sellos = m_curSellos
End Property

Public Property Let Sellos(NuevoSellos As Currency)
    m_curSellos = NuevoSellos
End Property

Public Property Get LP() As Currency
    LP = m_curLP
End Property

Public Property Let LP(NuevoLP As Currency)
    m_curLP = NuevoLP
End Property

Public Property Get IB() As Currency
    IB = m_curIB
End Property

Public Property Let IB(NuevoIB As Currency)
    m_curIB = NuevoIB
End Property

Public Property Get SUSS() As Currency
    SUSS = m_curSUSS
End Property

Public Property Let SUSS(NuevoSUSS As Currency)
    m_curSUSS = NuevoSUSS
End Property

Public Property Get FechaInicio() As String
    FechaInicio = m_datFechaInicio
End Property

Public Property Let FechaInicio(NuevoFechaInicio As String)
    m_datFechaInicio = NuevoFechaInicio
End Property

Public Property Get FechaCierre() As String
    FechaCierre = m_datFechaCierre
End Property

Public Property Let FechaCierre(NuevoFechaCierre As String)
    m_datFechaCierre = NuevoFechaCierre
End Property
