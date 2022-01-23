VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCertificados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_strProveedor As String
Private m_strObra As String
Private m_strCertificado As String
Private m_curFondoDeReparo As Currency
Private m_curMontoBruto As Currency
Private m_curIB As Currency
Private m_curLP As Currency
Private m_curSUSS As Currency
Private m_curGanancias As Currency
Private m_curINVICO As Currency
Private m_datFecha As Date

Public Property Get Proveedor() As String
    Proveedor = m_strProveedor
End Property

Public Property Let Proveedor(NuevoProveedor As String)
    m_strProveedor = NuevoProveedor
End Property

Public Property Get Obra() As String
    Obra = m_strObra
End Property

Public Property Let Obra(NuevaObra As String)
    m_strObra = left(NuevaObra, 100)
End Property

Public Property Get Certificado() As String
    Certificado = m_strCertificado
End Property

Public Property Let Certificado(NuevoCertificado As String)
    m_strCertificado = NuevoCertificado
End Property


Public Property Get FondoDeReparo() As Currency
    FondoDeReparo = m_curFondoDeReparo
End Property

Public Property Let FondoDeReparo(NuevoFondoDeReparo As Currency)
    m_curFondoDeReparo = NuevoFondoDeReparo
End Property

Public Property Get MontoBruto() As Currency
    MontoBruto = m_curMontoBruto
End Property

Public Property Let MontoBruto(NuevoMontoBruto As Currency)
    m_curMontoBruto = NuevoMontoBruto
End Property

Public Property Get IB() As Currency
    IB = m_curIB
End Property

Public Property Let IB(NuevoIB As Currency)
    m_curIB = NuevoIB
End Property

Public Property Get LP() As Currency
    LP = m_curLP
End Property

Public Property Let LP(NuevoLP As Currency)
    m_curLP = NuevoLP
End Property

Public Property Get SUSS() As Currency
    SUSS = m_curSUSS
End Property

Public Property Let SUSS(NuevoSUSS As Currency)
    m_curSUSS = NuevoSUSS
End Property

Public Property Get Ganancias() As Currency
    Ganancias = m_curGanancias
End Property

Public Property Let Ganancias(NuevoGanancias As Currency)
    m_curGanancias = NuevoGanancias
End Property

Public Property Get INVICO() As Currency
    INVICO = m_curINVICO
End Property

Public Property Let INVICO(NuevoINVICO As Currency)
    m_curINVICO = NuevoINVICO
End Property

Public Property Get Periodo() As Date
    Periodo = m_datFecha
End Property

Public Property Let Periodo(NuevoPeriodo As Date)
    
    Dim MesYAño As String
    MesYAño = Format(NuevoPeriodo, "mmmm/yyyy")
    m_datFecha = MesYAño
    
End Property

