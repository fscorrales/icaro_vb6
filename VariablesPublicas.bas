Attribute VB_Name = "VariablesPublicas"
Public strEjercicio As String
Public strEditandoEPAMConsolidado As String
Public strEditandoResolucionEPAM As String
Public strProveedor As String
Public strObra As String
Public strTipo As String
Public strCertificado As String
Public curFondoDeReparo As Currency
Public curMontoBruto As Currency
Public curIB As Currency
Public curLP As Currency
Public curSellos As Currency
Public curSUSS As Currency
Public curGanancias As Currency
Public curINVICO As Currency
Public bolAutocargaCertificado As Boolean
Public bolAutocargaEPAM As Boolean
Public bolObraDesdeEPAMConsolidado As Boolean
Public bolCUITDesdeCarga As Boolean
Public bolObraDesdeCarga As Boolean
Public bolActualizarEstructura As Boolean

Public Function ConexionDos(DataBase As String) As ADODB.Connection
    
    Set ConexionDos = New ADODB.Connection
    With ConexionDos
        .Provider = "Microsoft.Jet.OLEDB.4.0"
        .ConnectionString = "Data Source=" & DataBase
        .Open
    End With

End Function

Public Function RecordsetDos(Recordset As String, DataBase As ADODB.Connection)
   
    Set rstViejo = New ADODB.Recordset
    rstViejo.Open Recordset, DataBase, adOpenDynamic, adLockOptimistic

End Function

Public Sub VaciarVariablesAutocarga()
    
    curMontoBruto = "0"
    curFondoDeReparo = "0"
    curIB = "0"
    curLP = "0"
    curSUSS = "0"
    curGanancias = "0"
    curINVICO = "0"
    curSellos = "0"
    strProveedor = ""
    strCertificado = ""
    strObra = ""
    strTipo = ""
    bolAutocargaCertificado = False
    bolAutocargaEPAM = False
    
End Sub

Public Sub LlenarVariablesAutocargaCertificado()
    
    With AutocargaCertificados.dgAutocargaCertificado
        Dim i As Integer
        i = .Row
        strProveedor = .TextMatrix(i, 0)
        strObra = .TextMatrix(i, 1)
        strCertificado = .TextMatrix(i, 2)
        If Trim(strCertificado) = "FR" Then
            curFondoDeReparo = De_Txt_a_Num_01(.TextMatrix(i, 6))
            curFondoDeReparo = curFondoDeReparo * (-1)
        Else
            curFondoDeReparo = De_Txt_a_Num_01(.TextMatrix(i, 6))
        End If
        curMontoBruto = De_Txt_a_Num_01(.TextMatrix(i, 3))
        curIB = De_Txt_a_Num_01(.TextMatrix(i, 7))
        curLP = De_Txt_a_Num_01(.TextMatrix(i, 8))
        curSUSS = De_Txt_a_Num_01(.TextMatrix(i, 9))
        curGanancias = De_Txt_a_Num_01(.TextMatrix(i, 10))
        curINVICO = De_Txt_a_Num_01(.TextMatrix(i, 11))
    End With
    
End Sub

Public Sub LlenarVariablesAutocargaEPAM()
    
    With AutocargaEPAM.dgAutocargaEPAM
        Dim i As Integer
        i = .Row
        strObra = .TextMatrix(i, 0)
        'strTipo = .TextMatrix(i, 1)
        curMontoBruto = De_Txt_a_Num_01(.TextMatrix(i, 1))
        curGanancias = De_Txt_a_Num_01(.TextMatrix(i, 5))
        curSellos = De_Txt_a_Num_01(.TextMatrix(i, 6))
        curIB = De_Txt_a_Num_01(.TextMatrix(i, 7))
        curLP = De_Txt_a_Num_01(.TextMatrix(i, 8))
        curSUSS = De_Txt_a_Num_01(.TextMatrix(i, 9))
    End With
       
End Sub

Public Sub VaciarTodasLasVariables()

    strComprobante = ""
    strCargaEstructura = ""
    strEditandoEstructura = ""
    bolCUITDesdeCarga = False
    bolObraDesdeCarga = False
    bolAutocarga = False
    bolObraDesdeEPAM = False
    bolEditandoComprobante = False
    bolEditandoPartida = False
    bolEditandoFuente = False
    bolEditandoLocalidad = False
    bolEditandoCuentaBancaria = False
    bolEditandoRetencion = False
    bolEditandoImputacionEPAM = False
    bolEditandoResolucionesEPAM = False
    Nr = 0
    Np = 0
    VaciarVariablesAutocarga
    bolEditandoProveedores = False
    bolEditandoObras = False
    
End Sub
