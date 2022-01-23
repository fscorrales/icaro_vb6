Attribute VB_Name = "CalcularRetenciones"
Public Function CalcularGanancias(Regimen As String, CUIT As String, Fecha As Date, Importe As Currency, TipoDeObra As String, Partida As String, Optional Comprobante As String = "0") As Currency

    Dim SQL As String
    Dim curMinimoNoImponible As Currency
    
    Set dbAFIP = New ADODB.Connection
    dbAFIP.CursorLocation = adUseClient
    dbAFIP.Mode = adModeRead
    dbAFIP.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\AFIP.mdb"
    dbAFIP.Open
    
    Set rstRegistroAFIP = New ADODB.Recordset
    CalcularGanancias = 0
    Fecha = Format(CDate(Fecha), "MM/dd/yyyy")
    
    SQL = "Select Ganancias, Denominacion From CONDICIONTRIBUTARIA Where CUIT= '" & CUIT & "'"
    rstRegistroAFIP.Open SQL, dbAFIP, adOpenForwardOnly, adLockOptimistic
    If rstRegistroAFIP.BOF = False Then
        If rstRegistroAFIP!Ganancias = "AC" Or BuscarTexto(rstRegistroAFIP!Denominacion, "UTE") = True Then
            rstRegistroAFIP.Close
            dbAFIP.Close
            Set dbAFIP = Nothing
            Set dbAFIP = New ADODB.Connection
            dbAFIP.CursorLocation = adUseClient
            dbAFIP.Mode = adModeRead
            dbAFIP.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\ICARO.mdb"
            dbAFIP.Open
            If Regimen = "LOCACION" Then
                curMinimoNoImponible = 5000
                SQL = "Select CARGA.Comprobante, CARGA.Fecha, CARGA.Importe, CARGA.CUIT, ACTIVIDADES.TipoDeObra, IMPUTACION.Partida From CUIT INNER JOIN (CARGA INNER JOIN (ACTIVIDADES INNER JOIN IMPUTACION ON ACTIVIDADES.Actividad = IMPUTACION.Imputacion) ON CARGA.Comprobante = IMPUTACION.Comprobante) ON CUIT.CUIT = CARGA.CUIT Where CARGA.CUIT= '" & CUIT & "' And Year(Fecha) = '" & Year(Fecha) & "' And Month(Fecha) = '" & Month(Fecha) & "' And Fecha <= #" & Fecha & "# And Left(CARGA.Comprobante,4) < '" & Left(Comprobante, 4) & "' Order By CARGA.Comprobante"
                rstRegistroAFIP.Open SQL, dbAFIP, adOpenForwardOnly, adLockOptimistic
                If rstRegistroAFIP.BOF = False Then
                    rstRegistroAFIP.MoveFirst
                    Do While rstRegistroAFIP.EOF = False
                        If curMinimoNoImponible < 0 Then
                            Select Case TipoDeObra
                            Case "R" 'Reparaciones?
                                CalcularGanancias = Importe / 1.21 * 0.02
                            Case "V"
                                If Partida = "421" Then 'Construcción Viviendas?
                                    CalcularGanancias = Importe / 1.105 * 0.02
                                ElseIf Partida = "422" Then 'Infraestructura?
                                    CalcularGanancias = Importe / 1.21 * 0.02
                                End If
                            End Select
                            Exit Do
                        Else
                            Select Case rstRegistroAFIP!TipoDeObra
                            Case "R" 'Reparaciones?
                                curMinimoNoImponible = curMinimoNoImponible - (rstRegistroAFIP!Importe / 1.21)
                            Case "V"
                                If rstRegistroAFIP!Partida = "421" Then 'Construcción Viviendas?
                                    curMinimoNoImponible = curMinimoNoImponible - (rstRegistroAFIP!Importe / 1.105)
                                ElseIf rstRegistroAFIP!Partida = "422" Then 'Infraestructura?
                                    curMinimoNoImponible = curMinimoNoImponible - (rstRegistroAFIP!Importe / 1.21)
                                End If
                            End Select
                        End If
                    rstRegistroAFIP.MoveNext
                    Loop
                    If curMinimoNoImponible < 0 Then
                        Select Case TipoDeObra
                        Case "R" 'Reparaciones?
                            CalcularGanancias = Importe / 1.21 * 0.02
                        Case "V"
                            If Partida = "421" Then 'Construcción Viviendas?
                                CalcularGanancias = Importe / 1.105 * 0.02
                            ElseIf Partida = "422" Then 'Infraestructura?
                                CalcularGanancias = Importe / 1.21 * 0.02
                            End If
                        End Select
                    Else
                        Select Case TipoDeObra
                        Case "R" 'Reparaciones?
                            CalcularGanancias = (Importe / 1.21 - curMinimoNoImponible) * 0.02
                        Case "V"
                            If Partida = "421" Then 'Construcción Viviendas?
                                CalcularGanancias = (Importe / 1.105 - curMinimoNoImponible) * 0.02
                            ElseIf Partida = "422" Then 'Infraestructura?
                                CalcularGanancias = (Importe / 1.21 - curMinimoNoImponible) * 0.02
                            End If
                        End Select
                    End If
                Else
                    If curMinimoNoImponible < 0 Then
                        Select Case TipoDeObra
                        Case "R" 'Reparaciones?
                            CalcularGanancias = Importe / 1.21 * 0.02
                        Case "V"
                            If Partida = "421" Then 'Construcción Viviendas?
                                CalcularGanancias = Importe / 1.105 * 0.02
                            ElseIf Partida = "422" Then 'Infraestructura?
                                CalcularGanancias = Importe / 1.21 * 0.02
                            End If
                        End Select
                    Else
                        Select Case TipoDeObra
                        Case "R" 'Reparaciones?
                            CalcularGanancias = (Importe / 1.21 - curMinimoNoImponible) * 0.02
                        Case "V"
                            If Partida = "421" Then 'Construcción Viviendas?
                                CalcularGanancias = (Importe / 1.105 - curMinimoNoImponible) * 0.02
                            ElseIf Partida = "422" Then 'Infraestructura?
                                CalcularGanancias = (Importe / 1.21 - curMinimoNoImponible) * 0.02
                            End If
                        End Select
                        If CalcularGanancias < 0 Then
                            CalcularGanancias = 0
                        End If
                    End If
                End If
            ElseIf Regimen = "VENTA" Then
                'falta programar
            End If
        End If
    Else
        CalcularGanancias = 0
    End If
    
    rstRegistroAFIP.Close
    Set rstRegistroAFIP = Nothing
    dbAFIP.Close
    Set dbAFIP = Nothing

End Function

Public Function CalcularSUSS(Regimen As String, CUIT As String, Fecha As Date, Importe As Currency, TipoDeObra As String, Partida As String, Optional Comprobante As String = "0") As Currency

    Dim SQL As String
    Dim curMinimoExento As Currency
    
    Set dbAFIP = New ADODB.Connection
    dbAFIP.CursorLocation = adUseClient
    dbAFIP.Mode = adModeRead
    dbAFIP.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\AFIP.mdb"
    dbAFIP.Open
    
    Set rstRegistroAFIP = New ADODB.Recordset
    CalcularSUSS = 0
    Fecha = Format(CDate(Fecha), "MM/dd/yyyy")
    
    SQL = "Select Empleador From CONDICIONTRIBUTARIA Where CUIT= '" & CUIT & "'"
    rstRegistroAFIP.Open SQL, dbAFIP, adOpenForwardOnly, adLockOptimistic
    If rstRegistroAFIP.BOF = False Then
        If rstRegistroAFIP!Empleador = "S" Then
            rstRegistroAFIP.Close
            dbAFIP.Close
            Set dbAFIP = Nothing
            Set dbAFIP = New ADODB.Connection
            dbAFIP.CursorLocation = adUseClient
            dbAFIP.Mode = adModeRead
            dbAFIP.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\ICARO.mdb"
            dbAFIP.Open
            If Regimen = "CONSTRUCCION" Then
                SQL = "Select CARGA.Comprobante, CARGA.Fecha, CARGA.Importe, CARGA.CUIT, ACTIVIDADES.TipoDeObra, IMPUTACION.Partida From CUIT INNER JOIN (CARGA INNER JOIN (ACTIVIDADES INNER JOIN IMPUTACION ON ACTIVIDADES.Actividad = IMPUTACION.Imputacion) ON CARGA.Comprobante = IMPUTACION.Comprobante) ON CUIT.CUIT = CARGA.CUIT Where CARGA.CUIT= '" & CUIT & "' And Year(Fecha) = '" & Year(Fecha) & "' And Fecha <= #" & Fecha & "# And Left(CARGA.Comprobante,4) < '" & Left(Comprobante, 4) & "' Order By CARGA.Comprobante"
                rstRegistroAFIP.Open SQL, dbAFIP, adOpenForwardOnly, adLockOptimistic
                If rstRegistroAFIP.BOF = False Then
                    rstRegistroAFIP.MoveFirst
                    Do While rstRegistroAFIP.EOF = False
                        If curMinimoExento > 400000 Then
                            Select Case TipoDeObra
                            Case "R" 'Reparaciones?
                                If Partida = "421" Then 'Construcción Viviendas?
                                    CalcularSUSS = Importe / 1.21 * 0.025
                                ElseIf Partida = "422" Then 'Infraestructura?
                                    CalcularSUSS = Importe / 1.21 * 0.012
                                End If
                            Case "V"
                                If Partida = "421" Then 'Construcción Viviendas?
                                    CalcularSUSS = Importe / 1.105 * 0.025
                                ElseIf Partida = "422" Then 'Infraestructura?
                                    CalcularSUSS = Importe / 1.21 * 0.012
                                End If
                            End Select
                            Exit Do
                        Else
                            Select Case rstRegistroAFIP!TipoDeObra
                            Case "R" 'Reparaciones?
                                If rstRegistroAFIP!Partida = "421" Then 'Reparaciones?
                                    curMinimoExento = curMinimoExento + (rstRegistroAFIP!Importe / 1.21)
                                    CalcularSUSS = CalcularSUSS + (rstRegistroAFIP!Importe / 1.21 * 0.025)
                                ElseIf rstRegistroAFIP!Partida = "422" Then 'Infraestructura?
                                    curMinimoExento = curMinimoExento + (rstRegistroAFIP!Importe / 1.21)
                                    CalcularSUSS = CalcularSUSS + (rstRegistroAFIP!Importe / 1.21 * 0.012)
                                End If
                            Case "V"
                                If rstRegistroAFIP!Partida = "421" Then 'Construcción Viviendas?
                                    curMinimoExento = curMinimoExento + (rstRegistroAFIP!Importe / 1.105)
                                    CalcularSUSS = CalcularSUSS + (rstRegistroAFIP!Importe / 1.105 * 0.025)
                                ElseIf rstRegistroAFIP!Partida = "422" Then 'Infraestructura?
                                    curMinimoExento = curMinimoExento + (rstRegistroAFIP!Importe / 1.21)
                                    CalcularSUSS = CalcularSUSS + (rstRegistroAFIP!Importe / 1.21 * 0.012)
                                End If
                            End Select
                        End If
                        rstRegistroAFIP.MoveNext
                    Loop
                    If Not curMinimoExento > 400000 Then
                        Select Case TipoDeObra
                        Case "R" 'Reparaciones?
                            If Partida = "421" Then 'Reparaciones?
                                curMinimoExento = curMinimoExento + Importe / 1.21
                                CalcularSUSS = CalcularSUSS + (Importe / 1.21 * 0.025)
                            ElseIf Partida = "422" Then 'Infraestructura?
                                curMinimoExento = curMinimoExento + Importe / 1.21
                                CalcularSUSS = CalcularSUSS + (Importe / 1.21 * 0.012)
                            End If
                        Case "V"
                            If Partida = "421" Then 'Construcción Viviendas?
                                curMinimoExento = curMinimoExento + Importe / 1.105
                                CalcularSUSS = CalcularSUSS + (Importe / 1.105 * 0.025)
                            ElseIf Partida = "422" Then 'Infraestructura?
                                curMinimoExento = curMinimoExento + Importe / 1.21
                                CalcularSUSS = CalcularSUSS + (Importe / 1.21 * 0.012)
                            End If
                        End Select
                        If Not curMinimoExento > 400000 Then
                            CalcularSUSS = 0
                        End If
                    Else
                        Select Case TipoDeObra
                        Case "R" 'Reparaciones?
                            If Partida = "421" Then 'Construcción Viviendas?
                                CalcularSUSS = Importe / 1.21 * 0.025
                            ElseIf Partida = "422" Then 'Infraestructura?
                                CalcularSUSS = Importe / 1.21 * 0.012
                            End If
                        Case "V"
                            If Partida = "421" Then 'Construcción Viviendas?
                                CalcularSUSS = Importe / 1.105 * 0.025
                            ElseIf Partida = "422" Then 'Infraestructura?
                                CalcularSUSS = Importe / 1.21 * 0.012
                            End If
                        End Select
                    End If
                Else
                    Select Case TipoDeObra
                    Case "R" 'Reparaciones?
                        If Partida = "421" Then 'Reparaciones?
                            curMinimoExento = Importe / 1.21
                            CalcularSUSS = curMinimoExento * 0.025
                        ElseIf Partida = "422" Then 'Infraestructura?
                            curMinimoExento = Importe / 1.21
                            CalcularSUSS = curMinimoExento * 0.012
                        End If
                    Case "V"
                        If Partida = "421" Then 'Construcción Viviendas?
                            curMinimoExento = Importe / 1.105
                            CalcularSUSS = curMinimoExento * 0.025
                        ElseIf Partida = "422" Then 'Infraestructura?
                            curMinimoExento = Importe / 1.21
                            CalcularSUSS = curMinimoExento * 0.012
                        End If
                    End Select
                    If Not curMinimoExento > 400000 Then
                        CalcularSUSS = 0
                    End If
                End If
            ElseIf Regimen = "GENERAL" Then
                'falta programar
            End If
        End If
    Else
        CalcularSUSS = 0
    End If
    
    rstRegistroAFIP.Close
    Set rstRegistroAFIP = Nothing
    dbAFIP.Close
    Set dbAFIP = Nothing

End Function

