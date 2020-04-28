Module BlackScholes
    Public Function GetVol(symbol As String) As Double
        symbol = symbol.Trim()
        Select Case symbol
            Case "ABT"
                Return 0.745
            Case "AMD"
                Return 0.9505
            Case "APPN"
                Return 1.173
            Case "CCL"
                Return 1.888
            Case "F"
                Return 0.9155
            Case "JNJ"
                Return 0.629
            Case "KO"
                Return 0.6325
            Case "PENN"
                Return 2.6875
            Case "PSA"
                Return 0.6845
            Case "SHOP"
                Return 1.0
            Case "SPCE"
                Return 1.7685
            Case "ZM"
                Return 0.896
            Case Else
                Return 0.0
        End Select
    End Function

    Public Function CalcFamilyDelta(tkr As String, targetDate As Date, initialOnly As Boolean) As Double
        Dim tempFamDelta As Double = 0
        Dim sym As String
        tkr = tkr.Trim()
        Dim delta As Double = 0
        'IP 
        For Each row As DataRow In myDataSet.Tables("InitialPositionsTbl").Rows
            sym = row("Symbol").ToString().Trim
            If IsInTheFamily(sym, tkr) Then
                delta = CalcDelta(sym, targetDate)
                tempFamDelta += delta * row("units")
            End If
        Next
        If initialOnly Then
            Return tempFamDelta
        End If
        'AP
        For Each row As DataRow In myDataSet.Tables("AcquiredPositionsTbl").Rows
            sym = row("Symbol").ToString().Trim()
            If IsInTheFamily(sym, tkr) Then
                delta = CalcDelta(sym, targetDate)
                tempFamDelta += delta * row("Units")
            End If
        Next

        Return tempFamDelta

    End Function

    Public Function CalcFamilyGamma(tkr As String, targetDate As Date, initialOnly As Boolean) As Double
        Dim tempFamGamma As Double = 0
        Dim sym As String
        Dim gamma As Double = 0
        tkr = tkr.Trim()
        'IP 
        For Each row As DataRow In myDataSet.Tables("InitialPositionsTbl").Rows
            sym = row("Symbol").ToString().Trim
            If IsInTheFamily(sym, tkr) Then
                gamma = CalcGamma(sym, targetDate)
                tempFamGamma += gamma * row("units")
            End If
        Next
        If initialOnly Then
            Return tempFamGamma
        End If

        For Each row As DataRow In myDataSet.Tables("AcquiredPositionsTbl").Rows
            sym = row("Symbol").ToString().Trim()
            If IsInTheFamily(sym, tkr) Then
                gamma = CalcGamma(sym, targetDate)
                tempFamGamma += gamma * row("Units")
            End If
        Next
        Return tempFamGamma

    End Function

    Public Function IsInTheFamily(sym As String, familyTicker As String) As Boolean
        sym = sym.Trim()
        familyTicker = familyTicker.Trim()
        If sym = familyTicker Then
            Return True
        End If
        If sym = "CAccount" Or IsAStock(sym) Then
            Return False
        End If
        If GetUnderlier(sym) = familyTicker Then
            Return True
        End If
        Return False
    End Function

    Public Function CalcDelta(symbol As String, targetdate As Date) As Double
        Dim sigma, K, S, t, d1 As Double
        Dim ts As TimeSpan
        Dim r As Double = riskFreeRate
        Dim underlier As String

        If symbol = "CAccount" Then
            Return 0
        End If

        If IsAStock(symbol) Then
            Return 1
        End If

        If targetdate >= GetExpiration(symbol).Date Then
            Return 0
        End If

        If GetAsk(symbol, currentDate) = 0 Then
            Return 0
        End If
        underlier = GetUnderlier(symbol)
        sigma = GetVol(underlier)
        K = GetStrike(symbol)
        S = CalcMTM(underlier, targetdate)
        ts = GetExpiration(symbol).Date - targetdate.Date
        t = ts.Days / 365.25
        d1 = (Math.Log(S / K) + (r + sigma * sigma / 2) * t) / (sigma * Math.Sqrt(t))
        If IsACall(symbol) Then
            Return Globals.ThisWorkbook.Application.WorksheetFunction.Norm_S_Dist(d1, True)
        Else
            Return (Globals.ThisWorkbook.Application.WorksheetFunction.Norm_S_Dist(d1, True) - 1)
        End If
    End Function

    Public Function CalcGamma(symbol As String, targetdate As Date) As Double
        Dim sigma, K, S, t, d1 As Double
        Dim ts As TimeSpan
        Dim r As Double = riskFreeRate
        Dim underlier As String
        Dim answer As Double

        If symbol = "CAccount" Then
            Return 0
        End If

        If IsAStock(symbol) Then
            Return 0
        End If

        If targetdate >= GetExpiration(symbol).Date Then
            Return 0
        End If

        If GetAsk(symbol, currentDate) = 0 Then
            Return 0
        End If

        underlier = GetUnderlier(symbol)
        sigma = GetVol(underlier)
        K = GetStrike(symbol)
        S = CalcMTM(underlier, targetdate)
        ts = GetExpiration(symbol).Date - targetdate.Date
        t = ts.Days / 365.25
        d1 = (Math.Log(S / K) + (r + sigma * sigma / 2) * t) / (sigma * Math.Sqrt(t))

        answer = (Globals.ThisWorkbook.Application.WorksheetFunction.NormDist(d1, 0, 1, False)) / (S * sigma * Math.Sqrt(t)) 'TODO add dividend
        Return answer
    End Function

End Module
