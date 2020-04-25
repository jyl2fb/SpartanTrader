Module BlackScholes
    Public Function GetVol(symbol As String) As Double
        symbol = symbol.Trim()
        Select Case symbol
            Case "AAPL"
                Return 0.3
            Case "AMZ"
                Return 0.4
            Case "BIDU"
                Return 0.4
            Case "FB"
                Return 0.35
            Case "GOOG"
                Return 0.35
            Case "LUV"
                Return 0.25
            Case "MSFT"
                Return 0.35
            Case "SBUX"
                Return 0.2
            Case "SNAP"
                Return 1.1
            Case "TSLA"
                Return 0.4
            Case "TEVA"
                Return 0.4
            Case "WMT"
                Return 0.35
            Case Else
                Return 0.3
        End Select
    End Function

    Public Function CalcFamilyDelta(tkr As String, targetDate As Date) As Double
        Dim tempFamDelta As Double = 0
        Dim delta As Double = 0
        Dim sym As String
        tkr = tkr.Trim()
        'AP
        For Each row As DataRow In myDataSet.Tables("AcquiredPositionsTbl").Rows
            sym = row("Symbol").ToString().Trim()
            If IsInTheFamily(sym, tkr) Then
                delta = CalcDelta(sym, targetDate)
                tempFamDelta += delta * row("Units")
            End If
        Next
        'IP 
        For Each row As DataRow In myDataSet.Tables("InitialPositionsTbl").Rows
            sym = row("Symbol").ToString().Trim
            If IsInTheFamily(sym, tkr) Then
                delta = CalcDelta(sym, targetDate)
                tempFamDelta += delta * row("units")
            End If
        Next
        Return tempFamDelta

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
        Return Globals.ThisWorkbook.Application.WorksheetFunction.NormDist(d1, 0, 1, False) / (S * sigma * Math.Sqrt(t)) 'TODO add dividend
    End Function

End Module
