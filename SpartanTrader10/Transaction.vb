Public Class Transaction
    Public type As String
    Public qty As Integer
    Public symbol As String

    Public strike As Double
    Public delta As Double
    Public dividend As Double

    Public price As Double
    Public tcost As Double
    Public totValue As Double

    Public CAccountAT As Double
    Public interestSLT As Double
    Public marginAT As Double

    Public securityType As String

    Public underlier As String
    Public currentPositionInAP As Double
    Public underlierCurrentPositionInAP As Double

    Public expiration As Date
    Public mtm As Double
    Public mtmUnderlier As Double

    Public familyDelta As Double
    Public familyTicker As String
    Public score As Double
    Public hedgeQty As Double

    Public Sub Show()
        Globals.Dashboard.Range("C04").Value = type
        Globals.Dashboard.Range("C05").Value = qty
        Globals.Dashboard.Range("C06").Value = symbol

        Globals.Dashboard.Range("C8").Value = strike
        Globals.Dashboard.Range("C9").Value = delta
        Globals.Dashboard.Range("C10").Value = dividend

        Globals.Dashboard.Range("C12").Value = price
        Globals.Dashboard.Range("C13").Value = tcost
        Globals.Dashboard.Range("C14").Value = totValue

        Globals.Dashboard.Range("C16").Value = CAccountAT
        Globals.Dashboard.Range("C17").Value = interestSLT
        Globals.Dashboard.Range("C18").Value = marginAT
    End Sub

    Public Sub Clear()
        type = "-      "
        qty = 0
        symbol = "-      "
        strike = 0
        delta = 0
        dividend = 0
        price = 0
        tcost = 0
        totValue = 0
        CAccountAT = 0
        interestSLT = 0
        marginAT = 0
        securityType = ""
        currentPositionInAP = 0
        underlier = ""
        underlierCurrentPositionInAP = 0

        Globals.Dashboard.Range("C4:C6").Font.Color = System.Drawing.Color.White
    End Sub
    Public Sub CalcTransactionProperties(targetDate As Date)
        If IsAStock(symbol) Then ' set securityType and strike 
            securityType = "Stock"
            strike = 0
            dividend = GetDividend(symbol, targetDate)
            underlier = ""
            mtm = CalcMTM(symbol, targetDate)
        Else
            securityType = "Option"
            strike = GetStrike(symbol)
            dividend = 0
            underlier = GetUnderlier(symbol)
            expiration = GetExpiration(symbol)
            mtm = CalcMTM(underlier, targetDate)
            mtmUnderlier = CalcMTM(underlier, targetDate)
        End If
        FindTransactionPrice(symbol, targetDate)

        tcost = CalcTransactionCost()
        totValue = CalcTotValue()
        interestSLT = CalcInterestSLT(targetDate)
        CAccountAT = CAccount + totValue + interestSLT
        currentPositionInAP = GetCurrentPositionInAP(symbol)
        marginAT = margin + CalcEffectOfTransactionOnMargin(targetDate)
        delta = CalcDelta(symbol, targetDate)
    End Sub

    Public Function CalcTransactionCost() As Double
        Return GetTrCostCoefficient(securityType, type) * Math.Abs(qty) * price
    End Function

    Public Function CalcTotValue() As Double
        Select Case type
            Case "Buy"
                Return -(price * qty) - tcost
            Case "Sell"
                Return (price * qty) - tcost
            Case "SellShort"
                Return (price * qty) - tcost
            Case "CashDiv"
                Return (dividend * qty) - tcost
            Case "X-Put"
                Return (strike * qty) - tcost
            Case "X-Call"
                Return -(strike * qty) - tcost
            Case Else
                Return 0
        End Select
    End Function

    Public Sub Highlight()
        Globals.Dashboard.Range("C4:C6").Font.Color = System.Drawing.Color.Lime
    End Sub

    Public Function CalcEffectOfTransactionOnMargin(targdate As Date) As Double
        Select Case type
            Case "Sell"
                '  Sell has no effect on margin because you can only sell what you have long
                Return 0 ' effect of transaction on margin

            Case "Buy"
                If currentPositionInAP >= 0 Then
                    Return 0
                Else
                    If qty >= Math.Abs(currentPositionInAP) Then
                        Return -currentPositionInAP * mtm
                        ' buying eliminates all margin for this symbol
                    Else
                        Return (qty * mtm)
                        ' buying reduces the margin
                    End If
                End If

            Case "SellShort"
                Return -qty * mtm
                ' Selling short is easy: it always increases the margin

            Case "CashDiv"
                Return 0

            Case "X-Call"
                ' Two effects on margin: the change in options and the change in stocks
                Dim OptionEffect As Double = 0
                underlierCurrentPositionInAP = GetCurrentPositionInAP(underlier)

                ' first the effect of exercising the call on the call position
                If currentPositionInAP < 0 Then
                    OptionEffect = qty * mtm
                    ' i.e., it reduces the margin
                Else
                    OptionEffect = 0
                End If

                ' next, the effect of the called stock
                ' two cases: long and short calls.
                If currentPositionInAP >= 0 Then      ' long call is like buying
                    If underlierCurrentPositionInAP >= 0 Then
                        Return OptionEffect
                    Else
                        If qty >= Math.Abs(underlierCurrentPositionInAP) Then
                            Return OptionEffect - (underlierCurrentPositionInAP * mtmUnderlier)
                            ' X-call eliminates all margin for this symbol
                        Else
                            Return OptionEffect + (qty * mtmUnderlier)
                            ' X-call reduces the margin
                        End If
                    End If

                Else      '  exercising short calls is like selling

                    If underlierCurrentPositionInAP <= 0 Then
                        Return OptionEffect - (qty * mtmUnderlier)
                    Else   ' underlier positive
                        If underlierCurrentPositionInAP >= qty Then
                            Return OptionEffect
                        Else
                            Return OptionEffect - ((qty - underlierCurrentPositionInAP) * (qty * mtmUnderlier))
                        End If
                    End If
                End If

            Case "X-Put"
                ' Two effects on margin: the change in options and the change in stocks
                Dim OptionEffect As Double = 0
                underlierCurrentPositionInAP = GetCurrentPositionInAP(underlier)

                ' first the effect of exercising the option on the put position
                If currentPositionInAP < 0 Then
                    OptionEffect = qty * mtm
                    ' and it reduces the margin
                Else
                    OptionEffect = 0
                End If

                ' next, the effect of the stock
                ' two cases: long and short puts
                If currentPositionInAP < 0 Then      ' short put is like buying

                    If underlierCurrentPositionInAP >= 0 Then
                        Return OptionEffect
                    Else
                        If qty >= Math.Abs(underlierCurrentPositionInAP) Then
                            Return OptionEffect - (underlierCurrentPositionInAP * mtmUnderlier)
                            ' X-put eliminates all margin for this symbol
                        Else
                            Return OptionEffect + (qty * mtmUnderlier)
                            ' X-put reduces the margin
                        End If
                    End If

                Else      ' long put is like selling

                    If underlierCurrentPositionInAP <= 0 Then
                        Return OptionEffect - (qty * mtmUnderlier)
                    Else   ' underlier positive
                        If underlierCurrentPositionInAP >= qty Then
                            Return OptionEffect
                        Else
                            Return OptionEffect - ((qty - underlierCurrentPositionInAP) * mtmUnderlier)
                        End If
                    End If
                End If
        End Select
        MessageBox.Show("Holy BatApples! I could not figure out the impact of " + symbol + " on margin.  I returned $0.")
        Return 0
    End Function

    Public Sub FindTransactionPrice(symbol As String, targetdate As Date)
        Select Case type ' set the price based on the action 
            Case "Buy"
                price = GetAsk(symbol, targetdate)
            Case "Sell"
                price = GetBid(symbol, targetdate)
            Case "SellShort"
                price = GetBid(symbol, targetdate)
            Case "CashDiv"
                price = 0
            Case "X-Call"
                price = strike
            Case "X-Put"
                price = strike
            Case Else
                MessageBox.Show("Unknown transaction type, Dave.",
                                "Unknown trType", MessageBoxButtons.OK, MessageBoxIcon.Error)
                price = 0
        End Select
    End Sub

    Public Sub UpdatePosition()
        Dim newPosition As Double
        Dim newULPosition As Double
        Select Case type
            Case "Buy"
                newPosition = currentPositionInAP + Math.Abs(qty)
                UploadPosition(symbol, newPosition)
            Case "Sell"
                newPosition = currentPositionInAP - Math.Abs(qty)
                UploadPosition(symbol, newPosition)
            Case "SellShort"
                newPosition = currentPositionInAP - Math.Abs(qty)
                UploadPosition(symbol, newPosition)
            Case "CashDiv"
                ' only cash effects 
            Case "X-Put"
                underlierCurrentPositionInAP = GetCurrentPositionInAP(underlier)
                If currentPositionInAP > 0 Then
                    newPosition = currentPositionInAP - Math.Abs(qty)
                    UploadPosition(symbol, newPosition)
                    newULPosition = underlierCurrentPositionInAP - Math.Abs(qty)
                    UploadPosition(underlier, newULPosition)
                Else ' put position is short 
                    newPosition = currentPositionInAP + Math.Abs(qty)
                    UploadPosition(symbol, newPosition)
                    newULPosition = underlierCurrentPositionInAP + Math.Abs(qty)
                    UploadPosition(underlier, newULPosition)
                End If
        End Select

        UploadPosition("CAccount", CAccountAT)
        DownloadDataTableFromDB("Select * from PortfolioTeam" + teamID + " order by symbol", "AcquiredPositionsTbl")

    End Sub

    Public Sub MarkAsDone()
        type = "----"
        symbol = ""
        qty = 0
        totValue = 0
    End Sub
End Class
