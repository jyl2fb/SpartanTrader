Module ScoringRules
    '  NOTE: this code is given to you As-Is. Its parameters are suboptimal.
    '  You need To understand it and make all the necessary changes.

    Public Sub ScoreSellingStock(baseScore As Integer, recommendation As Transaction, tDate As Date)
        Dim adjustment As Double = 0
        Dim cRec As New Transaction  ' this is a CANDIDATE rec, do not confuse with recommendation!
        cRec.symbol = recommendation.familyTicker
        cRec.type = "Sell"
        cRec.familyDelta = recommendation.familyDelta

        If IsInIP(cRec.symbol) Then
            Exit Sub   ' cannot sell if in IP
        Else
            cRec.currentPositionInAP = GetCurrentPositionInAP(cRec.symbol)
            If cRec.currentPositionInAP <= 0 Then ' we cannot sell since we are not long
                Exit Sub
            Else
                cRec.delta = 1
                cRec.hedgeQty = CalcQtyNeededToHedge(cRec.symbol, cRec.delta, cRec.familyDelta)
                If cRec.hedgeQty = 0 Then
                    Exit Sub  ' nothing to sell
                Else
                    If cRec.hedgeQty > cRec.currentPositionInAP Then     ' you have fewer than needed, so
                        cRec.qty = cRec.currentPositionInAP              ' sell all you have.
                        adjustment = -50                                 ' Arbitrary adjustment!
                    Else
                        cRec.qty = cRec.hedgeQty
                    End If
                    cRec.score = baseScore + adjustment
                    CandidateRecList.Add(cRec)
                End If
            End If
        End If
    End Sub

    Public Sub ScoreSellingCall(baseScore As Integer, recommendation As Transaction, tDate As Date)
        Dim adjustment As Double = 0
        Dim cRec As Transaction  ' this is a CANDIDATE rec, do not confuse with recommendation!
        For Each APRow As DataRow In myDataSet.Tables("AcquiredPositionsTbl").Rows
            cRec = New Transaction
            cRec.type = "Sell"
            cRec.symbol = APRow("Symbol").Trim()
            cRec.familyDelta = recommendation.familyDelta
            adjustment = 0
            If IsACall(cRec.symbol) Then
                If IsInTheFamily(cRec.symbol, recommendation.familyTicker) Then
                    cRec.currentPositionInAP = APRow("Units")
                    If cRec.currentPositionInAP > 0 Then   ' to sell, you need to have 
                        cRec.delta = CalcDelta(cRec.symbol, tDate)
                        cRec.hedgeQty = CalcQtyNeededToHedge(cRec.symbol, cRec.delta, cRec.familyDelta)
                        If cRec.hedgeQty > 0 Then
                            If cRec.currentPositionInAP < cRec.hedgeQty Then
                                cRec.qty = cRec.currentPositionInAP ' sell all you have
                                adjustment = -50           ' because incomplete hedge
                            Else
                                cRec.qty = cRec.hedgeQty
                            End If
                            cRec.score = baseScore + adjustment
                            CandidateRecList.Add(cRec)
                        End If
                    End If
                End If
            End If
        Next
    End Sub

    Public Sub ScoreSellingShortCall(baseScore As Integer, recommendation As Transaction, tDate As Date)
        Dim adjustment, maxshort
        Dim cRec As Transaction
        If TooCloseToMaxMargins() Then
            Exit Sub ' we have no more credit
        End If
        For Each partialSymbol As String In {"_COCTE", "_COCTD", "_COCTC", "_COCTB", "_COCTA"}   ' only these options will be considered, in this order
            cRec = New Transaction                                        ' you might add/subtract To/from the list, for example adding the JULY otions
            cRec.type = "SellShort"
            cRec.symbol = recommendation.familyTicker + partialSymbol
            cRec.familyDelta = recommendation.familyDelta
            adjustment = 0
            If Not IsInIP(cRec.symbol) Then
                cRec.currentPositionInAP = GetCurrentPositionInAP(cRec.symbol)
                If cRec.currentPositionInAP <= 0 Then   ' because if long cannot sell short
                    cRec.delta = CalcDelta(cRec.symbol, tDate)
                    cRec.hedgeQty = CalcQtyNeededToHedge(cRec.symbol, cRec.delta, cRec.familyDelta)
                    maxshort = MaxShortWithinConstraints(cRec.symbol, tDate)
                    If cRec.hedgeQty > maxshort Then
                        cRec.qty = maxshort
                        adjustment = -50
                    Else
                        cRec.qty = cRec.hedgeQty
                    End If
                    If cRec.qty > 0 Then
                        Select Case partialSymbol
                            Case "_COCTE"
                                adjustment = adjustment + 5
                            Case "_COCTD"
                                adjustment = adjustment + 4
                            Case "_COCTC"
                                adjustment = adjustment + 3
                            Case "_COCTB"
                                adjustment = adjustment + 2
                            Case "_COCTA"
                                adjustment = adjustment + 1
                        End Select
                        cRec.score = baseScore + adjustment
                        CandidateRecList.Add(cRec)
                    End If
                End If
            End If
        Next
    End Sub

    Public Sub ScoreBuyingBackPut(baseScore As Integer, recommendation As Transaction, tDate As Date)
        Dim adjustment, maxbuy As Double
        Dim cRec As Transaction  ' this is a CANDIDATE rec, do not confuse with recommendation!
        If AvailableCashIsLow() Then
            Exit Sub
        End If
        For Each dr As DataRow In myDataSet.Tables("AcquiredPositionsTbl").Rows
            cRec = New Transaction
            cRec.familyDelta = recommendation.familyDelta
            cRec.type = "Buy"
            cRec.symbol = dr("Symbol").Trim()
            adjustment = 0
            If IsAStock(cRec.symbol) Or cRec.symbol = "CAccount" Then
                ' skip
            Else
                If IsAPut(cRec.symbol) Then
                    cRec.underlier = GetUnderlier(cRec.symbol)
                    If cRec.underlier = recommendation.familyTicker Then
                        cRec.currentPositionInAP = dr("Units")
                        If cRec.currentPositionInAP < 0 Then
                            cRec.delta = CalcDelta(cRec.symbol, tDate)
                            cRec.hedgeQty = CalcQtyNeededToHedge(cRec.symbol, cRec.delta, cRec.familyDelta)
                            If cRec.hedgeQty > Math.Abs(cRec.currentPositionInAP) Then
                                cRec.hedgeQty = Math.Abs(cRec.currentPositionInAP)         ' buyback all shorts
                                adjustment = -50
                            End If 'fixed bug
                            maxbuy = MaxPurchasePossible(cRec.symbol, tDate)           ' how much can you afford?
                            If maxbuy < cRec.hedgeQty Then
                                cRec.qty = maxbuy
                                adjustment = adjustment - 50
                            Else
                                cRec.qty = cRec.hedgeQty
                            End If
                            If cRec.qty > 0 Then
                                cRec.score = baseScore + adjustment
                                CandidateRecList.Add(cRec)
                            End If
                        End If
                    End If
                End If
            End If
        Next
    End Sub

    Public Sub ScoreBuyingPut(baseScore As Integer, recommendation As Transaction, tDate As Date)
        Dim adjustment As Double = 0
        Dim maxBuy As Double = 0
        Dim cRec As Transaction  ' this is a CANDIDATE rec, do not confuse with recommendation!
        If AvailableCashIsLow() Then
            Exit Sub
        End If
        ' only considers OCT options - you can easily change this: talk to me
        For Each partialSymbol As String In {"_POCTA", "_POCTB", "_POCTC", "_POCTD", "_POCTE"}
            cRec = New Transaction
            cRec.type = "Buy"
            cRec.symbol = recommendation.familyTicker + partialSymbol
            cRec.familyDelta = recommendation.familyDelta
            adjustment = 0
            If Not IsInIP(cRec.symbol) Then
                cRec.currentPositionInAP = GetCurrentPositionInAP(cRec.symbol)
                If cRec.currentPositionInAP >= 0 Then                      ' if short it is a buyback
                    cRec.delta = CalcDelta(cRec.symbol, tDate)
                    cRec.hedgeQty = CalcQtyNeededToHedge(cRec.symbol, cRec.delta, cRec.familyDelta)
                    maxBuy = MaxPurchasePossible(cRec.symbol, tDate)       ' how much can you afford?
                    If maxBuy < cRec.hedgeQty Then
                        cRec.qty = maxBuy
                        adjustment = -50
                    Else
                        cRec.qty = cRec.hedgeQty
                    End If
                    If cRec.qty > 0 Then
                        Select Case partialSymbol
                            Case "_POCTE"
                                adjustment = adjustment + 1               ' preferences - need work
                            Case "_POCTD"
                                adjustment = adjustment + 4
                            Case "_POCTC"
                                adjustment = adjustment + 5
                            Case "_POCTB"
                                adjustment = adjustment + 2
                            Case "_POCTA"
                                adjustment = adjustment + 3
                        End Select
                        cRec.score = baseScore + adjustment
                        CandidateRecList.Add(cRec)
                    End If
                End If
            End If
        Next
    End Sub

    Public Sub ScoreSellingShortStock(baseScore As Integer, recommendation As Transaction, tDate As Date)
        If TooCloseToMaxMargins() Then
            Exit Sub
        End If
        Dim adjustment, maxShort As Double
        Dim cRec As New Transaction  ' this is a CANDIDATE rec, do not confuse with recommendation!
        cRec.symbol = recommendation.familyTicker
        cRec.type = "SellShort"
        cRec.familyDelta = recommendation.familyDelta

        If IsInIP(cRec.symbol) Then
            Exit Sub                 ' cannot sellshort if in IP
        Else
            cRec.currentPositionInAP = GetCurrentPositionInAP(cRec.symbol)
            If cRec.currentPositionInAP > 0 Then      ' we cannot sellshort if long
                Exit Sub
            Else
                cRec.delta = 1
                cRec.hedgeQty = CalcQtyNeededToHedge(cRec.symbol, cRec.delta, cRec.familyDelta)
                maxShort = MaxShortWithinConstraints(cRec.symbol, tDate)
                If cRec.hedgeQty > maxShort Then
                    cRec.qty = maxShort
                    adjustment = -50
                Else
                    cRec.qty = cRec.hedgeQty
                End If
                If cRec.qty > 0 Then
                    cRec.score = baseScore + adjustment
                    CandidateRecList.Add(cRec)
                End If
            End If
        End If
    End Sub

    Public Sub ScoreSellingPut(baseScore As Integer, recommendation As Transaction, tDate As Date)
        Dim adjustment As Double = 0
        Dim cRec As Transaction  ' this is a CANDIDATE rec, do not confuse with recommendation!
        For Each APRow As DataRow In myDataSet.Tables("AcquiredPositionsTbl").Rows
            cRec = New Transaction
            cRec.type = "Sell"
            cRec.symbol = APRow("Symbol").Trim()
            cRec.familyDelta = recommendation.familyDelta
            adjustment = 0
            If IsAPut(cRec.symbol) Then
                If IsInTheFamily(cRec.symbol, recommendation.familyTicker) Then
                    cRec.currentPositionInAP = APRow("Units")
                    If cRec.currentPositionInAP > 0 Then   ' to sell, you need to have 
                        cRec.delta = CalcDelta(cRec.symbol, tDate)
                        cRec.hedgeQty = CalcQtyNeededToHedge(cRec.symbol, cRec.delta, cRec.familyDelta)
                        If cRec.hedgeQty > 0 Then
                            If cRec.currentPositionInAP < cRec.hedgeQty Then
                                cRec.qty = cRec.currentPositionInAP ' sell all you have
                                adjustment = -50           ' because incomplete hedge
                            Else
                                cRec.qty = cRec.hedgeQty
                            End If
                            cRec.score = baseScore + adjustment
                            CandidateRecList.Add(cRec)
                        End If
                    End If
                End If
            End If
        Next
    End Sub

    Public Sub ScoreSellingShortPut(baseScore As Integer, recommendation As Transaction, tDate As Date)
        Dim adjustment, maxshort
        Dim cRec As Transaction
        If TooCloseToMaxMargins() Then
            Exit Sub ' we have no more credit
        End If
        For Each partialSymbol As String In {"_POCTE", "_POCTD", "_POCTC", "_POCTB", "_POCTA"}   ' only these options will be considered, in this order
            cRec = New Transaction                                        ' you might add/subtract To/from the list, for example adding the JULY otions
            cRec.type = "SellShort"
            cRec.symbol = recommendation.familyTicker + partialSymbol
            cRec.familyDelta = recommendation.familyDelta
            adjustment = 0
            If Not IsInIP(cRec.symbol) Then
                cRec.currentPositionInAP = GetCurrentPositionInAP(cRec.symbol)
                If cRec.currentPositionInAP <= 0 Then   ' because if long cannot sell short
                    cRec.delta = CalcDelta(cRec.symbol, tDate)
                    cRec.hedgeQty = CalcQtyNeededToHedge(cRec.symbol, cRec.delta, cRec.familyDelta)
                    maxshort = MaxShortWithinConstraints(cRec.symbol, tDate)
                    If cRec.hedgeQty > maxshort Then
                        cRec.qty = maxshort
                        adjustment = -50
                    Else
                        cRec.qty = cRec.hedgeQty
                    End If
                    If cRec.qty > 0 Then
                        Select Case partialSymbol
                            Case "_POCTE"
                                adjustment = adjustment + 5
                            Case "_POCTD"
                                adjustment = adjustment + 4
                            Case "_POCTC"
                                adjustment = adjustment + 3
                            Case "_POCTB"
                                adjustment = adjustment + 2
                            Case "_POCTA"
                                adjustment = adjustment + 1
                        End Select
                        cRec.score = baseScore + adjustment
                        CandidateRecList.Add(cRec)
                    End If
                End If
            End If
        Next
    End Sub

    Public Sub ScoreBuyingBackCall(baseScore As Integer, recommendation As Transaction, tDate As Date)
        Dim adjustment, maxbuy As Double
        Dim cRec As Transaction  ' this is a CANDIDATE rec, do not confuse with recommendation!
        If AvailableCashIsLow() Then
            Exit Sub
        End If
        For Each dr As DataRow In myDataSet.Tables("AcquiredPositionsTbl").Rows
            cRec = New Transaction
            cRec.familyDelta = recommendation.familyDelta
            cRec.type = "Buy"
            cRec.symbol = dr("Symbol").Trim()
            adjustment = 0
            If IsAStock(cRec.symbol) Or cRec.symbol = "CAccount" Then
                ' skip
            Else
                If IsACall(cRec.symbol) Then
                    cRec.underlier = GetUnderlier(cRec.symbol)
                    If GetUnderlier(cRec.symbol) = recommendation.familyTicker Then
                        cRec.currentPositionInAP = dr("Units")
                        If cRec.currentPositionInAP < 0 Then
                            cRec.delta = CalcDelta(cRec.symbol, tDate)
                            cRec.hedgeQty = CalcQtyNeededToHedge(cRec.symbol, cRec.delta, cRec.familyDelta)
                            If Math.Abs(cRec.currentPositionInAP) < cRec.hedgeQty Then
                                cRec.hedgeQty = Math.Abs(cRec.currentPositionInAP)    ' buy back all that you have
                                adjustment = -50
                            End If
                            maxbuy = MaxPurchasePossible(cRec.symbol, tDate)           ' how much can you afford?
                            If maxbuy < cRec.hedgeQty Then
                                cRec.qty = maxbuy
                                adjustment = adjustment - 50
                            Else
                                cRec.qty = cRec.hedgeQty
                            End If
                            If cRec.qty > 0 Then
                                cRec.score = baseScore + adjustment
                                CandidateRecList.Add(cRec)
                            End If
                        End If
                    End If
                End If
            End If
        Next
    End Sub

    Public Sub ScoreBuyingBackStock(baseScore As Integer, recommendation As Transaction, tDate As Date)
        Dim cRec As New Transaction  ' this is the CANDIDATE rec, do not confuse with recommendation!
        cRec.symbol = recommendation.familyTicker
        cRec.type = "Buy"
        cRec.familyDelta = recommendation.familyDelta
        Dim adjustment, maxBuy As Double

        If IsInIP(cRec.symbol) Then
            Exit Sub   ' cannot buy if in IP
        End If
        If AvailableCashIsLow() Then
            Exit Sub
        End If
        cRec.currentPositionInAP = GetCurrentPositionInAP(cRec.symbol)
        If cRec.currentPositionInAP < 0 Then                     ' buy back requires a short
            cRec.delta = 1
            cRec.hedgeQty = CalcQtyNeededToHedge(cRec.symbol, cRec.delta, cRec.familyDelta)
            maxBuy = MaxPurchasePossible(cRec.symbol, tDate)      ' how many can we afford?
            If maxBuy < cRec.hedgeQty Then
                cRec.hedgeQty = maxBuy
                adjustment = -50
            End If
            If cRec.hedgeQty > Math.Abs(cRec.currentPositionInAP) Then
                cRec.qty = Math.Abs(cRec.currentPositionInAP)               ' buy back all
                adjustment = -50
            Else
                cRec.qty = cRec.hedgeQty
            End If
            If cRec.qty > 0 Then
                cRec.score = baseScore + adjustment
                CandidateRecList.Add(cRec)
            End If
        End If
    End Sub

    Public Sub ScoreBuyingCall(baseScore As Integer, recommendation As Transaction, tDate As Date)
        Dim adjustment As Double = 0
        Dim maxBuy As Double = 0
        Dim cRec As Transaction  ' this is a CANDIDATE rec, do not confuse with recommendation!
        If AvailableCashIsLow() Then
            Exit Sub
        End If
        ' only considers OCT options - you can easily change this: talk to me
        For Each partialSymbol As String In {"_COCTE", "_COCTD", "_COCTC", "_COCTB", "_COCTA"}
            cRec = New Transaction
            cRec.type = "Buy"
            cRec.symbol = recommendation.familyTicker + partialSymbol
            cRec.familyDelta = recommendation.familyDelta
            adjustment = 0

            If Not IsInIP(cRec.symbol) Then
                cRec.currentPositionInAP = GetCurrentPositionInAP(cRec.symbol)
                If cRec.currentPositionInAP >= 0 Then                          ' if short is a buyback
                    cRec.delta = CalcDelta(cRec.symbol, tDate)
                    cRec.hedgeQty = CalcQtyNeededToHedge(cRec.symbol, cRec.delta, cRec.familyDelta)
                    maxBuy = MaxPurchasePossible(cRec.symbol, tDate)       ' how much can you afford?
                    If maxBuy < cRec.hedgeQty Then
                        cRec.qty = maxBuy
                        adjustment = -50
                    Else
                        cRec.qty = cRec.hedgeQty
                    End If
                    If cRec.qty > 0 Then
                        Select Case partialSymbol                          ' think about these scores
                            Case "_COCTE"
                                adjustment = adjustment + 1
                            Case "_COCTD"
                                adjustment = adjustment + 2
                            Case "_COCTC"
                                adjustment = adjustment + 3
                            Case "_COCTB"
                                adjustment = adjustment + 4
                            Case "_COCTA"
                                adjustment = adjustment + 5
                        End Select
                        cRec.score = baseScore + adjustment
                        CandidateRecList.Add(cRec)
                    End If
                End If
            End If
        Next
    End Sub

    Public Sub ScoreBuyingStock(baseScore As Integer, recommendation As Transaction, tDate As Date)
        Dim cRec As New Transaction  ' this is the CANDIDATE rec, do not confuse with recommendation!
        cRec.symbol = recommendation.familyTicker
        cRec.type = "Buy"
        cRec.familyDelta = recommendation.familyDelta
        Dim adjustment As Double = 0
        Dim maxBuy As Double

        If IsInIP(cRec.symbol) Then
            Exit Sub   ' cannot buy if in IP
        End If
        If AvailableCashIsLow() Then
            Exit Sub
        End If
        cRec.currentPositionInAP = GetCurrentPositionInAP(cRec.symbol)
        If cRec.currentPositionInAP < 0 Then     ' if short then we need a buyback
            Exit Sub
        Else
            cRec.delta = 1
            cRec.hedgeQty = CalcQtyNeededToHedge(cRec.symbol, cRec.delta, cRec.familyDelta)
            maxBuy = MaxPurchasePossible(cRec.symbol, tDate) ' how many can we afford?
            If maxBuy < cRec.hedgeQty Then
                cRec.qty = maxBuy
                adjustment = -50
            Else
                cRec.qty = cRec.hedgeQty
            End If
            If cRec.qty > 0 Then
                cRec.score = baseScore + adjustment
                CandidateRecList.Add(cRec)
            End If
        End If
    End Sub

End Module