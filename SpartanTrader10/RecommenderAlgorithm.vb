﻿Module RecommenderAlgorithm
    Public Sub ResetAllRecommendations()
        'Dim ticker As String
        'For i = 0 To 11
        '    ticker = myDataSet.Tables("TickersTbl").Rows(i)("Ticker")
        '    RecommendationFamily(i) = New Transaction
        '    RecommendationFamily(i).familyTicker = ticker.Trim()
        '    Globals.Dashboard.Range("I4").Offset(i, 0).Value = ticker
        '    Globals.Dashboard.Range("J4").Offset(i, 0).Value = GetVol(ticker)
        'Next
        FinalRecList = New List(Of Transaction)
        FinalRecList.Clear()
    End Sub

    Public Sub CalcAllRecommendations(targetDate As Date)
        For i = 0 To 11
            CalcRecommendation(i, targetDate)
        Next

        If julyoptionssold = False Then
            If targetDate.Month = 7 Then
                SellJulyOptions(targetDate)
            End If
        End If

    End Sub

    Public Sub CalcRecommendation(i As Integer, targetDate As Date)
        Dim famtkr As String
        Dim famdelta, ipfamdelta, famgamma, ipfamgamma As Double
        famtkr = RecommendationFamily(i).Trim()
        famdelta = CalcFamilyDelta(famtkr, targetDate, False)
        ipfamdelta = CalcFamilyDelta(famtkr, targetDate, True) + deltaAdjustment
        famgamma = CalcFamilyGamma(famtkr, targetDate, False)
        ipfamgamma = CalcFamilyGamma(famtkr, targetDate, True)

        IntermediaryRecList = New List(Of Transaction)
        IntermediaryRecList.Clear()

        If HedgingToday(targetDate) = True Then
            'intreclisthere?
            If NeedToHedge(famdelta, famgamma, targetDate) = True Then
                'clear potentialist
                GetPotentialList(famtkr, targetDate)
                'Potential Functions
                'FillPotential(IntermediaryRecList, targetDate)
                SolvePotential(IntermediaryRecList, targetDate, ipfamdelta, ipfamgamma)
                GetSolvedTransaction(IntermediaryRecList)
                'CalcCandidateRecScores(Recommendations(i), targetDate)
                'FindBestCandidateRec(Recommendations(i), targetDate)
                Application.DoEvents()
            End If
        End If
    End Sub

    Public Function HedgingToday(targetDate As Date) As Boolean

        If targetDate.DayOfWeek = DayOfWeek.Saturday Or targetDate.DayOfWeek = DayOfWeek.Sunday Then
            Return False
        End If
        'If TPV - TaTPV > 0 Then
        '    Return True
        'End If

        'If TPV - TaTPV < -25000 Then
        '    Return True
        'End If

        Return True
    End Function

    Public Function NeedToHedge(famdelta As Double, famgamma As Double, targetDate As Date) As Boolean
        Dim difference As Double = TPV - TaTPV
        If difference > 0 Then
            If Math.Abs(famdelta) > 1000 Or Math.Abs(famgamma) > 250 Then
                Return True
            End If
            Return False
        ElseIf difference < 5000 Then
            If Math.Abs(famdelta) > 1000 Or Math.Abs(famgamma) > 250 Then
                Return True
            End If
            Return False
        End If
        Return False
    End Function

    Public Sub CalcCandidateRecScores(rec As Transaction, targetDate As Date)
        If rec.familyDelta > 0 Then
            ' baseScores are the key parameters of the smart trader 
            ' the baseScores are *intentionally suboptimal-. Improve them! 
            ScoreSellingStock(200, rec, targetDate)
            ScoreSellingCall(400, rec, targetDate)
            ScoreBuyingPut(500, rec, targetDate)
            ScoreBuyingBackPut(300, rec, targetDate)
            ScoreSellingShortCall(200, rec, targetDate)
            ScoreSellingShortStock(800, rec, targetDate)
        Else ' famdelta < 0 
            ScoreSellingPut(800, rec, targetDate)
            ScoreBuyingBackCall(400, rec, targetDate)
            ScoreBuyingBackStock(200, rec, targetDate)
            ScoreBuyingCall(500, rec, targetDate)
            ScoreSellingShortPut(700, rec, targetDate)
            ScoreBuyingStock(100, rec, targetDate)
        End If
    End Sub

    Public Function CalcQtyNeededToHedge(sym As String, delta As Double, familyDelta As Double) As Integer
        Dim familyDeltaTarget As Double = 0
        Dim q As Double

        If Math.Abs(delta) < 0.05 Then
            Return 0
        Else
            q = (familyDeltaTarget - familyDelta) / delta
            Return Math.Abs(Math.Round(q))
        End If
    End Function

    Public Function AvailableCashIsLow() As Boolean
        Dim availableCash As Double = CAccount - (Math.Abs(margin) * 0.3)
        If availableCash < 1000000 Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Function MaxPurchasePossible(sym As String, tdate As Date) As Double
        Dim ask As Double = 0
        Dim q As Double = 0
        Dim availableCash = 0
        availableCash = CAccount - (Math.Abs(margin) * 0.3) - 1000000
        ask = GetAsk(sym, tdate)
        If availableCash > 0 And ask > 0 Then
            q = availableCash / ask
            Return Math.Truncate(q)
        Else
            Return 0
        End If
    End Function

    Public Sub FindBestCandidateRec(rec As Transaction, targetDate As Date)
        Dim bestScoreSoFar As Double = -1000
        If MasterRecList.Count = 0 Then
            Exit Sub
        End If

        For Each cr As Transaction In MasterRecList
            If cr.score > bestScoreSoFar Then
                rec.type = cr.type
                rec.qty = cr.qty
                rec.symbol = cr.symbol
                bestScoreSoFar = cr.score
            End If
        Next

        If bestScoreSoFar > -1000 Then
            rec.CalcTransactionProperties(targetDate)
        End If
    End Sub

    Public Function TooCloseToMaxMargins() As Boolean
        If ((maxMargin - Math.Abs(margin)) < 2000000) Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Function MaxShortWithinConstraints(sym As String, tdate As Date) As Double
        Dim q As Double = 0
        Dim maxAllowableIncreaseInMargins As Double = 0
        If TooCloseToMaxMargins() Then
            Return 0
        Else
            maxAllowableIncreaseInMargins = (maxMargin - Math.Abs(margin)) - 1000000
            If maxAllowableIncreaseInMargins <= 0 Then
                Return 0
            Else
                q = maxAllowableIncreaseInMargins / GetBid(sym, tdate)
                Return Math.Truncate(q)
            End If
        End If
    End Function
End Module

