Module portfolio


    Public Function CalcMTM(symbol As String, targetDate As Date) As Double
        If symbol = "CAccount" Then
            Return 1
        Else
            Return (GetAsk(symbol, targetDate) + GetBid(symbol, targetDate)) / 2
        End If
    End Function
    Public Function CalcAPValue(targetDate As Date) As Double  ' edited
        Dim cumulativeAPValue As Double = 0
        Dim symbol As String
        Dim units As Double
        Dim posValue As Double

        For Each myRow As DataRow In myDataSet.Tables("AcquiredPositionsTbl").Rows
            symbol = myRow("Symbol").ToString().Trim
            units = myRow("Units")
            If symbol <> "CAccount" Then
                posValue = units * CalcMTM(symbol, targetDate)
                cumulativeAPValue += posValue
                myRow("Value") = posValue
            End If
        Next

        Return cumulativeAPValue
    End Function
    Public Function CalcIPValue(targetDate As Date) As Double
        Dim cumulativeIPValue As Double = 0
        Dim symbol As String
        Dim units As Double
        Dim posValue As Double

        For Each myRow As DataRow In myDataSet.Tables("InitialPositionsTbl").Rows
            symbol = myRow("Symbol").ToString().Trim
            units = myRow("Units")
            If symbol <> "CAccount" Then
                posValue = units * CalcMTM(symbol, targetDate)
                cumulativeIPValue = cumulativeIPValue + posValue
                myRow("Value") = posValue
            End If
        Next

        Return cumulativeIPValue
    End Function

    Public Function CalcMargin(targetDate As Date) As Double
        Return CalcAPMargin(targetDate) + CalcIPMargin(targetDate)
    End Function

    Public Function CalcAPMargin(targetDate As Date) As Double
        Dim cumulativeAPMValue As Double = 0
        Dim symbol As String
        Dim units As Double
        Dim posValue As Double
        For Each myRow As DataRow In myDataSet.Tables("AcquiredPositionsTbl").Rows
            symbol = myRow("Symbol").ToString().Trim
            units = myRow("Units")
            If symbol <> "CAccount" And units < 0 Then
                posValue = units * CalcMTM(symbol, targetDate)
                cumulativeAPMValue += posValue
            End If
        Next
        Return cumulativeAPMValue
    End Function

    Public Function CalcIPMargin(targetDate As Date) As Double
        Dim cumulativeIPMValue As Double = 0
        Dim symbol As String
        Dim units As Double
        Dim posValue As Double
        For Each myRow As DataRow In myDataSet.Tables("InitialPositionsTbl").Rows
            symbol = myRow("Symbol").ToString().Trim
            units = myRow("Units")
            If symbol <> "CAccount" And units < 0 Then
                posValue = units * CalcMTM(symbol, targetDate)
                cumulativeIPMValue += posValue
            End If
        Next
        Return cumulativeIPMValue
    End Function

    Public Function CalcTPVAtStart() As Double
        Return CalcIPValue(startDate) + initialCAccount
    End Function

    Public Function CalcTaTPV(targetDate As Date) As Double
        Dim ts As TimeSpan = targetDate.Date - startDate.Date
        Dim t As Double = ts.Days / 365.25
        Return TPVatStart * Math.Exp(riskFreeRate * t)
    End Function

    Public Function IsInIP(x As String) As Boolean
        x = x.Trim()
        For Each myRow As DataRow In myDataSet.Tables("InitialPositionsTbl").Rows
            If myRow("Symbol").trim() = x Then
                Return True
            End If
        Next
        Return False
    End Function

    Public Function CalcTPV(targetDate As Date) As Double
        Return IPvalue + APvalue + CAccount + CalcInterestSLT(currentDate)
    End Function

    Public Function CalcInterestSLT(toThisDay As Date)
        Dim ts As TimeSpan = toThisDay.Date - lastTransactionDate.Date
        Dim t As Double = ts.Days / 365.25
        Return CAccount * (Math.Exp(riskFreeRate * t) - 1)
    End Function

    Public Function CalcTE() As Double
        If TPV >= TaTPV Then
            Return (TPV - TaTPV) * 0.25
        Else
            Return TaTPV - TPV
        End If
    End Function

    Public Function UpdateSumTE(targetDate As Date) As Double
        If targetDate.DayOfWeek = DayOfWeek.Sunday And targetDate.Date > lastTEUpDate.Date Then
            lastTEUpDate = targetDate
            Return TE
        Else
            Return 0
        End If
    End Function

    Public Sub Execute(t As Transaction)
        Dim mySQL As String
        mySQL = String.Format("INSERT INTO TransactionQueue (date, TeamID, SYmbol, Type, Qty, Price, Cost, TotValue, " +
                              "InterestSinceLastTransaction, CashPositionAfterTransaction, TotMargin) VALUES " +
                              "('{0}', {1}, '{2}', '{3}', {4}, {5}, {6}, {7}, {8}, {9}, {10})",
                              currentDate.ToShortDateString,
                              teamID,
                              t.symbol,
                              t.type,
                              t.qty,
                              t.price,
                              t.tcost,
                              t.totValue,
                              t.interestSLT,
                              t.CAccountAT,
                              t.marginAT)
        RunNonQuery(mySQL)
        lastTransactionDate = currentDate
        CAccount = t.CAccountAT
        margin = t.marginAT
        t.UpdatePosition()

    End Sub

    Public Sub RoboExecuteAll(tdate As Date)

        Globals.ThisWorkbook.Application.ScreenUpdating = False
        Dim sortlist = FinalRecList.OrderByDescending(Function(x) x.type).ToList()

        'For i = 0 To 11
        '    Recommendations(i).MarkAsDone()
        '    Globals.Dashboard.DisplayRecommendation(i)
        'Next
        'For i = 0 To 11
        '    CalcRecommendation(i, tdate)
        '    Globals.Dashboard.DisplayRecommendation(i)
        '    RoboExecuteRec(i, tdate)
        '    Application.DoEvents()
        'Next
        'For i = 0 To 11
        '    FinalRecList(i).MarkAsDone()
        '    Globals.Dashboard.DisplayRecommendation(i)
        'Next
        'I want to short sell before anything else in case I need the margin.
        'For i = 0 To FinalRecList.Count() - 1
        '    If FinalRecList(i).type = "SellShort" Then
        '        RoboExecuteRec(i, tdate)
        '    End If
        'Next

        'For i = 0 To FinalRecList.Count() - 1
        '    If FinalRecList(i).type <> "SellShort" Then
        '        RoboExecuteRec(i, tdate)
        '    End If
        'Next

        For i = 0 To sortlist.Count - 1
            RoboExecuteRec(i, tdate)
        Next

        Globals.ThisWorkbook.Application.ScreenUpdating = True
    End Sub

    Public Sub RoboExecuteStepByStep(tdate As Date)
        For i = 0 To 11
            CalcRecommendation(i, tdate)
            Globals.Dashboard.DisplayRecommendation(i)
            If MessageBox.Show(RecommendationFamily(i) + " is next.",
                               "Simulation Mode", MessageBoxButtons.OKCancel,
                               MessageBoxIcon.Stop) = DialogResult.OK Then
                RoboExecuteRec(i, tdate)
                For j = i To 11
                    CalcRecommendation(j, tdate)
                    Globals.Dashboard.DisplayRecommendation(j)
                Next
            Else
                traderMode = "Manual"
                Exit Sub
            End If
            Application.DoEvents()
        Next
    End Sub

    Public Sub RoboExecuteRec(i As Integer, tdate As Date)
        'If Recommendations(i).type <> "Hold" Then
        '    If IsValid(Recommendations(i)) Then
        '        Execute(Recommendations(i))
        '        CalcFinancialMetrics(currentDate)
        '        DisplayFinancialMetrics(currentDate)
        '    End If
        'End If

        If FinalRecList(i).type <> "Hold" Then
            If IsValid(FinalRecList(i)) Then
                FinalRecList(i).CalcTransactionProperties(tdate)
                Execute(FinalRecList(i))
                CalcFinancialMetrics(currentDate)
                'DisplayFinancialMetrics(currentDate)
            End If
        End If
    End Sub

End Module
