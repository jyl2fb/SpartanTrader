Imports Microsoft.SolverFoundation.Common
Imports Microsoft.SolverFoundation.Solvers
Imports Microsoft.SolverFoundation.Services
Module DeltaGammaHedger

    Public Sub GetPotentialList(famtkr As String, targetdate As Date)
        IntermediaryRecList.Clear()

        For Each element In MasterRecList
            If targetdate.Month >= 7 Then
                If element.familyTicker = famtkr And ExpiresInJuly(element.symbol) = False Then
                    Dim var = element
                    var.delta = CalcDelta(var.symbol, targetdate)
                    var.gamma = CalcGamma(var.symbol, targetdate)
                    var.speed = CalcSpeed(var.symbol, targetdate)
                    var.mtm = CalcMTM(var.symbol, targetdate)
                    If var.mtm >= 0.02 Then
                        IntermediaryRecList.Add(var)
                    End If
                End If
            Else
                If element.familyTicker = famtkr Then
                    Dim var = element
                    var.delta = CalcDelta(var.symbol, targetdate)
                    var.gamma = CalcGamma(var.symbol, targetdate)
                    var.speed = CalcSpeed(var.symbol, targetdate)
                    var.mtm = CalcMTM(var.symbol, targetdate)
                    If var.mtm >= 0.02 Then
                        IntermediaryRecList.Add(var)
                    End If
                End If
            End If
        Next
        'For Each myRow As DataRow In myDataSet.Tables("TickersTbl").Rows
        '    If CheckAvailability(myRow("Ticker"), targetdate) Then
        '        If myRow("Ticker").ToString.Trim() = famtkr Then
        '            'mylist.Add(myRow("Ticker").ToString.Trim())
        '            Dim tempTransaction As New Transaction
        '            tempTransaction.symbol = famtkr
        '            tempTransaction.familyTicker = famtkr
        '            tempTransaction.delta = 1
        '            tempTransaction.gamma = 0
        '            tempTransaction.mtm = CalcMTM(famtkr, currentDate)
        '            IntermediaryRecList.Add(tempTransaction)
        '        End If
        '    End If
        'Next
        'For Each myRow As DataRow In myDataSet.Tables("SymbolsTbl").Rows
        '    If CheckAvailability(myRow("Symbol"), targetdate) Then
        '        If IsInTheFamily(myRow("Symbol").ToString.Trim(), famtkr) Then
        '            'mylist.Add(myRow("Symbol").ToString.Trim())
        '            Dim tempTransaction As New Transaction
        '            tempTransaction.symbol = myRow("Symbol").ToString.Trim()
        '            tempTransaction.familyTicker = famtkr
        '            tempTransaction.delta = CalcDelta(myRow("Symbol").ToString.Trim(), targetdate)
        '            tempTransaction.gamma = CalcGamma(myRow("Symbol").ToString.Trim(), targetdate)
        '            tempTransaction.mtm = CalcMTM(myRow("Symbol").ToString.Trim(), targetdate)
        '            IntermediaryRecList.Add(tempTransaction)
        '        End If
        '    End If
        'Next
    End Sub
    Public Sub SellJulyOptions(targetdate As Date)
        Dim symbol As String
        Dim units As Double

        Dim TempRecList As New List(Of Transaction)
        TempRecList.Clear()

        For Each myRow As DataRow In myDataSet.Tables("AcquiredPositionsTbl").Rows
            Dim temptransaction As New Transaction
            symbol = myRow("Symbol").ToString().Trim
            units = myRow("Units")
            If IsAnOption(symbol) And ExpiresInJuly(symbol) Then
                temptransaction.symbol = symbol
                temptransaction.weight = 0
                temptransaction.type = "WTF"
                TempRecList.Add(temptransaction)
            End If
        Next
        GetSolvedTransaction(TempRecList)
        julyoptionssold = True
    End Sub
    Public Sub FillMasterList()
        For Each myRow As DataRow In myDataSet.Tables("TickersTbl").Rows
            If IsInIP(myRow("Ticker")) = False Then
                Dim tempTransaction As New Transaction
                tempTransaction.symbol = myRow("Ticker").ToString.Trim()
                tempTransaction.familyTicker = myRow("Ticker").ToString.Trim()
                tempTransaction.type = "Hold"
                MasterRecList.Add(tempTransaction)
            End If
        Next

        For Each myRow As DataRow In myDataSet.Tables("SymbolsTbl").Rows
            If IsInIP(myRow("Symbol")) = False Then
                Dim tempTransaction2 As New Transaction
                tempTransaction2.symbol = myRow("Symbol").ToString.Trim()
                tempTransaction2.familyTicker = myRow("Underlier").ToString.Trim()
                tempTransaction2.type = "Hold"
                MasterRecList.Add(tempTransaction2)
            End If
        Next
    End Sub

    Public Function CheckAvailability(symbol As String, targetdate As Date) As Boolean
        If symbol = "CAccount" Then
            Return False
        End If

        If IsAStock(symbol) Then
            Return True
        End If

        If targetdate.AddDays(14) >= GetExpiration(symbol).Date Then
            Return False
        End If

        If GetAsk(symbol, currentDate) = 0 Then
            Return False
        End If

        Return True
    End Function

    Public Sub SolvePotential(potentialList As List(Of Transaction), targetDate As Date, famdelta As Double, famgamma As Double, famspeed As Double)
        Dim TrackingError = TPV - TaTPV
        Dim solver = SolverContext.GetContext()
        solver.ClearModel()
        Dim model = solver.CreateModel()
        Dim decisions = potentialList.[Select](Function(it) New Decision(Domain.Real, it.symbol))
        Dim tempname As String
        Dim Tvariable(potentialList.Count - 1)
        For i = 0 To potentialList.Count() - 1
            Dim tempcounter = i
            tempname = "T" + tempcounter.ToString()
            Tvariable(i) = tempname
            Dim tempDecision = New Decision(Domain.Real, tempname)
            decisions = decisions.Append(tempDecision)
        Next
        model.AddDecisions(decisions.ToArray())

        Dim objective = New SumTermBuilder(potentialList.Count())

        For Each t In Tvariable
            Dim TDecision = model.Decisions.First(Function(it) it.Name = t)
            objective.Add(TDecision)
        Next

        model.AddGoal("MinimizeT", GoalKind.Minimize, objective.ToTerm())

        Dim deltacomponent = New SumTermBuilder(potentialList.Count())
        For Each potential In potentialList
            Dim deltasum = model.Decisions.First(Function(it) it.Name = potential.symbol)
            deltacomponent.Add(deltasum * potential.delta)
        Next

        Dim deltaconstraint = deltacomponent.ToTerm() = -1 * famdelta
        model.AddConstraint("Delta", deltaconstraint)

        Dim gammacomponent = New SumTermBuilder(potentialList.Count())
        For Each potential In potentialList
            Dim gammasum = model.Decisions.First(Function(it) it.Name = potential.symbol)
            gammacomponent.Add(gammasum * potential.gamma)
        Next
        Dim gammaconstraint = gammacomponent.ToTerm() = -1 * famgamma
        model.AddConstraint("Gamma", gammaconstraint)

        'If TrackingError >= 100000 Then ' CHANGE HERE
        Dim speedcomponent = New SumTermBuilder(potentialList.Count())
            For Each potential In potentialList
                Dim speedsum = model.Decisions.First(Function(it) it.Name = potential.symbol)
                speedcomponent.Add(speedsum * potential.speed)
            Next
            Dim speedconstraint = speedcomponent.ToTerm() = -1 * famspeed
            model.AddConstraint("Speed", speedconstraint)
        'End If

        For var = 0 To potentialList.Count - 1
            Dim i = var
            Dim qvalue = model.Decisions.First(Function(it) it.Name = potentialList(i).symbol)
            Dim tvalue = model.Decisions.First(Function(it) it.Name = Tvariable(i))
            Dim qconstraint = qvalue * potentialList(i).mtm <= tvalue
            Dim qconstraintneg = -1 * qvalue * potentialList(i).mtm <= tvalue
            Dim marginconstraint = qvalue >= Math.Min(0, GetCurrentPositionInAP(potentialList(i).symbol))
            model.AddConstraint("TP" + i.ToString(), qconstraint)
            model.AddConstraint("TN" + i.ToString(), qconstraintneg)

            If TooCloseToMaxMargins() Then
                model.AddConstraint("margin" + i.ToString(), marginconstraint)
            End If
        Next



        Dim solution = solver.Solve()
        If (solution.Quality = SolverQuality.Optimal) Then
            For Each potentee In potentialList
                Dim decision = model.Decisions.First(Function(it) it.Name = potentee.symbol)
                potentee.weight = decision.ToDouble()
                potentee.familyDelta = famdelta
                potentee.familyGamma = famgamma
                potentee.familySpeed = famspeed
            Next
        End If

    End Sub

    Public Sub GetSolvedTransaction(reclist As List(Of Transaction))
        For Each transaction In reclist
            Dim appos = GetCurrentPositionInAP(transaction.symbol)
            If transaction.weight = appos Then
                transaction.type = "Hold"
                transaction.qty = 0
            End If

            If transaction.weight > appos Then ' we buy
                transaction.type = "Buy"
                If transaction.weight >= 0 And appos <= 0 Then
                    transaction.qty = Math.Abs(appos)
                Else
                    transaction.qty = Math.Abs(transaction.weight - appos)
                End If
            ElseIf transaction.weight < appos Then
                If transaction.weight >= 0 And appos >= 0 Then
                    transaction.type = "Sell"
                    transaction.qty = Math.Abs(appos - transaction.weight)
                ElseIf transaction.weight <= 0 And appos > 0 Then
                    transaction.type = "Sell"
                    transaction.qty = Math.Abs(appos)
                Else
                    transaction.type = "SellShort"
                    transaction.qty = Math.Abs(appos - transaction.weight)
                End If
            End If

            If transaction.qty <> 0 Then
                FinalRecList.Add(transaction)
            End If
        Next
    End Sub
    Public Function ExpiresInJuly(symbol As String) As Boolean
        If IsAStock(symbol) Or symbol = "CAccount" Then
            Return False
        End If
        If GetExpiration(symbol).Month = 7 Then
            Return True
        End If
        Return False
    End Function
End Module
