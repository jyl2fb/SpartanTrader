Imports Microsoft.SolverFoundation.Common
Imports Microsoft.SolverFoundation.Solvers
Imports Microsoft.SolverFoundation.Services
Module DeltaGammaHedger

    Public Sub GetPotentialList(famtkr As String, targetdate As Date)
        Dim mylist As New List(Of String)
        IntermediaryRecList.Clear()
        For Each myRow As DataRow In myDataSet.Tables("TickersTbl").Rows
            If CheckAvailability(myRow("Ticker"), targetdate) Then
                If myRow("Ticker").ToString.Trim() = famtkr Then
                    'mylist.Add(myRow("Ticker").ToString.Trim())
                    Dim tempTransaction As New Transaction
                    tempTransaction.symbol = famtkr
                    tempTransaction.familyTicker = famtkr
                    tempTransaction.delta = 1
                    tempTransaction.gamma = 0
                    tempTransaction.mtm = CalcMTM(famtkr, currentDate)
                    IntermediaryRecList.Add(tempTransaction)
                End If
            End If
        Next
        For Each myRow As DataRow In myDataSet.Tables("SymbolsTbl").Rows
            If CheckAvailability(myRow("Symbol"), targetdate) Then
                If IsInTheFamily(myRow("Symbol").ToString.Trim(), famtkr) Then
                    'mylist.Add(myRow("Symbol").ToString.Trim())
                    Dim tempTransaction As New Transaction
                    tempTransaction.symbol = myRow("Symbol").ToString.Trim()
                    tempTransaction.familyTicker = famtkr
                    tempTransaction.delta = CalcDelta(myRow("Symbol").ToString.Trim(), targetdate)
                    tempTransaction.gamma = CalcGamma(myRow("Symbol").ToString.Trim(), targetdate)
                    tempTransaction.mtm = CalcMTM(myRow("Symbol").ToString.Trim(), targetdate)
                    IntermediaryRecList.Add(tempTransaction)
                End If
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

        If targetdate >= GetExpiration(symbol).Date Then
            Return False
        End If

        If GetAsk(symbol, currentDate) = 0 Then
            Return False
        End If

        Return True
    End Function

    Public Sub FillPotential(potentialList As List(Of String), targetDate As Date)
        Dim zerofiller As Double = 0.0
        Globals.Dashboard.Range("AM4", "AR35").Clear()
        For i = 0 To (potentialList.Count - 1)
            Globals.Dashboard.Range("AM4").Offset(i, 0).Value = potentialList(i).ToString().Trim()
            Globals.Dashboard.Range("AM4").Offset(i, 1).Value = CalcDelta(potentialList(i).ToString, targetDate)
            Globals.Dashboard.Range("AM4").Offset(i, 2).Value = CalcGamma(potentialList(i).ToString(), targetDate)
            Globals.Dashboard.Range("AM4").Offset(i, 3).Value = CalcMTM(potentialList(i).ToString(), targetDate)
            Globals.Dashboard.Range("AM4").Offset(i, 4).Value = zerofiller
            Globals.Dashboard.Range("AM4").Offset(i, 5).Value = zerofiller
        Next
    End Sub

    Public Sub SolvePotential(potentialList As List(Of Transaction), targetDate As Date, famdelta As Double, famgamma As Double)

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
        'For i = 0 To potentialList.Count() - 1
        'tempname = "T" + i
        'Tvariable(i) = tempname
        'decisionx = New Decision(Domain.Real, tempname)
        'Next
        'Dim decisions2 = Tvariable.[Select](Function(it) New Decision(Domain.Real, it.ToString))
        'decisions = decisions1.Concat(decisions2)
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

        For var = 0 To potentialList.Count - 1
            Dim i = var
            Dim qvalue = model.Decisions.First(Function(it) it.Name = potentialList(i).symbol)
            Dim tvalue = model.Decisions.First(Function(it) it.Name = Tvariable(i))
            Dim qconstraint = qvalue * potentialList(i).mtm <= tvalue
            Dim qconstraintneg = -1 * qvalue * potentialList(i).mtm <= tvalue
            model.AddConstraint("TP" + i.ToString(), qconstraint)
            model.AddConstraint("TN" + i.ToString(), qconstraintneg)
        Next

        Dim solution = solver.Solve()
        System.Diagnostics.Debug.WriteLine(solution.GetReport())

        If (solution.Quality = SolverQuality.Optimal) Then
            For Each potentee In potentialList
                Dim decision = model.Decisions.First(Function(it) it.Name = potentee.symbol)
                potentee.weight = decision.ToDouble()
            Next
        End If


        'Dim solver As New model
        'Dim tempDic = CreateObject("Scripting.Dictionary")
        'Dim tempDic2 = CreateObject("Scripting.Dictionary")
        'Dim delta, gamma, tvariable As Double
        'Dim currentWeightstr, CurrentTstr As String

        'solver.AddRow("delta", delta)
        'solver.AddRow("gamma", gamma)
        'solver.AddRow("tvariable", tvariable)

        'For i = 0 To potentialList.Count - 1

        '    currentWeightstr = "Weight" + i
        '    CurrentTstr = "Dummy" + i


        '    solver.AddVariable(currentWeightstr, tempDic(i))
        '    'solver.AddVariable(CurrentTstr, CurrentT)
        '    'solver.SetBounds(CurrentT, CalcMTM(potentialList(i), targetDate) *
        '    solver.SetCoefficient(delta, tempDic(i), CalcDelta(potentialList(i).ToString, targetDate))
        '    solver.SetCoefficient(gamma, tempDic(i), CalcGamma(potentialList(i).ToString, targetDate))


        'Next


        'solver.SetBounds(delta, -1 * famdelta, -1 * famdelta)
        'solver.SetBounds(gamma, -1 * famgamma, -1 * famgamma)
        'solver.AddGoal(tvariable, 1, True)

        'clsBenchSphere
    End Sub
End Module
