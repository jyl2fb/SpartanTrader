Imports Microsoft.SolverFoundation.Common
Imports Microsoft.SolverFoundation.Solvers
Imports Microsoft.SolverFoundation.Services
Module DeltaGammaHedger

    Public Sub GetPotentialList(famtkr As String, targetdate As Date)

        Dim templist As List(Of Transaction)
        Dim familyDelta = CalcFamilyDelta(famtkr, targetdate, False)
        Dim familyGamma = CalcFamilyGamma(famtkr, targetdate, False)
        Dim familySpeed = CalcFamilySpeed(famtkr, targetdate, False)
        Dim ipfamilydelta = CalcFamilyDelta(famtkr, targetdate, True) + deltaAdjustment
        Dim ipfamilygamma = CalcFamilyGamma(famtkr, targetdate, True)
        Dim ipfamilyspeed = CalcFamilySpeed(famtkr, targetdate, True)
        If NeedToHedge(familyDelta, familyGamma, targetdate) Then

            If targetdate.Month >= 7 Then
                templist = MasterRecList.FindAll(Function(p) p.familyTicker.Trim() = famtkr.Trim() And ExpiresInJuly(p.symbol) = False)
            Else
                templist = MasterRecList.FindAll(Function(p) p.familyTicker.Trim() = famtkr.Trim())
            End If

            For Each var In templist
                var.delta = CalcDelta(var.symbol, targetdate)
                var.gamma = CalcGamma(var.symbol, targetdate)
                var.speed = CalcSpeed(var.symbol, targetdate)
                var.mtm = CalcMTM(var.symbol, targetdate)
                var.familyDelta = familyDelta
                var.ipfamilydelta = ipfamilydelta
                var.familyGamma = familyGamma
                var.ipfamilygamma = ipfamilygamma
                var.familySpeed = familySpeed
                var.ipfamilyspeed = ipfamilyspeed
                If var.mtm >= 0.01 Then
                    IntermediaryRecList.Add(var)
                End If
            Next

        End If
        'For Each element In MasterRecList
        '    If targetdate.Month >= 7 Then
        '        If element.familyTicker = famtkr And ExpiresInJuly(element.symbol) = False Then
        '            Dim var = element
        '            var.delta = CalcDelta(var.symbol, targetdate)
        '            var.gamma = CalcGamma(var.symbol, targetdate)
        '            var.speed = CalcSpeed(var.symbol, targetdate)
        '            var.mtm = CalcMTM(var.symbol, targetdate)
        '            If var.mtm >= 0.02 Then
        '                IntermediaryRecList.Add(var)
        '            End If
        '        End If
        '    Else
        '        If element.familyTicker = famtkr Then
        '            Dim var = element
        '            var.delta = CalcDelta(var.symbol, targetdate)
        '            var.gamma = CalcGamma(var.symbol, targetdate)
        '            var.speed = CalcSpeed(var.symbol, targetdate)
        '            var.mtm = CalcMTM(var.symbol, targetdate)
        '            If var.mtm >= 0.02 Then
        '                IntermediaryRecList.Add(var)
        '            End If
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

    Public Sub SolvePotential(potentialList As List(Of Transaction), targetDate As Date)
        Dim solver = SolverContext.GetContext()
        solver.ClearModel()
        Dim model = solver.CreateModel()
        Dim controls As New Directive With {
            .TimeLimit = 10000
        }
        'Dim tempname2 As String
        'Dim myval = 0
        'Dim Yvariable(potentialList.Count - 1)
        'Dim decisions2 As List(Of Decision) = Nothing
        'For i = 0 To potentialList.Count() - 1
        '    Dim tempcounter = i
        '    tempname2 = "Y" + tempcounter.ToString()
        '    Yvariable(i) = tempname2
        '    Dim tempDecision2 = New Decision(Domain.Real, tempname2)
        '    model.AddDecision(tempDecision2)
        'Next

        'Dim ycomponent As New SumTermBuilder(potentialList.Count() - 1)

        For Each family In RecommendationFamily
            Dim currentlist = potentialList.FindAll(Function(x) x.familyTicker.Trim() = family.Trim())
            If currentlist.Count > 0 Then
                Dim decisions = currentlist.[Select](Function(it) New Decision(Domain.Real, it.symbol))
                Dim tempname As String
                Dim famdelta = currentlist(0).ipfamilydelta
                Dim famgamma = currentlist(0).ipfamilygamma
                Dim famspeed = currentlist(0).ipfamilyspeed

                Dim Tvariable(currentlist.Count - 1)
                For i = 0 To currentlist.Count() - 1
                    Dim tempcounter = i
                    tempname = "T" + family + tempcounter.ToString()
                    Tvariable(i) = tempname
                    Dim tempDecision = New Decision(Domain.Real, tempname)
                    decisions = decisions.Append(tempDecision)
                Next

                model.AddDecisions(decisions.ToArray())

                Dim objective = New SumTermBuilder(20)

                For Each t In Tvariable
                    Dim TDecision = model.Decisions.First(Function(it) it.Name = t)
                    objective.Add(TDecision)
                Next

                model.AddGoal("MinimizeT" + family, GoalKind.Minimize, objective.ToTerm())
                Dim deltacomponent = New SumTermBuilder(currentlist.Count - 1)
                Dim gammacomponent = New SumTermBuilder(currentlist.Count - 1)


                For Each potential In currentlist
                    Dim weight = model.Decisions.First(Function(it) it.Name = potential.symbol)
                    deltacomponent.Add(weight * potential.delta)
                    gammacomponent.Add(weight * potential.gamma)
                Next

                Dim deltaconstraint = deltacomponent.ToTerm() = -1 * famdelta
                model.AddConstraint("Delta" + family, deltaconstraint)


                Dim gammaconstraint = gammacomponent.ToTerm() = -1 * famgamma

                If Math.Abs(margin) <= marginline Then
                    model.AddConstraint("Gamma" + family, gammaconstraint)
                End If

                If highenoughline Then ' CHANGE HERE
                    Dim speedcomponent = New SumTermBuilder(potentialList.Count())
                    For Each potential In currentlist
                        Dim speedsum = model.Decisions.First(Function(it) it.Name = potential.symbol)
                        speedcomponent.Add(speedsum * potential.speed)
                    Next
                    Dim speedconstraint = speedcomponent.ToTerm() = -1 * famspeed
                    model.AddConstraint("Speed" + family, speedconstraint)
                End If

                For var = 0 To currentlist.Count - 1
                    Dim i = var
                    Dim qvalue = model.Decisions.First(Function(it) it.Name = currentlist(i).symbol)
                    Dim tvalue = model.Decisions.First(Function(it) it.Name = Tvariable(i))
                    'Dim yvalue = model.Decisions.First(Function(it) it.Name = Yvariable(myval))
                    Dim qconstraint = qvalue * currentlist(i).mtm <= tvalue
                    Dim qconstraintneg = -1 * qvalue * currentlist(i).mtm <= tvalue
                    'Dim yconstraint = yvalue >= 0
                    'Dim yconstraintneg = yvalue >= -2 * qvalue * currentlist(i).mtm
                    'ycomponent.Add(yvalue)

                    model.AddConstraint("TP" + family + i.ToString(), qconstraint)
                    model.AddConstraint("TN" + family + i.ToString(), qconstraintneg)
                    'model.AddConstraint("YP" + family + i.ToString(), yconstraint)
                    'model.AddConstraint("YN" + family + i.ToString(), yconstraintneg)

                    'myval += 1
                Next
            End If
        Next

        'Dim yconstraintgeneral = ycomponent.ToTerm <= 30000000
        'model.AddConstraint("YNumberConstraint", yconstraintgeneral)



        Dim solution = solver.Solve()

        If (solution.Quality = SolverQuality.Optimal) Then
            For Each potentee In potentialList
                Dim decision = model.Decisions.First(Function(it) it.Name = potentee.symbol)
                potentee.weight = decision.ToDouble()
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
