Module Controls
    Public Function IsStockInputValid() As Boolean
        'to be complete, a transaction needs qty, symbol/ticker and a type.
        'First, check ticker
        If Globals.Dashboard.TickersCBox.SelectedItem = Nothing Then
            MessageBox.Show("Picking stocks is hard, I know. Do your best, Dave.",
                            "No ticker", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        Else
            CT.symbol = Globals.Dashboard.TickersCBox.SelectedItem
        End If

        'checks qty
        Try
            CT.qty = Integer.Parse(Globals.Dashboard.StockQtyBox.Text)
        Catch
            MessageBox.Show("Quantity, Dave?",
                            "No quantity", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End Try
        If CT.qty = 0 Then
            MessageBox.Show("Trading zero qty, Dave?",
                            "No quantity", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End If
        ' if all checks are passed
        Return True
    End Function

    Public Function IsValid(t As Transaction) As Boolean

        If t.type = "Hold" Then
            MessageBox.Show("Holy BatSmell! I told you to hold that!", "Controls", MessageBoxButtons.OK, MessageBoxIcon.Hand)
            Return False
        End If
        If (currentDate.DayOfWeek = DayOfWeek.Saturday Or currentDate.DayOfWeek = DayOfWeek.Sunday) And
            (t.type = "Buy" Or t.type = "Sell" Or t.type = "SellShort" Or t.type = "CashDiv") Then
            MessageBox.Show("Holy BatSmoke! Weekend. Can't do that.", "Controls", MessageBoxButtons.OK, MessageBoxIcon.Hand)
            Return False
        End If

        If t.qty = 0 Then
            MessageBox.Show("Holy BatSmog! Zero quantity. Not sent.", "Controls", MessageBoxButtons.OK, MessageBoxIcon.Hand)
            Return False
        End If

        If IsInIP(t.symbol) And (t.type = "Buy" Or t.type = "Sell" Or t.type = "SellShort") Then
            MessageBox.Show("Holy BatFog! You cannot trade securities in IP. Not sent.", "Accounting Controls", MessageBoxButtons.OK, MessageBoxIcon.Hand)
            Return False
        End If

        If t.type = "CashDiv" And t.dividend = 0 Then
            MessageBox.Show("Holy BatCloud! No dividend. Not sent.", "Accounting Controls", MessageBoxButtons.OK, MessageBoxIcon.Hand)
            Return False
        End If

        Return True ' if all controls are passed
    End Function
    Public Function IsOptionInputValid() As Boolean
        ' To be complete, a transaction needs qty, symbol/ticker And a type.
        ' first, check symbol
        If Globals.Dashboard.SymbolsCBox.SelectedItem = Nothing Then
            MessageBox.Show("Picking options is hard, I know. Do your best, Dave.",
                            "No symbol", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        Else
            CT.symbol = Globals.Dashboard.SymbolsCBox.SelectedItem
        End If

        'check type
        If CT.type = "" Then
            MessageBox.Show("To buy or not to buy, that is the question.",
                            "No transaction type", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End If

        'check qty
        Try
            CT.qty = Integer.Parse(Globals.Dashboard.OptionQtyBox.Text)
        Catch
            MessageBox.Show("Quantity, Dave?",
                            "No quantity", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End Try
        If CT.qty = 0 Then
            MessageBox.Show("Trading zero qty, Dave?",
                            "No quantity", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End If
        Return True
    End Function

    Public Function IsAPEntryValid(sym As String, unit As String) As Boolean
        If sym = "" Or unit = "" Then
            Return False
        End If
        If Not IsNumeric(unit) Then
            Return False
        End If
        sym = sym.Trim()
        If Double.Parse(unit) = 0 And sym <> "CAccount" Then
            Return False
        End If
        If Not (IsAStock(sym) Or IsAnOption(sym) Or sym = "CAccount") Then
            MessageBox.Show("Holy Batpencil! I am afraid I cannot process " + sym + ", Dave.", "Error",
                            MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End If
        Return True
    End Function
End Module
