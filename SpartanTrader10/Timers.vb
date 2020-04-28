Module Timers
    Public WithEvents SpyTimer As Timer
    Public WithEvents SecondsTimer As Timer

    Public Sub StartTimers()
        If IsNothing(SpyTimer) Then
            SpyTimer = New Timer
            SpyTimer.Interval = 2000
            SecondsTimer = New Timer
            SecondsTimer.Interval = 1000
            secondsLeft = 59
            currentDate = DownloadCurrentDate()
            Globals.Dashboard.DateLine.Value = "Waiting for online data from Athens...."
            SecondsTimer.Start()
            SpyTimer.Start()
        End If
    End Sub

    Private Sub SecondsTimer_Tick() Handles SecondsTimer.Tick
        Try
            secondsLeft = secondsLeft - 1
            If secondsLeft < 0 Then
                secondsLeft = 59
            End If
            If Globals.ThisWorkbook.ActiveSheet.Name = "Dashboard" Then
                Globals.Dashboard.SecondsCell.Value = secondsLeft
                Select Case secondsLeft
                    Case Is <= 5
                        Globals.Dashboard.SecondsCell.Font.Color = System.Drawing.Color.Red
                    Case Else
                        Globals.Dashboard.SecondsCell.Font.Color = System.Drawing.Color.DarkOrange
                End Select
            End If
        Catch ex As Exception
        End Try

    End Sub

    Public Sub StopTimers()
        If IsNothing(SpyTimer) Then
            'skip
        Else
            SpyTimer.Stop()
            SecondsTimer.Stop()
            Globals.Dashboard.SecondsCell.Value = "--"
            Globals.Dashboard.SecondsCell.Font.Color = System.Drawing.Color.DarkOrange
            SpyTimer = Nothing
            SecondsTimer = Nothing
        End If
    End Sub

    Private Sub SpyTimer_Tick() Handles SpyTimer.Tick
        Try
            tempNewDate = DownloadCurrentDate()

            If tempNewDate.Date <> currentDate.Date Then
                currentDate = tempNewDate
                secondsLeft = 59
                RunDailyRoutine(currentDate)

            Else
                If waitingForData Then
                    If IsThereData() = True Then
                        waitingForData = False
                        secondsLeft = 59
                        FirstDayStart()
                    End If
                End If
            End If
        Catch ex As Exception

        End Try
    End Sub

    Public Sub FirstDayStart()
        currentDate = DownloadCurrentDate()
        DownloadStaticData()
        DownloadTeamData(currentDate)
        SetFinancialConstants()
        ResetAllRecommendations()
        RunDailyRoutine(currentDate)
    End Sub
End Module
