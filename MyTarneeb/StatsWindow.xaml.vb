Public Class StatsWindow

    Private Sub Window_Loaded(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles MyBase.Loaded
        ResetButton.Content = LanguageManager.GetStrFromModel("Mettre à zéro")
        PeriodTabItem.Header = LanguageManager.GetStrFromModel("Période")
        TotalTabItem.Header = LanguageManager.GetStrFromModel("Total")
        LoadPeriodStats()
        LoadAllTimeStats()
    End Sub

    Private Sub ResetButton_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles ResetButton.Click
        Dim ResetConfirmation As MessageBoxResult = MessageBox.Show(LanguageManager.GetStrFromModel("Voulez-vous vraiment mettre à zéro les statistiques de période ?"), LanguageManager.GetStrFromModel("Mettre à zèro les stastiques"), MessageBoxButton.YesNo, MessageBoxImage.Exclamation, MessageBoxResult.No)
        If ResetConfirmation = MessageBoxResult.Yes Then
            My.Settings.ResetDate = My.Computer.Clock.LocalTime
            My.Settings.PeriodGameTime = New TimeSpan
            My.Settings.PeriodGamesPlayed = 0
            My.Settings.PeriodWonGames = 0
            My.Settings.PeriodLostGames = 0
            My.Settings.Save()
            LoadPeriodStats()
        End If
    End Sub

    Private Sub LoadAllTimeStats()
        Dim TotalPlayedGames As Integer = My.Settings.TotalGamesPlayed, TotalWonGames As Integer = My.Settings.TotalWonGames, TotalLostGames As Integer = My.Settings.TotalLostGames
        Dim PercentageTotalWon As Integer, PercentageTotalLost As Integer
        TotalGameTimeLabel.Content = LanguageManager.GetStrFromModel("Temps de jeu total") & " : " & My.Settings.TotalGameTime.Days & " " & LanguageManager.GetStrFromModel("jours") & ", " & My.Settings.TotalGameTime.Hours & " " & LanguageManager.GetStrFromModel("heures") & ", " & My.Settings.TotalGameTime.Minutes & " " & LanguageManager.GetStrFromModel("minutes et") & " " & My.Settings.TotalGameTime.Seconds & " " & LanguageManager.GetStrFromModel("secondes")
        TotalNumGamesPlayed.Content = LanguageManager.GetStrFromModel("Parties jouées") & " : " & TotalPlayedGames
        TotalNumGamesWon.Content = LanguageManager.GetStrFromModel("Parties gagnées") & " : " & TotalWonGames
        TotalNumGamesLost.Content = LanguageManager.GetStrFromModel("Parties perdues") & " : " & TotalLostGames
        If TotalPlayedGames = 0 Then
            PercentageTotalWonLabel.Content = "0%"
            PercentageTotalLostLabel.Content = "0%"
        Else
            PercentageTotalWon = Math.Round((TotalWonGames / TotalPlayedGames) * 100)
            PercentageTotalLost = Math.Round((TotalLostGames / TotalPlayedGames) * 100)
            PercentageTotalWonLabel.Content = PercentageTotalWon.ToString() & "%"
            PercentageTotalLostLabel.Content = PercentageTotalLost.ToString() & "%"
        End If
    End Sub

    Private Sub LoadPeriodStats()
        Dim PeriodPlayedGames As Integer = My.Settings.PeriodGamesPlayed, PeriodWonGames As Integer = My.Settings.PeriodWonGames, PeriodLostGames As Integer = My.Settings.PeriodLostGames
        Dim PercentagePeriodWon As Integer, PercentagePeriodLost As Integer
        PeriodGameTimeLabel.Content = LanguageManager.GetStrFromModel("Temps de jeu") & " : " & My.Settings.PeriodGameTime.Days & " " & LanguageManager.GetStrFromModel("jours") & ", " & My.Settings.PeriodGameTime.Hours & " " & LanguageManager.GetStrFromModel("heures") & ", " & My.Settings.PeriodGameTime.Minutes & " " & LanguageManager.GetStrFromModel("minutes et") & " " & My.Settings.PeriodGameTime.Seconds & " " & LanguageManager.GetStrFromModel("secondes")
        PeriodPlayedGamesLabel.Content = LanguageManager.GetStrFromModel("Parties jouées") & " : " & PeriodPlayedGames
        PeriodWonGamesLabel.Content = LanguageManager.GetStrFromModel("Parties gagnées") & " : " & PeriodWonGames
        PeriodLostGamesLabel.Content = LanguageManager.GetStrFromModel("Parties perdues") & " : " & PeriodLostGames
        If PeriodPlayedGames = 0 Then
            PercentagePeriodWonGames.Content = "0%"
            PercentagePeriodLostGames.Content = "0%"
        Else
            PercentagePeriodWon = Math.Round((PeriodWonGames / PeriodPlayedGames) * 100)
            PercentagePeriodLost = Math.Round((PeriodLostGames / PeriodPlayedGames) * 100)
            PercentagePeriodWonGames.Content = PercentagePeriodWon.ToString() & "%"
            PercentagePeriodLostGames.Content = PercentagePeriodLost.ToString() & "%"
        End If
        If My.Settings.ResetDate <> Nothing Then
            ResetCounterDateLabel.Content = LanguageManager.GetStrFromModel("Dernière mise à zéro") & " : " & My.Settings.ResetDate.ToString()
        Else
            ResetCounterDateLabel.Content = LanguageManager.GetStrFromModel("Dernière mise à zéro") & " : " & LanguageManager.GetStrFromModel("jamais")
        End If
    End Sub

    Private Sub OKButton_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles OKButton.Click
        Me.Close()
    End Sub
End Class
