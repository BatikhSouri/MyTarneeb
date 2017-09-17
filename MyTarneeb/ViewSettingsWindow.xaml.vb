Public Class ViewSettingsWindow

    Private Sub ViewSettingsWindow_Loaded(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles Me.Loaded
        ByDescendingCardsRadioButton.Content = LanguageManager.GetStrFromModel("Par ordre décroissant")
        ByAscendingCardsRadioButton.Content = LanguageManager.GetStrFromModel("Par ordre croissant")
        ByDescendingFamiliesRadioButton.Content = LanguageManager.GetStrFromModel("Par ordre décroissant")
        ByAscendingFamiliesRadioButton.Content = LanguageManager.GetStrFromModel("Par ordre croissant")
        FamiliesOrderGroupBox.Header = LanguageManager.GetStrFromModel("Classer les familles") & " :"
        CardsOrderGroupBox.Header = LanguageManager.GetStrFromModel("Classer les cartes") & " :"
        Button2.Content = LanguageManager.GetStrFromModel("Annuler")
        Me.Title = LanguageManager.GetStrFromModel("Réglages d'affichage")
        If My.Settings.FamiliesByDescending = True Then
            ByDescendingFamiliesRadioButton.IsChecked = True
        Else
            ByAscendingFamiliesRadioButton.IsChecked = True
        End If
        If My.Settings.CardsByDescending = True Then
            ByDescendingCardsRadioButton.IsChecked = True
        Else
            ByAscendingCardsRadioButton.IsChecked = True
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles Button2.Click
        Me.Close()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles Button1.Click
        My.Settings.FamiliesByDescending = ByDescendingFamiliesRadioButton.IsChecked
        My.Settings.CardsByDescending = ByDescendingCardsRadioButton.IsChecked
        My.Settings.Save()
        Me.Close()
    End Sub
End Class
