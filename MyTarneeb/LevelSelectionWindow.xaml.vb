Public Class LevelSelectionWindow

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles Button1.Click
        Dim ActualLevel As Byte = My.Settings.GameLevel
        Dim SelectedLevel As Byte
        If VeryEasyGameRadioButton.IsChecked = True Then
            SelectedLevel = 0
        ElseIf EasyGameRadioButton.IsChecked = True Then
            SelectedLevel = 1
        ElseIf MediumGameRadioButton.IsChecked = True Then
            SelectedLevel = 2
        ElseIf HardGameRadioButton.IsChecked = True Then
            SelectedLevel = 3
        ElseIf ExtremeGameRadioButton.IsChecked = True Then
            SelectedLevel = 4
        End If
        If SelectedLevel <> ActualLevel Then
            My.Settings.GameLevel = SelectedLevel
            'If MainWindow.GameStarted = True Then
            'MessageBox.Show("Une partie de Tarneeb est en cours. Le changement de niveau prendre effet lors de la prochaine partie", "Sélection du niveau de difficulté", MessageBoxButton.OK, MessageBoxImage.Information)
            'End If
            My.Settings.Save()
        End If
        Me.Close()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles Button2.Click
        Me.Close()
    End Sub

    Private Sub LevelSelectionWindow_Loaded(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles Me.Loaded
        If My.Settings.GameLevel = 0 Then
            VeryEasyGameRadioButton.IsChecked = True
        ElseIf My.Settings.GameLevel = 1 Then
            EasyGameRadioButton.IsChecked = True
        ElseIf My.Settings.GameLevel = 2 Then
            MediumGameRadioButton.IsChecked = True
        ElseIf My.Settings.GameLevel = 3 Then
            HardGameRadioButton.IsChecked = True
        ElseIf My.Settings.GameLevel = 4 Then
            ExtremeGameRadioButton.IsChecked = True
        End If
        VeryEasyGameRadioButton.Content = LanguageManager.GetStrFromModel("Très facile")
        EasyGameRadioButton.Content = LanguageManager.GetStrFromModel("Facile")
        MediumGameRadioButton.Content = LanguageManager.GetStrFromModel("Moyen")
        HardGameRadioButton.Content = LanguageManager.GetStrFromModel("Difficile")
        ExtremeGameRadioButton.Content = LanguageManager.GetStrFromModel("Extrême")
        Me.Title = LanguageManager.GetStrFromModel("Niveau de difficulté")
        GroupBox1.Header = LanguageManager.GetStrFromModel("Niveau de difficulté") & " :"
        Button2.Content = LanguageManager.GetStrFromModel("Annuler")
    End Sub
End Class