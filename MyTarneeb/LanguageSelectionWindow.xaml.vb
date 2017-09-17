Public Class LanguageSelectionWindow

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles Button2.Click
        Me.Close()
    End Sub

    Private Sub LanguageSelectionWindow_Loaded(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles Me.Loaded
        Me.Title = LanguageManager.GetStrFromModel("Langues")
        Button2.Content = LanguageManager.GetStrFromModel("Annuler")
        Dim LanguageList As String() = LanguageManager.GetLanguagesList()
        For Each s As String In LanguageList
            ListBox1.Items.Add(s)
        Next
    End Sub

    Private Sub ListBox1_SelectionChanged(ByVal sender As System.Object, ByVal e As System.Windows.Controls.SelectionChangedEventArgs) Handles ListBox1.SelectionChanged
        Dim SelectedLanguage As String = LanguageManager.GetEnglishLanguageName(ListBox1.Items(ListBox1.SelectedIndex))
        Me.Title = LanguageManager.GetStrFromModel("Langues", SelectedLanguage)
        Button2.Content = LanguageManager.GetStrFromModel("Annuler", SelectedLanguage)
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles Button1.Click
        My.Settings.Language = LanguageManager.GetEnglishLanguageName(ListBox1.Items(ListBox1.SelectedIndex))
        My.Settings.Save()
        Me.Close()
    End Sub
End Class