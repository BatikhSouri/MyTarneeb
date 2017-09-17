Public Class TarneebSelectionWindow

    Private _DidSelection As Boolean = False
    Public TarneebSelection As Tarneeb.CardFamilies

    Private Sub TarneebSelectionWindow_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles Me.Closing
        If _DidSelection = False Then
            e.Cancel = True
        End If
    End Sub

    Private Sub TarneebSelectionWindow_Loaded(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles Me.Loaded
        Me.Title = LanguageManager.GetStrFromModel("Sélectionnez le Tarneeb")
        Label1.Content = LanguageManager.GetStrFromModel("Vous avez gagné le pari; veuillez sélectionner le Tarneeb (atout)") & " :"
        DiamondsSButton.ToolTip = LanguageManager.GetStrFromModel(DiamondsSButton.ToolTip)
        HeartsSButton.ToolTip = LanguageManager.GetStrFromModel(HeartsSButton.ToolTip)
        ClubsSButton.ToolTip = LanguageManager.GetStrFromModel(ClubsSButton.ToolTip)
        SpadesSButton.ToolTip = LanguageManager.GetStrFromModel(SpadesSButton.ToolTip)
    End Sub

    Private Sub DiamondsSButton_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles DiamondsSButton.Click
        TarneebSelection = Tarneeb.CardFamilies.Diamond
        _DidSelection = True
        Me.Close()
    End Sub

    Private Sub SpadesSButton_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles SpadesSButton.Click
        TarneebSelection = Tarneeb.CardFamilies.Spade
        _DidSelection = True
        Me.Close()
    End Sub

    Private Sub HeartsSButton_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles HeartsSButton.Click
        TarneebSelection = Tarneeb.CardFamilies.Heart
        _DidSelection = True
        Me.Close()
    End Sub

    Private Sub ClubsSButton_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles ClubsSButton.Click
        TarneebSelection = Tarneeb.CardFamilies.Club
        _DidSelection = True
        Me.Close()
    End Sub
End Class
