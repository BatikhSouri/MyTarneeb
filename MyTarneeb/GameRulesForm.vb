Public Class GameRulesForm

    Private Sub GameRulesForm_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.Text = LanguageManager.GetStrFromModel("Règles du jeu")
        WebBrowser1.Navigate(My.Application.Info.DirectoryPath & "\Languages\" & LanguageManager.GetStrFromModel("Aide") & ".htm")
    End Sub
End Class