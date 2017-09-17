Public NotInheritable Class AboutBox

    Private Sub AboutBox_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ' Définissez le titre du formulaire.
        Dim ApplicationTitle As String
        If My.Application.Info.Title <> "" Then
            ApplicationTitle = My.Application.Info.Title
        Else
            ApplicationTitle = System.IO.Path.GetFileNameWithoutExtension(My.Application.Info.AssemblyName)
        End If
        Me.Text = LanguageManager.GetStrFromModel("À propos de") & " " & ApplicationTitle
        ' Initialisez tout le texte affiché dans la boîte de dialogue À propos de.
        ' TODO: personnalisez les informations d'assembly de l'application dans le volet "Application" de la 
        '    boîte de dialogue Propriétés du projet (sous le menu "Projet").
        Me.LabelProductName.Text = My.Application.Info.ProductName
        Me.LabelVersion.Text = String.Format("Version {0}", My.Application.Info.Version.ToString)
        Me.LabelCopyright.Text = My.Application.Info.Copyright
        Me.LabelCompanyName.Text = My.Application.Info.CompanyName
        If My.Settings.Language = "French" Or My.Settings.Language = "English" Then
            Me.DevLabel.Text = LanguageManager.GetStrFromModel("Développé par Ahmad Ben MRad")
            Me.SpecialThanksLabel.Text = LanguageManager.GetStrFromModel("Remerciements à") & " :" & vbCrLf & LanguageManager.GetStrFromModel("- Joseph Azouri, pour m'avoir appris les règles du Tarneeb") & vbCrLf & LanguageManager.GetStrFromModel("- Omar Itani, pour m'avoir aidé dans la conception de l'algorithme des joueurs ordinateur") & vbCrLf & LanguageManager.GetStrFromModel("- Nour Nachabe, pour m'avoir aidé dans la conception de l'algorithme des joueurs ordinateur mais aussi pour m'avoir encouragé en me disant : tu vas pas y arriver !") & vbCrLf & LanguageManager.GetStrFromModel("- Jesse Fuchs et Tom Hart pour le design des cartes")
            Me.CardsDesignLinkLabel.Text = LanguageManager.GetStrFromModel("Site web du design des cartes")
        End If
    End Sub

    Private Sub OKButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OKButton.Click
        Me.Close()
    End Sub

    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles CardsDesignLinkLabel.LinkClicked
        Process.Start("http://www.eludication.org/playingcards.html")
    End Sub
End Class
