Public Class BetForm

    Public Pass As Boolean = False
    Private _BetDone As Boolean = False

    Public Sub New(ByVal MinBet As Byte)
        InitializeComponent()
        If MinBet <> 0 Then
            NumericUpDown1.Value = MinBet
            NumericUpDown1.Minimum = MinBet
        End If
        'Me.ShowDialog(My.Application.MainWindow)
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Pass = True
        Me.Close()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        'If ComboBox1.SelectedItem <> Nothing Then
        _BetDone = True
        Me.Close()
        'Else
        'MessageBox.Show(LanguageManager.GetStrFromModel("Veuillez choisir votre famille d'atouts."), "Pari de MyTarneeb", MessageBoxButton.OK, MessageBoxImage.Information)
        'End If
    End Sub

    Private Sub BetForm_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If _BetDone = False Then
            Dim PassChoice As MessageBoxResult = MessageBox.Show(LanguageManager.GetStrFromModel("Voulez-vous vraiment passer ?"), "MyTarneeb", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No)
            If PassChoice = MessageBoxResult.No Then
                e.Cancel = True
            Else
                Pass = True
                Me.Dispose()
            End If
        Else
            Me.Dispose()
        End If
    End Sub

    Private Sub BetForm_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.Text = LanguageManager.GetStrFromModel("Pariez !")
        Label1.Text = LanguageManager.GetStrFromModel("Combien pariez-vous ?")
        'Label2.Text = LanguageManager.GetStrFromModel("Quelle est la famille de Tarneeb ?")
        Button1.Text = LanguageManager.GetStrFromModel("Je parie !")
        Button2.Text = LanguageManager.GetStrFromModel("Je passe !")
        'ComboBox1.Items.Clear()
        'ComboBox1.Items.Add(LanguageManager.GetStrFromModel("Carreaux"))
        'ComboBox1.Items.Add(LanguageManager.GetStrFromModel("Piques"))
        'ComboBox1.Items.Add(LanguageManager.GetStrFromModel("Coeurs"))
        'ComboBox1.Items.Add(LanguageManager.GetStrFromModel("Trèfles"))
    End Sub
End Class