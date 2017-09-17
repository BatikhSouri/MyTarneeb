Public Class BetWindow

    Private _Pass As Boolean = False
    Private _BetDone As Boolean = False
    Private _MinBet As Byte = 7

    Public Sub New()
        InitializeComponent()
        SyncLock SetWinOwner
            Me.Dispatcher.Invoke(SetWinOwner)
        End SyncLock
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        _Pass = True
        '5Me.Close()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If ComboBox1.SelectedIndex <> Nothing Then
            _BetDone = True
            Me.Close()
        Else
            MessageBox.Show("Veuillez choisir votre famille d'atouts.", "Pari de MyTarneeb", MessageBoxButton.OK, MessageBoxImage.Information)
        End If
    End Sub

    Private Sub BetForm_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.Closing
        If _BetDone = False Then
            Dim PassChoice As MessageBoxResult = MessageBox.Show("Voulez-vous vraiment passer ?", "MyTarneeb", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No)
            If PassChoice = MessageBoxResult.No Then
                e.Cancel = True
            Else
                _Pass = True
            End If
        End If
    End Sub

    Public Property Pass As Boolean
        Get
            Return _Pass
        End Get
        Set(ByVal value As Boolean)
            _Pass = value
        End Set
    End Property

    Public Property MinimumUserBet As Byte
        Get
            Return _MinBet
        End Get
        Set(ByVal value As Byte)
            _MinBet = value
            If _MinBet <> 0 Then
                Dim NumListControlCollection As ItemCollection
                For i = _MinBet To 13
                    Dim NumListItem As ComboBoxItem
                    NumListItem.Content = i.ToString()
                    NumListControlCollection.Add(NumListItem)
                Next
                ComboBox1.ItemsSource = NumListControlCollection
            End If
        End Set
    End Property

    Private Delegate Sub _DelSetWinOwner()

    Private SetWinOwner As New _DelSetWinOwner(AddressOf _SetWindowOwner)

    Private Sub _SetWindowOwner()
        'Me.Owner = My.Application.MainWindow
    End Sub
End Class
