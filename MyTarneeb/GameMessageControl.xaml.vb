Imports System.Threading

Public Class GameMessageControl

    Private _Message As String
    Private _MessageMillisecondsTimeout As Integer

    Private _ShowMessageThread As New Thread(AddressOf ShowNHideMessage)

    Private Property MessageShowingTime As Integer
        Set(ByVal value As Integer)
            _MessageMillisecondsTimeout = value
        End Set
        Get
            Return _MessageMillisecondsTimeout
        End Get
    End Property

    Private Property Message As String
        Set(ByVal value As String)
            _Message = value
            GameInfoLabel.Text = value
            '_ShowMessageThread.Start()
            SyncLock DelShowNHideMessage
                Me.Dispatcher.Invoke(DelShowNHideMessage)
            End SyncLock
        End Set
        Get
            Return _Message
        End Get
    End Property

    Public Sub ShowGameMessage(ByVal GameMessage As String, ByVal MessageTimeoutMs As Integer)
        SyncLock DelShowGameMessage
            Me.Dispatcher.Invoke(DelShowGameMessage, {GameMessage, MessageTimeoutMs})
        End SyncLock
    End Sub

    Private Sub _ShowGameMessage(ByVal GameMessage As String, ByVal MessageTimeoutMs As Integer)
        GameInfoLabel.Text = GameMessage
        Me.Visibility = Windows.Visibility.Visible
        Thread.Sleep(MessageTimeoutMs)
        Me.Visibility = Windows.Visibility.Collapsed
    End Sub

    Private Delegate Sub _DelSub()
    Private Delegate Sub _DelShowGameMessage(ByVal GameMessage As String, ByVal MessageTimeoutMs As Integer)

    Private DelShowGameMessage As New _DelShowGameMessage(AddressOf _ShowGameMessage)
    Private DelShowNHideMessage As New _DelSub(AddressOf _ShowNHideMessage)

    Private Sub _ShowNHideMessage()
        'GameInfoLabel.FontSize = Me.Height - 8
        Me.Visibility = Windows.Visibility.Visible
        Thread.Sleep(_MessageMillisecondsTimeout)
        Me.Visibility = Windows.Visibility.Collapsed
    End Sub

    Private Sub ShowNHideMessage()
        SyncLock DelShowNHideMessage
            Me.Dispatcher.Invoke(DelShowNHideMessage)
        End SyncLock
    End Sub
End Class
