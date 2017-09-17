Imports System.Windows.Media

Public Class GameProgressBar

    Private _BetValue As Byte = 7
    Private _HumanOwnerOfBet As Boolean = True
    Private _HumanWon As Byte
    Private _IAWon As Byte
    Private _PercentageProgressBarOwner As Decimal
    Private _PercentageProgressBarRival As Decimal
    Private _ProgressBarGradient As LinearGradientBrush
    Private _RedProgressBarGradient As RadialGradientBrush

    Private Sub _SetControl(ByVal HumanOwnerOfBet As Boolean, ByVal BetValue As Byte)
        _HumanOwnerOfBet = HumanOwnerOfBet
        _BetValue = BetValue
        OwnerProgressBar.Maximum = _BetValue
        RivalProgressBar.Maximum = 13 - _BetValue
        PreparePercentage()
        ResizeProgressBars()
        If _HumanOwnerOfBet = True Then
            OwnerProgressBar.Foreground = _ProgressBarGradient
            RivalProgressBar.Foreground = _RedProgressBarGradient
        Else
            OwnerProgressBar.Foreground = _RedProgressBarGradient
            RivalProgressBar.Foreground = _ProgressBarGradient
        End If
    End Sub

    Public Sub SetControl(ByVal HumanOwnerOfBet As Boolean, ByVal BetValue As Byte)
        SyncLock DelSetControl
            Me.Dispatcher.Invoke(DelSetControl, {HumanOwnerOfBet, BetValue})
        End SyncLock
    End Sub

    Public Delegate Sub _DelSetControl(ByVal HumanOwnerOfBet As Boolean, ByVal BetValue As Byte)
    Private Delegate Sub _DelHumanWon(ByVal HumanWon As Byte)
    Private Delegate Sub _DelIAWon(ByVal IAWon As Byte)

    Public DelSetControl As New _DelSetControl(AddressOf _SetControl)
    Private DelHumanWon As New _DelHumanWon(AddressOf SetHumanWon)
    Private DelIAWon As New _DelIAWon(AddressOf SetIAWon)

    Private Sub SetHumanWon(ByVal HumanWon As Byte)
        _HumanWon = HumanWon
        Dim _ToolTip As String
        If _HumanWon <> 0 Then
            _ToolTip = LanguageManager.GetStrFromModel("Tours gagnés :") & " " & _HumanWon
        Else
            _ToolTip = ""
        End If
        If HumanOwnerOfBet = True Then
            OwnerProgressBar.Value = _HumanWon
            OwnerProgressBar.ToolTip = _ToolTip
        Else
            RivalProgressBar.Value = _HumanWon
            RivalProgressBar.ToolTip = _ToolTip
        End If
    End Sub

    Private Sub SetIAWon(ByVal IAWon As Byte)
        _IAWon = IAWon
        Dim _ToolTip As String
        If _IAWon <> 0 Then
            _ToolTip = LanguageManager.GetStrFromModel("Tours perdus :") & " " & _IAWon
        Else
            _ToolTip = ""
        End If
        If HumanOwnerOfBet = True Then
            RivalProgressBar.Value = _IAWon
            RivalProgressBar.ToolTip = _ToolTip
        Else
            OwnerProgressBar.Value = _IAWon
            OwnerProgressBar.ToolTip = _ToolTip
        End If
    End Sub

    Public Property HumanOwnerOfBet As Boolean
        Set(ByVal value As Boolean)
            _HumanOwnerOfBet = value
            _SetControl(_HumanOwnerOfBet, _BetValue)
        End Set
        Get
            Return _HumanOwnerOfBet
        End Get
    End Property

    Public Property BetValue As Byte
        Set(ByVal value As Byte)
            _BetValue = value
            SetControl(_HumanOwnerOfBet, _BetValue)
        End Set
        Get
            Return _BetValue
        End Get
    End Property

    Public Property HumanWon As Byte
        Set(ByVal value As Byte)
            SyncLock DelHumanWon
                Me.Dispatcher.Invoke(DelHumanWon, {value})
            End SyncLock
        End Set
        Get
            Return _HumanWon
        End Get
    End Property

    Public Property IAWon As Byte
        Set(ByVal value As Byte)
            SyncLock DelIAWon
                Me.Dispatcher.Invoke(DelIAWon, {value})
            End SyncLock
        End Set
        Get
            Return _IAWon
        End Get
    End Property

    Private Sub GameProgressBar_Loaded(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles Me.Loaded
        Try
            _ProgressBarGradient = OwnerProgressBar.Foreground
            _RedProgressBarGradient = RivalProgressBar.Foreground
        Catch ex As InvalidCastException

        End Try
        ResizeProgressBars()
    End Sub

    Private Sub GameProgressBar_SizeChanged(ByVal sender As Object, ByVal e As System.Windows.SizeChangedEventArgs) Handles Me.SizeChanged
        PreparePercentage()
        ResizeProgressBars()
    End Sub

    Private Sub ResizeProgressBars()
        OwnerProgressBar.Height = _PercentageProgressBarOwner * Me.Height
        RivalProgressBar.Height = Math.Abs(_PercentageProgressBarRival * Me.Height)
        RivalProgressBar.Width = Me.Width
        Dim TGRivalProgressBar As New TransformGroup
        Dim RTRivalProgressBar As New RotateTransform(180, RivalProgressBar.Width / 2, RivalProgressBar.Height / 2)
        TGRivalProgressBar.Children.Add(RTRivalProgressBar)
        RivalProgressBar.RenderTransform = TGRivalProgressBar
    End Sub

    Private Sub PreparePercentage()
        _PercentageProgressBarOwner = _BetValue / 13
        _PercentageProgressBarRival = (13 - _BetValue) / 13
    End Sub
End Class
