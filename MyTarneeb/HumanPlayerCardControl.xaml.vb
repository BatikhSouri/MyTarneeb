Public Class HumanPlayerCardControl

    Private Shared _HumanTurn As Boolean

    Private _RotationAngle As Decimal

    Public Shared Property HumanTurn As Boolean
        Get
            Return _HumanTurn
        End Get
        Set(ByVal value As Boolean)
            _HumanTurn = value
        End Set
    End Property

    Public Property Card As Tarneeb.Card
        Get
            Return CardControl1.Card
        End Get
        Set(ByVal value As Tarneeb.Card)
            CardControl1.Card = value
            Me.ToolTip = value.CardName
        End Set
    End Property

    Public Sub SetControl(ByVal Card As Tarneeb.Card, ByVal ShowBackSide As Boolean)
        CardControl1.SetControl(Card, ShowBackSide)
        Me.ToolTip = Card.CardName
    End Sub

    Private Sub CardControl1_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Input.MouseButtonEventArgs) Handles CardControl1.MouseDown
        If _HumanTurn = True And MainWindow.GamePaused = False Then
            If My.Settings.TurnFamily = 255 Then
                My.Settings.HumanPlayFamily = CardControl1.Card.CardFamily
                My.Settings.HumanPlayValue = CardControl1.Card.CardNumber
                _HumanTurn = False
            Else
                Dim TurnFamilyList As Tarneeb.Card() = Tarneeb.CommonCardFunctions.GetFamilyList(MainWindow._Players(0).PlayerCards, My.Settings.TurnFamily)
                If IsNothing(TurnFamilyList) Then
                    My.Settings.HumanPlayFamily = CardControl1.Card.CardFamily
                    My.Settings.HumanPlayValue = CardControl1.Card.CardNumber
                    _HumanTurn = False
                Else
                    If CardControl1.Card.CardFamily <> My.Settings.TurnFamily Then
                        'MessageBox.Show("Vous n'avez pas le droit de jouer cette carte.", "MyTarneeb", MessageBoxButton.OK, MessageBoxImage.Information)
                    Else
                        My.Settings.HumanPlayFamily = CardControl1.Card.CardFamily
                        My.Settings.HumanPlayValue = CardControl1.Card.CardNumber
                        _HumanTurn = False
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub HumanPlayerCardControl_MouseEnter(ByVal sender As Object, ByVal e As System.Windows.Input.MouseEventArgs) Handles Me.MouseEnter
        Me.Height *= 1.1
        Me.Width *= 1.1
        'SetRotationAngle(_RotationAngle)
    End Sub

    Private Sub HumanPlayerCardControl_MouseLeave(ByVal sender As Object, ByVal e As System.Windows.Input.MouseEventArgs) Handles Me.MouseLeave
        Me.Height /= 1.1
        Me.Width /= 1.1
        'SetRotationAngle(_RotationAngle)
    End Sub

    Public Delegate Sub _DecimalSub(ByVal Dec As Decimal)

    Public DelSetRotationAngle As New _DecimalSub(AddressOf _SetRotationAngle)

    Public Sub SetRotationAngle(ByVal Angle As Decimal)
        SyncLock DelSetRotationAngle
            Me.Dispatcher.Invoke(DelSetRotationAngle, {Angle})
        End SyncLock
    End Sub

    Private Sub _SetRotationAngle(ByVal Angle As Decimal)
        If Angle >= 0 And Angle < 360 Then
            _RotationAngle = Angle
        ElseIf Angle < 0 Then
            Dim ExitF As Boolean = False
            Do While ExitF = False
                If Angle < 0 Then
                    Angle += 360
                Else
                    _RotationAngle = Angle
                    ExitF = True
                End If
            Loop
        ElseIf Angle >= 360 Then
            Dim ExitF As Boolean = False
            Do While ExitF = False
                If Angle > 360 Then
                    Angle -= 360
                Else
                    _RotationAngle = Angle
                    ExitF = True
                End If
            Loop
        End If
        Dim RT As New RotateTransform(_RotationAngle, 0, Me.Height)
        Me.RenderTransform = RT
    End Sub

    Public Sub ResetRotationCenter()
        Dim RT As RotateTransform = Me.RenderTransform
        RT.CenterY = Me.Height
        Me.RenderTransform = RT
    End Sub
End Class
