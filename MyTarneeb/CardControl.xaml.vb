Imports MyTarneeb.Tarneeb
Imports System.Threading

Public Class CardControl

    Private _ShowBack As Boolean = False
    Private _Card As Card
    Private _RotationAngle As Decimal

    Public Event NewCardValue()

    Public Sub SetControl(ByVal Card As Tarneeb.Card, ByVal ShowBack As Boolean)
        _Card = Card
        _ShowBack = ShowBack
        If _ShowBack = True Then
            ShowBackSide()
        Else
            ShowCard()
        End If
    End Sub

    Public Property ShowingBackSide As Boolean
        Set(ByVal value As Boolean)
            _ShowBack = value
            If _ShowBack = True Then
                ShowBackSide()
            Else
                ShowCard()
            End If
        End Set
        Get
            Return _ShowBack
        End Get
    End Property

    Public Property Card As Card
        Set(ByVal value As Card)
            _Card = value
            If _ShowBack = True Then
                ShowBackSide()
            Else
                ShowCard()
            End If
            If Not _Card.CardNumber = 0 Then
                RaiseEvent NewCardValue()
            End If
        End Set
        Get
            Return _Card
        End Get
    End Property

    Private Sub DelShowBackSide()
        Me.Visibility = Windows.Visibility.Visible
        Dim bmps As BitmapSource
        Dim BackImageFileName As String
        If My.Settings.UseCustomBackSide = False Then
            BackImageFileName = My.Settings.BackCardColor & "-" & My.Settings.BackCardDesign & ".png"
            bmps = BitmapFrame.Create(New Uri("pack://application:,,,/MyTarneeb;Component/Images/BackSides/" & BackImageFileName))
        Else
            BackImageFileName = My.Application.Info.DirectoryPath & "\UserBackSides\" & My.Settings.CustomBackSideFileName
            bmps = BitmapFrame.Create(New IO.FileStream(BackImageFileName, IO.FileMode.Open))
        End If
        CardImage.Source = bmps
    End Sub

    Private Sub DelShowCard()
        If Not _Card.CardNumber = 0 Then
            Me.Visibility = Windows.Visibility.Visible
            Dim ImageFileName As String
            If _Card.CardNumber >= 11 Then
                ImageFileName = _Card.CardFamily.ToString.ToLower & "s-" & _Card.CardNumber.ToString.ToLower.Chars(0).ToString & ".png"
            Else
                ImageFileName = _Card.CardFamily.ToString.ToLower & "s-" & Str(_Card.CardNumber) & ".png"
                ImageFileName = ImageFileName.Replace(" ", Nothing)
            End If
            Dim bmps As BitmapSource = BitmapFrame.Create(New Uri("pack://application:,,,/MyTarneeb;Component/Images/" & _Card.CardFamily.ToString() & "/" & ImageFileName))
            CardImage.Source = bmps
        Else
            CardImage.Source = Nothing
        End If
    End Sub

    Public Sub ShowCard()
        SyncLock DelImage
            Me.Dispatcher.Invoke(DelImage)
        End SyncLock
    End Sub

    Public Sub ShowBackSide()
        SyncLock DelBackSide
            Me.Dispatcher.Invoke(DelBackSide)
        End SyncLock
    End Sub

    Public Delegate Sub _DelSub()
    Public Delegate Sub _DelDecimalSub(ByVal Dec As Decimal)
    Public Delegate Sub _DelShortSub(ByVal Sho As Short)
    Private Delegate Sub _DelUseCard(ByVal Wait As Boolean)

    Public DelImage As New _DelSub(AddressOf DelShowCard)
    Public DelBackSide As New _DelSub(AddressOf DelShowBackSide)
    Private DelUseCard As New _DelUseCard(AddressOf _UseCard)
    Public DelSetRotationAngle As New _DelDecimalSub(AddressOf _SetRotationAngle)
    Public DelThrowEffect As New _DelShortSub(AddressOf _ThrowEffect)

    Public Sub UseCard(ByVal Wait As Boolean)
        SyncLock DelUseCard
            Me.Dispatcher.Invoke(DelUseCard, {Wait})
        End SyncLock
    End Sub

    Private Sub _UseCard(ByVal Wait As Boolean)
        Me.ShowingBackSide = False
        If Wait = True Then
            Thread.Sleep(500)
        End If
        Me.Visibility = Windows.Visibility.Collapsed
    End Sub

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

    Public Sub ThrowEffect(ByVal BaseAngle As Short)
        SyncLock DelThrowEffect
            Me.Dispatcher.Invoke(DelThrowEffect, {BaseAngle})
        End SyncLock
    End Sub

    Private Sub _ThrowEffect(ByVal BaseAngle As Short)
        Dim Angle = CalculateTurnCardRotationAngles(BaseAngle)
        Dim RT As RotateTransform = Me.RenderTransform
        RT.Angle = Angle
        Me.RenderTransform = RT
    End Sub

    Private Function CalculateTurnCardRotationAngles(ByVal BaseAngle As Short) As Short
        Dim Angle As Short
        Randomize()
        Dim AbsValueAngleDifference As Byte = Math.Round(Rnd() * 20)
        Randomize()
        Dim Positive As Boolean = Math.Round(Rnd())
        If Positive = True Then
            Angle = BaseAngle + AbsValueAngleDifference
        Else
            Angle = BaseAngle - AbsValueAngleDifference
        End If
        Return Angle
    End Function
End Class