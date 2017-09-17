Public Class BackSideSelectionControl

    Private _BackSideColor As BackSideColors = BackSideColors.Blue
    Private _BackSideID As Byte = 1

    Public Event IsCheckedChanged()

    Public Enum BackSideColors
        Blue = 0
        Red = 1
    End Enum

    Public Sub SetControl(ByVal BackSideID As Byte, ByVal BackSideColor As BackSideColors)
        If BackSideID > 0 And BackSideID < 5 Then
            _BackSideColor = BackSideColor
            _BackSideID = BackSideID
            ShowBackSide()
        End If
    End Sub

    Public Property BackSideID As Byte
        Set(ByVal value As Byte)
            If value > 0 And value < 5 Then
                _BackSideID = value
                ShowBackSide()
            End If
        End Set
        Get
            Return _BackSideID
        End Get
    End Property

    Public Property BackSideColor As BackSideColors
        Set(ByVal value As BackSideColors)
            _BackSideColor = value
            ShowBackSide()
        End Set
        Get
            Return _BackSideColor
        End Get
    End Property

    Public Property IsChecked As Boolean
        Set(ByVal value As Boolean)
            SelectionRadioButton.IsChecked = value
            RaiseEvent IsCheckedChanged()
        End Set
        Get
            Return SelectionRadioButton.IsChecked
        End Get
    End Property

    Private Sub ShowBackSide()
        If Not IsNothing(_BackSideID) And Not IsNothing(_BackSideColor) Then
            Dim BackImageFileName As String
            If _BackSideColor = BackSideColors.Red Then
                BackImageFileName = "red-" & _BackSideID & ".png"
            Else
                BackImageFileName = "blue-" & _BackSideID & ".png"
            End If
            Dim bmps As BitmapSource = BitmapFrame.Create(New Uri("pack://application:,,,/MyTarneeb;Component/Images/BackSides/" & BackImageFileName))
            BackSideImage.Source = bmps
        End If
    End Sub

    Private Sub BackSideSelectionControl_Loaded(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles Me.Loaded
        ShowBackSide()
    End Sub
End Class
