Imports System.IO

Public Class SettingsWindow

    Private CustomBackSideOpenFileDialog As Windows.Forms.OpenFileDialog
    Private BackSidesFileList As FileInfo()

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles Button2.Click
        Me.Close()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles Button1.Click
        'My.Settings.UseCustomBackSide = UserBackSideCheckBox.IsChecked
        'If UserBackSideCheckBox.IsChecked = False Then
        If BlueBackSideSelectionControl1.IsChecked = True Then
            My.Settings.BackCardColor = "blue"
            My.Settings.BackCardDesign = 1
        ElseIf BlueBackSideSelectionControl2.IsChecked = True Then
            My.Settings.BackCardColor = "blue"
            My.Settings.BackCardDesign = 2
        ElseIf BlueBackSideSelectionControl3.IsChecked = True Then
            My.Settings.BackCardColor = "blue"
            My.Settings.BackCardDesign = 3
        ElseIf BlueBackSideSelectionControl4.IsChecked = True Then
            My.Settings.BackCardColor = "blue"
            My.Settings.BackCardDesign = 4
        ElseIf RedBackSideSelectionControl1.IsChecked = True Then
            My.Settings.BackCardColor = "red"
            My.Settings.BackCardDesign = 1
        ElseIf RedBackSideSelectionControl2.IsChecked = True Then
            My.Settings.BackCardColor = "red"
            My.Settings.BackCardDesign = 2
        ElseIf RedBackSideSelectionControl3.IsChecked = True Then
            My.Settings.BackCardColor = "red"
            My.Settings.BackCardDesign = 3
        Else
            My.Settings.BackCardColor = "red"
            My.Settings.BackCardDesign = 4
        End If
        'Else
        'My.Settings.CustomBackSideFileName = FileComboBox.Items(FileComboBox.SelectedIndex).ToString()
        'End If
        My.Settings.Save()
        Me.Close()
    End Sub

    Private Sub Window_Loaded(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles MyBase.Loaded
        RedBackSideSelectionControl1.BackSideColor = BackSideSelectionControl.BackSideColors.Red
        RedBackSideSelectionControl2.BackSideColor = BackSideSelectionControl.BackSideColors.Red
        RedBackSideSelectionControl3.BackSideColor = BackSideSelectionControl.BackSideColors.Red
        RedBackSideSelectionControl4.BackSideColor = BackSideSelectionControl.BackSideColors.Red
        'CustomBackSideOpenFileDialog = New Forms.OpenFileDialog
        'CustomBackSideOpenFileDialog.Filter = "Fichiers BMP (*.bmp)|*.bmp|Fichiers GIF (*.gif)|*.gif|Fichiers JPEG (*.jpg, *.jpeg)|*.jpg|Fichiers PNG (*.png)|*.png|Fichiers TIFF (*.tif, *.tiff)|*.tif"
        'CustomBackSideOpenFileDialog.InitialDirectory = Environment.SpecialFolder.MyPictures
        'CustomBackSideOpenFileDialog.Multiselect = False
        'CustomBackSideOpenFileDialog.RestoreDirectory = True
        'CustomBackSideOpenFileDialog.Title = "Sélectionnez votre motif"
        If My.Settings.BackCardColor = "blue" Then
            If My.Settings.BackCardDesign = 1 Then
                BlueBackSideSelectionControl1.IsChecked = True
            ElseIf My.Settings.BackCardDesign = 2 Then
                BlueBackSideSelectionControl2.IsChecked = True
            ElseIf My.Settings.BackCardDesign = 3 Then
                BlueBackSideSelectionControl3.IsChecked = True
            Else
                BlueBackSideSelectionControl4.IsChecked = True
            End If
        Else
            If My.Settings.BackCardDesign = 1 Then
                RedBackSideSelectionControl1.IsChecked = True
            ElseIf My.Settings.BackCardDesign = 2 Then
                RedBackSideSelectionControl2.IsChecked = True
            ElseIf My.Settings.BackCardDesign = 3 Then
                RedBackSideSelectionControl3.IsChecked = True
            Else
                RedBackSideSelectionControl4.IsChecked = True
            End If
        End If
        Me.Title = LanguageManager.GetStrFromModel("Motif des cartes")
        GroupBox2.Header = LanguageManager.GetStrFromModel("Motif des cartes") & " :"
        Button2.Content = LanguageManager.GetStrFromModel("Annuler")
    End Sub

    Private Sub CheckBox1_Checked(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles UserBackSideCheckBox.Checked
        If UserBackSideCheckBox.IsChecked = True Then
            FileComboBox.IsEnabled = True
            BrowseButton.IsEnabled = True
            ScrollViewer1.IsEnabled = False
            LoadFileList()
        Else
            ScrollViewer1.IsEnabled = True
            FileComboBox.IsEnabled = False
            BrowseButton.IsEnabled = False
        End If
    End Sub

    Private Sub LoadFileList()
        Dim BackSidesDirInfo As New DirectoryInfo(My.Application.Info.DirectoryPath & "\UserBackSides")
        Dim PngBackSidesFileList As FileInfo() = BackSidesDirInfo.GetFiles("*.png")
        Dim BmpBackSidesFileList As FileInfo() = BackSidesDirInfo.GetFiles("*.bmp")
        Dim JgpBackSidesFileList As FileInfo() = BackSidesDirInfo.GetFiles("*.jpg")
        Dim GifBackSidesFileList As FileInfo() = BackSidesDirInfo.GetFiles("*.gif")
        Dim TifBackSidesFileList As FileInfo() = BackSidesDirInfo.GetFiles("*.tif")
        Dim iBackSides As Byte = 0
        For Each fi As FileInfo In PngBackSidesFileList
            BackSidesFileList(iBackSides) = fi
            iBackSides += 1
        Next
        For Each fi As FileInfo In BmpBackSidesFileList
            BackSidesFileList(iBackSides) = fi
            iBackSides += 1
        Next
        For Each fi As FileInfo In JgpBackSidesFileList
            BackSidesFileList(iBackSides) = fi
            iBackSides += 1
        Next
        For Each fi As FileInfo In GifBackSidesFileList
            BackSidesFileList(iBackSides) = fi
            iBackSides += 1
        Next
        For Each fi As FileInfo In TifBackSidesFileList
            BackSidesFileList(iBackSides) = fi
            iBackSides += 1
        Next
        BackSidesFileList = BackSidesFileList.OrderBy(Function(FileInfo) FileInfo.Name).ToArray()
        For Each fi As FileInfo In BackSidesFileList
            FileComboBox.Items.Add(fi.Name)
        Next
    End Sub

    Private Sub BrowseButton_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles BrowseButton.Click
BrowseStart: Dim BackSideConfirmation As Forms.DialogResult = CustomBackSideOpenFileDialog.ShowDialog(Me)
        If BackSideConfirmation = Forms.DialogResult.OK Then
            Dim SelectedFileFI As New FileInfo(CustomBackSideOpenFileDialog.FileName)
            If SelectedFileFI.Extension = "jpg" Or SelectedFileFI.Extension = "bmp" Or SelectedFileFI.Extension = "gif" Or SelectedFileFI.Extension = "png" Or SelectedFileFI.Extension = "tif" Then
                Try
                    IO.File.Copy(CustomBackSideOpenFileDialog.FileName, My.Application.Info.DirectoryPath & "\UserBackSides\" & CustomBackSideOpenFileDialog.SafeFileName)
                    FileComboBox.Items.Add(CustomBackSideOpenFileDialog.SafeFileName)
                    FileComboBox.SelectedIndex = FileComboBox.Items.Count - 1
                Catch ex As IOException
                    Dim UserChoice As MessageBoxResult = MessageBox.Show("Un fichier du même nom existe déjà. Souhaitez-vous le remplacer ?", "Ajout de motif", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No)
                    IO.File.Copy(CustomBackSideOpenFileDialog.FileName, My.Application.Info.DirectoryPath & "\UserBackSides\" & CustomBackSideOpenFileDialog.SafeFileName, True)
                    FileComboBox.Items.Add(CustomBackSideOpenFileDialog.SafeFileName)
                    FileComboBox.SelectedIndex = FileComboBox.Items.Count - 1
                End Try
            Else
                MessageBox.Show("Format de fichier d'image invalide. Veuillez réessayer.", "Sélection de motif", MessageBoxButton.OK, MessageBoxImage.Error)
                GoTo BrowseStart
            End If
        End If
    End Sub

    Private Sub FileComboBox_SelectionChanged(ByVal sender As Object, ByVal e As System.Windows.Controls.SelectionChangedEventArgs) Handles FileComboBox.SelectionChanged
        Dim bmps As BitmapSource = BitmapFrame.Create(New FileStream(My.Application.Info.DirectoryPath.ToString() & "\UserBackSides\" & FileComboBox.Items(FileComboBox.SelectedIndex).ToString(), FileMode.Open))
        BackSidePreviewImage.Source = bmps
    End Sub

    Private Sub BlueBackSideSelectionControl1_IsCheckedChanged() Handles BlueBackSideSelectionControl1.IsCheckedChanged
        If BlueBackSideSelectionControl1.IsChecked = True Then
            BackSidePreviewImage.Source = BlueBackSideSelectionControl1.BackSideImage.Source
        End If
    End Sub

    Private Sub BlueBackSideSelectionControl2_IsCheckedChanged() Handles BlueBackSideSelectionControl2.IsCheckedChanged
        If BlueBackSideSelectionControl2.IsChecked = True Then
            BackSidePreviewImage.Source = BlueBackSideSelectionControl2.BackSideImage.Source
        End If
    End Sub

    Private Sub BlueBackSideSelectionControl3_IsCheckedChanged() Handles BlueBackSideSelectionControl3.IsCheckedChanged
        If BlueBackSideSelectionControl3.IsChecked = True Then
            BackSidePreviewImage.Source = BlueBackSideSelectionControl3.BackSideImage.Source
        End If
    End Sub

    Private Sub BlueBackSideSelectionControl4_IsCheckedChanged() Handles BlueBackSideSelectionControl4.IsCheckedChanged
        If BlueBackSideSelectionControl4.IsChecked = True Then
            BackSidePreviewImage.Source = BlueBackSideSelectionControl4.BackSideImage.Source
        End If
    End Sub

    Private Sub RedBackSideSelectionControl1_IsCheckedChanged() Handles RedBackSideSelectionControl1.IsCheckedChanged
        If RedBackSideSelectionControl1.IsChecked = True Then
            BackSidePreviewImage.Source = RedBackSideSelectionControl1.BackSideImage.Source
        End If
    End Sub

    Private Sub RedBackSideSelectionControl2_IsCheckedChanged() Handles RedBackSideSelectionControl2.IsCheckedChanged
        If RedBackSideSelectionControl2.IsChecked = True Then
            BackSidePreviewImage.Source = RedBackSideSelectionControl2.BackSideImage.Source
        End If
    End Sub

    Private Sub RedBackSideSelectionControl3_IsCheckedChanged() Handles RedBackSideSelectionControl3.IsCheckedChanged
        If RedBackSideSelectionControl3.IsChecked = True Then
            BackSidePreviewImage.Source = RedBackSideSelectionControl3.BackSideImage.Source
        End If
    End Sub

    Private Sub RedBackSideSelectionControl4_IsCheckedChanged() Handles RedBackSideSelectionControl4.IsCheckedChanged
        If RedBackSideSelectionControl4.IsChecked = True Then
            BackSidePreviewImage.Source = RedBackSideSelectionControl4.BackSideImage.Source
        End If
    End Sub
End Class