Class Application

    ' Les événements de niveau application, par exemple Startup, Exit et DispatcherUnhandledException
    ' peuvent être gérés dans ce fichier.

    Public Const WebSrvRoot As String = "www.mytarneeb.com"

    Private Sub Application_Startup(ByVal sender As Object, ByVal e As System.Windows.StartupEventArgs) Handles Me.Startup
        If Not IO.Directory.Exists(My.Application.Info.DirectoryPath & "\Languages") Then
            Try
                Dim LanguagesWebFolderUrl As String = "http://" & WebSrvRoot & "/downloads/languages/" & My.Application.Info.Version.ToString & "/"
                My.Computer.Network.DownloadFile(LanguagesWebFolderUrl & "FilesList.txt", My.Application.Info.DirectoryPath & "\FilesList.txt")
                Dim FilesListReader As New IO.StreamReader(My.Application.Info.DirectoryPath & "\FilesList.txt")
                IO.Directory.CreateDirectory(My.Application.Info.DirectoryPath & "\Languages")
                Do While FilesListReader.EndOfStream = False
                    Dim FileName As String = FilesListReader.ReadLine()
                    My.Computer.Network.DownloadFile(LanguagesWebFolderUrl & FileName, My.Application.Info.DirectoryPath & "\Languages\" & FileName)
                Loop
                FilesListReader.Close()
                IO.File.Delete(My.Application.Info.DirectoryPath & "\FilesList.txt")
            Catch ex As Exception
                MessageBox.Show("MyTarneeb can't download the language files from the internet. Please retry after having verifed if you are connected to the intrnet and that no firewall or other security system is blocking MyTarneeb.", "First launch", MessageBoxButton.OK, MessageBoxImage.Error)
                Try
                    IO.Directory.Delete(My.Application.Info.DirectoryPath & "\Languages")
                Catch ex2 As Exception

                End Try
                Try
                    IO.File.Delete(My.Application.Info.DirectoryPath & "\FilesList.txt")
                Catch ex2 As Exception

                End Try
                My.Application.Shutdown()
            End Try
        End If
    End Sub
End Class
