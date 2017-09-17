'Fichier :      LanguageManager.vb
'Programmeur :  Ahmad Ben MRad
'Version :      1.1.0.0
'Description :  Gestionnaire de langues basé sur une langue modèle ou base. Dans chacun des fichiers de langue, les expressions de même signification se trouvent à la même ligne.

Public Module LanguageManager

    Private LanguageEnglishName As String
    Private Const ReferenceLanguage As String = "French"

    Public Function GetLanguagesList() As String()
        Dim _LanguagesList As String()
        Dim _LanguagesDI As New IO.DirectoryInfo(My.Application.Info.DirectoryPath & "\Languages\")
        Dim _LanguagesFI As IO.FileInfo() = _LanguagesDI.GetFiles("*.lang")
        Dim SR As IO.StreamReader
        Dim iLanguage As Integer = 0
        For Each file As IO.FileInfo In _LanguagesFI
            SR = New IO.StreamReader(file.FullName)
            SR.ReadLine()
            SR.ReadLine()
            ReDim Preserve _LanguagesList(iLanguage)
            _LanguagesList(iLanguage) = SR.ReadLine()
            SR.Close()
            iLanguage += 1
        Next
        Return _LanguagesList
    End Function

    Public Function GetEnglishLanguageName(ByVal LanguageName As String) As String
        Dim _LanguagesDI As New IO.DirectoryInfo(My.Application.Info.DirectoryPath & "\Languages\")
        Dim _LanguagesFI As IO.FileInfo() = _LanguagesDI.GetFiles("*.lang")
        Dim SR As IO.StreamReader
        Dim iLanguage As Integer = 0
        For Each file As IO.FileInfo In _LanguagesFI
            SR = New IO.StreamReader(file.FullName)
            SR.ReadLine()
            SR.ReadLine()
            If SR.ReadLine() = LanguageName Then
                SR.Close()
                Return IO.Path.GetFileNameWithoutExtension(file.Name)
                Exit Function
            Else
                SR.Close()
            End If
        Next
        Return Nothing
    End Function

    Public Function GetStrFromModel(ByVal Sentence As String) As String
        LanguageEnglishName = My.Settings.Language
        If LanguageEnglishName = ReferenceLanguage Then
            Return Sentence
        Else
            Dim Line As Integer = 0
            Dim srm As IO.StreamReader
            Try
                srm = New IO.StreamReader(My.Application.Info.DirectoryPath & "\Languages\" & ReferenceLanguage & ".lang")
                Do Until Sentence = srm.ReadLine()
                    Line = Line + 1
                Loop
            Catch ex As IO.EndOfStreamException
                Throw New InvalidOperationException("Invalid sentence")
                srm.Close()
                srm.Dispose()
                Exit Function
            Catch ex2 As IO.FileNotFoundException
                Throw ex2
                Exit Function
            End Try
            Dim srl As IO.StreamReader
            Dim S As String
            Try
                srl = New IO.StreamReader(My.Application.Info.DirectoryPath & "\Languages\" & LanguageEnglishName & ".lang")
                Dim i As Integer = 1
                For i = 1 To Line
                    srl.ReadLine()
                Next
                S = srl.ReadLine()
            Catch ex As IO.EndOfStreamException
                Throw New InvalidOperationException("Invalid sentence")
                srl.Close()
                srl.Dispose()
                srm.Close()
                srm.Dispose()
                Exit Function
            Catch ex2 As IO.FileNotFoundException
                Throw ex2
                srm.Close()
                srm.Dispose()
                Exit Function
            End Try
            srl.Close()
            srl.Dispose()
            srm.Close()
            srm.Dispose()
            Return S
        End If
    End Function

    Public Function GetStrFromModel(ByVal Sentence As String, ByVal Language As String) As String
        Dim _LanguageEnglishName As String = Language
        If _LanguageEnglishName = ReferenceLanguage Then
            Return Sentence
        Else
            Dim Line As Integer = 0
            Dim srm As IO.StreamReader
            Try
                srm = New IO.StreamReader(My.Application.Info.DirectoryPath & "\Languages\" & ReferenceLanguage & ".lang")
                Do Until Sentence = srm.ReadLine()
                    Line = Line + 1
                Loop
            Catch ex As IO.EndOfStreamException
                Throw New InvalidOperationException("Invalid sentence")
                srm.Close()
                srm.Dispose()
                Exit Function
            Catch ex2 As IO.FileNotFoundException
                Throw ex2
                Exit Function
            End Try
            Dim srl As IO.StreamReader
            Dim S As String
            Try
                srl = New IO.StreamReader(My.Application.Info.DirectoryPath & "\Languages\" & _LanguageEnglishName & ".lang")
                Dim i As Integer = 1
                For i = 1 To Line
                    srl.ReadLine()
                Next
                S = srl.ReadLine()
            Catch ex As IO.EndOfStreamException
                Throw New InvalidOperationException("Invalid sentence")
                srl.Close()
                srl.Dispose()
                srm.Close()
                srm.Dispose()
                Exit Function
            Catch ex2 As IO.FileNotFoundException
                Throw ex2
                srm.Close()
                srm.Dispose()
                Exit Function
            End Try
            srl.Close()
            srl.Dispose()
            srm.Close()
            srm.Dispose()
            Return S
        End If
    End Function

    Private Sub VerifyLanguage()
        LanguageEnglishName = My.Settings.Language
        Dim LInfo As New IO.DirectoryInfo(My.Application.Info.DirectoryPath & "\Languages")
        Dim LanguageDList As IO.FileInfo()
        LanguageDList = LInfo.GetFiles()
        For Each File As IO.FileInfo In LanguageDList
            If File.Name = LanguageEnglishName & ".lang" Then
                Exit Sub
            End If
        Next
        My.Settings.Language = ReferenceLanguage
        LanguageEnglishName = ReferenceLanguage
    End Sub

    Private Function VerifyFileVersion() As Boolean
        Dim VAssembly As String = My.Application.Info.Version.ToString()
        Dim VFile As String
        Dim VR As IO.StreamReader
        Try
            VR = New IO.StreamReader(My.Application.Info.DirectoryPath & "\Languages\" & LanguageEnglishName & ".lang")
            VFile = VR.ReadLine()
            VR.Close()
        Catch ex As IO.FileNotFoundException
            Throw ex
            Exit Function
        End Try
        If VFile = VAssembly Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Sub VerifyFiles()
        VerifyLanguage()
        If VerifyFileVersion() = False Then
            My.Settings.Language = ReferenceLanguage
        End If
    End Sub
End Module