'Fichier :      MainWindow.xaml.vb
'Programmeur :  Ahmad Ben MRad
'Version :      1.0.1.0
'Description :  Ce fichier est le code de la fenêtre principale. Il contient également le code de gestion de partie.

Imports MyTarneeb.Tarneeb
Imports System.Threading
Imports System.Collections.ObjectModel

Class MainWindow

    Private Shared _GameStarted As Boolean = False, _GamePause As Boolean = False
    Private _PreparationEnded As Boolean = True
    Public BettingPhase As Boolean = False, SeekingTarneebPhase As Boolean = False
    Private HumanWon As Byte = 0, IAWon As Byte = 0, TurnNumber As Byte = 0

    Private _GameTime As Stopwatch

    Private _GameThread As Thread
    Private _GamePreparationThread As Thread
    Private _TimerThread As Thread
    Private _LoadThread As Thread
    Private _ShowGameCommentThread As Thread

    Private Event GameStarting()
    Private Event GameFinished()

    Private _GameSet(51) As Card

    'Public Const WebSrvRoot As String = "192.168.1.53"

    Public Shared _Players(3) As Tarneeb.IPlayer

    Private Sub NewGame()
        'ShowGameMessage("Test", TimeSpan.FromSeconds(2))
        RaiseEvent GameStarting()
        _PreparationEnded = False
        'préparation du tirage au sort des cartes des joueurs
        'CleanGameTable()
        TurnNumber = 0
        _Players(0) = New Tarneeb.HumanPlayer
        _Players(1) = New Tarneeb.AIPlayer(My.Settings.GameLevel) 'le niveau défini pour les joueurs ordi
        _Players(2) = New Tarneeb.AIPlayer(2) 'le niveau moyen pour aider le joueur humain et conpenser son niveau
        _Players(3) = New Tarneeb.AIPlayer(My.Settings.GameLevel) 'le niveau défini pour les joueurs ordi
        DefinePlayersBetOrder()
        'Dim NumMixes As UShort = 1
CardMixing: Dim GameSet(51) As Tarneeb.Card
        GameSet = MixCardSet(Card.GameSet)
        Dim FuturePlayerCards(12) As Tarneeb.Card
        Dim ActualCard As Tarneeb.Card
        Dim BetCollection() As Bet, iBet As Byte = 0
        'Dim TirageText As String = ""
        Dim CardID As Byte
        Dim PlayersOrder As String = My.Settings.PreparationPlayersOrder
        Dim l As Byte = 0
        'tirage au sort des cartes des 4 joueurs
        For j = 0 To 3
            'TirageText &= vbCr & "Player " & j.ToString() & " : "
            ReDim FuturePlayerCards(12)
            For i = 1 To 13
                'Tirage au sort de la carte
                Randomize()
                CardID = Math.Ceiling(Rnd() * UBound(GameSet))
                FuturePlayerCards(i - 1) = GameSet(CardID)
                'TirageText &= GameSet(CardID).CardFamily.ToString & GameSet(CardID).CardNumber & " "
                ActualCard = GameSet(CardID)
                'Retrait de la carte du tas
                For k = 0 To UBound(GameSet)
                    If GameSet(k).CardFamily = ActualCard.CardFamily And GameSet(k).CardNumber = ActualCard.CardNumber Then
                        If UBound(GameSet) > 1 And k < UBound(GameSet) Then
                            For l = k To UBound(GameSet) - 1
                                GameSet(l) = GameSet(l + 1)
                            Next
                        End If
                        ReDim Preserve GameSet(UBound(GameSet) - 1)
                        Exit For
                    End If
                Next
            Next
            If CountHighCards(FuturePlayerCards) < 2 Then
                'NumMixes += 1
                GoTo CardMixing
            Else
                Dim MaxFam As CardFamilies = CommonCardFunctions.GetTheFamilyWhichHasTheMost(FuturePlayerCards)
                Dim MaxFamList As Card() = GetFamilyList(FuturePlayerCards, MaxFam)
                Dim NumMaxFam As Byte = UBound(MaxFamList) + 1
                Dim NumHighCardsMaxFam As Byte = CountHighCards(MaxFamList)
                'ajout dans le jeu du joueur
                If NumMaxFam >= 6 And NumHighCardsMaxFam >= 2 Then
                    _Players(Val(PlayersOrder(j))).PlayerCards = FuturePlayerCards
                Else
                    'NumMixes += 1
                    GoTo CardMixing
                End If
            End If
        Next
        'MessageBox.Show(TirageText)
        'Pari des joueurs
        'MessageBox.Show("Number of card mixing tries : " & NumMixes.ToString())
        PrepareCardControls()
        If Val(PlayersOrder(0)) = 0 Then
            MessageBox.Show(LanguageManager.GetStrFromModel("Vous pariez en premier"), "MyTarneeb", MessageBoxButton.OK, MessageBoxImage.Information)
        ElseIf Val(PlayersOrder(0)) = 1 Then
            MessageBox.Show(LanguageManager.GetStrFromModel("Le joueur à votre droite parie en premier"), "MyTarneeb", MessageBoxButton.OK, MessageBoxImage.Information)
        ElseIf Val(PlayersOrder(0)) = 2 Then
            MessageBox.Show(LanguageManager.GetStrFromModel("Votre coéquipier parie en premier"), "MyTarneeb", MessageBoxButton.OK, MessageBoxImage.Information)
        Else
            MessageBox.Show(LanguageManager.GetStrFromModel("Le joueur à votre gauche parie en premier"), "MyTarneeb", MessageBoxButton.OK, MessageBoxImage.Information)
        End If
        BettingPhase = True
        For i = 0 To 3
            ReDim Preserve BetCollection(iBet)
            BetCollection(i) = _Players(Val(PlayersOrder(i))).Bet(BetCollection)
            iBet += 1
            Dim PlayerID As Byte = Val(PlayersOrder(i))
            Dim PlayerDescription As String
            If PlayerID <> 0 Then
                If PlayerID = 1 Then
                    PlayerDescription = LanguageManager.GetStrFromModel("Le joueur à votre droite") & " "
                ElseIf PlayerID = 2 Then
                    PlayerDescription = LanguageManager.GetStrFromModel("Votre coéquipier") & " "
                Else
                    PlayerDescription = LanguageManager.GetStrFromModel("Le joueur à votre gauche") & " "
                End If
                If BetCollection(i).Bet = 0 Then
                    MessageBox.Show(PlayerDescription & LanguageManager.GetStrFromModel("passe son tour."), "MyTarneeb", MessageBoxButton.OK, MessageBoxImage.Information)
                Else
                    MessageBox.Show(PlayerDescription & LanguageManager.GetStrFromModel("parie") & " " & BetCollection(i).Bet & ".", "MyTarneeb", MessageBoxButton.OK, MessageBoxImage.Information)
                End If
            End If
        Next
        BettingPhase = False
        Dim HighestBet As Bet = GetHighestBet(BetCollection)
        If HighestBet.Bet < 7 Then
            MessageBox.Show(LanguageManager.GetStrFromModel("Aucun joueur n'a parié. Les cartes vont être redistribuées."), "MyTarneeb", MessageBoxButton.OK, MessageBoxImage.Information)
            GoTo CardMixing
        End If
        My.Settings.BetOwner = Val(My.Settings.PreparationPlayersOrder(GetBetID(BetCollection, HighestBet)))
        SeekingTarneebPhase = True
        My.Settings.Tarneeb = _Players(My.Settings.BetOwner).SeekForTarneeb
        SeekingTarneebPhase = False
        HighestBet = _Players(My.Settings.BetOwner).PlayerBet 'rechargement du pari après SeekForTarneeb (modif possibles ac HumanPlayer as bet owner)
        My.Settings.CurrentBet = HighestBet.Bet
        Dim MessageText As String
        If My.Settings.BetOwner = 0 Or My.Settings.BetOwner = 2 Then
            GameProgressBar1.SetControl(True, HighestBet.Bet)
            MessageText = LanguageManager.GetStrFromModel("Votre équipe est le propriétaire du pari.") & " "
        Else
            GameProgressBar1.SetControl(False, HighestBet.Bet)
            MessageText = LanguageManager.GetStrFromModel("L'équipe adverse est le propriétaire du pari.") & " "
        End If
        SyncLock _DelSetBetOwnerDisplay
            Me.Dispatcher.Invoke(_DelSetBetOwnerDisplay, {My.Settings.BetOwner})
        End SyncLock
        MessageText &= LanguageManager.GetStrFromModel("Le pari est de") & " " & HighestBet.Bet & ". "
        For i = 0 To 3
            _Players(i).SetTarneebFamily()
        Next
        If HighestBet.CardFamily = CardFamilies.Diamond Then
            MessageText &= LanguageManager.GetStrFromModel("Le Tarneeb est") & " " & LanguageManager.GetStrFromModel("carreau") & ". "
        ElseIf HighestBet.CardFamily = CardFamilies.Club Then
            MessageText &= LanguageManager.GetStrFromModel("Le Tarneeb est") & " " & LanguageManager.GetStrFromModel("trèfle") & ". "
        ElseIf HighestBet.CardFamily = CardFamilies.Heart Then
            MessageText &= LanguageManager.GetStrFromModel("Le Tarneeb est") & " " & LanguageManager.GetStrFromModel("coeur") & ". "
        ElseIf HighestBet.CardFamily = CardFamilies.Spade Then
            MessageText &= LanguageManager.GetStrFromModel("Le Tarneeb est") & " " & LanguageManager.GetStrFromModel("pique") & ". "
        End If
        If My.Settings.BetOwner = 0 Then
            MessageText &= LanguageManager.GetStrFromModel("Vous commencez la partie.")
        ElseIf My.Settings.BetOwner = 1 Then
            MessageText &= LanguageManager.GetStrFromModel("Le joueur à votre droite commence la partie.")
        ElseIf My.Settings.BetOwner = 2 Then
            MessageText &= LanguageManager.GetStrFromModel("Votre coéquipier commence la partie.")
        ElseIf My.Settings.BetOwner = 3 Then
            MessageText &= LanguageManager.GetStrFromModel("Le joueur à votre gauche commence la partie.")
        End If
        MessageBox.Show(MessageText, "MyTarneeb", MessageBoxButton.OK, MessageBoxImage.Information)
        SyncLock _DelSetBetDisplay
            Me.Dispatcher.Invoke(_DelSetBetDisplay, {HighestBet})
        End SyncLock
        SyncLock _DelSetGameLevelLabel
            Me.Dispatcher.Invoke(_DelSetGameLevelLabel)
        End SyncLock
        DefinePlayersOrder(My.Settings.BetOwner)

        _GameThread = New Thread(AddressOf PlayGame)
        _GameThread.IsBackground = True
        _GameThread.Start()
        _TimerThread = New Thread(AddressOf ShowTime)
        _TimerThread.IsBackground = True

        _PreparationEnded = True
    End Sub

    Private Sub PlayGame()
        _GameStarted = True
        SyncLock _DelSetLabelsVisibilty
            Me.Dispatcher.Invoke(_DelSetLabelsVisibilty, {Windows.Visibility.Visible})
        End SyncLock
        _GameTime = New Stopwatch
        _GameTime.Reset()
        _GameTime.Start()
        _TimerThread.Start()
        Dim TurnCards(0) As Card
        Dim PlayerChoice As Card
        Dim iCardGameSet As Byte = 0
        HumanWon = 0
        IAWon = 0
        Dim ActualPlayerID As Byte, ActualCardID As Byte
        Dim PlayersTurn, LastTurnPlayersOrder As String
        Dim WinningPlayerID, WinningPlayerInd As Byte
        SyncLock _DelSetTurnCardControl1Visibility
            Me.Dispatcher.Invoke(_DelSetTurnCardControl1Visibility, {Windows.Visibility.Visible})
        End SyncLock
        SyncLock _DelSetTurnCardControl2Visibility
            Me.Dispatcher.Invoke(_DelSetTurnCardControl2Visibility, {Windows.Visibility.Visible})
        End SyncLock
        SyncLock _DelSetTurnCardControl3Visibility
            Me.Dispatcher.Invoke(_DelSetTurnCardControl3Visibility, {Windows.Visibility.Visible})
        End SyncLock
        SyncLock _DelSetTurnCardControl4Visibility
            Me.Dispatcher.Invoke(_DelSetTurnCardControl4Visibility, {Windows.Visibility.Visible})
        End SyncLock
        For Me.TurnNumber = 1 To 13
            TurnCards = Nothing
            ReDim TurnCards(0)
            PlayersTurn = My.Settings.PlayersTurn
            SyncLock DelSetTurnNumDisplay
                Me.Dispatcher.Invoke(DelSetTurnNumDisplay, {TurnNumber})
            End SyncLock
            For j = 0 To 3
                If j = 0 Then
                    My.Settings.TurnFamily = 255
                Else
                    My.Settings.TurnFamily = TurnCards(0).CardFamily
                End If
                ActualPlayerID = PlayersTurn(j).ToString()
                ReDim Preserve TurnCards(j)
                TurnCards(j) = _Players(ActualPlayerID).Play(TurnCards)
                PlayerChoice = TurnCards(UBound(TurnCards))
                ActualCardID = _Players(ActualPlayerID).TurnCardID
                If ActualCardID <> 255 Then
                    ShowCard(ActualPlayerID)
                Else
                    Throw New IndexOutOfRangeException
                End If
                Thread.Sleep(250)
            Next
            Thread.Sleep(1250)
            LastTurnCard1.Card = Nothing
            LastTurnCard2.Card = Nothing
            LastTurnCard3.Card = Nothing
            LastTurnCard4.Card = Nothing
            PrepareCardControls()
            LastTurnCard1.Card = TurnCards(0)
            Thread.Sleep(75)
            LastTurnCard2.Card = TurnCards(1)
            Thread.Sleep(75)
            LastTurnCard3.Card = TurnCards(2)
            Thread.Sleep(75)
            LastTurnCard4.Card = TurnCards(3)
            Thread.Sleep(75)
            TurnCardControl1.Card = Nothing
            TurnCardControl2.Card = Nothing
            TurnCardControl3.Card = Nothing
            TurnCardControl4.Card = Nothing
            For Each c As Card In TurnCards
                _GameSet(iCardGameSet) = c
                iCardGameSet += 1
            Next
            SetLastTurnCardToolTip(TurnCards, My.Settings.PlayersTurn)
            LastTurnPlayersOrder = PlayersTurn
            WinningPlayerID = DefinePlayersOrder(TurnCards)
            WinningPlayerInd = GetWinningPlayerInd(LastTurnPlayersOrder, WinningPlayerID)
            If WinningPlayerID = 0 Or WinningPlayerID = 2 Then
                GameProgressBar1.HumanWon += 1
                HumanWon += 1
            Else
                GameProgressBar1.IAWon += 1
                IAWon += 1
            End If
            SyncLock DelSetGameStateDisplay
                Me.Dispatcher.Invoke(DelSetGameStateDisplay, {HumanWon, IAWon})
            End SyncLock
            For k = 1 To 3
                _Players(k).EndOfRound(TurnCards, WinningPlayerInd)
            Next
            'Dim iPlayer As Byte = 0, TarneebList As Card()
            'For Each p As IPlayer In _Players
            'TarneebList = GetFamilyList(p.PlayerCards, My.Settings.Tarneeb)
            'If Not IsNothing(TarneebList) Then
            'If UBound(TarneebList) = UBound(p.PlayerCards) Then
            'FinishGame(IPlayer)
            'GoTo EndGame
            'End If
            'End If
            'IPlayer += 1
            'Next
        Next
EndGame: _GameTime.Stop()
        CleanGameTable()
        My.Settings.TotalGameTime = My.Settings.TotalGameTime.Add(_GameTime.Elapsed)
        My.Settings.PeriodGameTime = My.Settings.PeriodGameTime.Add(_GameTime.Elapsed)
        My.Settings.TotalGamesPlayed += 1
        My.Settings.PeriodGamesPlayed += 1
        If My.Settings.BetOwner = 0 Or My.Settings.BetOwner = 2 Then
            If HumanWon >= My.Settings.CurrentBet Then
                My.Settings.TotalWonGames += 1
                My.Settings.PeriodWonGames += 1
                MessageBox.Show(LanguageManager.GetStrFromModel("Vous avez gagné !"), "MyTarneeb", MessageBoxButton.OK, MessageBoxImage.Information)
                'GameMessageControl1.ShowGameMessage("Vous avez gagné !", 5000)
                'Thread.Sleep(5000)
            Else
                My.Settings.TotalLostGames += 1
                My.Settings.PeriodLostGames += 1
                MessageBox.Show(LanguageManager.GetStrFromModel("Vous avez perdu !"), "MyTarneeb", MessageBoxButton.OK, MessageBoxImage.Information)
                'GameMessageControl1.ShowGameMessage("Vous avez perdu !", 5000)
                'Thread.Sleep(5000)
            End If
        Else
            If IAWon >= My.Settings.CurrentBet Then
                My.Settings.TotalLostGames += 1
                My.Settings.PeriodLostGames += 1
                MessageBox.Show(LanguageManager.GetStrFromModel("Vous avez perdu !"), "MyTarneeb", MessageBoxButton.OK, MessageBoxImage.Information)
                'GameMessageControl1.ShowGameMessage("Vous avez perdu !", 5000)
                'Thread.Sleep(5000)
            Else
                My.Settings.TotalWonGames += 1
                My.Settings.PeriodWonGames += 1
                MessageBox.Show(LanguageManager.GetStrFromModel("Vous avez gagné !"), "MyTarneeb", MessageBoxButton.OK, MessageBoxImage.Information)
                'GameMessageControl1.ShowGameMessage("Vous avez gagné !", 5000)
                'Thread.Sleep(5000)
            End If
        End If
        My.Settings.Save()
        _GameStarted = False
        RaiseEvent GameFinished()
        AskForNewGame()
    End Sub

    Private Function GetWinningPlayerInd(ByVal PlayersOrder As String, ByVal WinningPlayerID As Byte)
        For i = 0 To 3
            If Val(PlayersOrder(i)) = WinningPlayerID Then
                Return i
                Exit For
            End If
        Next
    End Function

    Private Delegate Sub _SetControlVisibility(ByVal ControlVisibility As Visibility)
    Private Delegate Sub _SetControlContent(ByVal Text As String)
    Private Delegate Sub _ByteSub(ByVal byt As Byte)
    Private Delegate Sub _SetBetDisplay(ByVal GameBet As Bet)
    Private Delegate Sub _ShowGameInfo(ByVal Text As String, ByVal Time As TimeSpan)

    Private _DelSetLabelsVisibilty As New _SetControlVisibility(AddressOf DelSetLabelsVisibility)
    Private _DelSetTurnCardControl1Visibility As New _SetControlVisibility(AddressOf DelSetTurnCardControl1Visibility)
    Private _DelSetTurnCardControl2Visibility As New _SetControlVisibility(AddressOf DelSetTurnCardControl2Visibility)
    Private _DelSetTurnCardControl3Visibility As New _SetControlVisibility(AddressOf DelSetTurnCardControl3Visibility)
    Private _DelSetTurnCardControl4Visibility As New _SetControlVisibility(AddressOf DelSetTurnCardControl4Visibility)
    Private _DelSetBetDisplay As New _SetBetDisplay(AddressOf DelSetBetDisplay)
    Private _DelSetBetOwnerDisplay As New _ByteSub(AddressOf DelSetBetOwnerDisplay)
    Private _DelSetGameLevelLabel As New _DelSub(AddressOf DelSetGameLevelLabel)
    Private _DelShowGameMessage As New _ShowGameInfo(AddressOf _ShowGameMessage)

    Private Sub ShowGameMessage(ByVal Message As String, ByVal Time As TimeSpan)
        SyncLock _DelShowGameMessage
            Me.Dispatcher.Invoke(_DelShowGameMessage, {Message, Time})
        End SyncLock
    End Sub

    Private Sub _ShowGameMessage(ByVal Message As String, ByVal Time As TimeSpan)
        GameMessageTextBlock.Text = Message
        Thread.Sleep(Time)
        GameMessageTextBlock.Text = ""
    End Sub

    Private Sub DelSetLabelsVisibility(ByVal LabelVisibility As Visibility)
        LastTurnCardsLabel.Visibility = LabelVisibility
        GraphicalTarneebLabel.Visibility = LabelVisibility
        TarneebImage.Visibility = LabelVisibility
    End Sub

    Private Sub DelSetTurnCardControl1Visibility(ByVal ControlVisibility As Visibility)
        TurnCardControl1.Visibility = ControlVisibility
    End Sub

    Private Sub DelSetTurnCardControl2Visibility(ByVal ControlVisibility As Visibility)
        TurnCardControl2.Visibility = ControlVisibility
    End Sub

    Private Sub DelSetTurnCardControl3Visibility(ByVal ControlVisibility As Visibility)
        TurnCardControl3.Visibility = ControlVisibility
    End Sub

    Private Sub DelSetTurnCardControl4Visibility(ByVal ControlVisibility As Visibility)
        TurnCardControl4.Visibility = ControlVisibility
    End Sub

    Private Sub DelSetBetDisplay(ByVal GameBet As Bet)
        TarneebLabel.Content = "Tarneeb : " & GameBet.FamilyName
        BetLabel.Content = LanguageManager.GetStrFromModel("Pari :") & " " & GameBet.Bet
        Dim TarneebImageUriStr As String = "pack://application:,,,/MyTarneeb;component/Images/Tarneebs/"
        If GameBet.CardFamily = CardFamilies.Diamond Then
            TarneebImageUriStr &= "diamonds"
        ElseIf GameBet.CardFamily = CardFamilies.Heart Then
            TarneebImageUriStr &= "hearts"
        ElseIf GameBet.CardFamily = CardFamilies.Club Then
            TarneebImageUriStr &= "clubs"
        ElseIf GameBet.CardFamily = CardFamilies.Spade Then
            TarneebImageUriStr &= "spades"
        End If
        TarneebImageUriStr &= ".png"
        TarneebImage.Source = BitmapFrame.Create(New Uri(TarneebImageUriStr))
        TarneebImage.ToolTip = GameBet.FamilyName
    End Sub

    Private Sub DelSetBetOwnerDisplay(ByVal BetOwnerID As Byte)
        If BetOwnerID = 0 Or BetOwnerID = 2 Then
            BetOwnerLabel.Content = LanguageManager.GetStrFromModel("Propriétaires du pari :") & " " & LanguageManager.GetStrFromModel("votre équipe").ToLower()
        Else
            BetOwnerLabel.Content = LanguageManager.GetStrFromModel("Propriétaires du pari :") & " " & LanguageManager.GetStrFromModel("vos adversaires").ToLower()
        End If
    End Sub

    Private Sub DelSetGameLevelLabel()
        GameLevelLabel.Content = LanguageManager.GetStrFromModel("Niveau :") & " "
        If My.Settings.GameLevel = 0 Then
            GameLevelLabel.Content &= LanguageManager.GetStrFromModel("Très facile")
        ElseIf My.Settings.GameLevel = 1 Then
            GameLevelLabel.Content &= LanguageManager.GetStrFromModel("Facile")
        ElseIf My.Settings.GameLevel = 2 Then
            GameLevelLabel.Content &= LanguageManager.GetStrFromModel("Moyen")
        ElseIf My.Settings.GameLevel = 3 Then
            GameLevelLabel.Content &= LanguageManager.GetStrFromModel("Difficile")
        Else
            GameLevelLabel.Content &= LanguageManager.GetStrFromModel("Extrême")
        End If
    End Sub

    Public Shared ReadOnly Property GameStarted As Boolean
        Get
            Return _GameStarted
        End Get
    End Property

    Public Shared ReadOnly Property GamePaused As Boolean
        Get
            Return _GamePause
        End Get
    End Property

    Private Function MixCardSet(ByVal CardSet As Card()) As Card()
        Dim MixedCardSet(UBound(CardSet)) As Card
        Dim CardID As Integer
        For i = 0 To UBound(CardSet)
            Randomize()
            CardID = Math.Ceiling(Rnd() * UBound(CardSet))
            'MessageBox.Show(CardID)
            MixedCardSet(i) = CardSet(CardID)
            For j = 0 To UBound(CardSet)
                If CardSet(j).CardFamily = CardSet(CardID).CardFamily And CardSet(j).CardNumber = CardSet(CardID).CardNumber Then
                    If UBound(CardSet) > 1 And j < UBound(CardSet) Then
                        For k = j To UBound(CardSet) - 1
                            CardSet(k) = CardSet(k + 1)
                        Next
                    End If
                    ReDim Preserve CardSet(UBound(CardSet) - 1)
                    Exit For
                End If
            Next
        Next
        Return MixedCardSet
    End Function

    Private Sub NewGameMenuItem_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles NewGameMenuItem.Click
        If _GameStarted = True Then
            _GameTime.Stop()
            _GameThread.Suspend()
            Dim AskGame As MessageBoxResult = MessageBox.Show("Voulez-vous vraiment recommencer le jeu ?", "MyTarneeb", MessageBoxButton.YesNo, MessageBoxImage.Question)
            If AskGame = MessageBoxResult.Yes Then
                _GameStarted = False
                _GamePreparationThread = New Thread(AddressOf NewGame)
                _GamePreparationThread.SetApartmentState(ApartmentState.STA)
                _GamePreparationThread.IsBackground = True
                _GamePreparationThread.Start()
                _GameThread = Nothing
            Else
                _GameTime.Start()
                _GameThread.Resume()
            End If
        Else
            If _PreparationEnded = True Then
                _GamePreparationThread = New Thread(AddressOf NewGame)
                _GamePreparationThread.SetApartmentState(ApartmentState.STA)
                _GamePreparationThread.IsBackground = True
                _GamePreparationThread.Start()
            End If
        End If
    End Sub

    Private Sub QuitMenuItem_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles QuitMenuItem.Click
        Me.Close()
    End Sub

    Private Sub MainWindow_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
        If BettingPhase = True Then
            Try
                AppActivate(LanguageManager.GetStrFromModel("Pariez !"))
            Catch ex As Exception

            End Try
        End If
        If SeekingTarneebPhase = True Then
            Try
                AppActivate(LanguageManager.GetStrFromModel("Sélectionnez le Tarneeb"))
            Catch ex As Exception

            End Try
        End If
    End Sub

    Private Sub MainWindow_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles Me.Closing
            If _GameStarted = True Then
                _GameTime.Stop()
                _GameThread.Suspend()
            Dim UserConfirmation As MessageBoxResult = MessageBox.Show(LanguageManager.GetStrFromModel("Voulez-vous vraiment quitter MyTarneeb ?"), "MyTarneeb", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No)
                If UserConfirmation = MessageBoxResult.Yes Then
                    My.Application.Shutdown()
                Else
                    e.Cancel = True
                    _GameTime.Start()
                    _GameThread.Resume()
                End If
            End If
    End Sub

    Private Sub AboutMenuItem_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles AboutMenuItem.Click
        Dim _AboutBox As New AboutBox
        _AboutBox.ShowDialog()
    End Sub

    Public Shared Function GetRotateRenderTransformGroup(ByVal CenterX As UInteger, ByVal CenterY As UInteger, ByVal Angle As Integer) As TransformGroup
        Dim _tg As New TransformGroup
        Dim _rotatetransform As New RotateTransform(Angle, CenterX, CenterY)
        _tg.Children.Add(_rotatetransform)
        Return _tg
    End Function

    Private Function GetHighestBet(ByVal BetCollection As Bet()) As Bet
        If Not IsNothing(BetCollection) Then
            Dim HighestBet As New Bet(0, CardFamilies.Diamond)
            For i = 0 To UBound(BetCollection)
                If BetCollection(i).Bet > HighestBet.Bet Then
                    HighestBet = BetCollection(i)
                End If
            Next
            Return HighestBet
        Else
            Throw New Exception("Invalid Bet Collection")
        End If
    End Function

    Private Function GetBetID(ByVal BetCollection As Bet(), ByVal TheBet As Bet) As Byte
        If Not IsNothing(BetCollection) And Not IsNothing(TheBet) Then
            For i = 0 To UBound(BetCollection)
                If BetCollection(i).CardFamily = TheBet.CardFamily And BetCollection(i).Bet = TheBet.Bet Then
                    Return i
                    Exit Function
                End If
            Next
            Throw New Exception("TheBet cannot be found in the BetCollection")
        Else
            Throw New Exception("Invalid Function Parameters")
        End If
    End Function

    Private Structure FamilyInfo
        Private _Family As CardFamilies
        Private _NumOfCard As Byte

        Public Property Family As CardFamilies
            Set(ByVal value As CardFamilies)
                _Family = value
            End Set
            Get
                Return _Family
            End Get
        End Property

        Public Property NumOfCards As Byte
            Set(ByVal value As Byte)
                _NumOfCard = value
            End Set
            Get
                Return _NumOfCard
            End Get
        End Property

        Public Sub New(ByVal Family As CardFamilies, ByVal NumOfCards As Byte)
            _Family = Family
            _NumOfCard = NumOfCards
        End Sub
    End Structure

    Private Function SortCards(ByVal _PlayerCards As Card(), ByVal PlayerID As Byte, ByVal OrderFamiles As Boolean) As Card()
        Dim SpadeList(0) As Card, ClubList(0) As Card, DiamondList(0) As Card, HeartList(0) As Card
        Dim iSpade As Byte = 0, iClub As Byte = 0, iDiamond As Byte = 0, iHeart As Byte = 0
        Dim NumList(3) As FamilyInfo
        Dim SortedPlayerCards(0) As Card
        For i = 0 To UBound(_PlayerCards)
            If _PlayerCards(i).CardFamily = CardFamilies.Diamond And _PlayerCards(i).CardNumber > 0 Then
                ReDim Preserve DiamondList(iDiamond)
                DiamondList(iDiamond) = _PlayerCards(i)
                iDiamond += 1
            ElseIf _PlayerCards(i).CardFamily = CardFamilies.Club And _PlayerCards(i).CardNumber > 0 Then
                ReDim Preserve ClubList(iClub)
                ClubList(iClub) = _PlayerCards(i)
                iClub += 1
            ElseIf _PlayerCards(i).CardFamily = CardFamilies.Heart And _PlayerCards(i).CardNumber > 0 Then
                ReDim Preserve HeartList(iHeart)
                HeartList(iHeart) = _PlayerCards(i)
                iHeart += 1
            ElseIf _PlayerCards(i).CardFamily = CardFamilies.Spade And _PlayerCards(i).CardNumber > 0 Then
                ReDim Preserve SpadeList(iSpade)
                SpadeList(iSpade) = _PlayerCards(i)
                iSpade += 1
            End If
        Next
        If My.Settings.CardsByDescending = True Then
            DiamondList = DiamondList.OrderByDescending(Function(Card) Card.CardNumber).ToArray()
            SpadeList = SpadeList.OrderByDescending(Function(Card) Card.CardNumber).ToArray()
            ClubList = ClubList.OrderByDescending(Function(Card) Card.CardNumber).ToArray()
            HeartList = HeartList.OrderByDescending(Function(Card) Card.CardNumber).ToArray()
        Else
            DiamondList = DiamondList.OrderBy(Function(Card) Card.CardNumber).ToArray()
            SpadeList = SpadeList.OrderBy(Function(Card) Card.CardNumber).ToArray()
            ClubList = ClubList.OrderBy(Function(Card) Card.CardNumber).ToArray()
            HeartList = HeartList.OrderBy(Function(Card) Card.CardNumber).ToArray()
        End If
        If OrderFamiles = True Then
            NumList(0) = New FamilyInfo(CardFamilies.Diamond, DiamondList.Length)
            NumList(1) = New FamilyInfo(CardFamilies.Spade, SpadeList.Length)
            NumList(2) = New FamilyInfo(CardFamilies.Heart, HeartList.Length)
            NumList(3) = New FamilyInfo(CardFamilies.Club, ClubList.Length)
            If My.Settings.FamiliesByDescending = True Then
                NumList = NumList.OrderByDescending(Function(FamilyInfo) FamilyInfo.NumOfCards).ToArray()
            Else
                NumList = NumList.OrderBy(Function(FamilyInfo) FamilyInfo.NumOfCards).ToArray()
            End If
            Dim iCard As Byte = 0
            For i = 0 To 3
                If NumList(i).Family = CardFamilies.Diamond Then
                    For Each c As Card In DiamondList
                        ReDim Preserve SortedPlayerCards(iCard)
                        SortedPlayerCards(iCard) = c
                        iCard += 1
                    Next
                    Dim FamiliesOrder As String = _Players(PlayerID).FamiliesOrder
                    Dim FamiliesOrderModified As String = ""
                    Dim FamiliyID As String = CardFamilies.Diamond
                    For j = 0 To 3
                        If j = i Then
                            FamiliesOrderModified &= FamiliyID
                        Else
                            FamiliesOrderModified &= FamiliesOrder(j).ToString()
                        End If
                    Next
                    _Players(PlayerID).SetFamiliesOrder(FamiliesOrderModified)
                ElseIf NumList(i).Family = CardFamilies.Club Then
                    For Each c As Card In ClubList
                        ReDim Preserve SortedPlayerCards(iCard)
                        SortedPlayerCards(iCard) = c
                        iCard += 1
                    Next
                    Dim FamiliesOrder As String = _Players(PlayerID).FamiliesOrder
                    Dim FamiliesOrderModified As String = ""
                    Dim FamiliyID As String = CardFamilies.Club
                    For j = 0 To 3
                        If j = i Then
                            FamiliesOrderModified &= FamiliyID
                        Else
                            FamiliesOrderModified &= FamiliesOrder(j).ToString()
                        End If
                    Next
                    _Players(PlayerID).SetFamiliesOrder(FamiliesOrderModified)
                ElseIf NumList(i).Family = CardFamilies.Heart Then
                    For Each c As Card In HeartList
                        ReDim Preserve SortedPlayerCards(iCard)
                        SortedPlayerCards(iCard) = c
                        iCard += 1
                    Next
                    Dim FamiliesOrder As String = _Players(PlayerID).FamiliesOrder
                    Dim FamiliesOrderModified As String = ""
                    Dim FamiliyID As String = CardFamilies.Heart
                    For j = 0 To 3
                        If j = i Then
                            FamiliesOrderModified &= FamiliyID
                        Else
                            FamiliesOrderModified &= FamiliesOrder(j).ToString()
                        End If
                    Next
                    _Players(PlayerID).SetFamiliesOrder(FamiliesOrderModified)
                ElseIf NumList(i).Family = CardFamilies.Spade Then
                    For Each c As Card In SpadeList
                        ReDim Preserve SortedPlayerCards(iCard)
                        SortedPlayerCards(iCard) = c
                        iCard += 1
                    Next
                    Dim FamiliesOrder As String = _Players(PlayerID).FamiliesOrder
                    Dim FamiliesOrderModified As String = ""
                    Dim FamiliyID As String = CardFamilies.Spade
                    For j = 0 To 3
                        If j = i Then
                            FamiliesOrderModified &= FamiliyID
                        Else
                            FamiliesOrderModified &= FamiliesOrder(j).ToString()
                        End If
                    Next
                    _Players(PlayerID).SetFamiliesOrder(FamiliesOrderModified)
                Else
                    Throw New Exception("Big Problem In Card Sorting And FamilyInfo Structures")
                End If
            Next
        Else
            Dim iCard As Byte = 0
            For i = 0 To 3
                If _Players(PlayerID).FamiliesOrder(i).ToString() = CardFamilies.Spade Then
                    For Each c As Card In SpadeList
                        ReDim Preserve SortedPlayerCards(iCard)
                        SortedPlayerCards(iCard) = c
                        iCard += 1
                    Next
                ElseIf _Players(PlayerID).FamiliesOrder(i).ToString() = CardFamilies.Heart Then
                    For Each c As Card In HeartList
                        ReDim Preserve SortedPlayerCards(iCard)
                        SortedPlayerCards(iCard) = c
                        iCard += 1
                    Next
                ElseIf _Players(PlayerID).FamiliesOrder(i).ToString() = CardFamilies.Club Then
                    For Each c As Card In ClubList
                        ReDim Preserve SortedPlayerCards(iCard)
                        SortedPlayerCards(iCard) = c
                        iCard += 1
                    Next
                ElseIf _Players(PlayerID).FamiliesOrder(i).ToString() = CardFamilies.Diamond Then
                    For Each c As Card In DiamondList
                        ReDim Preserve SortedPlayerCards(iCard)
                        SortedPlayerCards(iCard) = c
                        iCard += 1
                    Next
                Else
                    Throw New Exception("Big Problem In Card Sorting And FamilyInfo Structures")
                End If
            Next
        End If
        Return CleanCardDeck(SortedPlayerCards)
    End Function

    Private Function SetLastTurnCardToolTip(ByVal TurnCard As Card(), ByVal PlayersTurn As String)
        Dim ToolTipText As String = ""
        For i = 0 To 3
            ToolTipText = TurnCard(i).CardName
            If TurnCard(i).CardFamily = My.Settings.Tarneeb Then
                ToolTipText &= " (Tarneeb)"
            End If
            ToolTipText &= " " & LanguageManager.GetStrFromModel("jouée par") & " "
            If PlayersTurn(i).ToString = 0 Then
                ToolTipText &= LanguageManager.GetStrFromModel("vous")
            ElseIf PlayersTurn(i).ToString = 1 Then
                ToolTipText &= LanguageManager.GetStrFromModel("le joueur à votre droite")
            ElseIf PlayersTurn(i).ToString = 2 Then
                ToolTipText &= LanguageManager.GetStrFromModel("votre coéquipier")
            ElseIf PlayersTurn(i).ToString = 3 Then
                ToolTipText &= LanguageManager.GetStrFromModel("le joueur à votre gauche")
            End If
            If i = 0 Then
                SyncLock SetLastTurnCard1ToolTip
                    Me.Dispatcher.Invoke(SetLastTurnCard1ToolTip, {ToolTipText})
                End SyncLock
            ElseIf i = 1 Then
                SyncLock SetLastTurnCard2ToolTip
                    Me.Dispatcher.Invoke(SetLastTurnCard2ToolTip, {ToolTipText})
                End SyncLock
            ElseIf i = 2 Then
                SyncLock SetLastTurnCard3ToolTip
                    Me.Dispatcher.Invoke(SetLastTurnCard3ToolTip, {ToolTipText})
                End SyncLock
            ElseIf i = 3 Then
                SyncLock SetLastTurnCard4ToolTip
                    Me.Dispatcher.Invoke(SetLastTurnCard4ToolTip, {ToolTipText})
                End SyncLock
            End If
        Next
    End Function

    Private Sub DelSetLastTurnCard1ToolTip(ByVal Text As String)
        LastTurnCard1.ToolTip = Text
    End Sub

    Private Sub DelSetLastTurnCard2ToolTip(ByVal Text As String)
        LastTurnCard2.ToolTip = Text
    End Sub

    Private Sub DelSetLastTurnCard3ToolTip(ByVal Text As String)
        LastTurnCard3.ToolTip = Text
    End Sub

    Private Sub DelSetLastTurnCard4ToolTip(ByVal Text As String)
        LastTurnCard4.ToolTip = Text
    End Sub

    Private Delegate Sub _DelSetLastTurnCardToolTip(ByVal Text As String)
    Private Delegate Sub _DelSub()

    Private SetLastTurnCard1ToolTip As New _DelSetLastTurnCardToolTip(AddressOf DelSetLastTurnCard1ToolTip)
    Private SetLastTurnCard2ToolTip As New _DelSetLastTurnCardToolTip(AddressOf DelSetLastTurnCard2ToolTip)
    Private SetLastTurnCard3ToolTip As New _DelSetLastTurnCardToolTip(AddressOf DelSetLastTurnCard3ToolTip)
    Private SetLastTurnCard4ToolTip As New _DelSetLastTurnCardToolTip(AddressOf DelSetLastTurnCard4ToolTip)
    Private DelPrepareCardControls As New _DelSub(AddressOf _PrepareCardControls)
    Private DelCleanGameTable As New _DelSub(AddressOf _CleanGameTable)

    Private Sub CleanGameTable()
        SyncLock DelCleanGameTable
            Me.Dispatcher.Invoke(DelCleanGameTable)
        End SyncLock
    End Sub

    Private Sub _CleanGameTable()
        IA1CardControl1.Visibility = Windows.Visibility.Collapsed
        IA1CardControl2.Visibility = Windows.Visibility.Collapsed
        IA1CardControl3.Visibility = Windows.Visibility.Collapsed
        IA1CardControl4.Visibility = Windows.Visibility.Collapsed
        IA1CardControl5.Visibility = Windows.Visibility.Collapsed
        IA1CardControl6.Visibility = Windows.Visibility.Collapsed
        IA1CardControl7.Visibility = Windows.Visibility.Collapsed
        IA1CardControl8.Visibility = Windows.Visibility.Collapsed
        IA1CardControl9.Visibility = Windows.Visibility.Collapsed
        IA1CardControl10.Visibility = Windows.Visibility.Collapsed
        IA1CardControl11.Visibility = Windows.Visibility.Collapsed
        IA1CardControl12.Visibility = Windows.Visibility.Collapsed
        IA1CardControl13.Visibility = Windows.Visibility.Collapsed
        IA2CardControl1.Visibility = Windows.Visibility.Collapsed
        IA2CardControl2.Visibility = Windows.Visibility.Collapsed
        IA2CardControl3.Visibility = Windows.Visibility.Collapsed
        IA2CardControl4.Visibility = Windows.Visibility.Collapsed
        IA2CardControl5.Visibility = Windows.Visibility.Collapsed
        IA2CardControl6.Visibility = Windows.Visibility.Collapsed
        IA2CardControl7.Visibility = Windows.Visibility.Collapsed
        IA2CardControl8.Visibility = Windows.Visibility.Collapsed
        IA2CardControl9.Visibility = Windows.Visibility.Collapsed
        IA2CardControl10.Visibility = Windows.Visibility.Collapsed
        IA2CardControl11.Visibility = Windows.Visibility.Collapsed
        IA2CardControl12.Visibility = Windows.Visibility.Collapsed
        IA2CardControl13.Visibility = Windows.Visibility.Collapsed
        IA3CardControl1.Visibility = Windows.Visibility.Collapsed
        IA3CardControl2.Visibility = Windows.Visibility.Collapsed
        IA3CardControl3.Visibility = Windows.Visibility.Collapsed
        IA3CardControl4.Visibility = Windows.Visibility.Collapsed
        IA3CardControl5.Visibility = Windows.Visibility.Collapsed
        IA3CardControl6.Visibility = Windows.Visibility.Collapsed
        IA3CardControl7.Visibility = Windows.Visibility.Collapsed
        IA3CardControl8.Visibility = Windows.Visibility.Collapsed
        IA3CardControl9.Visibility = Windows.Visibility.Collapsed
        IA3CardControl10.Visibility = Windows.Visibility.Collapsed
        IA3CardControl11.Visibility = Windows.Visibility.Collapsed
        IA3CardControl12.Visibility = Windows.Visibility.Collapsed
        IA3CardControl13.Visibility = Windows.Visibility.Collapsed
        HumanPlayerCardControl1.Visibility = Windows.Visibility.Collapsed
        HumanPlayerCardControl2.Visibility = Windows.Visibility.Collapsed
        HumanPlayerCardControl3.Visibility = Windows.Visibility.Collapsed
        HumanPlayerCardControl4.Visibility = Windows.Visibility.Collapsed
        HumanPlayerCardControl5.Visibility = Windows.Visibility.Collapsed
        HumanPlayerCardControl6.Visibility = Windows.Visibility.Collapsed
        HumanPlayerCardControl7.Visibility = Windows.Visibility.Collapsed
        HumanPlayerCardControl8.Visibility = Windows.Visibility.Collapsed
        HumanPlayerCardControl9.Visibility = Windows.Visibility.Collapsed
        HumanPlayerCardControl10.Visibility = Windows.Visibility.Collapsed
        HumanPlayerCardControl11.Visibility = Windows.Visibility.Collapsed
        HumanPlayerCardControl12.Visibility = Windows.Visibility.Collapsed
        HumanPlayerCardControl13.Visibility = Windows.Visibility.Collapsed
        TurnCardControl1.Card = Nothing
        TurnCardControl2.Card = Nothing
        TurnCardControl3.Card = Nothing
        TurnCardControl4.Card = Nothing
        LastTurnCard1.Card = Nothing
        LastTurnCard2.Card = Nothing
        LastTurnCard3.Card = Nothing
        LastTurnCard4.Card = Nothing
        LastTurnCardsLabel.Visibility = Windows.Visibility.Collapsed
        GraphicalTarneebLabel.Visibility = Windows.Visibility.Collapsed
        TarneebImage.Visibility = Windows.Visibility.Collapsed
        GameProgressBar1.IAWon = 0
        GameProgressBar1.HumanWon = 0
    End Sub

    Private Sub PrepareCardControls()
        SyncLock DelPrepareCardControls
            Me.Dispatcher.Invoke(DelPrepareCardControls)
        End SyncLock
    End Sub

    Private Sub _PrepareCardControls()
        If _GameStarted = True Then
            If TurnNumber >= 1 Then
                IA1CardControl13.Visibility = Windows.Visibility.Collapsed
                IA2CardControl13.Visibility = Windows.Visibility.Collapsed
                IA3CardControl13.Visibility = Windows.Visibility.Collapsed
                HumanPlayerCardControl13.Visibility = Windows.Visibility.Collapsed
            End If
            If TurnNumber >= 2 Then
                IA1CardControl12.Visibility = Windows.Visibility.Collapsed
                IA2CardControl12.Visibility = Windows.Visibility.Collapsed
                IA3CardControl12.Visibility = Windows.Visibility.Collapsed
                HumanPlayerCardControl12.Visibility = Windows.Visibility.Collapsed
            End If
            If TurnNumber >= 3 Then
                IA1CardControl11.Visibility = Windows.Visibility.Collapsed
                IA2CardControl11.Visibility = Windows.Visibility.Collapsed
                IA3CardControl11.Visibility = Windows.Visibility.Collapsed
                HumanPlayerCardControl11.Visibility = Windows.Visibility.Collapsed
            End If
            If TurnNumber >= 4 Then
                IA1CardControl10.Visibility = Windows.Visibility.Collapsed
                IA2CardControl10.Visibility = Windows.Visibility.Collapsed
                IA3CardControl10.Visibility = Windows.Visibility.Collapsed
                HumanPlayerCardControl10.Visibility = Windows.Visibility.Collapsed
            End If
            If TurnNumber >= 5 Then
                IA1CardControl9.Visibility = Windows.Visibility.Collapsed
                IA2CardControl9.Visibility = Windows.Visibility.Collapsed
                IA3CardControl9.Visibility = Windows.Visibility.Collapsed
                HumanPlayerCardControl9.Visibility = Windows.Visibility.Collapsed
            End If
            If TurnNumber >= 6 Then
                IA1CardControl8.Visibility = Windows.Visibility.Collapsed
                IA2CardControl8.Visibility = Windows.Visibility.Collapsed
                IA3CardControl8.Visibility = Windows.Visibility.Collapsed
                HumanPlayerCardControl8.Visibility = Windows.Visibility.Collapsed
            End If
            If TurnNumber >= 7 Then
                IA1CardControl7.Visibility = Windows.Visibility.Collapsed
                IA2CardControl7.Visibility = Windows.Visibility.Collapsed
                IA3CardControl7.Visibility = Windows.Visibility.Collapsed
                HumanPlayerCardControl7.Visibility = Windows.Visibility.Collapsed
            End If
            If TurnNumber >= 8 Then
                IA1CardControl6.Visibility = Windows.Visibility.Collapsed
                IA2CardControl6.Visibility = Windows.Visibility.Collapsed
                IA3CardControl6.Visibility = Windows.Visibility.Collapsed
                HumanPlayerCardControl6.Visibility = Windows.Visibility.Collapsed
            End If
            If TurnNumber >= 9 Then
                IA1CardControl5.Visibility = Windows.Visibility.Collapsed
                IA2CardControl5.Visibility = Windows.Visibility.Collapsed
                IA3CardControl5.Visibility = Windows.Visibility.Collapsed
                HumanPlayerCardControl5.Visibility = Windows.Visibility.Collapsed
            End If
            If TurnNumber >= 10 Then
                IA1CardControl4.Visibility = Windows.Visibility.Collapsed
                IA2CardControl4.Visibility = Windows.Visibility.Collapsed
                IA3CardControl4.Visibility = Windows.Visibility.Collapsed
                HumanPlayerCardControl4.Visibility = Windows.Visibility.Collapsed
            End If
            If TurnNumber >= 11 Then
                IA1CardControl3.Visibility = Windows.Visibility.Collapsed
                IA2CardControl3.Visibility = Windows.Visibility.Collapsed
                IA3CardControl3.Visibility = Windows.Visibility.Collapsed
                HumanPlayerCardControl3.Visibility = Windows.Visibility.Collapsed
            End If
            If TurnNumber >= 12 Then
                IA1CardControl2.Visibility = Windows.Visibility.Collapsed
                IA2CardControl2.Visibility = Windows.Visibility.Collapsed
                IA3CardControl2.Visibility = Windows.Visibility.Collapsed
                HumanPlayerCardControl2.Visibility = Windows.Visibility.Collapsed
            End If
        Else
            IA1CardControl1.Visibility = Windows.Visibility.Visible
            IA1CardControl2.Visibility = Windows.Visibility.Visible
            IA1CardControl3.Visibility = Windows.Visibility.Visible
            IA1CardControl4.Visibility = Windows.Visibility.Visible
            IA1CardControl5.Visibility = Windows.Visibility.Visible
            IA1CardControl6.Visibility = Windows.Visibility.Visible
            IA1CardControl7.Visibility = Windows.Visibility.Visible
            IA1CardControl8.Visibility = Windows.Visibility.Visible
            IA1CardControl9.Visibility = Windows.Visibility.Visible
            IA1CardControl10.Visibility = Windows.Visibility.Visible
            IA1CardControl11.Visibility = Windows.Visibility.Visible
            IA1CardControl12.Visibility = Windows.Visibility.Visible
            IA1CardControl13.Visibility = Windows.Visibility.Visible
            IA2CardControl1.Visibility = Windows.Visibility.Visible
            IA2CardControl2.Visibility = Windows.Visibility.Visible
            IA2CardControl3.Visibility = Windows.Visibility.Visible
            IA2CardControl4.Visibility = Windows.Visibility.Visible
            IA2CardControl5.Visibility = Windows.Visibility.Visible
            IA2CardControl6.Visibility = Windows.Visibility.Visible
            IA2CardControl7.Visibility = Windows.Visibility.Visible
            IA2CardControl8.Visibility = Windows.Visibility.Visible
            IA2CardControl9.Visibility = Windows.Visibility.Visible
            IA2CardControl10.Visibility = Windows.Visibility.Visible
            IA2CardControl11.Visibility = Windows.Visibility.Visible
            IA2CardControl12.Visibility = Windows.Visibility.Visible
            IA2CardControl13.Visibility = Windows.Visibility.Visible
            IA3CardControl1.Visibility = Windows.Visibility.Visible
            IA3CardControl2.Visibility = Windows.Visibility.Visible
            IA3CardControl3.Visibility = Windows.Visibility.Visible
            IA3CardControl4.Visibility = Windows.Visibility.Visible
            IA3CardControl5.Visibility = Windows.Visibility.Visible
            IA3CardControl6.Visibility = Windows.Visibility.Visible
            IA3CardControl7.Visibility = Windows.Visibility.Visible
            IA3CardControl8.Visibility = Windows.Visibility.Visible
            IA3CardControl9.Visibility = Windows.Visibility.Visible
            IA3CardControl10.Visibility = Windows.Visibility.Visible
            IA3CardControl11.Visibility = Windows.Visibility.Visible
            IA3CardControl12.Visibility = Windows.Visibility.Visible
            IA3CardControl13.Visibility = Windows.Visibility.Visible
            HumanPlayerCardControl1.Visibility = Windows.Visibility.Visible
            HumanPlayerCardControl2.Visibility = Windows.Visibility.Visible
            HumanPlayerCardControl3.Visibility = Windows.Visibility.Visible
            HumanPlayerCardControl4.Visibility = Windows.Visibility.Visible
            HumanPlayerCardControl5.Visibility = Windows.Visibility.Visible
            HumanPlayerCardControl6.Visibility = Windows.Visibility.Visible
            HumanPlayerCardControl7.Visibility = Windows.Visibility.Visible
            HumanPlayerCardControl8.Visibility = Windows.Visibility.Visible
            HumanPlayerCardControl9.Visibility = Windows.Visibility.Visible
            HumanPlayerCardControl10.Visibility = Windows.Visibility.Visible
            HumanPlayerCardControl11.Visibility = Windows.Visibility.Visible
            HumanPlayerCardControl12.Visibility = Windows.Visibility.Visible
            HumanPlayerCardControl13.Visibility = Windows.Visibility.Visible
        End If
        ReorganizeCardControls()
        Dim iCard As Byte
        For i = 0 To 3
            iCard = 0
            If _GameStarted = True Then
                _Players(i).PlayerCards = SortCards(_Players(i).PlayerCards, i, False)
            Else
                _Players(i).PlayerCards = SortCards(_Players(i).PlayerCards, i, True)
            End If
            For Each c As Card In _Players(i).PlayerCards
                Select Case i
                    Case 0
                        Select Case iCard
                            Case 0
                                HumanPlayerCardControl1.SetControl(c, False)
                            Case 1
                                HumanPlayerCardControl2.SetControl(c, False)
                            Case 2
                                HumanPlayerCardControl3.SetControl(c, False)
                            Case 3
                                HumanPlayerCardControl4.SetControl(c, False)
                            Case 4
                                HumanPlayerCardControl5.SetControl(c, False)
                            Case 5
                                HumanPlayerCardControl6.SetControl(c, False)
                            Case 6
                                HumanPlayerCardControl7.SetControl(c, False)
                            Case 7
                                HumanPlayerCardControl8.SetControl(c, False)
                            Case 8
                                HumanPlayerCardControl9.SetControl(c, False)
                            Case 9
                                HumanPlayerCardControl10.SetControl(c, False)
                            Case 10
                                HumanPlayerCardControl11.SetControl(c, False)
                            Case 11
                                HumanPlayerCardControl12.SetControl(c, False)
                            Case 12
                                HumanPlayerCardControl13.SetControl(c, False)
                        End Select
                    Case 1
                        Select Case iCard
                            Case 0
                                IA1CardControl1.SetControl(c, True)
                            Case 1
                                IA1CardControl2.SetControl(c, True)
                            Case 2
                                IA1CardControl3.SetControl(c, True)
                            Case 3
                                IA1CardControl4.SetControl(c, True)
                            Case 4
                                IA1CardControl5.SetControl(c, True)
                            Case 5
                                IA1CardControl6.SetControl(c, True)
                            Case 6
                                IA1CardControl7.SetControl(c, True)
                            Case 7
                                IA1CardControl8.SetControl(c, True)
                            Case 8
                                IA1CardControl9.SetControl(c, True)
                            Case 9
                                IA1CardControl10.SetControl(c, True)
                            Case 10
                                IA1CardControl11.SetControl(c, True)
                            Case 11
                                IA1CardControl12.SetControl(c, True)
                            Case 12
                                IA1CardControl13.SetControl(c, True)
                        End Select
                    Case 2
                        Select Case iCard
                            Case 0
                                IA2CardControl1.SetControl(c, True)
                            Case 1
                                IA2CardControl2.SetControl(c, True)
                            Case 2
                                IA2CardControl3.SetControl(c, True)
                            Case 3
                                IA2CardControl4.SetControl(c, True)
                            Case 4
                                IA2CardControl5.SetControl(c, True)
                            Case 5
                                IA2CardControl6.SetControl(c, True)
                            Case 6
                                IA2CardControl7.SetControl(c, True)
                            Case 7
                                IA2CardControl8.SetControl(c, True)
                            Case 8
                                IA2CardControl9.SetControl(c, True)
                            Case 9
                                IA2CardControl10.SetControl(c, True)
                            Case 10
                                IA2CardControl11.SetControl(c, True)
                            Case 11
                                IA2CardControl12.SetControl(c, True)
                            Case 12
                                IA2CardControl13.SetControl(c, True)
                        End Select
                    Case 3
                        Select Case iCard
                            Case 0
                                IA3CardControl1.SetControl(c, True)
                            Case 1
                                IA3CardControl2.SetControl(c, True)
                            Case 2
                                IA3CardControl3.SetControl(c, True)
                            Case 3
                                IA3CardControl4.SetControl(c, True)
                            Case 4
                                IA3CardControl5.SetControl(c, True)
                            Case 5
                                IA3CardControl6.SetControl(c, True)
                            Case 6
                                IA3CardControl7.SetControl(c, True)
                            Case 7
                                IA3CardControl8.SetControl(c, True)
                            Case 8
                                IA3CardControl9.SetControl(c, True)
                            Case 9
                                IA3CardControl10.SetControl(c, True)
                            Case 10
                                IA3CardControl11.SetControl(c, True)
                            Case 11
                                IA3CardControl12.SetControl(c, True)
                            Case 12
                                IA3CardControl13.SetControl(c, True)
                        End Select
                End Select
                iCard += 1
            Next
        Next
    End Sub

    Private Sub ReorganizeCardControls()
        If TurnNumber = 12 Then
            HumanPlayerCardControl1.SetRotationAngle(0)
            IA1CardControl1.SetRotationAngle(270)
            IA2CardControl1.SetRotationAngle(180)
            IA3CardControl1.SetRotationAngle(90)
        Else
            Dim AngleBetweenCards As Decimal = My.Settings.AngleBetweenEachCard
            Dim HumanBaseAngle As Integer = -(AngleBetweenCards * (12 - TurnNumber)) / 2, TeammateBaseAngle As Integer = HumanBaseAngle + 180, AI1BaseAngle As Integer = HumanBaseAngle + 270, AI3BaseAngle As Integer = HumanBaseAngle + 90
            Dim NextAngles(3) As Decimal
            NextAngles(0) = HumanBaseAngle
            NextAngles(1) = AI1BaseAngle
            NextAngles(2) = TeammateBaseAngle
            NextAngles(3) = AI3BaseAngle
            HumanPlayerCardControl1.SetRotationAngle(NextAngles(0))
            IA1CardControl1.SetRotationAngle(NextAngles(1))
            IA2CardControl1.SetRotationAngle(NextAngles(2))
            IA3CardControl1.SetRotationAngle(NextAngles(3))
            For i = 0 To 3
                NextAngles(i) += AngleBetweenCards
            Next
            HumanPlayerCardControl2.SetRotationAngle(NextAngles(0))
            IA1CardControl2.SetRotationAngle(NextAngles(1))
            IA2CardControl2.SetRotationAngle(NextAngles(2))
            IA3CardControl2.SetRotationAngle(NextAngles(3))
            For i = 0 To 3
                NextAngles(i) += AngleBetweenCards
            Next
            HumanPlayerCardControl3.SetRotationAngle(NextAngles(0))
            IA1CardControl3.SetRotationAngle(NextAngles(1))
            IA2CardControl3.SetRotationAngle(NextAngles(2))
            IA3CardControl3.SetRotationAngle(NextAngles(3))
            For i = 0 To 3
                NextAngles(i) += AngleBetweenCards
            Next
            HumanPlayerCardControl4.SetRotationAngle(NextAngles(0))
            IA1CardControl4.SetRotationAngle(NextAngles(1))
            IA2CardControl4.SetRotationAngle(NextAngles(2))
            IA3CardControl4.SetRotationAngle(NextAngles(3))
            For i = 0 To 3
                NextAngles(i) += AngleBetweenCards
            Next
            HumanPlayerCardControl5.SetRotationAngle(NextAngles(0))
            IA1CardControl5.SetRotationAngle(NextAngles(1))
            IA2CardControl5.SetRotationAngle(NextAngles(2))
            IA3CardControl5.SetRotationAngle(NextAngles(3))
            For i = 0 To 3
                NextAngles(i) += AngleBetweenCards
            Next
            HumanPlayerCardControl6.SetRotationAngle(NextAngles(0))
            IA1CardControl6.SetRotationAngle(NextAngles(1))
            IA2CardControl6.SetRotationAngle(NextAngles(2))
            IA3CardControl6.SetRotationAngle(NextAngles(3))
            For i = 0 To 3
                NextAngles(i) += AngleBetweenCards
            Next
            HumanPlayerCardControl7.SetRotationAngle(NextAngles(0))
            IA1CardControl7.SetRotationAngle(NextAngles(1))
            IA2CardControl7.SetRotationAngle(NextAngles(2))
            IA3CardControl7.SetRotationAngle(NextAngles(3))
            For i = 0 To 3
                NextAngles(i) += AngleBetweenCards
            Next
            HumanPlayerCardControl8.SetRotationAngle(NextAngles(0))
            IA1CardControl8.SetRotationAngle(NextAngles(1))
            IA2CardControl8.SetRotationAngle(NextAngles(2))
            IA3CardControl8.SetRotationAngle(NextAngles(3))
            For i = 0 To 3
                NextAngles(i) += AngleBetweenCards
            Next
            HumanPlayerCardControl9.SetRotationAngle(NextAngles(0))
            IA1CardControl9.SetRotationAngle(NextAngles(1))
            IA2CardControl9.SetRotationAngle(NextAngles(2))
            IA3CardControl9.SetRotationAngle(NextAngles(3))
            For i = 0 To 3
                NextAngles(i) += AngleBetweenCards
            Next
            HumanPlayerCardControl10.SetRotationAngle(NextAngles(0))
            IA1CardControl10.SetRotationAngle(NextAngles(1))
            IA2CardControl10.SetRotationAngle(NextAngles(2))
            IA3CardControl10.SetRotationAngle(NextAngles(3))
            For i = 0 To 3
                NextAngles(i) += AngleBetweenCards
            Next
            HumanPlayerCardControl11.SetRotationAngle(NextAngles(0))
            IA1CardControl11.SetRotationAngle(NextAngles(1))
            IA2CardControl11.SetRotationAngle(NextAngles(2))
            IA3CardControl11.SetRotationAngle(NextAngles(3))
            For i = 0 To 3
                NextAngles(i) += AngleBetweenCards
            Next
            HumanPlayerCardControl12.SetRotationAngle(NextAngles(0))
            IA1CardControl12.SetRotationAngle(NextAngles(1))
            IA2CardControl12.SetRotationAngle(NextAngles(2))
            IA3CardControl12.SetRotationAngle(NextAngles(3))
            For i = 0 To 3
                NextAngles(i) += AngleBetweenCards
            Next
            HumanPlayerCardControl13.SetRotationAngle(NextAngles(0))
            IA1CardControl13.SetRotationAngle(NextAngles(1))
            IA2CardControl13.SetRotationAngle(NextAngles(2))
            IA3CardControl13.SetRotationAngle(NextAngles(3))
        End If
    End Sub

    Private Sub SetCardControlAngles()
        Dim AngleBetweenCards As Decimal = My.Settings.AngleBetweenEachCard
        Dim BaseAngle As Decimal = -(My.Settings.AngleBetweenEachCard * 12) / 2
        'Dim AngleBetweenCards As Decimal = 120 / (12 - TurnNumber)
        Dim HumanBaseAngle As Integer = BaseAngle, TeammateBaseAngle As Integer = HumanBaseAngle + 180, AI1BaseAngle As Integer = HumanBaseAngle + 270, AI3BaseAngle As Integer = HumanBaseAngle + 90
        Dim NextAngles(3) As Decimal
        NextAngles(0) = HumanBaseAngle
        NextAngles(1) = AI1BaseAngle
        NextAngles(2) = TeammateBaseAngle
        NextAngles(3) = AI3BaseAngle
        HumanPlayerCardControl1.SetRotationAngle(NextAngles(0))
        IA1CardControl1.SetRotationAngle(NextAngles(1))
        IA2CardControl1.SetRotationAngle(NextAngles(2))
        IA3CardControl1.SetRotationAngle(NextAngles(3))
        For i = 0 To 3
            NextAngles(i) += AngleBetweenCards
        Next
        HumanPlayerCardControl2.SetRotationAngle(NextAngles(0))
        IA1CardControl2.SetRotationAngle(NextAngles(1))
        IA2CardControl2.SetRotationAngle(NextAngles(2))
        IA3CardControl2.SetRotationAngle(NextAngles(3))
        For i = 0 To 3
            NextAngles(i) += AngleBetweenCards
        Next
        HumanPlayerCardControl3.SetRotationAngle(NextAngles(0))
        IA1CardControl3.SetRotationAngle(NextAngles(1))
        IA2CardControl3.SetRotationAngle(NextAngles(2))
        IA3CardControl3.SetRotationAngle(NextAngles(3))
        For i = 0 To 3
            NextAngles(i) += AngleBetweenCards
        Next
        HumanPlayerCardControl4.SetRotationAngle(NextAngles(0))
        IA1CardControl4.SetRotationAngle(NextAngles(1))
        IA2CardControl4.SetRotationAngle(NextAngles(2))
        IA3CardControl4.SetRotationAngle(NextAngles(3))
        For i = 0 To 3
            NextAngles(i) += AngleBetweenCards
        Next
        HumanPlayerCardControl5.SetRotationAngle(NextAngles(0))
        IA1CardControl5.SetRotationAngle(NextAngles(1))
        IA2CardControl5.SetRotationAngle(NextAngles(2))
        IA3CardControl5.SetRotationAngle(NextAngles(3))
        For i = 0 To 3
            NextAngles(i) += AngleBetweenCards
        Next
        HumanPlayerCardControl6.SetRotationAngle(NextAngles(0))
        IA1CardControl6.SetRotationAngle(NextAngles(1))
        IA2CardControl6.SetRotationAngle(NextAngles(2))
        IA3CardControl6.SetRotationAngle(NextAngles(3))
        For i = 0 To 3
            NextAngles(i) += AngleBetweenCards
        Next
        HumanPlayerCardControl7.SetRotationAngle(NextAngles(0))
        IA1CardControl7.SetRotationAngle(NextAngles(1))
        IA2CardControl7.SetRotationAngle(NextAngles(2))
        IA3CardControl7.SetRotationAngle(NextAngles(3))
        For i = 0 To 3
            NextAngles(i) += AngleBetweenCards
        Next
        HumanPlayerCardControl8.SetRotationAngle(NextAngles(0))
        IA1CardControl8.SetRotationAngle(NextAngles(1))
        IA2CardControl8.SetRotationAngle(NextAngles(2))
        IA3CardControl8.SetRotationAngle(NextAngles(3))
        For i = 0 To 3
            NextAngles(i) += AngleBetweenCards
        Next
        HumanPlayerCardControl9.SetRotationAngle(NextAngles(0))
        IA1CardControl9.SetRotationAngle(NextAngles(1))
        IA2CardControl9.SetRotationAngle(NextAngles(2))
        IA3CardControl9.SetRotationAngle(NextAngles(3))
        For i = 0 To 3
            NextAngles(i) += AngleBetweenCards
        Next
        HumanPlayerCardControl10.SetRotationAngle(NextAngles(0))
        IA1CardControl10.SetRotationAngle(NextAngles(1))
        IA2CardControl10.SetRotationAngle(NextAngles(2))
        IA3CardControl10.SetRotationAngle(NextAngles(3))
        For i = 0 To 3
            NextAngles(i) += AngleBetweenCards
        Next
        HumanPlayerCardControl11.SetRotationAngle(NextAngles(0))
        IA1CardControl11.SetRotationAngle(NextAngles(1))
        IA2CardControl11.SetRotationAngle(NextAngles(2))
        IA3CardControl11.SetRotationAngle(NextAngles(3))
        For i = 0 To 3
            NextAngles(i) += AngleBetweenCards
        Next
        HumanPlayerCardControl12.SetRotationAngle(NextAngles(0))
        IA1CardControl12.SetRotationAngle(NextAngles(1))
        IA2CardControl12.SetRotationAngle(NextAngles(2))
        IA3CardControl12.SetRotationAngle(NextAngles(3))
        For i = 0 To 3
            NextAngles(i) += AngleBetweenCards
        Next
        HumanPlayerCardControl13.SetRotationAngle(NextAngles(0))
        IA1CardControl13.SetRotationAngle(NextAngles(1))
        IA2CardControl13.SetRotationAngle(NextAngles(2))
        IA3CardControl13.SetRotationAngle(NextAngles(3))
    End Sub

    Private Sub MainWindow_GameFinished() Handles Me.GameFinished
        Me.Dispatcher.Invoke(DelSetGameFinishedMenuItems)
    End Sub

    Private Sub MainWindow_GameStarting() Handles Me.GameStarting
        Me.Dispatcher.Invoke(DelSetGameStartingMenuItems)
    End Sub

    Private DelSetGameFinishedMenuItems As New _DelSub(AddressOf _SetGameFinishedMenuItems)
    Private DelSetGameStartingMenuItems As New _DelSub(AddressOf _SetGameStartingMenuItems)

    Private Sub _SetGameFinishedMenuItems()
        FinishGameMenuItem.IsEnabled = False
        If _GameStarted = False Then
            PauseMenuItem.IsEnabled = False
        End If
        NewGameMenuItem.IsEnabled = True
        LanguageMenuItem.IsEnabled = True
        ViewMenuItem.IsEnabled = True
        GameLevelMenuItem.IsEnabled = True
        StatsMenuItem.IsEnabled = True
        SettingsMenuItem.IsEnabled = True
        GameRulesMenuItem.IsEnabled = True
        AboutMenuItem.IsEnabled = True
    End Sub

    Private Sub _SetGameStartingMenuItems()
        FinishGameMenuItem.IsEnabled = True
        PauseMenuItem.IsEnabled = True
        NewGameMenuItem.IsEnabled = False
        LanguageMenuItem.IsEnabled = False
        ViewMenuItem.IsEnabled = False
        GameLevelMenuItem.IsEnabled = False
        StatsMenuItem.IsEnabled = False
        SettingsMenuItem.IsEnabled = False
        GameRulesMenuItem.IsEnabled = False
        AboutMenuItem.IsEnabled = False
    End Sub

    Private Sub MainWindow_Loaded(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles Me.Loaded
        My.Application.MainWindow = Me
        _LoadThread = New Thread(AddressOf LoadApp)
        _LoadThread.IsBackground = True
        _LoadThread.Start()
    End Sub

    Private Sub LoadApp()
        If My.Settings.UseCustomBackSide = True And IO.File.Exists(My.Application.Info.DirectoryPath & "\UserBackSides\" & My.Settings.CustomBackSideFileName) = False Then
            MessageBox.Show("Le fichier de motif sélectionné est introuvable. Veuillez vérifier vos paramètres.", "Motif personnalisé", MessageBoxButton.OK, MessageBoxImage.Error)
            My.Settings.UseCustomBackSide = False
            My.Settings.Save()
        End If
        'ResetGameSet()
        My.Settings.PreparationPlayersOrder = ""
        My.Settings.Save()
        LanguageManager.VerifyFiles()
        Me.Dispatcher.Invoke(LoadLanguageFiles)
        '_GamePreparationThread = New Thread(AddressOf NewGame)
        '_GamePreparationThread.IsBackground = True
        '_GamePreparationThread.SetApartmentState(ApartmentState.STA)
        'Dim GameStartChoice As MessageBoxResult = MessageBox.Show(LanguageManager.GetStrFromModel("Voulez-vous commencer une nouvelle partie ?"), "MyTarneeb", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.Yes)
        'If GameStartChoice = MessageBoxResult.Yes Then
        '_GamePreparationThread.Start()
        'End If
    End Sub

    Private Sub ResetGameSet()
        _GameSet = MixCardSet(Card.GameSet)
    End Sub

    Private Sub AskForNewGame()
        _GamePreparationThread = New Thread(AddressOf NewGame)
        _GamePreparationThread.IsBackground = True
        _GamePreparationThread.SetApartmentState(ApartmentState.STA)
        Dim GameStartChoice As MessageBoxResult = MessageBox.Show(LanguageManager.GetStrFromModel("Voulez-vous commencer une nouvelle partie ?"), "MyTarneeb", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.Yes)
        If GameStartChoice = MessageBoxResult.Yes Then
            _GameThread = Nothing
            _GamePreparationThread.Start()
        End If
    End Sub

    Private Sub MainWindow_StateChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.StateChanged
        If Me.WindowState = Windows.WindowState.Minimized And _GameStarted = True Then
            _GameTime.Stop()
            _GameThread.Suspend()
        Else
            If _GameStarted = True Then
                _GameTime.Start()
                Try
                    _GameThread.Resume()
                Catch ex As Exception

                End Try
            End If
        End If
    End Sub

    Private Function DefinePlayersOrder(ByVal TurnCards As Card()) As Byte
        Dim TurnFamily As CardFamilies = TurnCards(0).CardFamily
        Dim WinningCard As Card
        Dim WinningPlayerID As Byte, WinningCardID As Byte
        If TurnFamily <> My.Settings.Tarneeb Then
            Dim TurnFamilyList As Card() = GetFamilyList(TurnCards, TurnFamily)
            Dim TarneebList As Card() = GetFamilyList(TurnCards, My.Settings.Tarneeb)
            If IsNothing(TarneebList) = True Then
                WinningCard = GetHighestCard(TurnFamilyList)
            Else
                WinningCard = GetHighestCard(TarneebList)
            End If
        Else
            Dim TarneebList As Card() = GetFamilyList(TurnCards, My.Settings.Tarneeb)
            WinningCard = GetHighestCard(TarneebList)
        End If
        WinningCardID = GetCardID(TurnCards, WinningCard)
        WinningPlayerID = My.Settings.PlayersTurn(WinningCardID).ToString()
        DefinePlayersOrder(WinningPlayerID)
        Return WinningPlayerID
    End Function

    Private Sub DefinePlayersOrder(ByVal StartingPlayerID As Byte)
        Dim PlayersTurn As String
        Dim PreviousID As Integer = StartingPlayerID
        PlayersTurn = PreviousID.ToString()
        For i = 0 To 2
            If PreviousID = 3 Then
                PreviousID = 0
            Else
                PreviousID += 1
            End If
            PlayersTurn &= PreviousID.ToString()
        Next
        My.Settings.PlayersTurn = PlayersTurn
    End Sub

    Private Sub DefinePlayersBetOrder()
        If My.Settings.PreparationPlayersOrder = "" Then
            Randomize()
            Dim FirstPlayerID As Byte = Math.Round(Rnd() * 3)
            Dim PlayersOrder As String = FirstPlayerID.ToString()
            Dim PlayerID As Byte = FirstPlayerID
            For i = 1 To 3
                If PlayerID = 3 Then
                    PlayerID = 0
                Else
                    PlayerID += 1
                End If
                PlayersOrder &= PlayerID.ToString()
            Next
            My.Settings.PreparationPlayersOrder = PlayersOrder
        Else
            Dim PreviousPlayersOrder As String = My.Settings.PreparationPlayersOrder, PlayersOrder As String
            Dim PreviousFirstPlayerID As Byte = Val(PreviousPlayersOrder(0)), FirstPlayerID As Byte, PlayerID As Byte
            If PreviousFirstPlayerID = 3 Then
                FirstPlayerID = 0
            Else
                FirstPlayerID = PreviousFirstPlayerID + 1
            End If
            PlayersOrder = FirstPlayerID.ToString()
            PlayerID = FirstPlayerID
            For i = 0 To 2
                If PlayerID = 3 Then
                    PlayerID = 0
                Else
                    PlayerID += 1
                End If
                PlayersOrder &= PlayerID.ToString()
            Next
            My.Settings.PreparationPlayersOrder = PlayersOrder
        End If
        My.Settings.Save()
    End Sub

    Private Sub ShowCard(ByVal PlayerID As Byte)
        Select Case PlayerID
            Case 0
                Select Case _Players(PlayerID).TurnCardID
                    Case 0
                        HumanPlayerCardControl1.CardControl1.UseCard(False)
                    Case 1
                        HumanPlayerCardControl2.CardControl1.UseCard(False)
                    Case 2
                        HumanPlayerCardControl3.CardControl1.UseCard(False)
                    Case 3
                        HumanPlayerCardControl4.CardControl1.UseCard(False)
                    Case 4
                        HumanPlayerCardControl5.CardControl1.UseCard(False)
                    Case 5
                        HumanPlayerCardControl6.CardControl1.UseCard(False)
                    Case 6
                        HumanPlayerCardControl7.CardControl1.UseCard(False)
                    Case 7
                        HumanPlayerCardControl8.CardControl1.UseCard(False)
                    Case 8
                        HumanPlayerCardControl9.CardControl1.UseCard(False)
                    Case 9
                        HumanPlayerCardControl10.CardControl1.UseCard(False)
                    Case 10
                        HumanPlayerCardControl11.CardControl1.UseCard(False)
                    Case 11
                        HumanPlayerCardControl12.CardControl1.UseCard(False)
                    Case 12
                        HumanPlayerCardControl13.CardControl1.UseCard(False)
                End Select
                TurnCardControl1.Card = _Players(PlayerID).TurnCard
            Case 1
                Select Case _Players(PlayerID).TurnCardID
                    Case 0
                        IA1CardControl1.UseCard(True)
                    Case 1
                        IA1CardControl2.UseCard(True)
                    Case 2
                        IA1CardControl3.UseCard(True)
                    Case 3
                        IA1CardControl4.UseCard(True)
                    Case 4
                        IA1CardControl5.UseCard(True)
                    Case 5
                        IA1CardControl6.UseCard(True)
                    Case 6
                        IA1CardControl7.UseCard(True)
                    Case 7
                        IA1CardControl8.UseCard(True)
                    Case 8
                        IA1CardControl9.UseCard(True)
                    Case 9
                        IA1CardControl10.UseCard(True)
                    Case 10
                        IA1CardControl11.UseCard(True)
                    Case 11
                        IA1CardControl12.UseCard(True)
                    Case 12
                        IA1CardControl13.UseCard(True)
                End Select
                TurnCardControl2.Card = _Players(PlayerID).TurnCard
            Case 2
                Select Case _Players(PlayerID).TurnCardID
                    Case 0
                        IA2CardControl1.UseCard(True)
                    Case 1
                        IA2CardControl2.UseCard(True)
                    Case 2
                        IA2CardControl3.UseCard(True)
                    Case 3
                        IA2CardControl4.UseCard(True)
                    Case 4
                        IA2CardControl5.UseCard(True)
                    Case 5
                        IA2CardControl6.UseCard(True)
                    Case 6
                        IA2CardControl7.UseCard(True)
                    Case 7
                        IA2CardControl8.UseCard(True)
                    Case 8
                        IA2CardControl9.UseCard(True)
                    Case 9
                        IA2CardControl10.UseCard(True)
                    Case 10
                        IA2CardControl11.UseCard(True)
                    Case 11
                        IA2CardControl12.UseCard(True)
                    Case 12
                        IA2CardControl13.UseCard(True)
                End Select
                TurnCardControl3.Card = _Players(PlayerID).TurnCard
            Case 3
                Select Case _Players(PlayerID).TurnCardID
                    Case 0
                        IA3CardControl1.UseCard(True)
                    Case 1
                        IA3CardControl2.UseCard(True)
                    Case 2
                        IA3CardControl3.UseCard(True)
                    Case 3
                        IA3CardControl4.UseCard(True)
                    Case 4
                        IA3CardControl5.UseCard(True)
                    Case 5
                        IA3CardControl6.UseCard(True)
                    Case 6
                        IA3CardControl7.UseCard(True)
                    Case 7
                        IA3CardControl8.UseCard(True)
                    Case 8
                        IA3CardControl9.UseCard(True)
                    Case 9
                        IA3CardControl10.UseCard(True)
                    Case 10
                        IA3CardControl11.UseCard(True)
                    Case 11
                        IA3CardControl12.UseCard(True)
                    Case 12
                        IA3CardControl13.UseCard(True)
                End Select
                TurnCardControl4.Card = _Players(PlayerID).TurnCard
        End Select
    End Sub

    Private Sub ShowTime()
        Dim ExitShowTime As Boolean = False
        Do While ExitShowTime = False
            Thread.Sleep(1000)
            SyncLock DelSetTimeLabelText
                Dim ElapsedTime As TimeSpan = _GameTime.Elapsed
                Dim TimeText As String = LanguageManager.GetStrFromModel("Temps de jeu") & " : "
                If ElapsedTime.Hours = 0 Then
                    TimeText &= "00"
                ElseIf ElapsedTime.Hours > 0 And ElapsedTime.Hours < 10 Then
                    TimeText &= "0" & ElapsedTime.Hours.ToString()
                Else
                    TimeText &= ElapsedTime.Hours.ToString()
                End If
                TimeText &= ":"
                If ElapsedTime.Minutes = 0 Then
                    TimeText &= "00"
                ElseIf ElapsedTime.Minutes > 0 And ElapsedTime.Minutes < 10 Then
                    TimeText &= "0" & ElapsedTime.Minutes.ToString()
                Else
                    TimeText &= ElapsedTime.Minutes.ToString()
                End If
                TimeText &= ":"
                If ElapsedTime.Seconds = 0 Then
                    TimeText &= "00"
                ElseIf ElapsedTime.Seconds > 0 And ElapsedTime.Seconds < 10 Then
                    TimeText &= "0" & ElapsedTime.Seconds.ToString()
                Else
                    TimeText &= ElapsedTime.Seconds.ToString()
                End If
                Me.Dispatcher.Invoke(DelSetTimeLabelText, {TimeText})
            End SyncLock
        Loop
    End Sub

    Private Delegate Sub _DelSetTimeLabelText(ByVal Text As String)
    Private Delegate Sub _DelSetGameStateDisplay(ByVal WonTurns As Byte, ByVal LostTurns As Byte)
    Private Delegate Sub _DelSetTurnNumDisplay(ByVal TurnNum As Byte)

    Private DelSetTimeLabelText As New _DelSetTimeLabelText(AddressOf _SetTimeLabelText)
    Private DelSetGameStateDisplay As New _DelSetGameStateDisplay(AddressOf _SetGameStateDisplay)
    Private DelSetTurnNumDisplay As New _DelSetTurnNumDisplay(AddressOf _SetTurnNumDisplay)

    Private Sub _SetTimeLabelText(ByVal Text As String)
        GameTimeLabel.Content = Text
    End Sub

    Private Sub FinishGame(ByVal OnlyTarneebOwnerID As Byte)
        Dim TurnCards(3) As Card
        For i = 0 To UBound(_Players(OnlyTarneebOwnerID).PlayerCards)
            For j = 0 To 3
                If My.Settings.PlayersTurn(j).ToString() = OnlyTarneebOwnerID Then
                    TurnCards(j) = GetHighestCard(_Players(OnlyTarneebOwnerID).PlayerCards)
                Else
                    TurnCards(j) = GetLowestCard(_Players(My.Settings.PlayersTurn(j).ToString()).PlayerCards)
                End If
                Dim PlayerID As Byte = My.Settings.PlayersTurn(j).ToString()
                Dim CardID As Byte = GetCardID(_Players(PlayerID).PlayerCards, TurnCards(j))
                _Players(My.Settings.PlayersTurn(j).ToString()).PlayerCards = RemoveCardFromPlayersHands(_Players(My.Settings.PlayersTurn(j).ToString()).PlayerCards, CardID)
                ShowCard(PlayerID)
            Next
            If OnlyTarneebOwnerID = 0 Or OnlyTarneebOwnerID = 2 Then
                HumanWon += 1
            Else
                IAWon += 1
            End If
        Next
    End Sub

    Private Function RemoveCardFromPlayersHands(ByVal PlayerCards As Card(), ByVal TheIDOfCardToRemove As Byte) As Card() 'Retourne les cartes du joueur sans les valeurs à supprimer
        If Not IsNothing(PlayerCards) And TheIDOfCardToRemove <= UBound(PlayerCards) Then
            For i = 0 To UBound(PlayerCards)
                If i = TheIDOfCardToRemove Then
                    If UBound(PlayerCards) > 1 And i < UBound(PlayerCards) Then
                        For j = i To UBound(PlayerCards) - 1
                            PlayerCards(j) = PlayerCards(j + 1)
                        Next
                    End If
                    ReDim Preserve PlayerCards(UBound(PlayerCards) - 1)
                    Exit For
                End If
            Next
            Return PlayerCards
        Else
            Throw New ArgumentNullException("Invalid Card ID or PlayerCards()")
        End If
    End Function

    Private Function CleanCardDeck(ByVal CardDeck As Card()) As Card()
        Dim CleanedCardDeck(0) As Card
        Dim iCard As Byte = 0
        For Each c As Card In CardDeck
            If c.CardNumber > 0 Then
                ReDim Preserve CleanedCardDeck(iCard)
                CleanedCardDeck(iCard) = c
                iCard += 1
            End If
        Next
        Return CleanedCardDeck
    End Function

    Private Sub _SetGameStateDisplay(ByVal WonTurns As Byte, ByVal LostTurns As Byte)
        WonTurnsLabel.Content = LanguageManager.GetStrFromModel("Tours gagnés :") & " " & WonTurns
        LostTurnsLabel.Content = LanguageManager.GetStrFromModel("Tours perdus :") & " " & LostTurns
    End Sub

    Private Sub _SetTurnNumDisplay(ByVal TurnNum As Byte)
        TurnNumberLabel.Content = LanguageManager.GetStrFromModel("Tour :") & " " & TurnNum
    End Sub

    Private Sub LastTurnCard1_MouseEnter(ByVal sender As Object, ByVal e As System.Windows.Input.MouseEventArgs) Handles LastTurnCard1.MouseEnter
        LastTurnCard1.Height *= 1.2
        LastTurnCard1.Width *= 1.2
    End Sub

    Private Sub LastTurnCard1_MouseLeave(ByVal sender As Object, ByVal e As System.Windows.Input.MouseEventArgs) Handles LastTurnCard1.MouseLeave
        LastTurnCard1.Height /= 1.2
        LastTurnCard1.Width /= 1.2
    End Sub

    Private Sub LastTurnCard2_MouseEnter(ByVal sender As Object, ByVal e As System.Windows.Input.MouseEventArgs) Handles LastTurnCard2.MouseEnter
        LastTurnCard2.Height *= 1.2
        LastTurnCard2.Width *= 1.2
    End Sub

    Private Sub LastTurnCard2_MouseLeave(ByVal sender As Object, ByVal e As System.Windows.Input.MouseEventArgs) Handles LastTurnCard2.MouseLeave
        LastTurnCard2.Height /= 1.2
        LastTurnCard2.Width /= 1.2
    End Sub

    Private Sub LastTurnCard3_MouseEnter(ByVal sender As Object, ByVal e As System.Windows.Input.MouseEventArgs) Handles LastTurnCard3.MouseEnter
        LastTurnCard3.Height *= 1.2
        LastTurnCard3.Width *= 1.2
    End Sub

    Private Sub LastTurnCard3_MouseLeave(ByVal sender As Object, ByVal e As System.Windows.Input.MouseEventArgs) Handles LastTurnCard3.MouseLeave
        LastTurnCard3.Height /= 1.2
        LastTurnCard3.Width /= 1.2
    End Sub

    Private Sub LastTurnCard4_MouseEnter(ByVal sender As Object, ByVal e As System.Windows.Input.MouseEventArgs) Handles LastTurnCard4.MouseEnter
        LastTurnCard4.Height *= 1.2
        LastTurnCard4.Width *= 1.2
    End Sub

    Private Sub LastTurnCard4_MouseLeave(ByVal sender As Object, ByVal e As System.Windows.Input.MouseEventArgs) Handles LastTurnCard4.MouseLeave
        LastTurnCard4.Height /= 1.2
        LastTurnCard4.Width /= 1.2
    End Sub

    Private Sub GameLevelMenuItem_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles GameLevelMenuItem.Click
        Dim LvSelection As New LevelSelectionWindow
        LvSelection.ShowDialog()
    End Sub

    Private Sub StatsMenuItem_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles StatsMenuItem.Click
        Dim StatWin As New StatsWindow
        StatWin.ShowDialog()
    End Sub

    Private Sub SettingsMenuItem_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles SettingsMenuItem.Click
        Dim BackSideSetWin As New SettingsWindow
        BackSideSetWin.ShowDialog()
    End Sub

    Private Sub ViewMenuItem_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles ViewMenuItem.Click
        Dim ViewSettingsWin As New ViewSettingsWindow
        ViewSettingsWin.ShowDialog()
    End Sub

    Private Sub HumanPlayerCardControl1_MouseEnter(ByVal sender As Object, ByVal e As System.Windows.Input.MouseEventArgs) Handles HumanPlayerCardControl1.MouseEnter
        HumanPlayerCardControl1.ResetRotationCenter()
    End Sub

    Private Sub HumanPlayerCardControl1_MouseLeave(ByVal sender As Object, ByVal e As System.Windows.Input.MouseEventArgs) Handles HumanPlayerCardControl1.MouseLeave
        HumanPlayerCardControl1.ResetRotationCenter()
    End Sub

    Private Sub HumanPlayerCardControl2_MouseEnter(ByVal sender As Object, ByVal e As System.Windows.Input.MouseEventArgs) Handles HumanPlayerCardControl2.MouseEnter
        HumanPlayerCardControl2.ResetRotationCenter()
    End Sub

    Private Sub HumanPlayerCardControl2_MouseLeave(ByVal sender As Object, ByVal e As System.Windows.Input.MouseEventArgs) Handles HumanPlayerCardControl2.MouseLeave
        HumanPlayerCardControl2.ResetRotationCenter()
    End Sub

    Private Sub HumanPlayerCardControl3_MouseEnter(ByVal sender As Object, ByVal e As System.Windows.Input.MouseEventArgs) Handles HumanPlayerCardControl3.MouseEnter
        HumanPlayerCardControl3.ResetRotationCenter()
    End Sub

    Private Sub HumanPlayerCardControl3_MouseLeave(ByVal sender As Object, ByVal e As System.Windows.Input.MouseEventArgs) Handles HumanPlayerCardControl3.MouseLeave
        HumanPlayerCardControl3.ResetRotationCenter()
    End Sub

    Private Sub HumanPlayerCardControl4_MouseEnter(ByVal sender As Object, ByVal e As System.Windows.Input.MouseEventArgs) Handles HumanPlayerCardControl4.MouseEnter
        HumanPlayerCardControl4.ResetRotationCenter()
    End Sub

    Private Sub HumanPlayerCardControl4_MouseLeave(ByVal sender As Object, ByVal e As System.Windows.Input.MouseEventArgs) Handles HumanPlayerCardControl4.MouseLeave
        HumanPlayerCardControl4.ResetRotationCenter()
    End Sub

    Private Sub HumanPlayerCardControl5_MouseEnter(ByVal sender As Object, ByVal e As System.Windows.Input.MouseEventArgs) Handles HumanPlayerCardControl5.MouseEnter
        HumanPlayerCardControl5.ResetRotationCenter()
    End Sub

    Private Sub HumanPlayerCardControl5_MouseLeave(ByVal sender As Object, ByVal e As System.Windows.Input.MouseEventArgs) Handles HumanPlayerCardControl5.MouseLeave
        HumanPlayerCardControl5.ResetRotationCenter()
    End Sub

    Private Sub HumanPlayerCardControl6_MouseEnter(ByVal sender As Object, ByVal e As System.Windows.Input.MouseEventArgs) Handles HumanPlayerCardControl6.MouseEnter
        HumanPlayerCardControl6.ResetRotationCenter()
    End Sub

    Private Sub HumanPlayerCardControl6_MouseLeave(ByVal sender As Object, ByVal e As System.Windows.Input.MouseEventArgs) Handles HumanPlayerCardControl6.MouseLeave
        HumanPlayerCardControl6.ResetRotationCenter()
    End Sub

    Private Sub HumanPlayerCardControl7_MouseEnter(ByVal sender As Object, ByVal e As System.Windows.Input.MouseEventArgs) Handles HumanPlayerCardControl7.MouseEnter
        HumanPlayerCardControl7.ResetRotationCenter()
    End Sub

    Private Sub HumanPlayerCardControl7_MouseLeave(ByVal sender As Object, ByVal e As System.Windows.Input.MouseEventArgs) Handles HumanPlayerCardControl7.MouseLeave
        HumanPlayerCardControl7.ResetRotationCenter()
    End Sub

    Private Sub HumanPlayerCardControl8_MouseEnter(ByVal sender As Object, ByVal e As System.Windows.Input.MouseEventArgs) Handles HumanPlayerCardControl8.MouseEnter
        HumanPlayerCardControl8.ResetRotationCenter()
    End Sub

    Private Sub HumanPlayerCardControl8_MouseLeave(ByVal sender As Object, ByVal e As System.Windows.Input.MouseEventArgs) Handles HumanPlayerCardControl8.MouseLeave
        HumanPlayerCardControl8.ResetRotationCenter()
    End Sub

    Private Sub HumanPlayerCardControl9_MouseEnter(ByVal sender As Object, ByVal e As System.Windows.Input.MouseEventArgs) Handles HumanPlayerCardControl9.MouseEnter
        HumanPlayerCardControl9.ResetRotationCenter()
    End Sub

    Private Sub HumanPlayerCardControl9_MouseLeave(ByVal sender As Object, ByVal e As System.Windows.Input.MouseEventArgs) Handles HumanPlayerCardControl9.MouseLeave
        HumanPlayerCardControl9.ResetRotationCenter()
    End Sub

    Private Sub HumanPlayerCardControl10_MouseEnter(ByVal sender As Object, ByVal e As System.Windows.Input.MouseEventArgs) Handles HumanPlayerCardControl10.MouseEnter
        HumanPlayerCardControl10.ResetRotationCenter()
    End Sub

    Private Sub HumanPlayerCardControl10_MouseLeave(ByVal sender As Object, ByVal e As System.Windows.Input.MouseEventArgs) Handles HumanPlayerCardControl10.MouseLeave
        HumanPlayerCardControl10.ResetRotationCenter()
    End Sub

    Private Sub HumanPlayerCardControl11_MouseEnter(ByVal sender As Object, ByVal e As System.Windows.Input.MouseEventArgs) Handles HumanPlayerCardControl11.MouseEnter
        HumanPlayerCardControl11.ResetRotationCenter()
    End Sub

    Private Sub HumanPlayerCardControl11_MouseLeave(ByVal sender As Object, ByVal e As System.Windows.Input.MouseEventArgs) Handles HumanPlayerCardControl11.MouseLeave
        HumanPlayerCardControl11.ResetRotationCenter()
    End Sub

    Private Sub HumanPlayerCardControl12_MouseEnter(ByVal sender As Object, ByVal e As System.Windows.Input.MouseEventArgs) Handles HumanPlayerCardControl12.MouseEnter
        HumanPlayerCardControl12.ResetRotationCenter()
    End Sub

    Private Sub HumanPlayerCardControl12_MouseLeave(ByVal sender As Object, ByVal e As System.Windows.Input.MouseEventArgs) Handles HumanPlayerCardControl12.MouseLeave
        HumanPlayerCardControl12.ResetRotationCenter()
    End Sub

    Private Sub HumanPlayerCardControl13_MouseEnter(ByVal sender As Object, ByVal e As System.Windows.Input.MouseEventArgs) Handles HumanPlayerCardControl13.MouseEnter
        HumanPlayerCardControl13.ResetRotationCenter()
    End Sub

    Private Sub HumanPlayerCardControl13_MouseLeave(ByVal sender As Object, ByVal e As System.Windows.Input.MouseEventArgs) Handles HumanPlayerCardControl13.MouseLeave
        HumanPlayerCardControl13.ResetRotationCenter()
    End Sub

    Private Sub TurnCardControl1_NewCardValue() Handles TurnCardControl1.NewCardValue
        TurnCardControl1.ThrowEffect(0)
    End Sub

    Private Sub TurnCardControl2_NewCardValue() Handles TurnCardControl2.NewCardValue
        TurnCardControl2.ThrowEffect(-120)
    End Sub

    Private Sub TurnCardControl3_NewCardValue() Handles TurnCardControl3.NewCardValue
        TurnCardControl3.ThrowEffect(-150)
    End Sub

    Private Sub TurnCardControl4_NewCardValue() Handles TurnCardControl4.NewCardValue
        TurnCardControl4.ThrowEffect(80)
    End Sub

    Private Sub FinishGameMenuItem_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles FinishGameMenuItem.Click
        _GameThread.Abort()
        _TimerThread.Abort()
        _GameTime.Stop()
        'Dim UserChoice As MessageBoxResult = MessageBox.Show(LanguageManager.GetStrFromModel("Voulez-vous vraiment arrêter la partie en cours ?"), "MyTarneeb", MessageBoxButton.YesNo, MessageBoxImage.Question)
        'If UserChoice = MessageBoxResult.Yes Then
        _GameStarted = False
        _GameThread = Nothing
        CleanGameTable()
        'ResetGameSet()
        _GameTime.Reset()
        RaiseEvent GameFinished()
        'ElseIf UserChoice = MessageBoxResult.No Then
        '_GameTime.Start()
        '_GameThread.Resume()
        'End If
    End Sub

    Private Sub LanguageMenuItem_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles LanguageMenuItem.Click
        Dim AcutalLanguage As String = My.Settings.Language
        Dim LangSelect As New LanguageSelectionWindow
        LangSelect.ShowDialog()
        If AcutalLanguage <> My.Settings.Language Then
            Me.Dispatcher.Invoke(LoadLanguageFiles)
        End If
    End Sub

    Private LoadLanguageFiles As New _DelSub(AddressOf _LoadLanguageFiles)

    Private Sub _LoadLanguageFiles()
        GameMenuItem.Header = LanguageManager.GetStrFromModel("Jeu")
        NewGameMenuItem.Header = LanguageManager.GetStrFromModel("Nouvelle partie")
        NewGameButton.Content = NewGameMenuItem.Header
        FinishGameMenuItem.Header = LanguageManager.GetStrFromModel("Finir partie")
        PauseMenuItem.Header = LanguageManager.GetStrFromModel("Pause")
        StatsMenuItem.Header = LanguageManager.GetStrFromModel("Statistiques")
        LanguageMenuItem.Header = LanguageManager.GetStrFromModel("Langue")
        ViewMenuItem.Header = LanguageManager.GetStrFromModel("Affichage")
        GameLevelMenuItem.Header = LanguageManager.GetStrFromModel("Niveau de difficulté")
        SettingsMenuItem.Header = LanguageManager.GetStrFromModel("Motif des cartes")
        QuitMenuItem.Header = LanguageManager.GetStrFromModel("Quitter")
        GameRulesMenuItem.Header = LanguageManager.GetStrFromModel("Règles du jeu")
        MyTarneebWebSiteMenuItem.Header = LanguageManager.GetStrFromModel("Site web de MyTarneeb")
        AboutMenuItem.Header = LanguageManager.GetStrFromModel("À propos de")
        LastTurnCardsLabel.Content = LanguageManager.GetStrFromModel("Cartes jouées lors du dernier tour :")
        GameTimeLabel.Content = LanguageManager.GetStrFromModel("Temps de jeu : ")
        TurnNumberLabel.Content = LanguageManager.GetStrFromModel("Tour :")
        WonTurnsLabel.Content = LanguageManager.GetStrFromModel("Tours gagnés :")
        LostTurnsLabel.Content = LanguageManager.GetStrFromModel("Tours perdus :")
        TarneebLabel.Content &= " :"
        BetLabel.Content = LanguageManager.GetStrFromModel("Pari :")
        BetOwnerLabel.Content = LanguageManager.GetStrFromModel("Propriétaires du pari :")
        GameLevelLabel.Content = LanguageManager.GetStrFromModel("Niveau :")
    End Sub

    Private Sub PauseMenuItem_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles PauseMenuItem.Click
        If _GameStarted = True Then
            If PauseMenuItem.IsChecked = False Then 'lors du clic le IsChecked change donc si le jeu etait en pause et que ce MenuItem a été cliqué il devient False donc effectue reprise
                _GameThread.Resume()
                _GameTime.Start()
                _TimerThread.Resume()
                _SetGameStartingMenuItems()
                _GamePause = False
            Else
                _SetGameFinishedMenuItems()
                _TimerThread.Suspend()
                _GameTime.Stop()
                _GameThread.Suspend()
                _GamePause = True
            End If
        End If
    End Sub

    Private Sub GameRulesMenuItem_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles GameRulesMenuItem.Click
        Dim GR As New GameRulesForm
        GR.ShowDialog()
    End Sub

    Private Sub MyTarneebWebSiteMenuItem_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles MyTarneebWebSiteMenuItem.Click
        Process.Start("http://www.mytarneeb.com")
    End Sub

    'Private Sub PlayKnockSound()
    'Dim SP As New System.Media.SoundPlayer(My.Application.Info.DirectoryPath & "\knock.wav")
    'SP.Load()
    'SP.Play()
    'SP.Dispose()
    'End Sub

    Private Sub NewGameButton_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles NewGameButton.Click
        NewGameButton.Visibility = Windows.Visibility.Collapsed
        _GamePreparationThread = New Thread(AddressOf NewGame)
        _GamePreparationThread.SetApartmentState(ApartmentState.STA)
        _GamePreparationThread.IsBackground = True
        _GamePreparationThread.Start()
    End Sub
End Class