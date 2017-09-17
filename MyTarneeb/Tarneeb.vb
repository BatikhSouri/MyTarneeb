'Fichier :      Tarneeb.vb
'Programmeur :  Ahmad Ben MRad
'Version :      1.0.0.0
'Description :  Ce fichier contient les structures de bases du jeu MyTarneeb. Il y a aussi les classes des joueurs humains et ordinateurs.

Namespace Tarneeb

    Public Enum CardFamilies
        Diamond = 0
        Spade = 1
        Heart = 2
        Club = 3
    End Enum

    Public Enum Cards
        Two = 2
        Three = 3
        Four = 4
        Five = 5
        Six = 6
        Seven = 7
        Eight = 8
        Nine = 9
        Ten = 10
        Jack = 11
        Queen = 12
        King = 13
        Ace = 14
    End Enum

    Public Structure Card

        Private _CardFamily As CardFamilies
        Private _CardNumber As Cards

        Public Sub New(ByVal CardNumber As Cards, ByVal CardFamily As CardFamilies)
            _CardNumber = CardNumber
            _CardFamily = CardFamily
        End Sub

        Public Shared Operator =(ByVal Card1 As Card, ByVal Card2 As Card) As Boolean
            If Card1.CardFamily = Card2.CardFamily And Card1.CardNumber = Card2.CardNumber Then
                Return True
            Else
                Return False
            End If
        End Operator

        Public Shared Operator <>(ByVal Card1 As Card, ByVal Card2 As Card) As Boolean
            If Not Card1.CardFamily = Card2.CardFamily Or Not Card1.CardNumber = Card2.CardNumber Then
                Return True
            Else
                Return False
            End If
        End Operator

        Public Property CardFamily As CardFamilies
            Get
                Return _CardFamily
            End Get
            Set(ByVal value As CardFamilies)
                _CardFamily = value
            End Set
        End Property

        Public Property CardNumber As Cards
            Get
                Return _CardNumber
            End Get
            Set(ByVal value As Cards)
                _CardNumber = value
            End Set
        End Property

        Public ReadOnly Property CardName As String
            Get
                Dim FamilyPart As String
                If _CardFamily = Tarneeb.CardFamilies.Diamond Then
                    FamilyPart = " " & LanguageManager.GetStrFromModel("de") & " " & LanguageManager.GetStrFromModel("carreaux")
                    Select Case _CardNumber
                        Case Tarneeb.Cards.Two
                            Return LanguageManager.GetStrFromModel("Deux") & FamilyPart
                        Case Tarneeb.Cards.Three
                            Return LanguageManager.GetStrFromModel("Trois") & FamilyPart
                        Case Tarneeb.Cards.Four
                            Return LanguageManager.GetStrFromModel("Quatre") & FamilyPart
                        Case Tarneeb.Cards.Five
                            Return LanguageManager.GetStrFromModel("Cinq") & FamilyPart
                        Case Tarneeb.Cards.Six
                            Return LanguageManager.GetStrFromModel("Six") & FamilyPart
                        Case Tarneeb.Cards.Seven
                            Return LanguageManager.GetStrFromModel("Sept") & FamilyPart
                        Case Tarneeb.Cards.Eight
                            Return LanguageManager.GetStrFromModel("Huit") & FamilyPart
                        Case Tarneeb.Cards.Nine
                            Return LanguageManager.GetStrFromModel("Neuf") & FamilyPart
                        Case Tarneeb.Cards.Ten
                            Return LanguageManager.GetStrFromModel("Dix") & FamilyPart
                        Case Tarneeb.Cards.Jack
                            Return LanguageManager.GetStrFromModel("Valet") & FamilyPart
                        Case Tarneeb.Cards.Queen
                            Return LanguageManager.GetStrFromModel("Reine") & FamilyPart
                        Case Tarneeb.Cards.King
                            Return LanguageManager.GetStrFromModel("Roi") & FamilyPart
                        Case Tarneeb.Cards.Ace
                            Return LanguageManager.GetStrFromModel("As") & FamilyPart
                    End Select
                ElseIf _CardFamily = Tarneeb.CardFamilies.Spade Then
                    FamilyPart = " " & LanguageManager.GetStrFromModel("de") & " " & LanguageManager.GetStrFromModel("piques")
                    Select Case _CardNumber
                        Case Tarneeb.Cards.Two
                            Return LanguageManager.GetStrFromModel("Deux") & FamilyPart
                        Case Tarneeb.Cards.Three
                            Return LanguageManager.GetStrFromModel("Trois") & FamilyPart
                        Case Tarneeb.Cards.Four
                            Return LanguageManager.GetStrFromModel("Quatre") & FamilyPart
                        Case Tarneeb.Cards.Five
                            Return LanguageManager.GetStrFromModel("Cinq") & FamilyPart
                        Case Tarneeb.Cards.Six
                            Return LanguageManager.GetStrFromModel("Six") & FamilyPart
                        Case Tarneeb.Cards.Seven
                            Return LanguageManager.GetStrFromModel("Sept") & FamilyPart
                        Case Tarneeb.Cards.Eight
                            Return LanguageManager.GetStrFromModel("Huit") & FamilyPart
                        Case Tarneeb.Cards.Nine
                            Return LanguageManager.GetStrFromModel("Neuf") & FamilyPart
                        Case Tarneeb.Cards.Ten
                            Return LanguageManager.GetStrFromModel("Dix") & FamilyPart
                        Case Tarneeb.Cards.Jack
                            Return LanguageManager.GetStrFromModel("Valet") & FamilyPart
                        Case Tarneeb.Cards.Queen
                            Return LanguageManager.GetStrFromModel("Reine") & FamilyPart
                        Case Tarneeb.Cards.King
                            Return LanguageManager.GetStrFromModel("Roi") & FamilyPart
                        Case Tarneeb.Cards.Ace
                            Return LanguageManager.GetStrFromModel("As") & FamilyPart
                    End Select
                ElseIf _CardFamily = Tarneeb.CardFamilies.Heart Then
                    FamilyPart = " " & LanguageManager.GetStrFromModel("de") & " " & LanguageManager.GetStrFromModel("coeurs")
                    Select Case _CardNumber
                        Case Tarneeb.Cards.Two
                            Return LanguageManager.GetStrFromModel("Deux") & FamilyPart
                        Case Tarneeb.Cards.Three
                            Return LanguageManager.GetStrFromModel("Trois") & FamilyPart
                        Case Tarneeb.Cards.Four
                            Return LanguageManager.GetStrFromModel("Quatre") & FamilyPart
                        Case Tarneeb.Cards.Five
                            Return LanguageManager.GetStrFromModel("Cinq") & FamilyPart
                        Case Tarneeb.Cards.Six
                            Return LanguageManager.GetStrFromModel("Six") & FamilyPart
                        Case Tarneeb.Cards.Seven
                            Return LanguageManager.GetStrFromModel("Sept") & FamilyPart
                        Case Tarneeb.Cards.Eight
                            Return LanguageManager.GetStrFromModel("Huit") & FamilyPart
                        Case Tarneeb.Cards.Nine
                            Return LanguageManager.GetStrFromModel("Neuf") & FamilyPart
                        Case Tarneeb.Cards.Ten
                            Return LanguageManager.GetStrFromModel("Dix") & FamilyPart
                        Case Tarneeb.Cards.Jack
                            Return LanguageManager.GetStrFromModel("Valet") & FamilyPart
                        Case Tarneeb.Cards.Queen
                            Return LanguageManager.GetStrFromModel("Reine") & FamilyPart
                        Case Tarneeb.Cards.King
                            Return LanguageManager.GetStrFromModel("Roi") & FamilyPart
                        Case Tarneeb.Cards.Ace
                            Return LanguageManager.GetStrFromModel("As") & FamilyPart
                    End Select
                ElseIf _CardFamily = Tarneeb.CardFamilies.Club Then
                    FamilyPart = " " & LanguageManager.GetStrFromModel("de") & " " & LanguageManager.GetStrFromModel("trèfles")
                    Select Case _CardNumber
                        Case Tarneeb.Cards.Two
                            Return LanguageManager.GetStrFromModel("Deux") & FamilyPart
                        Case Tarneeb.Cards.Three
                            Return LanguageManager.GetStrFromModel("Trois") & FamilyPart
                        Case Tarneeb.Cards.Four
                            Return LanguageManager.GetStrFromModel("Quatre") & FamilyPart
                        Case Tarneeb.Cards.Five
                            Return LanguageManager.GetStrFromModel("Cinq") & FamilyPart
                        Case Tarneeb.Cards.Six
                            Return LanguageManager.GetStrFromModel("Six") & FamilyPart
                        Case Tarneeb.Cards.Seven
                            Return LanguageManager.GetStrFromModel("Sept") & FamilyPart
                        Case Tarneeb.Cards.Eight
                            Return LanguageManager.GetStrFromModel("Huit") & FamilyPart
                        Case Tarneeb.Cards.Nine
                            Return LanguageManager.GetStrFromModel("Neuf") & FamilyPart
                        Case Tarneeb.Cards.Ten
                            Return LanguageManager.GetStrFromModel("Dix") & FamilyPart
                        Case Tarneeb.Cards.Jack
                            Return LanguageManager.GetStrFromModel("Valet") & FamilyPart
                        Case Tarneeb.Cards.Queen
                            Return LanguageManager.GetStrFromModel("Reine") & FamilyPart
                        Case Tarneeb.Cards.King
                            Return LanguageManager.GetStrFromModel("Roi") & FamilyPart
                        Case Tarneeb.Cards.Ace
                            Return LanguageManager.GetStrFromModel("As") & FamilyPart
                    End Select
                End If
            End Get
        End Property

        Public Shared ReadOnly Property GameSet() As Card()
            Get
                Dim GSet(51) As Card
                Dim k As Byte = 0
                For j = 0 To 3
                    For i = 2 To 14
                        GSet(k) = New Card(i, j)
                        k += 1
                    Next
                Next
                Return GSet
            End Get
        End Property
    End Structure

    Public Structure Bet

        Private _Bet As Byte
        Private _CardFamily As CardFamilies

        Public Sub New(ByVal Bet As Byte, ByVal CardFamily As CardFamilies)
            _Bet = Bet
            _CardFamily = CardFamily
        End Sub

        Public Shared Operator =(ByVal Bet1 As Bet, ByVal Bet2 As Bet) As Boolean
            If Bet1.Bet = Bet2.Bet Then
                Return True
            Else
                Return False
            End If
        End Operator

        Public Shared Operator <>(ByVal Bet1 As Bet, ByVal Bet2 As Bet) As Boolean
            If Bet1.Bet <> Bet2.Bet Then
                Return True
            Else
                Return False
            End If
        End Operator

        Public Shared Operator >(ByVal Bet1 As Bet, ByVal Bet2 As Bet) As Boolean
            If Bet1.Bet > Bet2.Bet Then
                Return True
            Else
                Return False
            End If
        End Operator

        Public Shared Operator <(ByVal Bet1 As Bet, ByVal Bet2 As Bet) As Boolean
            If Bet1.Bet < Bet2.Bet Then
                Return True
            Else
                Return False
            End If
        End Operator

        Public ReadOnly Property Bet As Byte
            Get
                Return _Bet
            End Get
        End Property

        Public ReadOnly Property CardFamily As CardFamilies
            Get
                Return _CardFamily
            End Get
        End Property

        Public ReadOnly Property FamilyName As String
            Get
                If _CardFamily = CardFamilies.Diamond Then
                    Return LanguageManager.GetStrFromModel("carreaux")
                ElseIf _CardFamily = CardFamilies.Club Then
                    Return LanguageManager.GetStrFromModel("trèfles")
                ElseIf _CardFamily = CardFamilies.Heart Then
                    Return LanguageManager.GetStrFromModel("coeurs")
                ElseIf _CardFamily = CardFamilies.Spade Then
                    Return LanguageManager.GetStrFromModel("piques")
                End If
            End Get
        End Property
    End Structure

    Public Interface IPlayer

        'Subs & Functions
        Function Bet(ByVal PlayersBets As Bet()) As Bet
        Function SeekForTarneeb() As CardFamilies
        Function Play(ByVal RoundCards As Card()) As Card
        'Function WantToKnock(ByVal PreviousRoundFamily) As Boolean
        Sub SetTarneebFamily()
        Sub SetFamiliesOrder(ByVal OrderFamilies As String)
        Sub EndOfRound(ByVal RoundCards As Card(), ByVal LastTurnWinnerInd As Byte)
        'Properties
        ReadOnly Property PlayerBet As Bet
        ReadOnly Property TurnCardID As Byte
        ReadOnly Property TurnCard As Card
        ReadOnly Property FamiliesOrder As String
        Property PlayerCards As Card()

    End Interface

    Public Class HumanPlayer
        Implements IPlayer

        Private _PlayerCards As Card()
        Private _Tarneeb As CardFamilies
        Private _Bet As Bet
        Private _TurnCard As Card
        Private _TurnCardID As Byte
        Private _FamiliesOrder As String = "0123"

        Private Function Bet(ByVal PlayersBets As Bet()) As Bet Implements IPlayer.Bet
            Dim HighestBetValue = GetHighestBetValue(PlayersBets)
            Dim BetWin As BetForm
            If HighestBetValue = 0 Then
                BetWin = New BetForm(7)
            Else
                BetWin = New BetForm(HighestBetValue + 1)
            End If
            'BetWin.MinimumUserBet =
            Try
                BetWin.ShowDialog()
            Catch ex As AccessViolationException

            End Try
            Try
                BetWin.Dispose()
            Catch ex As Exception

            End Try
            If BetWin.Pass = True Then
                _Bet = New Bet(0, CardFamilies.Diamond)
            Else
                'MessageBox.Show(BetWin.ComboBox1.Items(BetWin.ComboBox1.SelectedIndex).ToString())
                _Bet = New Bet(BetWin.NumericUpDown1.Value, CardFamilies.Diamond)
                Return _Bet
            End If
        End Function

        Public Function SeekForTarneeb() As CardFamilies Implements IPlayer.SeekForTarneeb
            Dim STW As New TarneebSelectionWindow
            STW.ShowDialog()
            _Bet = New Bet(_Bet.Bet, STW.TarneebSelection)
            Return _Bet.CardFamily
        End Function

        Public Sub EndOfRound(ByVal RoundCards() As Card, ByVal LastTurnWinnerInd As Byte) Implements IPlayer.EndOfRound

        End Sub

        Public Function Play(ByVal RoundCards() As Card) As Card Implements IPlayer.Play
            My.Settings.HumanPlayFamily = 255
            My.Settings.HumanPlayValue = 0
            HumanPlayerCardControl.HumanTurn = True
            Dim ExitI As Boolean = False
            Do While ExitI = False
                If My.Settings.HumanPlayFamily <> 255 And My.Settings.HumanPlayValue <> 0 Then
                    ExitI = True
                Else
                    System.Threading.Thread.Sleep(250)
                End If
            Loop
            Dim HumanPlay As Card = New Card(My.Settings.HumanPlayValue, My.Settings.HumanPlayFamily)
            _TurnCard = HumanPlay
            _TurnCardID = GetCardID(_PlayerCards, HumanPlay)
            _PlayerCards = RemoveCardFromPlayersHands(_PlayerCards, HumanPlay)
            Return HumanPlay
        End Function

        Public ReadOnly Property PlayerBet As Bet Implements IPlayer.PlayerBet
            Get
                Return _Bet
            End Get
        End Property

        Public ReadOnly Property TurnCardID As Byte Implements IPlayer.TurnCardID
            Get
                Return _TurnCardID
            End Get
        End Property

        Public ReadOnly Property TurnCard As Card Implements IPlayer.TurnCard
            Get
                Return _TurnCard
            End Get
        End Property

        Public Property PlayerCards As Card() Implements IPlayer.PlayerCards
            Set(ByVal value As Card())
                _PlayerCards = value
            End Set
            Get
                Return _PlayerCards
            End Get
        End Property

        Public ReadOnly Property FamiliesOrder As String Implements IPlayer.FamiliesOrder
            Get
                Return _FamiliesOrder
            End Get
        End Property

        Public Sub SetFamilesOrder(ByVal FamiliesOrder As String) Implements IPlayer.SetFamiliesOrder
            _FamiliesOrder = FamiliesOrder
        End Sub

        Public Sub SetTarneebFamily() Implements IPlayer.SetTarneebFamily
            _Tarneeb = My.Settings.Tarneeb
        End Sub

        Private Function GetHighestBetValue(ByVal Bets As Bet()) As Byte
            Dim max As Byte = 0
            For Each b As Bet In Bets
                If b.Bet > max Then
                    max = b.Bet
                End If
            Next
            Return max
        End Function
    End Class

    Public Class AIPlayer
        Implements IPlayer

        Private Const NumberOfLevel As Byte = 5
        'Mémoristion des cartes sorties
        Private OutHighDiamondCardsOnTable(12) As Boolean 'liste des cartes hautes de carreaux sortis : Valet, Dame, Roi et As
        Private OutHighSpadeCardsOnTable(12) As Boolean
        Private OutHighHeartCardsOnTable(12) As Boolean
        Private OutHighClubCardsOnTable(12) As Boolean
        'Déductions sur les familles restantes de chez les autres joueurs
        Private DeductionsFamiliesFinishedNextPlayer(3) As Boolean
        Private DeductionsFamiliesFinishedPreviousPlayer(3) As Boolean
        Private DeductionsFamiliesFinishedTeammate(3) As Boolean
        'Suppositions sur les familles
        Private SuppositionsFamiliesFinishedNextPlayer(3) As Boolean
        Private SuppositionsFamiliesFinishedPreviousPlayer(3) As Boolean
        Private SuppositionsFamiliesFinishedTeammate(3) As Boolean
        'Autres variables
        Private _MemorizingProbability As Decimal
        Private _DeductionsProbability As Decimal
        Private _FamiliesOrder As String = "0123"
        Private _TurnCardID As Byte
        Private _TurnCard As Card
        Private _Tarneeb As CardFamilies
        Private _Bet As Bet
        Private _PlayerCards As Card()

        Public Sub New(ByVal GameLevel As Byte)
            If GameLevel > (NumberOfLevel - 1) Then
                Throw New Exception("Invalid Game Level")
            Else
                _MemorizingProbability = GameLevel / (NumberOfLevel - 2)
                _DeductionsProbability = GameLevel / (NumberOfLevel - 1)
                For Each dmemboo As Boolean In OutHighDiamondCardsOnTable
                    dmemboo = False
                Next
                For Each smemboo As Boolean In OutHighSpadeCardsOnTable
                    smemboo = False
                Next
                For Each hmemboo As Boolean In OutHighHeartCardsOnTable
                    hmemboo = False
                Next
                For Each cmemboo As Boolean In OutHighClubCardsOnTable
                    cmemboo = False
                Next
                For Each boo As Boolean In DeductionsFamiliesFinishedNextPlayer
                    boo = False
                Next
                For Each boo As Boolean In DeductionsFamiliesFinishedPreviousPlayer
                    boo = False
                Next
                For Each boo As Boolean In DeductionsFamiliesFinishedTeammate
                    boo = False
                Next
                For Each boo As Boolean In SuppositionsFamiliesFinishedNextPlayer
                    boo = False
                Next
                For Each boo As Boolean In SuppositionsFamiliesFinishedPreviousPlayer
                    boo = False
                Next
                For Each boo As Boolean In SuppositionsFamiliesFinishedTeammate
                    boo = False
                Next
            End If
        End Sub

        Public Sub SetTarneebFamily() Implements IPlayer.SetTarneebFamily
            _Tarneeb = My.Settings.Tarneeb
        End Sub

        Public Property PlayerCards As Card() Implements IPlayer.PlayerCards
            Set(ByVal value As Card())
                _PlayerCards = value
            End Set
            Get
                Return _PlayerCards
            End Get
        End Property

        Public ReadOnly Property PlayerBet As Bet Implements IPlayer.PlayerBet
            Get
                Return _Bet
            End Get
        End Property

        Public ReadOnly Property TurnCardID As Byte Implements IPlayer.TurnCardID
            Get
                Return _TurnCardID
            End Get
        End Property

        Public ReadOnly Property TurnCard As Card Implements IPlayer.TurnCard
            Get
                Return _TurnCard
            End Get
        End Property

        Public ReadOnly Property FamiliesOrder As String Implements IPlayer.FamiliesOrder
            Get
                Return _FamiliesOrder
            End Get
        End Property

        Public Function SeekForTarneeb() As CardFamilies Implements IPlayer.SeekForTarneeb
            Return _Bet.CardFamily
        End Function

        Public Sub SetFamiliesOrder(ByVal FamiliesOrder As String) Implements IPlayer.SetFamiliesOrder
            _FamiliesOrder = FamiliesOrder
        End Sub

        Public Function Play(ByVal TurnCards As Card()) As Card Implements IPlayer.Play
            Dim CardToPlay As Card 'la carte à jouer (à renvoyer à la fin de la fonction Play())
            Dim NumTurnCards As Byte = UBound(TurnCards) 'nombre de cartes sur la table
            Dim NumPlayerCards As Byte = _PlayerCards.Length 'nombre de cartes dans la main du joueur
            Dim TarneebFamily As CardFamilies = My.Settings.Tarneeb 'famille du tarneeb de cette partie
            If NumTurnCards <= 3 Then 'si le nombres de cartes présentes sur la table est plus petit ou égal à 3
                If NumTurnCards = 0 Then 'si il y a zéro cartes sur la table
                    'Essai de jouer la carte la plus haute où il n'y a pas de cartes supérieures de sa famille encore dans les mains des joueurs
                    Try
                        Dim SurePlays As Card() = GetPossibleSureCards(_PlayerCards)
                        Dim ReallySurePlays As Card() = GetReallySureCards(SurePlays)
                        If Not IsNothing(ReallySurePlays) Then
                            Dim HighestReallySureCard As Card = GetHighestCard(ReallySurePlays)
                            Dim HighestReallySureCardID As Byte = GetCardID(_PlayerCards, HighestReallySureCard)
                            CardToPlay = HighestReallySureCard
                        ElseIf Not IsNothing(SurePlays) Then
                            Dim HighestSureCard As Card = GetHighestCard(SurePlays)
                            Dim HighestSureCardID As Byte = GetCardID(_PlayerCards, HighestSureCard)
                            CardToPlay = HighestSureCard
                        Else
                            Dim CardsWithoutTarneebList As Card() = GetCardsWithoutTarneeb(_PlayerCards)
                            If Not IsNothing(CardsWithoutTarneebList) Then
                                CardToPlay = GetHighestCard(CardsWithoutTarneebList)
                            Else 'Autrement une carte d'une famille en infériorité numérique
                                Dim NumericalInfFamily As CardFamilies = GetWhichFamilyHasTheLess(_PlayerCards)
                                Dim InfFamilyCardList As Card() = GetFamilyList(_PlayerCards, NumericalInfFamily)
                                CardToPlay = GetLowestCard(InfFamilyCardList)
                            End If
                        End If
                    Catch ex As Exception
                        Dim CardsWithoutTarneebList As Card() = GetCardsWithoutTarneeb(_PlayerCards)
                        If Not IsNothing(CardsWithoutTarneebList) Then
                            CardToPlay = GetHighestCard(CardsWithoutTarneebList)
                        Else 'Autrement une carte d'une famille en infériorité numérique
                            Dim NumericalInfFamily As CardFamilies = GetWhichFamilyHasTheLess(_PlayerCards)
                            Dim InfFamilyCardList As Card() = GetFamilyList(_PlayerCards, NumericalInfFamily)
                            CardToPlay = GetLowestCard(InfFamilyCardList)
                        End If
                    End Try
                ElseIf NumTurnCards = 1 Then 'si il y a une carte sur la table
                    Dim TurnFamily As CardFamilies = TurnCards(0).CardFamily
                    Dim FamilyList As Card() = GetFamilyList(_PlayerCards, TurnFamily)
                    If Not IsNothing(FamilyList) Then
                        Dim HighFamilyList As Card() = GetHighSerieList(FamilyList)
                        If Not IsNothing(HighFamilyList) Then
                            'Try
                            Dim SurePlays() As Card = GetPossibleSureCards(HighFamilyList, TurnFamily)
                            If Not IsNothing(SurePlays) Then
                                Dim HigherCard As Card = GetHigherCard(SurePlays, TurnCards(0))
                                If Not HigherCard.CardNumber = 0 Then
                                    CardToPlay = HigherCard
                                Else
                                    CardToPlay = GetLowestCard(FamilyList)
                                End If
                            Else
                                Dim HigherCard As Card = GetHigherCard(FamilyList, TurnCards(0))
                                If HigherCard <> Nothing Then
                                    CardToPlay = HigherCard
                                Else
                                    CardToPlay = GetLowestCard(FamilyList)
                                End If
                            End If
                            'Catch ex As Exception
                            'Dim HigherCard As Card = GetHigherCard(FamilyList, TurnCards(0))
                            'If HigherCard.CardNumber <> Nothing And HigherCard.CardFamily <> Nothing Then
                            'CardToPlay = HigherCard
                            'Else
                            '   CardToPlay = GetLowestCard(FamilyList)
                            'End If
                            'End Try
                        Else
                            CardToPlay = GetLowestCard(FamilyList)
                        End If
                    Else
                        Dim TarneebList As Card() = GetFamilyList(_PlayerCards, TarneebFamily)
                        If Not IsNothing(TarneebList) Then
                            CardToPlay = GetLowestCard(TarneebList)
                        Else
                            Dim FamilyWhichHasTheLessList As Card() = GetFamilyList(_PlayerCards, GetWhichFamilyHasTheLess(_PlayerCards))
                            Dim SecondFamilyWhichHasTheLessList As Card()
                            Dim LowestCard As Card
                            If Not IsNothing(FamilyWhichHasTheLessList) Then
                                LowestCard = GetLowestCard(FamilyWhichHasTheLessList)
                                If LowestCard.CardNumber < 11 Then
                                    CardToPlay = LowestCard
                                Else
                                    If CountFamilies(_PlayerCards) >= 2 Then
                                        SecondFamilyWhichHasTheLessList = GetFamilyList(_PlayerCards, GetSecondFamilyWhichHasTheLess(_PlayerCards))
                                        If Not IsNothing(SecondFamilyWhichHasTheLessList) Then
                                            LowestCard = GetLowestCard(SecondFamilyWhichHasTheLessList)
                                            If LowestCard.CardNumber < 11 Then
                                                CardToPlay = LowestCard
                                            Else
                                                CardToPlay = GetLowestCard(_PlayerCards)
                                            End If
                                        End If
                                    Else
                                        CardToPlay = LowestCard
                                    End If
                                End If
                            End If
                            'Dim LowestCard As Card = GetLowestCard(GetFamilyList(_PlayerCards, GetWhichFamilyHasTheLess(_PlayerCards)))
                            'If LowestCard.CardNumber < 11 Then
                            'CardToPlay = LowestCard
                            'Else
                            'CardToPlay = GetLowestCard(GetFamilyList(_PlayerCards, GetSecondFamilyWhichHasTheLess(_PlayerCards)))
                            'End If
                        End If
                    End If
                ElseIf NumTurnCards = 2 Then 'si il y a deux cartes sur la table
                    Dim IsHeWinningB As Boolean = IsHeWinning(TurnCards)
                    Dim TurnFamily As CardFamilies = TurnCards(0).CardFamily
                    Dim FamilyList As Card() = GetFamilyList(_PlayerCards, TurnFamily)
                    If IsHeWinningB = True Then
                        If Not IsNothing(FamilyList) Then
                            If IsHigh(TurnCards(0)) Then
                                CardToPlay = GetLowestCard(FamilyList)
                            Else
                                Dim HigherCard As Card = GetHigherCard(FamilyList, TurnCards(1))
                                If HigherCard.CardNumber > 0 Then
                                    CardToPlay = HigherCard
                                Else
                                    CardToPlay = GetLowestCard(FamilyList)
                                End If
                            End If
                        Else
                            If IsHigh(TurnCards(0)) Then
                                Dim CardsWOTarneeb As Card() = GetCardsWithoutTarneeb(_PlayerCards)
                                If Not IsNothing(CardsWOTarneeb) Then
                                    CardToPlay = GetLowestCard(CardsWOTarneeb)
                                Else
                                    CardToPlay = GetLowestCard(_PlayerCards)
                                End If
                            Else
                                Dim TarneebList As Card() = GetFamilyList(_PlayerCards, TarneebFamily)
                                If Not IsNothing(TarneebList) Then
                                    Dim TurnTarneeb
                                    CardToPlay = GetLowestCard(TarneebList)
                                Else
                                    CardToPlay = GetLowestCard(_PlayerCards)
                                End If
                            End If
                        End If
                    Else
                        If Not IsNothing(FamilyList) Then
                            If TurnCards(1).CardFamily = TurnFamily Then
                                'Dim HigherCard As Card = GetHigherCard(FamilyList, TurnCards(1))
                                'If HigherCard.CardNumber > 0 Then
                                'CardToPlay = HigherCard
                                'Else
                                'CardToPlay = GetLowestCard(FamilyList)
                                'End If
                                CardToPlay = GetHighestCard(FamilyList)
                            Else
                                CardToPlay = GetLowestCard(FamilyList)
                            End If
                        Else
                            Dim TarneebList As Card() = GetFamilyList(_PlayerCards, TarneebFamily)
                            If Not IsNothing(TarneebList) Then
                                Dim TurnTarneebList As Card() = GetFamilyList(TurnCards, TarneebFamily)
                                If Not IsNothing(TurnTarneebList) Then
                                    Dim HigherTarneebCard As Card = GetHigherCard(_PlayerCards, GetHighestCard(TurnTarneebList))
                                    If HigherTarneebCard.CardNumber <> 0 Then
                                        CardToPlay = HigherTarneebCard
                                    Else
                                        CardToPlay = GetLowestCard(_PlayerCards)
                                    End If
                                Else
                                    CardToPlay = GetLowestCard(TarneebList)
                                End If
                            Else
                                CardToPlay = GetLowestCard(_PlayerCards)
                            End If
                        End If
                    End If
                ElseIf NumTurnCards = 3 Then 'si il y a trois cartes sur la table
                    Dim IsHeWinningB As Boolean = IsHeWinning(TurnCards)
                    Dim TurnFamily As CardFamilies = TurnCards(0).CardFamily
                    Dim FamilyList As Card() = GetFamilyList(_PlayerCards, TurnFamily)
                    If IsHeWinningB = True Then
                        If Not IsNothing(FamilyList) Then
                            CardToPlay = GetLowestCard(FamilyList)
                        Else
                            Dim CardsWOTarneeb As Card() = GetCardsWithoutTarneeb(_PlayerCards)
                            If Not IsNothing(CardsWOTarneeb) Then
                                CardToPlay = GetLowestCard(CardsWOTarneeb)
                            Else
                                CardToPlay = GetLowestCard(_PlayerCards)
                            End If
                        End If
                    Else
                        If Not IsNothing(FamilyList) Then
                            If Not TurnCards(0).CardFamily = TarneebFamily And TurnCards(2).CardFamily = TarneebFamily Then
                                CardToPlay = GetLowestCard(FamilyList)
                            Else
                                Dim HighCard As Card
                                If TurnCards(0).CardFamily = TurnCards(2).CardFamily Then
                                    If TurnCards(0).CardNumber > TurnCards(2).CardNumber Then
                                        HighCard = TurnCards(0)
                                    Else
                                        HighCard = TurnCards(2)
                                    End If
                                Else
                                    HighCard = TurnCards(0)
                                End If
                                Dim HigherCard As Card = GetHigherCard(FamilyList, HighCard)
                                If Not HigherCard.CardNumber = 0 Then
                                    CardToPlay = HigherCard
                                Else
                                    CardToPlay = GetLowestCard(FamilyList)
                                End If
                            End If
                        Else
                            Dim TarneebList As Card() = GetFamilyList(_PlayerCards, TarneebFamily)
                            If Not IsNothing(TarneebList) Then
                                If TurnCards(2).CardFamily = TarneebFamily Then
                                    Dim HigherTarneebCard As Card = GetHigherCard(TarneebList, TurnCards(2))
                                    If HigherTarneebCard.CardNumber <> 0 Then
                                        CardToPlay = HigherTarneebCard
                                    Else
                                        CardToPlay = GetLowestCard(_PlayerCards)
                                    End If
                                Else
                                    CardToPlay = GetLowestCard(TarneebList)
                                End If
                            Else
                                CardToPlay = GetLowestCard(_PlayerCards)
                            End If
                        End If
                    End If
                End If
                _TurnCard = CardToPlay
                If CardToPlay.CardNumber = 0 Then
                    Throw New NullReferenceException
                End If
                _TurnCardID = GetCardID(_PlayerCards, CardToPlay)
                _PlayerCards = RemoveCardFromPlayersHands(_PlayerCards, CardToPlay)
                Return CardToPlay
            Else
                'autrement, si le nombre de cartes est plus grand que trois, dire qu'il y a un problème
                Throw New Exception("Invalid Number Of Cards On Table")
                Exit Function
            End If
        End Function

        Private Function IsHeWinning(ByVal CardsOnTable As Card()) As Boolean
            Dim NumPlayer As Byte = UBound(CardsOnTable)
            If NumPlayer < 2 Or NumPlayer > 3 Then
                Throw New Exception("Invalid Player Number")
            Else
                Dim TurnFamily As CardFamilies = CardsOnTable(0).CardFamily
                If NumPlayer = 2 Then
                    If CardsOnTable(0).CardFamily = CardsOnTable(1).CardFamily Then
                        If CardsOnTable(0).CardNumber > CardsOnTable(1).CardNumber Then
                            Return True
                        Else
                            Return False
                        End If
                    Else
                        If CardsOnTable(1).CardFamily = _Tarneeb Then
                            Return False
                        Else
                            Return True
                        End If
                    End If
                Else
                    If CardsOnTable(0).CardFamily = CardsOnTable(1).CardFamily Then
                        If CardsOnTable(0).CardFamily = CardsOnTable(2).CardFamily Then
                            If CardsOnTable(1).CardNumber > CardsOnTable(0).CardNumber And CardsOnTable(1).CardNumber > CardsOnTable(2).CardNumber Then
                                Return True
                            Else
                                Return False
                            End If
                        Else
                            If CardsOnTable(2).CardFamily = _Tarneeb Then
                                Return False
                            Else
                                If CardsOnTable(1).CardNumber > CardsOnTable(0).CardNumber Then
                                    Return True
                                Else
                                    Return False
                                End If
                            End If

                        End If
                    Else
                        If CardsOnTable(1).CardFamily = _Tarneeb Then
                            If CardsOnTable(2).CardFamily = _Tarneeb Then
                                If CardsOnTable(1).CardNumber > CardsOnTable(2).CardNumber Then
                                    Return True
                                Else
                                    Return False
                                End If
                            Else
                                Return True
                            End If
                        Else
                            Return False
                        End If
                    End If
                End If
            End If
        End Function

        Private Function IsHigh(ByVal Card As Card) As Boolean
            If Card.CardNumber >= 8 Then
                Return True
            Else
                Return False
            End If
        End Function

        Private Function IsFromHighSerie(ByVal Card As Card) As Boolean
            If Card.CardNumber >= 11 Then
                Return True
            Else
                Return False
            End If
        End Function

        Private Function CountFamilies(ByVal PlayerCards As Card()) As Byte
            If Not IsNothing(PlayerCards) Then
                Dim FamilyPresences(3) As Boolean
                Dim NumFamilies As Byte = 0
                For Each c As Card In PlayerCards
                    FamilyPresences(c.CardFamily) = True
                Next
                For i = 0 To 3
                    If FamilyPresences(i) = True Then
                        NumFamilies += 1
                    End If
                Next
                Return NumFamilies
            Else
                Throw New ArgumentNullException
            End If
        End Function

        Public Function Bet(ByVal PlayersBets As Bet()) As Bet Implements IPlayer.Bet
            Dim HighestBetValue As Byte = GetHighestBetValue(PlayersBets)
            If HighestBetValue >= 7 And UBound(PlayersBets) >= 2 Then
                If PlayersBets(UBound(PlayersBets) - 2).Bet = HighestBetValue Then
                    GoTo Pass
                End If
            End If
            Dim FamNum(4) As Byte
            Dim SelectedFamily As CardFamilies
            Dim NumBet As Byte
            Dim FirstBet As New Bet(0, 0), SecondBet As New Bet(0, 0), SelectedBet As New Bet(0, 0)
            FamNum(0) = 0
            FamNum(1) = 0
            FamNum(2) = 0
            FamNum(3) = 0
            For Each c As Card In _PlayerCards
                Select Case c.CardFamily
                    Case CardFamilies.Diamond
                        FamNum(0) += 1
                    Case CardFamilies.Spade
                        FamNum(1) += 1
                    Case CardFamilies.Heart
                        FamNum(2) += 1
                    Case CardFamilies.Club
                        FamNum(3) += 1
                End Select
            Next
            If FamNum(0) = FamNum.Max Then
                SelectedFamily = CardFamilies.Diamond
            ElseIf FamNum(1) = FamNum.Max Then
                SelectedFamily = CardFamilies.Spade
            ElseIf FamNum(2) = FamNum.Max Then
                SelectedFamily = CardFamilies.Heart
            ElseIf FamNum(3) = FamNum.Max Then
                SelectedFamily = CardFamilies.Club
            End If
            Dim HighCardsCount As Byte = CountHighCards(GetFamilyList(_PlayerCards, SelectedFamily))
            NumBet = FamNum.Max
            For Each c As Card In _PlayerCards
                If c.CardFamily <> SelectedFamily And c.CardNumber >= 11 Then
                    NumBet += 1
                End If
            Next
            If HighCardsCount > 2 Then
                NumBet += Math.Round(3.5 / NumBet)
            Else
                NumBet = Math.Round(NumBet * 0.9)
            End If
            FirstBet = New Bet(NumBet, SelectedFamily)
            'Deuxième tentative de pari
            Dim SecondTryList As Card() = GetPlayersCardsListWithoutSpecificFamily(_PlayerCards, SelectedFamily)
            NumBet = 0
            FamNum(0) = 0
            FamNum(1) = 0
            FamNum(2) = 0
            FamNum(3) = 0
            For Each c As Card In SecondTryList
                Select Case c.CardFamily
                    Case CardFamilies.Diamond
                        FamNum(0) += 1
                    Case CardFamilies.Spade
                        FamNum(1) += 1
                    Case CardFamilies.Heart
                        FamNum(2) += 1
                    Case CardFamilies.Club
                        FamNum(3) += 1
                End Select
            Next
            If FamNum(0) = FamNum.Max Then
                SelectedFamily = CardFamilies.Diamond
            ElseIf FamNum(1) = FamNum.Max Then
                SelectedFamily = CardFamilies.Spade
            ElseIf FamNum(2) = FamNum.Max Then
                SelectedFamily = CardFamilies.Heart
            ElseIf FamNum(3) = FamNum.Max Then
                SelectedFamily = CardFamilies.Club
            End If
            Dim HighCardsCount2ndTry As Byte = CountHighCards(GetFamilyList(_PlayerCards, SelectedFamily))
            NumBet = FamNum.Max
            For Each c As Card In _PlayerCards
                If c.CardFamily <> SelectedFamily And c.CardNumber >= 11 Then
                    NumBet += 1
                End If
            Next
            If HighCardsCount2ndTry > 2 Then
                NumBet += Math.Round(3.5 / NumBet)
            Else
                NumBet = Math.Round(NumBet * 0.9)
            End If
            SecondBet = New Bet(NumBet, SelectedFamily)
            If FirstBet.Bet > SecondBet.Bet Then
                SelectedBet = FirstBet
            ElseIf FirstBet.Bet = SecondBet.Bet Then
                Dim FirstCardSum As Byte = CardValuesSum(GetFamilyList(_PlayerCards, FirstBet.CardFamily))
                Dim SecondCardSum As Byte = CardValuesSum(GetFamilyList(_PlayerCards, SecondBet.CardFamily))
                If FirstCardSum > SecondCardSum Then
                    SelectedBet = FirstBet
                Else
                    SelectedBet = SecondBet
                End If
            Else
                SelectedBet = SecondBet
            End If
            If SelectedBet.Bet < 7 Then
Pass:           Me._Bet = New Bet(0, 0)
            Else
                If HighestBetValue = 0 And PlayersBets.Length = 4 Then
                    If SelectedBet.Bet >= 7 Then
                        Me._Bet = New Bet(7, SelectedBet.CardFamily)
                    Else
                        Me._Bet = New Bet(0, SelectedBet.CardFamily)
                    End If
                Else
                    If SelectedBet.Bet > HighestBetValue Then
                        Me._Bet = New Bet(SelectedBet.Bet, SelectedBet.CardFamily)
                    Else
                        Me._Bet = New Bet(0, SelectedBet.CardFamily)
                    End If
                End If
            End If
            Return _Bet
        End Function

        Private Function GetHighestBetValue(ByVal Bets As Bet()) As Byte
            Dim max As Byte = 0
            For Each b As Bet In Bets
                If b.Bet > max Then
                    max = b.Bet
                End If
            Next
            Return max
        End Function

        Private Function CardValuesSum(ByVal Cards As Card()) As Byte
            If Not IsNothing(Cards) Then
                Dim ValuesSum As Byte = 0
                For Each c As Card In Cards
                    ValuesSum += c.CardNumber
                Next
                Return ValuesSum
            Else
                Return 0
            End If
        End Function

        Private Function GetPlayersCardsListWithoutSpecificFamily(ByVal PlayerCards As Card(), ByVal ExcludedFamily As CardFamilies) As Card()
            If Not IsNothing(PlayerCards) And Not IsNothing(ExcludedFamily) Then
                Dim ResultCardsList(0) As Card, j As Byte = 0
                For i = 0 To UBound(PlayerCards)
                    If PlayerCards(i).CardFamily <> ExcludedFamily Then
                        ReDim Preserve ResultCardsList(j)
                        ResultCardsList(j) = PlayerCards(i)
                        j += 1
                    End If
                Next
                If Not IsNothing(ResultCardsList) Then
                    Return ResultCardsList
                Else
                    Return Nothing
                End If
            Else
                If IsNothing(PlayerCards) Then
                    Throw New Exception("Invalid Function Parameter : PlayerCards")
                Else
                    Throw New Exception("Invalid Function Parameter : ExcludedFamily")
                End If
            End If
        End Function

        Private Function GetCardsWithoutTarneeb(ByVal Cards As Card()) As Card()
            If IsNothing(Cards) = False Then
                Dim CardsList() As Card
                Dim i As Integer = 0, iCards As Integer = 0
                For i = 0 To UBound(Cards)
                    If Cards(i).CardFamily <> _Tarneeb Then
                        ReDim Preserve CardsList(iCards)
                        CardsList(iCards) = Cards(i)
                        iCards += 1
                    End If
                Next
                If Not IsNothing(CardsList) Then
                    Return CardsList
                Else
                    Return Nothing
                End If
            Else
                Throw New Exception("Invalid Function Parameter")
            End If
        End Function

        Private Function GetWhichFamilyHasTheLess(ByVal Cards As Card()) As CardFamilies
            Dim NumHeart As Byte = 0, NumClub As Byte = 0, NumSpade As Byte = 0, NumDiamond As Byte = 0
            For i = 0 To UBound(Cards)
                If Cards(i).CardFamily = CardFamilies.Club Then
                    NumClub += 1
                ElseIf Cards(i).CardFamily = CardFamilies.Diamond Then
                    NumDiamond += 1
                ElseIf Cards(i).CardFamily = CardFamilies.Heart Then
                    NumHeart += 1
                Else
                    NumSpade += 1
                End If
            Next
            Dim NumList(3) As Byte
            NumList(0) = NumDiamond
            NumList(1) = NumSpade
            NumList(2) = NumHeart
            NumList(3) = NumClub
            Dim MinNum As Byte = 13
            For i = 0 To 3
                If NumList(i) <= MinNum And NumList(i) > 0 Then
                    MinNum = NumList(i)
                End If
            Next
            If MinNum = NumDiamond Then
                Return CardFamilies.Diamond
            ElseIf MinNum = NumSpade Then
                Return CardFamilies.Spade
            ElseIf MinNum = NumHeart Then
                Return CardFamilies.Heart
            ElseIf MinNum = NumClub Then
                Return CardFamilies.Club
            Else
                Throw New Exception("Unknown Problem")
            End If
        End Function

        Private Function GetSecondFamilyWhichHasTheLess(ByVal Cards As Card()) As CardFamilies
            Dim FamilyWhichHasTheLess As CardFamilies = GetWhichFamilyHasTheLess(Cards)
            Dim NumHeart As Byte = 0, NumClub As Byte = 0, NumSpade As Byte = 0, NumDiamond As Byte = 0
            For i = 0 To UBound(Cards)
                If Cards(i).CardFamily <> FamilyWhichHasTheLess Then
                    If Cards(i).CardFamily = CardFamilies.Club Then
                        NumClub += 1
                    ElseIf Cards(i).CardFamily = CardFamilies.Diamond Then
                        NumDiamond += 1
                    ElseIf Cards(i).CardFamily = CardFamilies.Heart Then
                        NumHeart += 1
                    Else
                        NumSpade += 1
                    End If
                End If
            Next
            Dim NumList(3) As Byte
            NumList(0) = NumDiamond
            NumList(1) = NumSpade
            NumList(2) = NumHeart
            NumList(3) = NumClub
            Dim MinNum As Byte = 13
            For i = 0 To 3
                If NumList(i) <= MinNum And NumList(i) > 0 Then
                    MinNum = NumList(i)
                End If
            Next
            If MinNum = NumDiamond Then
                Return CardFamilies.Diamond
            ElseIf MinNum = NumClub Then
                Return CardFamilies.Club
            ElseIf MinNum = NumHeart Then
                Return CardFamilies.Heart
            ElseIf MinNum = NumSpade Then
                Return CardFamilies.Spade
            Else
                Throw New Exception("Unknown Problem")
            End If
        End Function

        Private Function HasHighSerieFromFamily(ByVal CardsOfSameFamily As Card()) As Boolean 'retourne si Ace(0), King(1), Queen(2) et Jack(3)
            Dim HasCard(3) As Boolean
            For j = 0 To 3
                HasCard(j) = False
            Next
            Dim i As Integer = 0
            For i = 0 To UBound(CardsOfSameFamily)
                If CardsOfSameFamily(i).CardNumber = Tarneeb.Cards.Ace Then
                    HasCard(0) = True
                ElseIf CardsOfSameFamily(i).CardNumber = Tarneeb.Cards.King Then
                    HasCard(1) = True
                ElseIf CardsOfSameFamily(i).CardNumber = Tarneeb.Cards.Queen Then
                    HasCard(2) = True
                ElseIf CardsOfSameFamily(i).CardNumber = Tarneeb.Cards.Jack Then
                    HasCard(3) = True
                End If
            Next
            For k = 0 To 3
                If HasCard(k) = False Then
                    Return False
                    Exit Function
                End If
            Next
            Return True
        End Function

        Private Function GetHighSerieList(ByVal Cards As Card()) As Card()
            If IsNothing(Cards) = False Then
                Dim HighSerieList(0) As Card
                Dim i As Integer = 0
                For Each hcard As Card In Cards
                    If hcard.CardNumber >= 11 Then
                        ReDim Preserve HighSerieList(i)
                        HighSerieList(i) = hcard
                        i += 1
                    End If
                Next
                If IsNothing(HighSerieList) = False Then
                    Return HighSerieList
                Else
                    Return Nothing
                End If
            Else
                Throw New Exception("Invalid Function Parameter")
            End If
        End Function

        Private Function GetHigherCard(ByVal PlayerCards As Card(), ByVal TheCardToBeat As Card) As Card
            Dim HigherCards() As Card
            Dim i As Byte = 0
            For Each c As Card In PlayerCards
                If c.CardNumber > TheCardToBeat.CardNumber Then
                    ReDim Preserve HigherCards(i)
                    HigherCards(i) = c
                    i += 1
                End If
            Next
            If Not IsNothing(HigherCards) Then
                Return GetLowestCard(HigherCards)
            Else
                Return Nothing
            End If
        End Function

        Private Function HasSpecificCard(ByVal Cards As Card(), ByVal SpecificCard As Card) As Boolean
            For Each c As Card In Cards
                If c.CardNumber = SpecificCard.CardNumber And c.CardFamily = SpecificCard.CardFamily Then
                    Return True
                    Exit Function
                End If
            Next
            Return False
        End Function

        Private Function IsSureValue(ByVal IsSureCard As Card) As Boolean
            Dim iSure As Byte = IsSureCard.CardNumber - 2 + 1
            If IsSureCard.CardFamily = CardFamilies.Diamond Then
                For i = iSure To 12
                    If OutHighDiamondCardsOnTable(i) = False Then
                        Return False
                        Exit Function
                    End If
                Next
                Return True
            ElseIf IsSureCard.CardFamily = CardFamilies.Spade Then
                For i = iSure To 12
                    If OutHighSpadeCardsOnTable(i) = False Then
                        Return False
                        Exit Function
                    End If
                Next
                Return True
            ElseIf IsSureCard.CardFamily = CardFamilies.Heart Then
                For i = iSure To 12
                    If OutHighHeartCardsOnTable(i) = False Then
                        Return False
                        Exit Function
                    End If
                Next
                Return True
            Else
                For i = iSure To 12
                    If OutHighClubCardsOnTable(i) = False Then
                        Return False
                        Exit Function
                    End If
                Next
                Return True
            End If
        End Function

        Private Function GetPossibleSureCards(ByVal PlayerCards As Card()) As Card()
            If IsNothing(PlayerCards) = False Then
                Dim PossibleCards As Card()
                Dim SureCard As Card
                Dim i As Byte = 0
                For j = 2 To 14
                    Dim IsSureCardV As Byte = j
                    For k = 0 To 3
                        If IsSureValue(New Card(IsSureCardV, k)) = True Then
                            SureCard = New Card(IsSureCardV, k)
                            If HasSpecificCard(PlayerCards, SureCard) = True Then
                                ReDim Preserve PossibleCards(i)
                                PossibleCards(i) = SureCard
                                i += 1
                            End If
                        End If
                    Next
                Next
                If Not IsNothing(PossibleCards) Then
                    Return PossibleCards
                Else
                    Return Nothing
                End If
            Else
                Throw New Exception("Invalid Parameter : HighPlayerCards = Nothing")
            End If
        End Function

        Private Function GetPossibleSureCards(ByVal FamilyCards As Card(), ByVal SelectedFamily As CardFamilies) As Card()
            If IsNothing(FamilyCards) = False Then
                Dim PossibleCards As Card()
                Dim SureCard As Card
                Dim i As Byte = 0
                If SelectedFamily = CardFamilies.Diamond Then
                    For j = 2 To 14
                        If IsSureValue(New Card(j, CardFamilies.Diamond)) = True Then
                            SureCard = New Card(j, CardFamilies.Diamond)
                            If HasSpecificCard(FamilyCards, SureCard) = True Then
                                ReDim Preserve PossibleCards(i)
                                PossibleCards(i) = SureCard
                                i += 1
                            End If
                        End If
                    Next
                ElseIf SelectedFamily = CardFamilies.Spade Then
                    For j = 2 To 14
                        If IsSureValue(New Card(j, CardFamilies.Spade)) = True Then
                            SureCard = New Card(j, CardFamilies.Spade)
                            If HasSpecificCard(FamilyCards, SureCard) = True Then
                                ReDim Preserve PossibleCards(i)
                                PossibleCards(i) = SureCard
                                i += 1
                            End If
                        End If
                    Next
                ElseIf SelectedFamily = CardFamilies.Heart Then
                    For j = 2 To 14
                        If IsSureValue(New Card(j, CardFamilies.Heart)) = True Then
                            SureCard = New Card(j, CardFamilies.Heart)
                            If HasSpecificCard(FamilyCards, SureCard) = True Then
                                ReDim Preserve PossibleCards(i)
                                PossibleCards(i) = SureCard
                                i += 1
                            End If
                        End If
                    Next
                Else
                    For j = 2 To 14
                        If IsSureValue(New Card(j, CardFamilies.Club)) = True Then
                            SureCard = New Card(j, CardFamilies.Club)
                            If HasSpecificCard(FamilyCards, SureCard) = True Then
                                ReDim Preserve PossibleCards(i)
                                PossibleCards(i) = SureCard
                                i += 1
                            End If
                        End If
                    Next
                End If
                If Not IsNothing(PossibleCards) Then
                    Return PossibleCards
                Else
                    Return Nothing
                End If
            Else
                Dim ExMessage As String = "Invalid Parameter(s) For GetPossibleSureCards() Function : "
                If SelectedFamily = Nothing And IsNothing(FamilyCards) Then
                    Throw New Exception(ExMessage & "SelectedFamily = Nothing And HighFamilyCards = Nothing")
                ElseIf SelectedFamily = Nothing Then
                    Throw New Exception(ExMessage & "SelectedFamily = Nothing")
                Else
                    Throw New Exception(ExMessage & "HighFamilyCards = Nothing")
                End If
            End If
        End Function

        Private Function GetReallySureCards() As Card()
            Dim PossibleSureCards As Card() = GetPossibleSureCards(_PlayerCards)
            Return GetReallySureCards(PossibleSureCards)
        End Function

        Private Function GetReallySureCards(ByVal PossibleSureCards As Card()) As Card()
            Dim ReallySureCards As Card(), iReallySureCard As Byte = 0
            If Not IsNothing(PossibleSureCards) Then
                For Each c As Card In PossibleSureCards
                    Dim OpponentsHaveCardsOfFamily As Boolean = DoOpponentsHaveCardsOfFamily(c.CardFamily, True)
                    Dim OpponentsHaveTarneeb As Boolean = DoOpponentsHaveCardsOfFamily(_Tarneeb, True)
                    If OpponentsHaveCardsOfFamily = False And OpponentsHaveTarneeb = False Then
                        ReDim Preserve ReallySureCards(iReallySureCard)
                        ReallySureCards(iReallySureCard) = c
                        iReallySureCard += 1
                    End If
                Next
            End If
            For Each c As Card In _PlayerCards
                If c.CardNumber >= 11 Then
                    Dim OpponentsHaveCardsOfFamily As Boolean = DoOpponentsHaveCardsOfFamily(c.CardFamily, True)
                    If OpponentsHaveCardsOfFamily = True Then
                        ReDim Preserve ReallySureCards(iReallySureCard)
                        ReallySureCards(iReallySureCard) = c
                        iReallySureCard += 1
                    End If
                End If
            Next
            If _DeductionsProbability = 1 Then
                For Each c As Card In _PlayerCards
                    Dim TeammateHasCardsOfFamily As Boolean = DoTeammateHaveCardsOfFamily(c.CardFamily, False)
                    Dim TeammateHasTarneeb As Boolean = DoTeammateHaveCardsOfFamily(_Tarneeb, False)
                    Dim OpponentsHaveTarneeb As Boolean = DoOpponentsHaveCardsOfFamily(_Tarneeb, False)
                    If TeammateHasCardsOfFamily = False And TeammateHasTarneeb = True And OpponentsHaveTarneeb = False Then
                        ReDim Preserve ReallySureCards(iReallySureCard)
                        ReallySureCards(iReallySureCard) = c
                        iReallySureCard += 1
                    End If
                Next
            End If
            Return RemoveCardsPlayerShouldntPlay(ReallySureCards)
        End Function

        Private Function RemoveCardsPlayerShouldntPlay(ByVal ReallySureCards As Card()) As Card()
            For Each c As Card In ReallySureCards
                Dim OpponentsHaveFromFamily As Boolean = DoOpponentsHaveCardsOfFamily(c.CardFamily, False)
                Dim OpponentsHaveFromTarneeb As Boolean = DoOpponentsHaveCardsOfFamily(_Tarneeb, False)
                If OpponentsHaveFromFamily = False And OpponentsHaveFromTarneeb = True Then
                    ReallySureCards = RemoveCardFromPlayersHands(ReallySureCards, c)
                End If
            Next
            Return ReallySureCards
        End Function

        Public Sub EndOfRound(ByVal RoundCards As Card(), ByVal LastTurnWinnerInd As Byte) Implements IPlayer.EndOfRound
            If _MemorizingProbability > 0 Then
                For Each CardToMemorize As Card In RoundCards
                    If _MemorizingProbability < 1 Then
                        Dim ChanceNum As Decimal
                        Randomize()
                        ChanceNum = Rnd()
                        If ChanceNum < _MemorizingProbability Then
                            Memorize(CardToMemorize)
                        End If
                    Else
                        Memorize(CardToMemorize)
                    End If
                Next
                If _DeductionsProbability < 1 Then
                    Dim ChanceNum As Decimal
                    Randomize()
                    ChanceNum = Rnd()
                    If ChanceNum < _DeductionsProbability Then
                        MakeDeductions(RoundCards)
                        MakeSuppositions(RoundCards, LastTurnWinnerInd)
                    End If
                Else
                    MakeDeductions(RoundCards)
                    MakeSuppositions(RoundCards, LastTurnWinnerInd)
                End If
            End If
        End Sub

        Private Sub Memorize(ByVal CardToMemorize As Card)
            If CardToMemorize.CardNumber <> 0 Then
                'If CardToMemorize.CardNumber >= 11 Then
                If CardToMemorize.CardFamily = CardFamilies.Diamond Then
                    OutHighDiamondCardsOnTable(CardToMemorize.CardNumber - 2) = True '11
                ElseIf CardToMemorize.CardFamily = CardFamilies.Spade Then
                    OutHighSpadeCardsOnTable(CardToMemorize.CardNumber - 2) = True
                ElseIf CardToMemorize.CardFamily = CardFamilies.Heart Then
                    OutHighHeartCardsOnTable(CardToMemorize.CardNumber - 2) = True
                Else
                    OutHighClubCardsOnTable(CardToMemorize.CardNumber - 2) = True
                End If
                'End If
            Else
                Throw New Exception("Parameter CardToMemorize = Nothing")
            End If
        End Sub

        Private Sub MakeDeductions(ByVal RoundCards As Card())
            Dim MyPlayedCardID As Byte = GetCardID(RoundCards, _TurnCard)
            Dim PreviousCardID, NextCardID, TeamMateCardID As Byte
            Dim PreviousCard, NextCard, TeamMateCard As Card
            If MyPlayedCardID = 3 Then
                NextCardID = 0
            Else
                NextCardID = MyPlayedCardID + 1
            End If
            If MyPlayedCardID = 0 Then
                PreviousCardID = 3
            Else
                PreviousCardID = MyPlayedCardID - 1
            End If
            If MyPlayedCardID >= 2 Then
                TeamMateCardID = MyPlayedCardID - 2
            Else
                TeamMateCardID = MyPlayedCardID + 2
            End If
            PreviousCard = RoundCards(PreviousCardID)
            NextCard = RoundCards(NextCardID)
            TeamMateCard = RoundCards(TeamMateCardID)
            Dim TurnFamily As CardFamilies = RoundCards(0).CardFamily
            If NextCard.CardFamily <> TurnFamily Then
                DeductionsFamiliesFinishedNextPlayer(TurnFamily) = True
            End If
            If PreviousCard.CardFamily <> TurnFamily Then
                DeductionsFamiliesFinishedPreviousPlayer(TurnFamily) = True
            End If
            If TeamMateCard.CardFamily <> TurnFamily Then
                DeductionsFamiliesFinishedTeammate(TurnFamily) = True
            End If
        End Sub

        Private Sub MakeSuppositions(ByVal RoundCards As Card(), ByVal WinningPlayerInd As Byte)
            Dim WinningCard As Card = RoundCards(WinningPlayerInd)
            Dim TurnFamily As CardFamilies = RoundCards(0).CardFamily
            Dim AfterWinningCardList As Card() = GetCardsListAfterCard(RoundCards, WinningCard)
            If Not IsNothing(AfterWinningCardList) Then
                For i = 0 To UBound(AfterWinningCardList)
                    If AfterWinningCardList(i).CardFamily = TurnFamily And AfterWinningCardList(i).CardNumber >= 9 Then
                        Dim PlayerCardID As SByte = GetCardID(RoundCards, AfterWinningCardList(i))
                        Dim MyCardID As SByte = GetCardID(RoundCards, _TurnCard)
                        Dim PlayersDifference As SByte = PlayerCardID - MyCardID
                        If PlayersDifference = -1 Then
                            SuppositionsFamiliesFinishedPreviousPlayer(TurnFamily) = True
                        ElseIf PlayersDifference = 1 Then
                            SuppositionsFamiliesFinishedNextPlayer(TurnFamily) = True
                        ElseIf PlayersDifference = 2 Or PlayersDifference = -2 Then
                            SuppositionsFamiliesFinishedTeammate(TurnFamily) = True
                        End If
                    End If
                Next
            End If
        End Sub

        Private Function GetCardsListAfterCard(ByVal CardsList As Card(), ByVal CardBeforeList As Card) As Card()
            Dim CardBeforeListID As Byte = GetCardID(CardsList, CardBeforeList)
            Dim iCard As Byte = 0, NewCardsList As Card()
            For i = (CardBeforeListID + 1) To UBound(CardsList)
                ReDim Preserve NewCardsList(iCard)
                NewCardsList(iCard) = CardsList(i)
                iCard += 1
            Next
            Return NewCardsList
        End Function

        Private Function DoOpponentsHaveCardsOfFamily(ByVal CF As CardFamilies, ByVal WithSuppositions As Boolean) As Boolean
            If DeductionsFamiliesFinishedNextPlayer(CF) = False Then
                Return True
                Exit Function
            End If
            If DeductionsFamiliesFinishedPreviousPlayer(CF) = False Then
                Return True
                Exit Function
            End If
            If WithSuppositions = True Then
                If SuppositionsFamiliesFinishedNextPlayer(CF) = True Or SuppositionsFamiliesFinishedPreviousPlayer(CF) = True Then
                    Dim TurnNumber As Byte = 13 - _PlayerCards.Length + 1
                    Dim Prob As Decimal = TurnNumber / 13
                    Dim Rand As Decimal
                    Randomize()
                    Rand = Rnd()
                    If Rand < Prob Then
                        Return False
                        Exit Function
                    Else
                        Return True
                        Exit Function
                    End If
                End If
            End If
            Return False
        End Function

        Private Function DoTeammateHaveCardsOfFamily(ByVal CF As CardFamilies, ByVal WithSuppositions As Boolean) As Boolean
            If DeductionsFamiliesFinishedTeammate(CF) = False Then
                Return True
                Exit Function
            End If
            If WithSuppositions = True Then
                If SuppositionsFamiliesFinishedTeammate(CF) = True Then
                    Dim TurnNumber As Byte = 13 - _PlayerCards.Length + 1
                    Dim Prob As Decimal = TurnNumber / 13
                    Dim Rand As Decimal
                    Randomize()
                    Rand = Rnd()
                    If Rand < Prob Then
                        Return False
                        Exit Function
                    Else
                        Return True
                        Exit Function
                    End If
                End If
            End If
            Return False
        End Function
    End Class

    Public Module CommonCardFunctions

        Public Function GetCardID(ByVal PlayerCards As Card(), ByVal TheCardToFind As Card) As Byte
            If Not IsNothing(PlayerCards) And Not IsNothing(TheCardToFind) Then
                For i = 0 To UBound(PlayerCards)
                    If PlayerCards(i) = TheCardToFind Then
                        Return i
                        Exit Function
                    End If
                Next
                Return 255
            Else
                Throw New ArgumentNullException
            End If
        End Function

        Public Function GetTheFamilyWhichHasTheMost(ByVal PlayerCards As Card()) As CardFamilies
            If Not IsNothing(PlayerCards) Then
                Dim NumCards(3) As Byte
                For i = 0 To 3
                    NumCards(i) = 0
                Next
                For Each c As Card In PlayerCards
                    NumCards(c.CardFamily) += 1
                Next
                Dim Max As Byte = NumCards.Max
                For i = 0 To 3
                    If NumCards(i) = Max Then
                        Return i
                        Exit Function
                    End If
                Next
            Else
                Throw New ArgumentNullException
            End If
        End Function

        Public Function GetFamilyList(ByVal Cards As Card(), ByVal CFamily As CardFamilies) As Card()
            Dim FamilyList() As Card
            Dim i As Integer = 0, iFamily As Integer = 0
            For i = 0 To UBound(Cards)
                If Cards(i).CardFamily = CFamily And Cards(i).CardNumber > 0 Then
                    ReDim Preserve FamilyList(iFamily)
                    FamilyList(iFamily) = Cards(i)
                    iFamily += 1
                End If
            Next
            Return FamilyList
        End Function

        Public Function GetHighestCard(ByVal Cards As Card()) As Card
            If Not IsNothing(Cards) Then
                If Cards.Length = 1 Then
                    Return Cards(0)
                Else
                    Dim HighestCardSelection As Card
                    Dim HighNum As Byte = 0
                    For i = 0 To UBound(Cards)
                        If Cards(i).CardNumber > HighNum Then
                            HighestCardSelection = Cards(i)
                            HighNum = Cards(i).CardNumber
                        End If
                    Next
                    Return HighestCardSelection
                End If
            Else
                Throw New Exception("Invalid Function Parameter")
            End If
        End Function

        Public Function GetLowestCard(ByVal Cards As Card()) As Card
            If Not IsNothing(Cards) Then
                Dim LowestCardSelection As Card
                Dim LowNum As Byte = 14
                For i = 0 To UBound(Cards)
                    If Cards(i).CardNumber <= LowNum And Cards(i).CardNumber >= 2 Then
                        LowestCardSelection = Cards(i)
                        LowNum = Cards(i).CardNumber
                    End If
                Next
                Return LowestCardSelection
            Else
                Throw New ArgumentNullException("Invalid Function Parameter")
            End If
        End Function

        Public Function CountHighCards(ByVal FamilyList As Card()) As Byte
            If Not IsNothing(FamilyList) Then
                Dim NumHighCards As Byte = 0
                For Each c As Card In FamilyList
                    If c.CardNumber >= 11 Then
                        NumHighCards += 1
                    End If
                Next
                Return NumHighCards
            Else
                Throw New ArgumentNullException
            End If
        End Function

        Public Function RemoveCardFromPlayersHands(ByVal PlayerCards As Card(), ByVal TheIDOfCardToRemove As Byte) As Card() 'Retourne les cartes du joueur sans les valeurs à supprimer
            If Not IsNothing(PlayerCards) Then
                If TheIDOfCardToRemove <= UBound(PlayerCards) Then
                    If TheIDOfCardToRemove = 0 And UBound(PlayerCards) = 1 Then
                        PlayerCards(0) = PlayerCards(1)
                        ReDim Preserve PlayerCards(0)
                    Else
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
                    End If
                    Return PlayerCards
                    'ElseIf TheIDOfCardToRemove = UBound(PlayerCards) Then
                    'ReDim Preserve PlayerCards(UBound(PlayerCards) - 1)
                    'Return PlayerCards
                Else
                    Throw New Exception("Invalid Card ID")
                End If
            Else
                Throw New ArgumentNullException
            End If
        End Function

        Public Function RemoveCardFromPlayersHands(ByVal PlayerCards As Card(), ByVal TheCardToRemove As Card) As Card()
            Dim CrdID As Byte = GetCardID(PlayerCards, TheCardToRemove)
            If CrdID <> 255 Then
                Return RemoveCardFromPlayersHands(PlayerCards, CrdID)
            Else
                Throw New IndexOutOfRangeException
            End If
        End Function
    End Module
End Namespace