Attribute VB_Name = "modHandleData"
Option Explicit

' ******************************************
' **              BMO Source              **
' ******************************************

Sub HandleData(ByVal Index As Long, ByVal Data As String)
'On Error Resume Next
Dim Parse() As String
        
    ' Handle incoming data
    Parse = Split(Data, SEP_CHAR)
    
    Select Case Parse(0)
        Case CGetClasses
            HandleGetClasses Index
        Case CNewAccount
            HandleNewAccount Index, Parse
        Case CDelAccount
            HandleDelAccount Index, Parse
        Case CLogin
            HandleLogin Index, Parse
        Case CAddChar
            HandleAddChar Index, Parse
        Case CDelChar
            HandleDelChar Index, Parse
        Case CUseChar
            HandleUseChar Index, Parse
        Case CSayMsg
            HandleSayMsg Index, Parse
        Case CEmoteMsg
            HandleEmoteMsg Index, Parse
        Case CBroadcastMsg
            HandleBroadcastMsg Index, Parse
        Case CGlobalMsg
            HandleGlobalMsg Index, Parse
        Case CAdminMsg
            HandleAdminMsg Index, Parse
        Case CPlayerMsg
            HandlePlayerMsg Index, Parse
        Case CPlayerMove
            HandlePlayerMove Index, Parse
        Case CPlayerDir
            HandlePlayerDir Index, Parse
        Case CUseItem
            HandleUseItem Index, Parse
        Case CAttack
            HandleAttack Index
        Case CUseStatPoint
            HandleUseStatPoint Index, Parse
        Case CPlayerInfoRequest
            HandlePlayerInfoRequest Index, Parse
        Case CWarpMeTo
            HandleWarpMeTo Index, Parse
        Case CWarpToMe
            HandleWarpToMe Index, Parse
        Case CWarpTo
            HandleWarpTo Index, Parse
        Case CSetSprite
            HandleSetSprite Index, Parse
        Case CGetStats
            HandleGetStats Index
        Case CRequestNewMap
            HandleRequestNewMap Index, Parse
        Case CMapData
            HandleMapData Index, Parse
        Case CNeedMap
            HandleNeedMap Index, Parse
        Case CMapGetItem
            HandleMapGetItem Index
        Case CMapDropItem
            HandleMapDropItem Index, Parse
        Case CMapRespawn
            HandleMapRespawn Index
        Case CMapReport
            HandleMapReport Index
        Case CKickPlayer
            HandleKickPlayer Index, Parse
        Case CBanList
            HandleBanList Index
        Case CBanDestroy
            HandleBanDestroy Index
        Case CBanPlayer
            HandleBanPlayer Index, Parse
        Case CRequestEditMap
            HandleRequestEditMap Index
        Case CRequestEditItem
            HandleRequestEditItem Index
        Case CEditItem
            HandleEditItem Index, Parse
        Case CSaveItem
            HandleSaveItem Index, Parse
        Case CDelete
            HandleDelete Index, Parse
        Case CRequestEditNpc
            HandleRequestEditNpc Index
        Case CEditNpc
            HandleEditNpc Index, Parse
        Case CSaveNpc
            HandleSaveNpc Index, Parse
        Case CRequestEditShop
            HandleRequestEditShop Index
        Case CEditShop
            HandleEditShop Index, Parse
        Case CSaveShop
            HandleSaveShop Index, Parse
        Case CRequestEditSpell
            HandleRequestEditSpell Index
        Case CEditSpell
            HandleEditSpell Index, Parse
        Case CSaveSpell
            HandleSaveSpell Index, Parse
        Case CSetAccess
            HandleSetAccess Index, Parse
        Case CWhosOnline
            HandleWhosOnline Index
        Case CSetMotd
            HandleSetMotd Index, Parse
        Case CTrade
            HandleTrade Index
        Case CTradeRequest
            HandleTradeRequest Index, Parse
        Case CFixItem
            HandleFixItem Index, Parse
        Case CSearch
            HandleSearch Index, Parse
        Case CParty
            HandleParty Index, Parse
        Case CJoinParty
            HandleJoinParty Index, Parse
        Case CLeaveParty
            HandleLeaveParty Index, Parse
        Case CSpells
            HandleSpells Index
        Case CCast
            HandleCast Index, Parse
        Case CQuit
            HandleQuit Index
        Case CFireDeath
            HandleFireDeath Index, Parse
        Case CJump
            HandleJump Index
        Case CRequestItemSpawn
            HandleRequestItemSpawn Index, Parse
        Case CRequestRoomList
            HandleRequestRoomList Index
        Case CRequestJoinRoom
            HandleRequestJoinRoom Index, Parse
        Case CLeaveRoom
            HandleLeaveRoom Index
        Case CRequestHighScores
            HandleRequestHighScores Index
        Case CRequestStats
            HandleRequestStats Index
        Case CAddFriend
            HandleAddFriend Index, Parse
        Case CRemoveFriend
            HandleRemoveFriend Index, Parse
        Case CUpdateFriendList
            UpdateFriendsList Index
        Case CAddStatus
             HandleAddStatus Index, Parse
        Case CDealBP
             HandleDealBP Index, Parse
        Case CDealBC
             HandleDealBC Index, Parse
        Case CSaveSprite
             HandleSaveSprite Index, Parse
        Case CThrow
             HandleThrow Index
    End Select
End Sub
' :::::::::::::::::::::::::::::::::::::::::::::::
' :: Requesting classes for making a character ::
' :::::::::::::::::::::::::::::::::::::::::::::::
Sub HandleGetClasses(ByVal Index As Long)
    If Not IsPlaying(Index) Then
        Call SendNewCharClasses(Index)
    End If
End Sub
' ::::::::::::::::::::::::
' :: New account packet ::
' ::::::::::::::::::::::::
Sub HandleNewAccount(ByVal Index As Long, ByRef Parse() As String)
Dim Name As String, Password As String
Dim i As Long, n As Long

    If Not IsPlaying(Index) And Not IsLoggedIn(Index) Then
        ' Get the data
        Name = Parse(1)
        Password = Parse(2)
        
        ' Prevent hacking
        If Len(Trim$(Name)) < 3 Or Len(Trim$(Password)) < 3 Then
            Call AlertMsg(Index, "Your name and password must be at least three characters in length")
            Exit Sub
        End If
        
        ' Prevent hacking
        For i = 1 To Len(Name)
            n = AscW(Mid$(Name, i, 1))
            
            If (n >= 65 And n <= 90) Or (n >= 97 And n <= 122) Or (n = 95) Or (n = 32) Or (n >= 48 And n <= 57) Then
            Else
                Call AlertMsg(Index, "Invalid name, only letters, numbers, spaces, and _ allowed in names.")
                Exit Sub
            End If
        Next
        
        ' Check to see if account already exists
        If Not AccountExist(Name) Then
            Call AddAccount(Index, Name, Password)
            AddChar Index, Name, Val(Parse(3)), Val(Parse(4)), 1
            Call TextAdd(frmServer.txtText, "Conta " & Name & " foi criada.")
            Call AddLog("Account " & Name & " foi criada.", PLAYER_LOG)
            Call AlertMsg(Index, "Conta criada com sucesso!")
        Else
            Call AlertMsg(Index, "Login não disponível !")
        End If
        
    End If
End Sub
' :::::::::::::::::::::::::::
' :: Delete account packet ::
' :::::::::::::::::::::::::::
Sub HandleDelAccount(ByVal Index As Long, ByRef Parse() As String)
Dim Name As String, Password As String
Dim i As Long

    If Not IsPlaying(Index) And Not IsLoggedIn(Index) Then
        ' Get the data
        Name = Parse(1)
        Password = Parse(2)
        
        ' Prevent hacking
        If Len(Trim$(Name)) < 3 Or Len(Trim$(Password)) < 3 Then
            Call AlertMsg(Index, "The name and password must be at least three characters in length")
            Exit Sub
        End If
        
        If Not AccountExist(Name) Then
            Call AlertMsg(Index, "That account name does not exist.")
            Exit Sub
        End If
        
        If Not PasswordOK(Name, Password) Then
            Call AlertMsg(Index, "Incorrect password.")
            Exit Sub
        End If
                    
        ' Delete names from master name file
        Call LoadPlayer(Index, Name)
        For i = 1 To MAX_CHARS
            If LenB(Trim$(Player(Index).Char(i).Name)) > 0 Then
                Call DeleteName(Player(Index).Char(i).Name)
            End If
        Next
        Call ClearPlayer(Index)
        
        ' Everything went ok
        Call Kill(App.Path & "\Accounts\" & Trim$(Name) & ".bin")
        Call AddLog("Account " & Trim$(Name) & " has been deleted.", PLAYER_LOG)
        Call AlertMsg(Index, "Your account has been deleted.")
    End If
End Sub
' ::::::::::::::::::
' :: Login packet ::
' ::::::::::::::::::
Sub HandleLogin(ByVal Index As Long, ByRef Parse() As String)
Dim Name As String, Password As String
Dim i As Long, n As Long, f As Long

    If Not IsPlaying(Index) And Not IsLoggedIn(Index) Then
        ' Get the data
        Name = Parse(1)
        Password = Parse(2)
    
        ' Check versions
        If Val(Parse(3)) <> App.Major Or Val(Parse(4)) <> App.Minor Or Val(Parse(5)) <> App.Revision Then
            Call AlertMsg(Index, "Versão desatualizada,visite o site " & GAME_WEBSITE)
            Exit Sub
        End If
        
        If Len(Trim$(Name)) < 3 Or Len(Trim$(Password)) < 3 Then
            Call AlertMsg(Index, "Seu login e senha devem conter mais de três caracteres")
            Exit Sub
        End If
        
        If Not AccountExist(Name) Then
            Call AlertMsg(Index, "Login ou senha inválidos.")
            Exit Sub
        End If
    
        If Not PasswordOK(Name, Password) Then
            Call AlertMsg(Index, "Login ou senha inválidos.")
            Exit Sub
        End If
    
        If IsMultiAccounts(Name) Then
            Call AlertMsg(Index, "A conta já está conectada.")
            Exit Sub
        End If
        
        ' Prevent Dupeing
        For i = 1 To Len(Name)
            n = AscW(Mid(Name, i, 1))
            If (n >= 65 And n <= 90) Or (n >= 97 And n <= 122) Or (n = 95) Or (n = 32) Or (n >= 48 And n <= 57) Then
                'ok!
            Else
                Call AlertMsg(Index, "Login não autorizado!")
                Exit Sub
            End If
        Next

        ' Everything went ok

        ' Load the player
        Call LoadPlayer(Index, Name)
        'Call SendChars(Index)
    
        ' Check to make sure the character exists and if so, set it as its current char
        If CharExist(Index, 1) Then
            TempPlayer(Index).CharNum = 1
            SendDataTo Index, SInLobby & END_CHAR
            
            frmServer.lvwInfo.ListItems(Index).SubItems(1) = GetPlayerIP(Index)
            frmServer.lvwInfo.ListItems(Index).SubItems(2) = GetPlayerLogin(Index)
            frmServer.lvwInfo.ListItems(Index).SubItems(3) = GetPlayerName(Index)
            
            TempPlayer(Index).InGame = True
            
            SetPlayerMap Index, 100
            SendPlayerData Index
            
            'Call JoinGame(Index)
            
            ' Send an ok to client to start receiving in game data
            Call SendDataTo(Index, SLoginOk & SEP_CHAR & Index & SEP_CHAR & Player_HighIndex & END_CHAR)
            
            TotalPlayersOnline = TotalPlayersOnline + 1
            Call UpdateHighIndex

            GlobalMsg GetPlayerName(Index) & " está online!", Green
            
            ' Send welcome messages
            Call SendWelcome(Index)
            Call SendItems(Index)
            Call AddLog(GetPlayerLogin(Index) & "/" & GetPlayerName(Index) & " has began playing " & GAME_NAME & ".", PLAYER_LOG)
            Call TextAdd(frmServer.txtText, GetPlayerLogin(Index) & "/" & GetPlayerName(Index) & " has began playing " & GAME_NAME & ".")
            Call UpdateCaption
            Call SetPlayerBPoints(Index, GetPlayerBPoints(Index))
            Call SetPlayerBCash(Index, GetPlayerBCash(Index))
            If GetPlayerBCash(Index) = 0 Then
            Call SetPlayerBCash(Index, 101)
            End If
            
            ' Now we want to check if they are already on the master list (this makes it add the user if they already haven't been added to the master list for older accounts)
            If Not FindChar(GetPlayerName(Index)) Then
                f = FreeFile
                Open App.Path & "\accounts\charlist.txt" For Append As #f
                    Print #f, GetPlayerName(Index)
                Close #f
            End If
        Else
            Call AlertMsg(Index, "Personagem não existe!")
        End If

        ' Show the player up on the socket status
        Call AddLog(GetPlayerLogin(Index) & " se logou com o ip " & GetPlayerIP(Index) & ".", PLAYER_LOG)
        Call TextAdd(frmServer.txtText, GetPlayerLogin(Index) & " se logou com ip " & GetPlayerIP(Index) & ".")
    End If
End Sub
' ::::::::::::::::::::::::::
' :: Add character packet ::
' ::::::::::::::::::::::::::
Sub HandleAddChar(ByVal Index As Long, ByRef Parse() As String)
Dim Name As String, Password As String
Dim Sex As Long, Class As Long, CharNum As Long
Dim i As Long, n As Long

    If Not IsPlaying(Index) Then
        Name = Parse(1)
        Sex = Val(Parse(2))
        Class = Val(Parse(3))
        CharNum = Val(Parse(4))
    
        ' Prevent hacking
        If Len(Trim$(Name)) < 3 Then
            Call AlertMsg(Index, "Character name must be at least three characters in length.")
            Exit Sub
        End If
        
        ' Prevent hacking
        For i = 1 To Len(Name)
            n = AscW(Mid$(Name, i, 1))
            
            If (n >= 65 And n <= 90) Or (n >= 97 And n <= 122) Or (n = 95) Or (n = 32) Or (n >= 48 And n <= 57) Then
            Else
                Call AlertMsg(Index, "Invalid name, only letters, numbers, spaces, and _ allowed in names.")
                Exit Sub
            End If
        Next
                                
        ' Prevent hacking
        If CharNum < 1 Or CharNum > MAX_CHARS Then
            Call HackingAttempt(Index, "Invalid CharNum")
            Exit Sub
        End If
    
        ' Prevent hacking
        If (Sex < SEX_MALE) Or (Sex > SEX_FEMALE) Then
            Call HackingAttempt(Index, "Invalid Sex (dont laugh)")
            Exit Sub
        End If
        
        ' Prevent hacking
        If Class < 1 Or Class > Max_Classes Then
            Call HackingAttempt(Index, "Invalid Class")
            Exit Sub
        End If
    
        ' Check if char already exists in slot
        If CharExist(Index, CharNum) Then
            Call AlertMsg(Index, "Character already exists!")
            Exit Sub
        End If
        
        ' Check if name is already in use
        If FindChar(Name) Then
            Call AlertMsg(Index, "Sorry, but that name is in use!")
            Exit Sub
        End If
    
        ' Everything went ok, add the character
        Call AddChar(Index, Name, Sex, Class, CharNum)
        Call AddLog("Character " & Name & " added to " & GetPlayerLogin(Index) & "'s account.", PLAYER_LOG)
        Call AlertMsg(Index, "Character has been created!")
    End If
End Sub
' :::::::::::::::::::::::::::::::
' :: Deleting character packet ::
' :::::::::::::::::::::::::::::::
Sub HandleDelChar(ByVal Index As Long, ByRef Parse() As String)
Dim CharNum As Long

    If Not IsPlaying(Index) Then
        CharNum = Val(Parse(1))
    
        ' Prevent hacking
        If CharNum < 1 Or CharNum > MAX_CHARS Then
            Call HackingAttempt(Index, "Invalid CharNum")
            Exit Sub
        End If
        
        Call DelChar(Index, CharNum)
        Call AddLog("Character deleted on " & GetPlayerLogin(Index) & "'s account.", PLAYER_LOG)
        Call AlertMsg(Index, "Character has been deleted!")
    End If
End Sub
' ::::::::::::::::::::::::::::
' :: Using character packet ::
' ::::::::::::::::::::::::::::
Sub HandleUseChar(ByVal Index As Long, ByRef Parse() As String)
Dim CharNum As Long
Dim f As Long

    If Not IsPlaying(Index) Then
        CharNum = Val(Parse(1))
    
        ' Prevent hacking
        If CharNum < 1 Or CharNum > MAX_CHARS Then
            Call HackingAttempt(Index, "Invalid CharNum")
            Exit Sub
        End If
    
        ' Check to make sure the character exists and if so, set it as its current char
        If CharExist(Index, CharNum) Then
            TempPlayer(Index).CharNum = CharNum
            TempPlayer(Index).Distance = 1 + Player(Index).Char(TempPlayer(Index).CharNum).BonusDistance
        TempPlayer(Index).Throw = False
        TempPlayer(Index).Speed = False
            Call JoinGame(Index)
        
            CharNum = TempPlayer(Index).CharNum
            Call AddLog(GetPlayerLogin(Index) & "/" & GetPlayerName(Index) & " has began playing " & GAME_NAME & ".", PLAYER_LOG)
            Call TextAdd(frmServer.txtText, GetPlayerLogin(Index) & "/" & GetPlayerName(Index) & " has began playing " & GAME_NAME & ".")
            Call UpdateCaption
            
            ' Now we want to check if they are already on the master list (this makes it add the user if they already haven't been added to the master list for older accounts)
            If Not FindChar(GetPlayerName(Index)) Then
                f = FreeFile
                Open App.Path & "\accounts\charlist.txt" For Append As #f
                    Print #f, GetPlayerName(Index)
                Close #f
            End If
        Else
            Call AlertMsg(Index, "Character does not exist!")
        End If
    End If
End Sub
' ::::::::::::::::::::
' :: Social packets ::
' ::::::::::::::::::::
Sub HandleSayMsg(ByVal Index As Long, ByRef Parse() As String)
Dim Msg As String
Dim i As Long

    Msg = Parse(1)
    
    ' Prevent hacking
    For i = 1 To Len(Msg)
        If AscW(Mid$(Msg, i, 1)) < 32 Or AscW(Mid$(Msg, i, 1)) > 126 Then
            'Call HackingAttempt(Index, "Modificação de texto")
            Exit Sub
        End If
    Next
    
    Call AddLog("Map #" & GetPlayerMap(Index) & ": " & GetPlayerName(Index) & " says, '" & Msg & "'", PLAYER_LOG)
    Call MapMsg(GetPlayerMap(Index), GetPlayerName(Index) & ": " & Msg, SayColor)
End Sub

Sub HandleEmoteMsg(ByVal Index As Long, ByRef Parse() As String)
Dim Msg As String
Dim i As Long

    Msg = Parse(1)
    
    ' Prevent hacking
    For i = 1 To Len(Msg)
        If AscW(Mid$(Msg, i, 1)) < 32 Or AscW(Mid$(Msg, i, 1)) > 126 Then
            Call HackingAttempt(Index, "Emote Text Modification")
            Exit Sub
        End If
    Next
    
    Call AddLog("Map #" & GetPlayerMap(Index) & ": " & GetPlayerName(Index) & " " & Msg, PLAYER_LOG)
    Call MapMsg(GetPlayerMap(Index), GetPlayerName(Index) & " " & Right$(Msg, Len(Msg) - 1), EmoteColor)
End Sub

Sub HandleBroadcastMsg(ByVal Index As Long, ByRef Parse() As String)
Dim Msg As String, s As String
Dim i As Long
Dim RoomName As String

    Msg = Parse(1)
    
    ' Prevent hacking
    For i = 1 To Len(Msg)
        If AscW(Mid$(Msg, i, 1)) < 32 Or AscW(Mid$(Msg, i, 1)) > 126 Then
            Call HackingAttempt(Index, "Broadcast Text Modification")
            Exit Sub
        End If
    Next
    
    If GetPlayerMap(Index) = 100 Then RoomName = "Lobby" Else RoomName = "Sala " & GetPlayerMap(Index)
    
    s = "[" & RoomName & "] " & GetPlayerName(Index) & ": " & Msg
    Call AddLog(s, PLAYER_LOG)
    Call GlobalMsg(s, BroadcastColor)
    Call TextAdd(frmServer.txtText, s)
End Sub

Sub HandleGlobalMsg(ByVal Index As Long, ByRef Parse() As String)
Dim Msg As String, s As String
Dim i As Long

    Msg = Parse(1)
    
    ' Prevent hacking
    For i = 1 To Len(Msg)
        If AscW(Mid$(Msg, i, 1)) < 32 Or AscW(Mid$(Msg, i, 1)) > 126 Then
            Call HackingAttempt(Index, "Global Text Modification")
            Exit Sub
        End If
    Next
    
    If GetPlayerAccess(Index) > 0 Then
        s = "[ALERTA] " & GetPlayerName(Index) & ": " & Msg
        Call AddLog(s, ADMIN_LOG)
        Call GlobalMsg(s, GlobalColor)
        Call TextAdd(frmServer.txtText, s)
    End If
End Sub

Sub HandleAdminMsg(ByVal Index As Long, ByRef Parse() As String)
Dim Msg As String
Dim i As Long

    Msg = Parse(1)
    
    ' Prevent hacking
    For i = 1 To Len(Msg)
        If AscW(Mid$(Msg, i, 1)) < 32 Or AscW(Mid$(Msg, i, 1)) > 126 Then
            Call HackingAttempt(Index, "Admin Text Modification")
            Exit Sub
        End If
    Next
    
    If GetPlayerAccess(Index) > 0 Then
        Call AddLog("[ADMIN] " & GetPlayerName(Index) & Msg, ADMIN_LOG)
        Call AdminMsg("[ADMIN] " & GetPlayerName(Index) & Msg, AdminColor)
    End If
End Sub

Sub HandlePlayerMsg(ByVal Index As Long, ByRef Parse() As String)
Dim Msg As String
Dim i As Long, MsgTo As Long

    MsgTo = FindPlayer(Parse(1))
    Msg = Parse(2)
    
    ' Prevent hacking
    For i = 1 To Len(Msg)
        If AscW(Mid$(Msg, i, 1)) < 32 Or AscW(Mid$(Msg, i, 1)) > 126 Then
            Call HackingAttempt(Index, "Player Msg Text Modification")
            Exit Sub
        End If
    Next
    
    ' Check if they are trying to talk to themselves
    'If MsgTo <> Index Then
        If MsgTo > 0 Then
            Call AddLog(GetPlayerName(Index) & " diz " & GetPlayerName(MsgTo) & Msg & "'", PLAYER_LOG)
            Call PlayerMsg(MsgTo, GetPlayerName(Index) & " Fala para você: " & Msg, TellColor)
            Call PlayerMsg(Index, "Você fala: " & GetPlayerName(MsgTo) & Msg & "'", TellColor)
        Else
            Call PlayerMsg(Index, "O jogador não está online.", Red)
        End If
    'Else
    '    Call AddLog("Map #" & GetPlayerMap(Index) & ": " & GetPlayerName(Index) & " begins to mumble to himself, what a wierdo...", PLAYER_LOG)
    '    Call MapMsg(GetPlayerMap(Index), GetPlayerName(Index) & " begins to mumble to himself, what a wierdo...", Green)
    'End If
End Sub
' :::::::::::::::::::::::::::::
' :: Moving character packet ::
' :::::::::::::::::::::::::::::
Sub HandlePlayerMove(ByVal Index As Long, ByRef Parse() As String)
Dim Dir As Long, Movement As Long

    If TempPlayer(Index).GettingMap = YES Then
        Exit Sub
    End If

    Dir = Val(Parse(1))
    Movement = Val(Parse(2))
    
    ' Prevent hacking
    If Dir < DIR_UP Or Dir > DIR_RIGHT Then
        Call HackingAttempt(Index, "Invalid Direction")
        Exit Sub
    End If
    
    ' Prevent hacking
    If Movement < 1 Or Movement > 2 Then
        Call HackingAttempt(Index, "Invalid Movement")
        Exit Sub
    End If
    
    ' Prevent player from moving if they have casted a spell
    If TempPlayer(Index).CastedSpell = YES Then
        ' Check if they have already casted a spell, and if so we can't let them move
        If GetTickCount > TempPlayer(Index).AttackTimer + 1000 Then
            TempPlayer(Index).CastedSpell = NO
        Else
            Call SendPlayerXY(Index)
            Exit Sub
        End If
    End If
    
    Call PlayerMove(Index, Dir, Movement)
End Sub

' :::::::::::::::::::::::::::::
' :: Moving character packet ::
' :::::::::::::::::::::::::::::
Sub HandlePlayerDir(ByVal Index As Long, ByRef Parse() As String)
Dim Dir As Long

    If TempPlayer(Index).GettingMap = YES Then
        Exit Sub
    End If

    Dir = Val(Parse(1))
    
    ' Prevent hacking
    If Dir < DIR_UP Or Dir > DIR_RIGHT Then
        Call HackingAttempt(Index, "Invalid Direction")
        Exit Sub
    End If
    
    Call SetPlayerDir(Index, Dir)
    'Call SendDataToMapBut(Index, GetPlayerMap(Index), SPlayerDir & SEP_CHAR & Index & SEP_CHAR & GetPlayerDir(Index) & END_CHAR)
End Sub

' :::::::::::::::::::::
' :: Use item packet ::
' :::::::::::::::::::::
Sub HandleUseItem(ByVal Index As Long, ByRef Parse() As String)
Dim InvNum As Long, CharNum As Long
Dim i As Long, n As Long, x As Long, y As Long

    InvNum = Val(Parse(1))
    CharNum = TempPlayer(Index).CharNum
    
    ' Prevent hacking
    If InvNum < 1 Or InvNum > MAX_ITEMS Then
        Call HackingAttempt(Index, "Invalid InvNum")
        Exit Sub
    End If
    
    ' Prevent hacking
    If CharNum < 1 Or CharNum > MAX_CHARS Then
        Call HackingAttempt(Index, "Invalid CharNum")
        Exit Sub
    End If
    
    If (GetPlayerInvItemNum(Index, InvNum) > 0) And (GetPlayerInvItemNum(Index, InvNum) <= MAX_ITEMS) Then
        n = Item(GetPlayerInvItemNum(Index, InvNum)).Data2
        
        ' Find out what kind of item it is
        Select Case Item(GetPlayerInvItemNum(Index, InvNum)).Type
            Case ITEM_TYPE_ARMOR
                If InvNum <> GetPlayerEquipmentSlot(Index, Armor) Then
                    If Int(GetPlayerStat(Index, Stats.Defense)) < n Then
                        Call PlayerMsg(Index, "Your defense is to low to wear this armor!  Required DEF (" & n & ")", BrightRed)
                        Exit Sub
                    End If
                    Call SetPlayerEquipmentSlot(Index, InvNum, Armor)
                Else
                    Call SetPlayerEquipmentSlot(Index, 0, Armor)
                End If
                Call SendWornEquipment(Index)
            
            Case ITEM_TYPE_WEAPON
                If InvNum <> GetPlayerEquipmentSlot(Index, Weapon) Then
                    If Int(GetPlayerStat(Index, Stats.Strength)) < n Then
                        Call PlayerMsg(Index, "Your strength is to low to hold this weapon!  Required STR (" & n & ")", BrightRed)
                        Exit Sub
                    End If
                    Call SetPlayerEquipmentSlot(Index, InvNum, Weapon)
                Else
                    Call SetPlayerEquipmentSlot(Index, 0, Weapon)
                End If
                Call SendWornEquipment(Index)
                    
            Case ITEM_TYPE_HELMET
                If InvNum <> GetPlayerEquipmentSlot(Index, Helmet) Then
                    If Int(GetPlayerStat(Index, Stats.Speed)) < n Then
                        Call PlayerMsg(Index, "Your speed coordination is to low to wear this helmet!  Required SPEED (" & n & ")", BrightRed)
                        Exit Sub
                    End If
                    Call SetPlayerEquipmentSlot(Index, InvNum, Helmet)
                Else
                    Call SetPlayerEquipmentSlot(Index, 0, Helmet)
                End If
                Call SendWornEquipment(Index)
        
            Case ITEM_TYPE_SHIELD
                If InvNum <> GetPlayerEquipmentSlot(Index, Shield) Then
                    Call SetPlayerEquipmentSlot(Index, InvNum, Shield)
                Else
                    Call SetPlayerEquipmentSlot(Index, 0, Shield)
                End If
                Call SendWornEquipment(Index)
        
            Case ITEM_TYPE_POTIONADDHP
                Call SetPlayerVital(Index, Vitals.HP, GetPlayerVital(Index, Vitals.HP) + Item(Player(Index).Char(CharNum).Inv(InvNum).Num).Data1)
                Call TakeItem(Index, Player(Index).Char(CharNum).Inv(InvNum).Num, 0)
                Call SendVital(Index, Vitals.HP)
        
            Case ITEM_TYPE_POTIONADDMP
                Call SetPlayerVital(Index, Vitals.MP, GetPlayerVital(Index, Vitals.MP) + Item(Player(Index).Char(CharNum).Inv(InvNum).Num).Data1)
                Call TakeItem(Index, Player(Index).Char(CharNum).Inv(InvNum).Num, 0)
                Call SendVital(Index, Vitals.MP)
    
            Case ITEM_TYPE_POTIONADDSP
                Call SetPlayerVital(Index, Vitals.SP, GetPlayerVital(Index, Vitals.SP) + Item(Player(Index).Char(CharNum).Inv(InvNum).Num).Data1)
                Call TakeItem(Index, Player(Index).Char(CharNum).Inv(InvNum).Num, 0)
                Call SendVital(Index, Vitals.SP)

            Case ITEM_TYPE_POTIONSUBHP
                Call SetPlayerVital(Index, Vitals.HP, GetPlayerVital(Index, Vitals.HP) - Item(Player(Index).Char(CharNum).Inv(InvNum).Num).Data1)
                Call TakeItem(Index, Player(Index).Char(CharNum).Inv(InvNum).Num, 0)
                Call SendVital(Index, Vitals.HP)
            
            Case ITEM_TYPE_POTIONSUBMP
                Call SetPlayerVital(Index, Vitals.MP, GetPlayerVital(Index, Vitals.MP) - Item(Player(Index).Char(CharNum).Inv(InvNum).Num).Data1)
                Call TakeItem(Index, Player(Index).Char(CharNum).Inv(InvNum).Num, 0)
                Call SendVital(Index, Vitals.MP)
    
            Case ITEM_TYPE_POTIONSUBSP
                Call SetPlayerVital(Index, Vitals.SP, GetPlayerVital(Index, Vitals.SP) - Item(Player(Index).Char(CharNum).Inv(InvNum).Num).Data1)
                Call TakeItem(Index, Player(Index).Char(CharNum).Inv(InvNum).Num, 0)
                Call SendVital(Index, Vitals.SP)
                
            Case ITEM_TYPE_KEY
                Select Case GetPlayerDir(Index)
                    Case DIR_UP
                        If GetPlayerY(Index) > 0 Then
                            x = GetPlayerX(Index)
                            y = GetPlayerY(Index) - 1
                        Else
                            Exit Sub
                        End If
                        
                    Case DIR_DOWN
                        If GetPlayerY(Index) < MAX_MAPY Then
                            x = GetPlayerX(Index)
                            y = GetPlayerY(Index) + 1
                        Else
                            Exit Sub
                        End If
                            
                    Case DIR_LEFT
                        If GetPlayerX(Index) > 0 Then
                            x = GetPlayerX(Index) - 1
                            y = GetPlayerY(Index)
                        Else
                            Exit Sub
                        End If
                            
                    Case DIR_RIGHT
                        If GetPlayerX(Index) < MAX_MAPX Then
                            x = GetPlayerX(Index) + 1
                            y = GetPlayerY(Index)
                        Else
                            Exit Sub
                        End If
                End Select
                
                ' Check if a key exists
                If Map(GetPlayerMap(Index)).Tile(x, y).Type = TILE_TYPE_KEY Then
                    ' Check if the key they are using matches the map key
                    If GetPlayerInvItemNum(Index, InvNum) = Map(GetPlayerMap(Index)).Tile(x, y).Data1 Then
                        TempTile(GetPlayerMap(Index)).DoorOpen(x, y) = YES
                        TempTile(GetPlayerMap(Index)).DoorTimer = GetTickCount
                        
                        Call SendDataToMap(GetPlayerMap(Index), SMapKey & SEP_CHAR & x & SEP_CHAR & y & SEP_CHAR & 1 & END_CHAR)
                        Call MapMsg(GetPlayerMap(Index), "A porta foi aberta.", Red)
                        
                        ' Check if we are supposed to take away the item
                        If Map(GetPlayerMap(Index)).Tile(x, y).Data2 = 1 Then
                            Call TakeItem(Index, GetPlayerInvItemNum(Index, InvNum), 0)
                            Call PlayerMsg(Index, "A chave quebrou.", Yellow)
                        End If
                    End If
                End If
                
            Case ITEM_TYPE_SPELL
                ' Get the spell num
                n = Item(GetPlayerInvItemNum(Index, InvNum)).Data1
                
                If n > 0 Then
                    ' Make sure they are the right class
                    If Spell(n).ClassReq - 1 = GetPlayerClass(Index) Or Spell(n).ClassReq = 0 Then
                        ' Make sure they are the right level
                        i = GetSpellReqLevel(n)
                        If i <= GetPlayerLevel(Index) Then
                            i = FindOpenSpellSlot(Index)
                            
                            ' Make sure they have an open spell slot
                            If i > 0 Then
                                ' Make sure they dont already have the spell
                                If Not HasSpell(Index, n) Then
                                    Call SetPlayerSpell(Index, i, n)
                                    Call TakeItem(Index, GetPlayerInvItemNum(Index, InvNum), 0)
                                    'Call PlayerMsg(Index, "Você estuda a magia cuidadosamente...", Yellow)
                                    'Call PlayerMsg(Index, "Você aprendeu uma nova magia!", red)
                                Else
                                    Call TakeItem(Index, GetPlayerInvItemNum(Index, InvNum), 0)
                                    'Call PlayerMsg(Index, "Você já aprendeu essa magia!.", BrightRed)
                                End If
                            Else
                                'Call PlayerMsg(Index, "Você já aprendeu tudo o que podia!", BrightRed)
                            End If
                        Else
                            'Call PlayerMsg(Index, "Você precisa ser nível " & I & " para aprender essa magia.", White)
                        End If
                    Else
                       ' Call PlayerMsg(Index, "Essa magia só pode ser aprendida por " & GetClassName(Spell(n).ClassReq - 1) & ".", White)
                    End If
                Else
                    'Call PlayerMsg(Index, "Esse pergaminho não funciona!", Red)
                End If
                
        End Select
    End If
End Sub
' ::::::::::::::::::::::::::
' :: Player attack packet ::
' ::::::::::::::::::::::::::
Sub HandleAttack(ByVal Index As Long)
'On Error Resume Next
'Dim i As Long, n As Long
'Dim Damage As Long
'Dim TempIndex As Long

    If Not Bomb(GetPlayerX(Index), GetPlayerY(Index)).Here Then
        If Not TempPlayer(Index).Dieing Then
            If Not TempPlayer(Index).TotalBombs + 1 > TempPlayer(Index).Bombs Then
            
                TempPlayer(Index).TotalBombs = TempPlayer(Index).TotalBombs + 1
                
                With Bomb(GetPlayerX(Index), GetPlayerY(Index))
                    .Maker = GetPlayerName(Index)
                    .MakerIndex = Index
                    .Timer = GetTickCount + 2000
                    .Map = GetPlayerMap(Index)
                    .Distance = TempPlayer(Index).Distance
                    .Here = True
                End With
                
                SendDataToMap GetPlayerMap(Index), SBomb & SEP_CHAR & 0 & SEP_CHAR & GetPlayerX(Index) & SEP_CHAR & GetPlayerY(Index) & END_CHAR
            
            End If
        End If
    End If
End Sub
Sub HandleThrow(ByVal Index As Long)
Dim x As Long
Dim y As Long
On Error Resume Next
If TempPlayer(Index).Throw = False Then Exit Sub
x = GetPlayerX(Index)
y = GetPlayerY(Index)

        Select Case GetPlayerDir(Index)
       Case DIR_UP
       If Bomb(GetPlayerX(Index), GetPlayerY(Index) - 1).Here And Not Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) - 2).Type = TILE_TYPE_WALL Then
        Bomb(x, y - 1).Here = False
        Bomb(x, y - 2).Maker = GetPlayerName(Index)
        Bomb(x, y - 2).MakerIndex = Index
        Bomb(x, y - 2).Timer = GetTickCount + 2000
        Bomb(x, y - 2).Map = GetPlayerMap(Index)
        Bomb(x, y - 2).Distance = TempPlayer(Index).Distance
        Bomb(x, y - 2).Here = True
        SendDataToMap GetPlayerMap(Index), SBomb & SEP_CHAR & 1 & SEP_CHAR & GetPlayerX(Index) & SEP_CHAR & GetPlayerY(Index) - 1 & END_CHAR
        SendDataToMap GetPlayerMap(Index), SBomb & SEP_CHAR & 0 & SEP_CHAR & GetPlayerX(Index) & SEP_CHAR & GetPlayerY(Index) - 2 & END_CHAR
       End If
        Case DIR_DOWN
       If Bomb(x, y + 1).Here Then
        If Bomb(GetPlayerX(Index), GetPlayerY(Index) + 1).Here And Not Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) + 2).Type = TILE_TYPE_WALL Then
        Bomb(x, y + 1).Here = False
        Bomb(x, y + 2).Maker = GetPlayerName(Index)
        Bomb(x, y + 2).MakerIndex = Index
        Bomb(x, y + 2).Timer = GetTickCount + 2000
        Bomb(x, y + 2).Map = GetPlayerMap(Index)
        Bomb(x, y + 2).Distance = TempPlayer(Index).Distance
        Bomb(x, y + 2).Here = True
        SendDataToMap GetPlayerMap(Index), SBomb & SEP_CHAR & 1 & SEP_CHAR & GetPlayerX(Index) & SEP_CHAR & GetPlayerY(Index) + 1 & END_CHAR
        SendDataToMap GetPlayerMap(Index), SBomb & SEP_CHAR & 0 & SEP_CHAR & GetPlayerX(Index) & SEP_CHAR & GetPlayerY(Index) + 2 & END_CHAR
       End If
       End If
        Case DIR_LEFT
        If Bomb(GetPlayerX(Index) - 1, GetPlayerY(Index)).Here And Not Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) - 2, GetPlayerY(Index)).Type = TILE_TYPE_WALL Then
        Bomb(x - 1, y).Here = False
        Bomb(x - 2, y).Maker = GetPlayerName(Index)
        Bomb(x - 2, y).MakerIndex = Index
        Bomb(x - 2, y).Timer = GetTickCount + 2000
        Bomb(x - 2, y).Map = GetPlayerMap(Index)
        Bomb(x - 2, y).Distance = TempPlayer(Index).Distance
        Bomb(x - 2, y).Here = True
        SendDataToMap GetPlayerMap(Index), SBomb & SEP_CHAR & 1 & SEP_CHAR & GetPlayerX(Index) - 1 & SEP_CHAR & GetPlayerY(Index) & END_CHAR
        SendDataToMap GetPlayerMap(Index), SBomb & SEP_CHAR & 0 & SEP_CHAR & GetPlayerX(Index) - 2 & SEP_CHAR & GetPlayerY(Index) & END_CHAR
       End If
        Case DIR_RIGHT
        If Bomb(GetPlayerX(Index) + 1, GetPlayerY(Index)).Here And Not Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) + 2, GetPlayerY(Index)).Type = TILE_TYPE_WALL Then
        Bomb(x + 1, y).Here = False
        Bomb(x + 2, y).Maker = GetPlayerName(Index)
        Bomb(x + 2, y).MakerIndex = Index
        Bomb(x + 2, y).Timer = GetTickCount + 2000
        Bomb(x + 2, y).Map = GetPlayerMap(Index)
        Bomb(x + 2, y).Distance = TempPlayer(Index).Distance
        Bomb(x + 2, y).Here = True
        SendDataToMap GetPlayerMap(Index), SBomb & SEP_CHAR & 1 & SEP_CHAR & GetPlayerX(Index) + 1 & SEP_CHAR & GetPlayerY(Index) & END_CHAR
        SendDataToMap GetPlayerMap(Index), SBomb & SEP_CHAR & 0 & SEP_CHAR & GetPlayerX(Index) + 2 & SEP_CHAR & GetPlayerY(Index) & END_CHAR
       End If
    End Select

End Sub
' ::::::::::::::::::::::
' :: Use stats packet ::
' ::::::::::::::::::::::
Sub HandleUseStatPoint(ByVal Index As Long, ByRef Parse() As String)
Dim PointType As Long

    PointType = Val(Parse(1))
    
    ' Prevent hacking
    If (PointType < 0) Or (PointType > 3) Then
        Call HackingAttempt(Index, "Invalid Point Type")
        Exit Sub
    End If
            
    ' Make sure they have points
    If GetPlayerPOINTS(Index) > 0 Then
        ' Take away a stat point
        Call SetPlayerPOINTS(Index, GetPlayerPOINTS(Index) - 1)
        
        ' Everything is ok
        Select Case PointType
            Case 0
                Call SetPlayerStat(Index, Stats.Strength, GetPlayerStat(Index, Stats.Strength) + 1)
                Call PlayerMsg(Index, "Você ganhou força!", Red)
            Case 1
                Call SetPlayerStat(Index, Stats.Defense, GetPlayerStat(Index, Stats.Defense) + 1)
                Call PlayerMsg(Index, "Você ganhou defesa!", Red)
            Case 2
                Call SetPlayerStat(Index, Stats.Magic, GetPlayerStat(Index, Stats.Magic) + 1)
                Call PlayerMsg(Index, "Você ganhou inteligência!", Red)
            Case 3
                Call SetPlayerStat(Index, Stats.Speed, GetPlayerStat(Index, Stats.Speed) + 1)
                Call PlayerMsg(Index, "Você ganhou velocidade!", Red)
        End Select
    Else
        Call PlayerMsg(Index, "Você não tem mais pontos!", BrightRed)
    End If
    
    ' Send the update
    Call SendStats(Index)
End Sub

' ::::::::::::::::::::::::::::::::
' :: Player info request packet ::
' ::::::::::::::::::::::::::::::::
Sub HandlePlayerInfoRequest(ByVal Index As Long, ByRef Parse() As String)
Dim Name As String
Dim i As Long, n As Long

    Name = Parse(1)
    
    i = FindPlayer(Name)
    If i > 0 Then
        'Call PlayerMsg(Index, "Account: " & Trim$(Player(I).Login) & ", Name: " & GetPlayerName(I), BrightGreen)
        If GetPlayerAccess(Index) > ADMIN_MONITOR Then
            'Call PlayerMsg(Index, "-=- Stats for " & GetPlayerName(I) & " -=-", BrightGreen)
            'Call PlayerMsg(Index, "Nível: " & GetPlayerLevel(I) & "  Exp: " & GetPlayerExp(I) & "/" & GetPlayerNextLevel(I), BrightGreen)
            'Call PlayerMsg(Index, "HP: " & GetPlayerVital(I, Vitals.HP) & "/" & GetPlayerMaxVital(I, Vitals.HP) & "  MP: " & GetPlayerVital(I, Vitals.MP) & "/" & GetPlayerMaxVital(I, Vitals.MP) & "  SP: " & GetPlayerVital(I, Vitals.SP) & "/" & GetPlayerMaxVital(I, Vitals.SP), BrightGreen)
            'Call PlayerMsg(Index, "Força: " & GetPlayerStat(I, Stats.Strength) & "  Defesa: " & GetPlayerStat(I, Stats.Defense) & "  Magia: " & GetPlayerStat(I, Stats.Magic) & "  Velocidade: " & GetPlayerStat(I, Stats.Speed), BrightGreen)
            n = Int(GetPlayerStat(i, Stats.Strength) / 2) + Int(GetPlayerLevel(i) / 2)
            i = Int(GetPlayerStat(i, Stats.Defense) / 2) + Int(GetPlayerLevel(i) / 2)
            If n > 100 Then n = 100
            If i > 100 Then i = 100
            'Call PlayerMsg(Index, "Chance de crítico: " & n & "%, Block Chance: " & I & "%", BrightGreen)
        End If
    Else
        'Call PlayerMsg(Index, "Jogador não está online.", Red)
    End If
End Sub
' :::::::::::::::::::::::
' :: Warp me to packet ::
' :::::::::::::::::::::::
Sub HandleWarpMeTo(ByVal Index As Long, ByRef Parse() As String)
Dim n As Long

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_MAPPER Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If
    
    ' The player
    n = FindPlayer(Parse(1))
    
    If n <> Index Then
        If n > 0 Then
            Call PlayerWarp(Index, GetPlayerMap(n), GetPlayerX(n), GetPlayerY(n))
            Call PlayerMsg(n, GetPlayerName(Index) & " Se teleportou até você.", BrightBlue)
            Call PlayerMsg(Index, "Você foi invocador por " & GetPlayerName(n) & ".", BrightBlue)
            Call AddLog(GetPlayerName(Index) & " foi teleportado por " & GetPlayerName(n) & ", map #" & GetPlayerMap(n) & ".", ADMIN_LOG)
        Else
            Call PlayerMsg(Index, "Jogador não está online.", Red)
        End If
    Else
        Call PlayerMsg(Index, "Você não pode se teleportar para você mesmo!", Red)
    End If
End Sub
' :::::::::::::::::::::::
' :: Warp to me packet ::
' :::::::::::::::::::::::
Sub HandleWarpToMe(ByVal Index As Long, ByRef Parse() As String)
Dim n As Long

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_MAPPER Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If
    
    ' The player
    n = FindPlayer(Parse(1))
    
    If n <> Index Then
        If n > 0 Then
            Call PlayerWarp(n, GetPlayerMap(Index), GetPlayerX(Index), GetPlayerY(Index))
            Call PlayerMsg(n, "Você foi invocado por " & GetPlayerName(Index) & ".", BrightBlue)
            Call PlayerMsg(Index, GetPlayerName(n) & " foi invocado.", BrightBlue)
            Call AddLog(GetPlayerName(Index) & " teleportou " & GetPlayerName(n) & " para ele mesmo, map #" & GetPlayerMap(Index) & ".", ADMIN_LOG)
        Else
            Call PlayerMsg(Index, "Jogador não está online.", Red)
        End If
    Else
        Call PlayerMsg(Index, "Você não pode se teleportar para você mesmo!", Red)
    End If
End Sub
' ::::::::::::::::::::::::
' :: Warp to map packet ::
' ::::::::::::::::::::::::
Sub HandleWarpTo(ByVal Index As Long, ByRef Parse() As String)
Dim n As Long

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_MAPPER Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If
    
    ' The map
    n = Val(Parse(1))
    
    ' Prevent hacking
    If n < 0 Or n > MAX_MAPS Then
        Call HackingAttempt(Index, "Invalid map")
        Exit Sub
    End If
    
    Call PlayerWarp(Index, n, GetPlayerX(Index), GetPlayerY(Index))
    Call PlayerMsg(Index, "You have been warped to map #" & n, BrightBlue)
    Call AddLog(GetPlayerName(Index) & " warped to map #" & n & ".", ADMIN_LOG)
End Sub
' :::::::::::::::::::::::
' :: Set sprite packet ::
' :::::::::::::::::::::::
Sub HandleSetSprite(ByVal Index As Long, ByRef Parse() As String)
Dim n As Long

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_MAPPER Then
        Call HackingAttempt(Index, "Clonagem")
        Exit Sub
    End If
    
    ' The sprite
    n = Val(Parse(1))
    
    Call SetPlayerSprite(Index, n)
    Call SendPlayerData(Index)
    Exit Sub
End Sub
' ::::::::::::::::::::::::::
' :: Stats request packet ::
' ::::::::::::::::::::::::::
Sub HandleGetStats(ByVal Index As Long)
Dim i As Long, n As Long

    'Call PlayerMsg(Index, "-=- Stats for " & GetPlayerName(Index) & " -=-", Red)
    'Call PlayerMsg(Index, "Nível: " & GetPlayerLevel(Index) & "  Exp: " & GetPlayerExp(Index) & "/" & GetPlayerNextLevel(Index), White)
    'Call PlayerMsg(Index, "HP: " & GetPlayerVital(Index, Vitals.HP) & "/" & GetPlayerMaxVital(Index, Vitals.HP) & "  MP: " & GetPlayerVital(Index, Vitals.MP) & "/" & GetPlayerMaxVital(Index, Vitals.MP) & "  SP: " & GetPlayerVital(Index, Vitals.SP) & "/" & GetPlayerMaxVital(Index, Vitals.SP), White)
    'Call PlayerMsg(Index, "Força: " & GetPlayerStat(Index, Stats.Strength) & "  Defesa: " & GetPlayerStat(Index, Stats.Defense) & "  Magia: " & GetPlayerStat(Index, Stats.Magic) & "  Velocidade: " & GetPlayerStat(Index, Stats.Speed), Red)
    n = Int(GetPlayerStat(Index, Stats.Strength) / 2) + Int(GetPlayerLevel(Index) / 2)
    i = Int(GetPlayerStat(Index, Stats.Defense) / 2) + Int(GetPlayerLevel(Index) / 2)
    If n > 100 Then n = 100
    If i > 100 Then i = 100
    'Call PlayerMsg(Index, "Chance de crítico: " & n & "%, Block Chance: " & I & "%", Red)
End Sub
' ::::::::::::::::::::::::::::::::::
' :: Player request for a new map ::
' ::::::::::::::::::::::::::::::::::
Sub HandleRequestNewMap(ByVal Index As Long, ByRef Parse() As String)
Dim Dir As Long

    Dir = Val(Parse(1))
    
    ' Prevent hacking
    If Dir < DIR_UP Or Dir > DIR_RIGHT Then
        Call HackingAttempt(Index, "Invalid Direction")
        Exit Sub
    End If
            
    Call PlayerMove(Index, Dir, 1)
End Sub
' :::::::::::::::::::::
' :: Map data packet ::
' :::::::::::::::::::::
Sub HandleMapData(ByVal Index As Long, ByRef Parse() As String)
Dim n As Long, i As Long, MapNum As Long
Dim x As Long, y As Long

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_MAPPER Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If

    n = 0
    
    MapNum = GetPlayerMap(Index)
    
    i = Map(MapNum).Revision + 1
    
    Call ClearMap(MapNum)
    
    MapNum = GetPlayerMap(Index)
    Map(MapNum).Name = Parse(n + 1)
    Map(MapNum).PlayerLimit = Parse(n + 2)
    Map(MapNum).Revision = i
    Map(MapNum).Moral = Val(Parse(n + 3))
    Map(MapNum).TileSet = Val(Parse(n + 4))
    Map(MapNum).Up = Val(Parse(n + 5))
    Map(MapNum).Down = Val(Parse(n + 6))
    Map(MapNum).Left = Val(Parse(n + 7))
    Map(MapNum).Right = Val(Parse(n + 8))
    Map(MapNum).Music = Val(Parse(n + 9))
    Map(MapNum).BootMap = Val(Parse(n + 10))
    Map(MapNum).BootX = Val(Parse(n + 11))
    Map(MapNum).BootY = Val(Parse(n + 12))
    Map(MapNum).Shop = Val(Parse(n + 13))
    
    n = n + 14
    
    For x = 0 To MAX_MAPX
        For y = 0 To MAX_MAPY
            Map(MapNum).Tile(x, y).Ground = Val(Parse(n))
            Map(MapNum).Tile(x, y).Mask = Val(Parse(n + 1))
            Map(MapNum).Tile(x, y).Anim = Val(Parse(n + 2))
            Map(MapNum).Tile(x, y).Fringe = Val(Parse(n + 3))
            Map(MapNum).Tile(x, y).Type = Val(Parse(n + 4))
            Map(MapNum).Tile(x, y).Data1 = Val(Parse(n + 5))
            Map(MapNum).Tile(x, y).Data2 = Val(Parse(n + 6))
            Map(MapNum).Tile(x, y).Data3 = Val(Parse(n + 7))
            
            n = n + 8
        Next
    Next
    
    For x = 1 To MAX_MAP_NPCS
        Map(MapNum).Npc(x) = Val(Parse(n))
        n = n + 1
        Call ClearMapNpc(x, MapNum)
    Next
    
    Call SendMapNpcsToMap(MapNum)
    Call SpawnMapNpcs(MapNum)
    
    ' Clear out it all
    For i = 1 To MAX_MAP_ITEMS
        Call SpawnItemSlot(i, 0, 0, 0, GetPlayerMap(Index), MapItem(GetPlayerMap(Index), i).x, MapItem(GetPlayerMap(Index), i).y)
        Call ClearMapItem(i, GetPlayerMap(Index))
    Next
    
    ' Respawn
    Call SpawnMapItems(GetPlayerMap(Index))
    
    ' Save the map
    Call SaveMap(MapNum)
    
    Call MapCache_Create(MapNum)
    
    ' Refresh map for everyone online
    Dim e As Long
    For i = 1 To Player_HighIndex
        e = i
        If IsPlaying(e) And GetPlayerMap(e) = MapNum Then
            Call PlayerWarp(e, MapNum, GetPlayerX(e), GetPlayerY(e))
        End If
    Next
    
    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            HandleRequestRoomList i
        End If
    Next
    
End Sub
' ::::::::::::::::::::::::::::
' :: Need map yes/no packet ::
' ::::::::::::::::::::::::::::
Sub HandleNeedMap(ByVal Index As Long, ByRef Parse() As String)
Dim s As String

    ' Get yes/no value
    s = Parse(1)
            
    ' Check if map data is needed to be sent
    If s = "yes" Then
        Call SendMap(Index, GetPlayerMap(Index))
    End If
    
    Call SendMapItemsTo(Index, GetPlayerMap(Index))
    Call SendMapNpcsTo(Index, GetPlayerMap(Index))
    Call SendJoinMap(Index)
    TempPlayer(Index).GettingMap = NO
    Call SendDataTo(Index, SMapDone & END_CHAR)
End Sub
' :::::::::::::::::::::::::::::::::::::::::::::::
' :: Player trying to pick up something packet ::
' :::::::::::::::::::::::::::::::::::::::::::::::
Sub HandleMapGetItem(ByVal Index As Long)
    Call PlayerMapGetItem(Index)
End Sub
' ::::::::::::::::::::::::::::::::::::::::::::
' :: Player trying to drop something packet ::
' ::::::::::::::::::::::::::::::::::::::::::::
Sub HandleMapDropItem(ByVal Index As Long, ByRef Parse() As String)
Dim InvNum As Long, Ammount As Long

    InvNum = Val(Parse(1))
    Ammount = Val(Parse(2))
    
    ' Prevent hacking
    If InvNum < 1 Or InvNum > MAX_INV Then
        Call HackingAttempt(Index, "Invalid InvNum")
        Exit Sub
    End If
    
    ' Prevent hacking
    If Ammount > GetPlayerInvItemValue(Index, InvNum) Then
        Call HackingAttempt(Index, "Item ammount modification")
        Exit Sub
    End If
    
    ' Prevent hacking
    If Item(GetPlayerInvItemNum(Index, InvNum)).Type = ITEM_TYPE_CURRENCY Then
        ' Check if money and if it is we want to make sure that they aren't trying to drop 0 value
        If Ammount <= 0 Then
            Call HackingAttempt(Index, "Trying to drop 0 ammount of currency")
            Exit Sub
        End If
    End If
        
    Call PlayerMapDropItem(Index, InvNum, Ammount)
End Sub
' ::::::::::::::::::::::::
' :: Respawn map packet ::
' ::::::::::::::::::::::::
Sub HandleMapRespawn(ByVal Index As Long)
Dim i As Long

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_MAPPER Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If
    
    ' Clear out it all
    For i = 1 To MAX_MAP_ITEMS
        Call SpawnItemSlot(i, 0, 0, 0, GetPlayerMap(Index), MapItem(GetPlayerMap(Index), i).x, MapItem(GetPlayerMap(Index), i).y)
        Call ClearMapItem(i, GetPlayerMap(Index))
    Next
    
    ' Respawn
    Call SpawnMapItems(GetPlayerMap(Index))
    
    ' Respawn NPCS
    For i = 1 To MAX_MAP_NPCS
        Call SpawnNpc(i, GetPlayerMap(Index))
    Next
    
    Call PlayerMsg(Index, "Map respawned.", Blue)
    Call AddLog(GetPlayerName(Index) & " has respawned map #" & GetPlayerMap(Index), ADMIN_LOG)
End Sub
' :::::::::::::::::::::::
' :: Map report packet ::
' :::::::::::::::::::::::
Sub HandleMapReport(ByVal Index As Long)
Dim s As String
Dim i As Long
Dim tMapStart As Long, tMapEnd As Long

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_MAPPER Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If

    s = "Free Maps: "
    tMapStart = 1
    tMapEnd = 1
    
    For i = 1 To MAX_MAPS
        If LenB(Trim$(Map(i).Name)) = 0 Then
            tMapEnd = tMapEnd + 1
        Else
            If tMapEnd - tMapStart > 0 Then
                s = s & Trim$(CStr(tMapStart)) & "-" & Trim$(CStr(tMapEnd - 1)) & ", "
            End If
            tMapStart = i + 1
            tMapEnd = i + 1
        End If
    Next
    
    s = s & Trim$(CStr(tMapStart)) & "-" & Trim$(CStr(tMapEnd - 1)) & ", "
    s = Mid$(s, 1, Len(s) - 2)
    s = s & "."
    
    Call PlayerMsg(Index, s, Brown)
End Sub
' ::::::::::::::::::::::::
' :: Kick player packet ::
' ::::::::::::::::::::::::
Sub HandleKickPlayer(ByVal Index As Long, ByRef Parse() As String)
Dim n As Long

    ' Prevent hacking
    If GetPlayerAccess(Index) <= 0 Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If
    
    ' The player index
    n = FindPlayer(Parse(1))
    
    If n <> Index Then
        If n > 0 Then
            If GetPlayerAccess(n) < GetPlayerAccess(Index) Then
                Call GlobalMsg(GetPlayerName(n) & " foi retirado do " & GAME_NAME & " por " & GetPlayerName(Index) & "!", Red)
                Call AddLog(GetPlayerName(Index) & " retirou " & GetPlayerName(n) & ".", ADMIN_LOG)
                Call AlertMsg(n, "Você foi retirado por " & GetPlayerName(Index) & "!")
            Else
                Call PlayerMsg(Index, "????", Red)
            End If
        Else
            Call PlayerMsg(Index, "Jogador não está online.", Red)
        End If
    Else
        Call PlayerMsg(Index, "Você não pode retirar você mesmo!", Red)
    End If
End Sub
' :::::::::::::::::::::
' :: Ban list packet ::
' :::::::::::::::::::::
Sub HandleBanList(ByVal Index As Long)
Dim n As Long, f As Long
Dim s As String, Name As String

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_MAPPER Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If
    
    n = 1
    f = FreeFile
    Open App.Path & "\data\banlist.txt" For Input As #f
    Do While Not EOF(f)
        Input #f, s
        Input #f, Name
        
        Call PlayerMsg(Index, n & ": Banned IP " & s & " by " & Name, Red)
        n = n + 1
    Loop
    Close #f
End Sub
' ::::::::::::::::::::::::
' :: Ban destroy packet ::
' ::::::::::::::::::::::::
Sub HandleBanDestroy(ByVal Index As Long)
Dim FileName As String
Dim File As Long

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_CREATOR Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If
    
    FileName = App.Path & "\data\banlist.txt"

    Kill FileName

    Call PlayerMsg(Index, "Ban list destruída.", Red)
    
End Sub
' :::::::::::::::::::::::
' :: Ban player packet ::
' :::::::::::::::::::::::
Sub HandleBanPlayer(ByVal Index As Long, ByRef Parse() As String)
Dim n As Long

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_MAPPER Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If
    
    ' The player index
    n = FindPlayer(Parse(1))
    
    If n <> Index Then
        If n > 0 Then
            If GetPlayerAccess(n) < GetPlayerAccess(Index) Then
                Call BanIndex(n, Index)
            Else
                Call PlayerMsg(Index, "????", Red)
            End If
        Else
            Call PlayerMsg(Index, "Jogador não está online.", Red)
        End If
    Else
        Call PlayerMsg(Index, "Você não pode banir você mesmo!", Red)
    End If
End Sub
' :::::::::::::::::::::::::::::
' :: Request edit map packet ::
' :::::::::::::::::::::::::::::
Sub HandleRequestEditMap(ByVal Index As Long)
    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_MAPPER Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If
    
    Call SendDataTo(Index, SEditMap & END_CHAR)
End Sub
' ::::::::::::::::::::::::::::::
' :: Request edit item packet ::
' ::::::::::::::::::::::::::::::
Sub HandleRequestEditItem(ByVal Index As Long)
    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If
    
    Call SendDataTo(Index, SItemEditor & END_CHAR)
End Sub
' ::::::::::::::::::::::
' :: Edit item packet ::
' ::::::::::::::::::::::
Sub HandleEditItem(ByVal Index As Long, ByRef Parse() As String)
Dim n As Long

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If
    
    ' The item #
    n = Val(Parse(1))
    
    ' Prevent hacking
    If n < 0 Or n > MAX_ITEMS Then
        Call HackingAttempt(Index, "Invalid Item Index")
        Exit Sub
    End If
    
    Call AddLog(GetPlayerName(Index) & " editing item #" & n & ".", ADMIN_LOG)
    Call SendEditItemTo(Index, n)
End Sub

Sub HandleDelete(ByVal Index As Long, ByRef Parse() As String)
    Dim Editor As Byte
    Dim EditorIndex As Long

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If

    Editor = CByte(Parse(1))
    EditorIndex = CLng(Parse(2))

    Select Case Editor
    
        Case EDITOR_ITEM
            ' Prevent hacking
            If EditorIndex < 1 Or EditorIndex > MAX_ITEMS Then
                Call HackingAttempt(Index, "Invalid Item Index")
                Exit Sub
            End If
            
            Call ClearItem(EditorIndex)
            
            Call SendUpdateItemToAll(EditorIndex)
            Call SaveItem(EditorIndex)
            Call AddLog(GetPlayerName(Index) & "Deleted item #" & EditorIndex & ".", ADMIN_LOG)
        
        Case EDITOR_NPC
            ' Prevent hacking
            If EditorIndex < 1 Or EditorIndex > MAX_NPCS Then
                Call HackingAttempt(Index, "Invalid NPC Index")
                Exit Sub
            End If
        
            Call ClearNpc(EditorIndex)
        
            Call SendUpdateNpcToAll(EditorIndex)
            Call SaveNpc(EditorIndex)
            Call AddLog(GetPlayerName(Index) & "Deleted npc #" & EditorIndex & ".", ADMIN_LOG)
        
        Case EDITOR_SPELL
            ' Prevent hacking
            If EditorIndex < 1 Or EditorIndex > MAX_SPELLS Then
                Call HackingAttempt(Index, "Invalid Spell Index")
                Exit Sub
            End If
        
            Call ClearSpell(EditorIndex)
            
            Call SendUpdateSpellToAll(EditorIndex)
            Call SaveSpell(EditorIndex)
            Call AddLog(GetPlayerName(Index) & "Deleted spell #" & EditorIndex & ".", ADMIN_LOG)
        
        Case EDITOR_SHOP
            ' Prevent hacking
            If EditorIndex < 1 Or EditorIndex > MAX_SHOPS Then
                Call HackingAttempt(Index, "Invalid Shop Index")
                Exit Sub
            End If
            
            Call ClearShop(EditorIndex)
            
            Call SendUpdateShopToAll(EditorIndex)
            Call SaveShop(EditorIndex)
            Call AddLog(GetPlayerName(Index) & "Deleted shop #" & EditorIndex & ".", ADMIN_LOG)
    End Select
    
    Call SendDataTo(Index, SREditor & END_CHAR)

End Sub


' ::::::::::::::::::::::
' :: Save item packet ::
' ::::::::::::::::::::::
Sub HandleSaveItem(ByVal Index As Long, ByRef Parse() As String)
Dim n As Long

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(Index, "Clonagem de acesso")
        Exit Sub
    End If
    
    n = Val(Parse(1))
    If n < 0 Or n > MAX_ITEMS Then
        Call HackingAttempt(Index, "Valor invalido")
        Exit Sub
    End If
    
    ' Update the item
    Item(n).Name = Parse(2)
    Item(n).Pic = Val(Parse(3))
    Item(n).Type = Val(Parse(4))
    Item(n).Data1 = Val(Parse(5))
    Item(n).Data2 = Val(Parse(6))
    Item(n).Data3 = Val(Parse(7))
    
    ' Save it
    Call SendUpdateItemToAll(n)
    Call SaveItem(n)
    Call AddLog(GetPlayerName(Index) & " saved item #" & n & ".", ADMIN_LOG)
End Sub
' :::::::::::::::::::::::::::::
' :: Request edit npc packet ::
' :::::::::::::::::::::::::::::
Sub HandleRequestEditNpc(ByVal Index As Long)
    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If
    
    Call SendDataTo(Index, SNpcEditor & END_CHAR)
End Sub
' :::::::::::::::::::::
' :: Edit npc packet ::
' :::::::::::::::::::::
Sub HandleEditNpc(ByVal Index As Long, ByRef Parse() As String)
Dim n As Long

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If
    
    ' The npc #
    n = Val(Parse(1))
    
    ' Prevent hacking
    If n < 0 Or n > MAX_NPCS Then
        Call HackingAttempt(Index, "Invalid NPC Index")
        Exit Sub
    End If
    
    Call AddLog(GetPlayerName(Index) & " editing npc #" & n & ".", ADMIN_LOG)
    Call SendEditNpcTo(Index, n)
End Sub
' :::::::::::::::::::::
' :: Save npc packet ::
' :::::::::::::::::::::
Sub HandleSaveNpc(ByVal Index As Long, ByRef Parse() As String)
Dim n As Long

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If
    
    n = Val(Parse(1))
    
    ' Prevent hacking
    If n < 0 Or n > MAX_NPCS Then
        Call HackingAttempt(Index, "Invalid NPC Index")
        Exit Sub
    End If
    
    ' Update the npc
    Npc(n).Name = Parse(2)
    Npc(n).AttackSay = Parse(3)
    Npc(n).Sprite = Val(Parse(4))
    Npc(n).SpawnSecs = Val(Parse(5))
    Npc(n).Behavior = Val(Parse(6))
    Npc(n).Range = Val(Parse(7))
    Npc(n).DropChance = Val(Parse(8))
    Npc(n).DropItem = Val(Parse(9))
    Npc(n).DropItemValue = Val(Parse(10))
    
    Npc(n).Stat(Stats.Strength) = Val(Parse(11))
    Npc(n).Stat(Stats.Defense) = Val(Parse(12))
    Npc(n).Stat(Stats.Speed) = Val(Parse(13))
    Npc(n).Stat(Stats.Magic) = Val(Parse(14))
    
    ' Save it
    Call SendUpdateNpcToAll(n)
    Call SaveNpc(n)
    Call AddLog(GetPlayerName(Index) & " saved npc #" & n & ".", ADMIN_LOG)
End Sub
' ::::::::::::::::::::::::::::::
' :: Request edit shop packet ::
' ::::::::::::::::::::::::::::::
Sub HandleRequestEditShop(ByVal Index As Long)
    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If
    
    Call SendDataTo(Index, SShopEditor & END_CHAR)
End Sub
' ::::::::::::::::::::::
' :: Edit shop packet ::
' ::::::::::::::::::::::
Sub HandleEditShop(ByVal Index As Long, ByRef Parse() As String)
Dim n As Long

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If
    
    ' The shop #
    n = Val(Parse(1))
    
    ' Prevent hacking
    If n < 0 Or n > MAX_SHOPS Then
        Call HackingAttempt(Index, "Invalid Shop Index")
        Exit Sub
    End If
    
    Call AddLog(GetPlayerName(Index) & " editing shop #" & n & ".", ADMIN_LOG)
    Call SendEditShopTo(Index, n)
End Sub
' ::::::::::::::::::::::
' :: Save shop packet ::
' ::::::::::::::::::::::
Sub HandleSaveShop(ByVal Index As Long, ByRef Parse() As String)
Dim ShopNum As Long, n As Long, i As Long

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If
    
    ShopNum = Val(Parse(1))
    
    ' Prevent hacking
    If ShopNum < 0 Or ShopNum > MAX_SHOPS Then
        Call HackingAttempt(Index, "Invalid Shop Index")
        Exit Sub
    End If
    
    ' Update the shop
    Shop(ShopNum).Name = Parse(2)
    Shop(ShopNum).JoinSay = Parse(3)
    Shop(ShopNum).LeaveSay = Parse(4)
    Shop(ShopNum).FixesItems = Val(Parse(5))
    
    n = 6
    For i = 1 To MAX_TRADES
        Shop(ShopNum).TradeItem(i).GiveItem = Val(Parse(n))
        Shop(ShopNum).TradeItem(i).GiveValue = Val(Parse(n + 1))
        Shop(ShopNum).TradeItem(i).GetItem = Val(Parse(n + 2))
        Shop(ShopNum).TradeItem(i).GetValue = Val(Parse(n + 3))
        n = n + 4
    Next
    
    ' Save it
    Call SendUpdateShopToAll(ShopNum)
    Call SaveShop(ShopNum)
    Call AddLog(GetPlayerName(Index) & " salvando shop #" & ShopNum & ".", ADMIN_LOG)
End Sub
' :::::::::::::::::::::::::::::::
' :: Request edit spell packet ::
' :::::::::::::::::::::::::::::::
Sub HandleRequestEditSpell(ByVal Index As Long)
    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If
    
    Call SendDataTo(Index, SSpellEditor & END_CHAR)
End Sub
' :::::::::::::::::::::::
' :: Edit spell packet ::
' :::::::::::::::::::::::
Sub HandleEditSpell(ByVal Index As Long, ByRef Parse() As String)
Dim n As Long

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If
    
    ' The spell #
    n = Val(Parse(1))
    
    ' Prevent hacking
    If n < 0 Or n > MAX_SPELLS Then
        Call HackingAttempt(Index, "Invalid Spell Index")
        Exit Sub
    End If
    
    Call AddLog(GetPlayerName(Index) & " editing spell #" & n & ".", ADMIN_LOG)
    Call SendEditSpellTo(Index, n)
End Sub
' :::::::::::::::::::::::
' :: Save spell packet ::
' :::::::::::::::::::::::
Sub HandleSaveSpell(ByVal Index As Long, ByRef Parse() As String)
Dim n As Long

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If
    
    ' Spell #
    n = Val(Parse(1))
    
    ' Prevent hacking
    If n < 0 Or n > MAX_SPELLS Then
        Call HackingAttempt(Index, "Invalid Spell Index")
        Exit Sub
    End If
    
    ' Update the spell
    Spell(n).Name = Parse(2)
    Spell(n).ClassReq = Val(Parse(3))
    Spell(n).LevelReq = Val(Parse(4))
    Spell(n).Type = Val(Parse(5))
    Spell(n).Data1 = Val(Parse(6))
    Spell(n).Data2 = Val(Parse(7))
    Spell(n).Data3 = Val(Parse(8))
            
    ' Save it
    Call SendUpdateSpellToAll(n)
    Call SaveSpell(n)
    Call AddLog(GetPlayerName(Index) & " salvando spell #" & n & ".", ADMIN_LOG)
End Sub
' :::::::::::::::::::::::
' :: Set access packet ::
' :::::::::::::::::::::::
Sub HandleSetAccess(ByVal Index As Long, ByRef Parse() As String)
Dim n As Long, i As Long

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_CREATOR Then
        Call HackingAttempt(Index, "Trying to use powers not available")
        Exit Sub
    End If
    
    ' The index
    n = FindPlayer(Parse(1))
    ' The access
    i = Val(Parse(2))
    
    ' Check for invalid access level
    If i >= 0 Or i <= 3 Then
        ' Check if player is on
        If n > 0 Then
            If GetPlayerAccess(n) <= 0 Then
                Call GlobalMsg(GetPlayerName(n) & " foi abençoado com os poderes do deus Bomber.", BrightBlue)
            End If
            
            Call SetPlayerAccess(n, i)
            Call SendPlayerData(n)
            Call AddLog(GetPlayerName(Index) & " modificou " & GetPlayerName(n) & " (acesso).", ADMIN_LOG)
        Else
            Call PlayerMsg(Index, "Jogador não está online.", Red)
        End If
    Else
        Call PlayerMsg(Index, "Valor inválido.", Red)
    End If
End Sub
' :::::::::::::::::::::::
' :: Who online packet ::
' :::::::::::::::::::::::
Sub HandleWhosOnline(ByVal Index As Long)
    Call SendWhosOnline(Index)
End Sub
' :::::::::::::::::::::
' :: Set MOTD packet ::
' :::::::::::::::::::::
Sub HandleSetMotd(ByVal Index As Long, ByRef Parse() As String)

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_MAPPER Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If
    
    MOTD = Trim$(Parse(1))
    Call PutVar(App.Path & "\data\motd.ini", "MOTD", "Msg", MOTD)
    Call GlobalMsg("MOTD changed to: " & MOTD, BrightCyan)
    Call AddLog(GetPlayerName(Index) & " changed MOTD to: " & MOTD, ADMIN_LOG)
End Sub
' ::::::::::::::::::
' :: Trade packet ::
' ::::::::::::::::::
Sub HandleTrade(ByVal Index As Long)
    If Map(GetPlayerMap(Index)).Shop > 0 Then
        Call SendTrade(Index, Map(GetPlayerMap(Index)).Shop)
    Else
        Call PlayerMsg(Index, "There is no shop here.", BrightRed)
    End If
End Sub
' ::::::::::::::::::::::::::
' :: Trade request packet ::
' ::::::::::::::::::::::::::
Sub HandleTradeRequest(ByVal Index As Long, ByRef Parse() As String)
Dim n As Long, i As Long, x As Long

    ' Trade num
    n = Val(Parse(1))
    
    ' Prevent hacking
    If (n <= 0) Or (n > MAX_TRADES) Then
        Call HackingAttempt(Index, "Trade Request Modification")
        Exit Sub
    End If
    
    ' Index for shop
    i = Map(GetPlayerMap(Index)).Shop
    
    ' Check if inv full
    x = FindOpenInvSlot(Index, Shop(i).TradeItem(n).GetItem)
    If x = 0 Then
        Call PlayerMsg(Index, "Trade unsuccessful, inventory full.", BrightRed)
        Exit Sub
    End If
    
    ' Check if they have the item
    If HasItem(Index, Shop(i).TradeItem(n).GiveItem) >= Shop(i).TradeItem(n).GiveValue Then
        Call TakeItem(Index, Shop(i).TradeItem(n).GiveItem, Shop(i).TradeItem(n).GiveValue)
        Call GiveItem(Index, Shop(i).TradeItem(n).GetItem, Shop(i).TradeItem(n).GetValue)
        Call PlayerMsg(Index, "The trade was successful!", Yellow)
    Else
        Call PlayerMsg(Index, "Trade unsuccessful.", BrightRed)
    End If
End Sub
' :::::::::::::::::::::
' :: Fix item packet ::
' :::::::::::::::::::::
Sub HandleFixItem(ByVal Index As Long, ByRef Parse() As String)
Dim n As Long, i As Long
Dim ItemNum As Long, DurNeeded As Long, GoldNeeded As Long

    ' Inv num
    n = Val(Parse(1))
    
    ' Make sure its a equipable item
    If Item(GetPlayerInvItemNum(Index, n)).Type < ITEM_TYPE_WEAPON Or Item(GetPlayerInvItemNum(Index, n)).Type > ITEM_TYPE_SHIELD Then
        Call PlayerMsg(Index, "You can only fix weapons, armors, helmets, and shields.", BrightRed)
        Exit Sub
    End If
    
    ' Check if they have a full inventory
    If FindOpenInvSlot(Index, GetPlayerInvItemNum(Index, n)) <= 0 Then
        Call PlayerMsg(Index, "You have no inventory space left!", BrightRed)
        Exit Sub
    End If
    
    ' Now check the rate of pay
    ItemNum = GetPlayerInvItemNum(Index, n)
    i = Int(Item(GetPlayerInvItemNum(Index, n)).Data2 / 5)
    If i <= 0 Then i = 1
    
    DurNeeded = Item(ItemNum).Data1 - GetPlayerInvItemDur(Index, n)
    GoldNeeded = Int(DurNeeded * i / 2)
    If GoldNeeded <= 0 Then GoldNeeded = 1
    
    ' Check if they even need it repaired
    If DurNeeded <= 0 Then
        Call PlayerMsg(Index, "O item está em perfeita condição!", Red)
        Exit Sub
    End If
    
    ' Check if they have enough for at least one point
    If HasItem(Index, 1) >= i Then
        ' Check if they have enough for a total restoration
        If HasItem(Index, 1) >= GoldNeeded Then
            Call TakeItem(Index, 1, GoldNeeded)
            Call SetPlayerInvItemDur(Index, n, Item(ItemNum).Data1)
            Call PlayerMsg(Index, "Item has been totally restored for " & GoldNeeded & " gold!", BrightBlue)
        Else
            ' They dont so restore as much as we can
            DurNeeded = (HasItem(Index, 1) / i)
            GoldNeeded = Int(DurNeeded * i / 2)
            If GoldNeeded <= 0 Then GoldNeeded = 1
            
            Call TakeItem(Index, 1, GoldNeeded)
            Call SetPlayerInvItemDur(Index, n, GetPlayerInvItemDur(Index, n) + DurNeeded)
            Call PlayerMsg(Index, "Item has been partially fixed for " & GoldNeeded & " gold!", BrightBlue)
        End If
    Else
        Call PlayerMsg(Index, "Insufficient gold to fix this item!", BrightRed)
    End If
End Sub
' :::::::::::::::::::
' :: Search packet ::
' :::::::::::::::::::
Sub HandleSearch(ByVal Index As Long, ByRef Parse() As String)
Dim x As Long, y As Long, i As Long

    x = Val(Parse(1))
    y = Val(Parse(2))
    
    ' Prevent subscript out of range
    If x < 0 Or x > MAX_MAPX Or y < 0 Or y > MAX_MAPY Then
        Exit Sub
    End If
    
    ' Check for a player
    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) And GetPlayerMap(Index) = GetPlayerMap(i) And GetPlayerX(i) = x And GetPlayerY(i) = y Then
            
            ' Consider the player
            'If GetPlayerLevel(I) >= GetPlayerLevel(Index) + 5 Then
               ' Call PlayerMsg(Index, "You wouldn't stand a chance.", BrightRed)
            'Else
                'If GetPlayerLevel(I) > GetPlayerLevel(Index) Then
                  '  Call PlayerMsg(Index, "This one seems to have an advantage over you.", Yellow)
               ' Else
                    'If GetPlayerLevel(I) = GetPlayerLevel(Index) Then
                        'Call PlayerMsg(Index, "This would be an even fight.", White)
                    'Else
                       ' If GetPlayerLevel(Index) >= GetPlayerLevel(I) + 5 Then
                          '  Call PlayerMsg(Index, "You could slaughter that player.", BrightBlue)
                        'Else
                           ' If GetPlayerLevel(Index) > GetPlayerLevel(I) Then
                                'Call PlayerMsg(Index, "You would have an advantage over that player.", Yellow)
                           ' End If
                        'End If
                    'End If
                'End If
            'End If
        
            ' Change target
            TempPlayer(Index).Target = i
            TempPlayer(Index).TargetType = TARGET_TYPE_PLAYER
            Call PlayerMsg(Index, "Your target is now " & GetPlayerName(i) & ".", Yellow)
            Exit Sub
        End If
    Next
    
    ' Check for an item
    For i = 1 To MAX_MAP_ITEMS
        If MapItem(GetPlayerMap(Index), i).Num > 0 Then
            If MapItem(GetPlayerMap(Index), i).x = x And MapItem(GetPlayerMap(Index), i).y = y Then
                Call PlayerMsg(Index, "You see a " & Trim$(Item(MapItem(GetPlayerMap(Index), i).Num).Name) & ".", Yellow)
                Exit Sub
            End If
        End If
    Next
    
    ' Check for an npc
    For i = 1 To MAX_MAP_NPCS
        If MapNpc(GetPlayerMap(Index), i).Num > 0 Then
            If MapNpc(GetPlayerMap(Index), i).x = x And MapNpc(GetPlayerMap(Index), i).y = y Then
                ' Change target
                TempPlayer(Index).Target = i
                TempPlayer(Index).TargetType = TARGET_TYPE_NPC
                Call PlayerMsg(Index, "Your target is now a " & Trim$(Npc(MapNpc(GetPlayerMap(Index), i).Num).Name) & ".", Yellow)
                Exit Sub
            End If
        End If
    Next
End Sub
' ::::::::::::::::::
' :: Party packet ::
' ::::::::::::::::::
Sub HandleParty(ByVal Index As Long, ByRef Parse() As String)
Dim n As Long

    n = FindPlayer(Parse(1))
    
    ' Prevent partying with self
    If n = Index Then
        Exit Sub
    End If
            
    ' Check for a previous party and if so drop it
    If TempPlayer(Index).InParty = YES Then
        Call PlayerMsg(Index, "You are already in a party!", Pink)
        Exit Sub
    End If
    
    If n > 0 Then
        ' Check if its an admin
        If GetPlayerAccess(Index) > ADMIN_MONITOR Then
            'Call PlayerMsg(Index, "You can't join a party, you are an admin!", BrightBlue)
            Exit Sub
        End If
    
        If GetPlayerAccess(n) > ADMIN_MONITOR Then
            'Call PlayerMsg(Index, "Admins cannot join parties!", BrightBlue)
            Exit Sub
        End If
        
        ' Make sure they are in right level range
        If GetPlayerLevel(Index) + 5 < GetPlayerLevel(n) Or GetPlayerLevel(Index) - 5 > GetPlayerLevel(n) Then
            Call PlayerMsg(Index, "There is more then a 5 level gap between you two, party failed.", Pink)
            Exit Sub
        End If
        
        ' Check to see if player is already in a party
        If TempPlayer(n).InParty = NO Then
           ' Call PlayerMsg(Index, "Party request has been sent to " & GetPlayerName(n) & ".", Pink)
           ' Call PlayerMsg(n, GetPlayerName(Index) & " wants you to join their party.  Type /join to join, or /leave to decline.", Pink)
        
            TempPlayer(Index).PartyStarter = YES
            TempPlayer(Index).PartyPlayer = n
            TempPlayer(n).PartyPlayer = Index
        Else
            'Call PlayerMsg(Index, "Player is already in a party!", Pink)
        End If
    Else
        Call PlayerMsg(Index, "Jogador não está online.", Red)
    End If
End Sub
' :::::::::::::::::::::::
' :: Join party packet ::
' :::::::::::::::::::::::
Sub HandleJoinParty(ByVal Index As Long, ByRef Parse() As String)
Dim n As Long

    n = TempPlayer(Index).PartyPlayer
    
    If n > 0 Then
        ' Check to make sure they aren't the starter
        If TempPlayer(Index).PartyStarter = NO Then
            ' Check to make sure that each of there party players match
            If TempPlayer(n).PartyPlayer = Index Then
                'Call PlayerMsg(Index, "You have joined " & GetPlayerName(n) & "'s party!", Pink)
                'Call PlayerMsg(n, GetPlayerName(Index) & " has joined your party!", Pink)
                
                TempPlayer(Index).InParty = YES
                TempPlayer(n).InParty = YES
            Else
                'Call PlayerMsg(Index, "Party failed.", Pink)
            End If
        Else
            'Call PlayerMsg(Index, "You have not been invited to join a party!", Pink)
        End If
    Else
        'Call PlayerMsg(Index, "You have not been invited into a party!", Pink)
    End If
End Sub
' ::::::::::::::::::::::::
' :: Leave party packet ::
' ::::::::::::::::::::::::
Sub HandleLeaveParty(ByVal Index As Long, ByRef Parse() As String)
Dim n As Long

    n = TempPlayer(Index).PartyPlayer
    
    If n > 0 Then
        If TempPlayer(Index).InParty = YES Then
            Call PlayerMsg(Index, "You have left the party.", Pink)
            Call PlayerMsg(n, GetPlayerName(Index) & " has left the party.", Pink)
            
            TempPlayer(Index).PartyPlayer = 0
            TempPlayer(Index).PartyStarter = NO
            TempPlayer(Index).InParty = NO
            TempPlayer(n).PartyPlayer = 0
            TempPlayer(n).PartyStarter = NO
            TempPlayer(n).InParty = NO
        Else
            Call PlayerMsg(Index, "Declined party request.", Pink)
            Call PlayerMsg(n, GetPlayerName(Index) & " declined your request.", Pink)
            
            TempPlayer(Index).PartyPlayer = 0
            TempPlayer(Index).PartyStarter = NO
            TempPlayer(Index).InParty = NO
            TempPlayer(n).PartyPlayer = 0
            TempPlayer(n).PartyStarter = NO
            TempPlayer(n).InParty = NO
        End If
    Else
        Call PlayerMsg(Index, "You are not in a party!", Pink)
    End If
End Sub
' :::::::::::::::::::
' :: Spells packet ::
' :::::::::::::::::::
Sub HandleSpells(ByVal Index As Long)
    Call SendPlayerSpells(Index)
End Sub
' :::::::::::::::::
' :: Cast packet ::
' :::::::::::::::::
Sub HandleCast(ByVal Index As Long, ByRef Parse() As String)
'Dim n As Long
'
'    ' Spell slot
'    n = Val(Parse(1))
'
'    Call CastSpell(Index, n)
End Sub
' ::::::::::::::::::::::
' :: Quit game packet ::
' ::::::::::::::::::::::
Sub HandleQuit(ByVal Index As Long)
   
    Call CloseSocket(Index)
End Sub

Sub HandleFireDeath(ByVal Index As Long, ByRef Parse() As String)
Dim Attacker As Long

    Attacker = Val(Parse(1))

    If IsPlaying(Attacker) Then
        If Attacker <> Index Then
            Call GlobalMsg(GetPlayerName(Index) & " foi morto por " & GetPlayerName(Attacker) & "!", BrightRed)
            SetPlayerKills Attacker, GetPlayerKills(Attacker) + 1
            SetPlayerBPoints Attacker, GetPlayerBPoints(Attacker) + 1
            SetPlayerDeaths Index, GetPlayerDeaths(Index) + 1
        Else
            Call GlobalMsg(GetPlayerName(Index) & " se matou!", BrightRed)
            SetPlayerDeaths Index, GetPlayerDeaths(Index) + 1
        End If
    Else
        Call GlobalMsg(GetPlayerName(Index) & " foi morto pelo fogo!", BrightRed)
    End If
    
    OnDeath Index

End Sub

Sub HandleJump(ByVal Index As Long)

    SetPlayerDir Index, DIR_DOWN
    SendPlayerData Index
    SendDataToMap GetPlayerMap(Index), SJump & SEP_CHAR & Index & END_CHAR

End Sub

Sub HandleRequestItemSpawn(ByVal Index As Long, ByRef Parse() As String)
Dim LuckyNum As Integer

    LuckyNum = Random(0, Total_BonusItems)

    If LuckyNum > 0 Then SpawnItem LuckyNum, 1, GetPlayerMap(Index), Val(Parse(1)), Val(Parse(2))

End Sub

Public Sub HandleRequestRoomList(ByVal Index As Long)
Dim i As Long
Dim ii As Long
Dim RoomCount As Long
Dim packet As String

    packet = SSendRoomList & SEP_CHAR
    
    For i = 1 To MAX_MAPS
        If Trim$(Map(i).Name) <> vbNullString Then
            RoomCount = RoomCount + 1
        End If
    Next
    
    If GetPlayerAccess(Index) > 0 Then RoomCount = MAX_MAPS
    
    If RoomCount > 0 Then
        packet = packet & RoomCount & SEP_CHAR
        For i = 1 To MAX_MAPS
            If GetPlayerAccess(Index) < 1 Then
                If Trim$(Map(i).Name) <> vbNullString Then
                If Map(i).Moral = MAP_MORAL_NONE Then
                        packet = packet & "[Em edição] Sala " & i & SEP_CHAR & GetTotalMapPlayers(i) & SEP_CHAR & Map(i).PlayerLimit & SEP_CHAR & Map(i).Name & SEP_CHAR
                   End If
                If Map(i).Moral = MAP_MORAL_FFA Then
                        packet = packet & "[Sobrevivência] Sala " & i & SEP_CHAR & GetTotalMapPlayers(i) & SEP_CHAR & Map(i).PlayerLimit & SEP_CHAR & Map(i).Name & SEP_CHAR
                End If
                If Map(i).Moral = MAP_MORAL_2X2 Then
                        packet = packet & "[Duelo entre times] Sala " & i & SEP_CHAR & GetTotalMapPlayers(i) & SEP_CHAR & Map(i).PlayerLimit & SEP_CHAR & Map(i).Name & SEP_CHAR
                End If
                End If
            Else
                If Map(i).Moral = MAP_MORAL_FFA Then
                    packet = packet & "[Em edição] Sala " & i & SEP_CHAR & GetTotalMapPlayers(i) & SEP_CHAR & Map(i).PlayerLimit & SEP_CHAR & Map(i).Name & SEP_CHAR
                End If
                If Map(i).Moral = MAP_MORAL_NONE Then
                    packet = packet & "[Sobrevivência] Sala " & i & SEP_CHAR & GetTotalMapPlayers(i) & SEP_CHAR & Map(i).PlayerLimit & SEP_CHAR & Map(i).Name & SEP_CHAR
                End If
                If Map(i).Moral = MAP_MORAL_2X2 Then
                        packet = packet & "[Duelo entre times] Sala " & i & SEP_CHAR & GetTotalMapPlayers(i) & SEP_CHAR & Map(i).PlayerLimit & SEP_CHAR & Map(i).Name & SEP_CHAR
                End If
            End If
        Next
    Else
        packet = packet & 0
    End If
    
    packet = packet & END_CHAR
    
    SendDataTo Index, packet

End Sub

Sub HandleRequestJoinRoom(ByVal Index As Long, ByRef Parse() As String)
Dim i As Long, ii As Long
Dim packet As String
Dim Players() As String
Dim p As Long

    i = Val(Parse(1))
    
    If GetTotalMapPlayers(i) = Map(i).PlayerLimit Then
        PlayerMsg Index, "A sala está cheia!", BrightRed
        Exit Sub
    End If
    
    If TempMap(i).InProgress And Map(i).Moral <> MAP_MORAL_FFA Then
        PlayerMsg Index, "A partida já começou!", BrightRed
        Exit Sub
    End If
    
    SetPlayerMap Index, i
    SendPlayerData Index
    
    If Map(i).Moral = MAP_MORAL_FFA Then
        Call SendDataTo(Index, SInGame & END_CHAR)
        Call JoinGame(Index)
        'PlayerMsg Index, "Se quiser voltar ao lobby,digite /sair", HelpColor
    End If
    
    SetPlayerDir Index, DIR_DOWN
    
    PlayerWarp Index, i, START_X, START_Y
    
    SendPlayerData Index
    
    Call GlobalMsg(GetPlayerName(Index) & " entrou na sala " & GetPlayerMap(Index) & "!", JoinLeftColor)
    
    If Map(i).Moral = MAP_MORAL_FFA Then Exit Sub
    
    If Map(i).Moral = MAP_MORAL_2X2 And GetTotalMapPlayers(Index) < 3 Then
    TempPlayer(Index).Team = 1
    End If
    If Map(i).Moral = MAP_MORAL_2X2 And GetTotalMapPlayers(Index) > 2 Then
    TempPlayer(Index).Team = 2
    End If
    
    ReDim Players(1 To Map(GetPlayerMap(Index)).PlayerLimit) As String
    
    ii = 1
    
    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            If GetPlayerMap(i) = GetPlayerMap(Index) Then
            Players(ii) = GetPlayerName(i)
                ii = ii + 1
            End If
        End If
    Next
    
    packet = SStartMatch & SEP_CHAR & Map(i).PlayerLimit & SEP_CHAR
    
    For i = 1 To Map(Parse(1)).PlayerLimit
        If Players(i) = vbNullString Then Players(i) = "Esperando..."
        packet = packet & Players(i) & SEP_CHAR
    Next
    
    packet = packet & END_CHAR
    
    SendDataToMap GetPlayerMap(Index), packet
    
    i = GetPlayerMap(Index)
    
    If GetTotalMapPlayers(i) = Map(i).PlayerLimit Then
        SendDataToMap i, SStartTimer & END_CHAR
        TempMap(i).Timer = GetTickCount + 6000
    End If

End Sub

Sub HandleLeaveRoom(ByVal Index As Long)
Dim Players() As String
Dim ii As Long, i As Long
Dim packet As String
Dim OldMap As Long

    If TempMap(GetPlayerMap(Index)).InProgress And Map(GetPlayerMap(Index)).Moral <> MAP_MORAL_FFA Then
        If Not TempPlayer(Index).Watching Then
            SetPlayerDeaths Index, GetPlayerDeaths(Index) + 1
            GlobalMsg GetPlayerName(Index) & " ganha 1 death por sair antes.", BrightRed
        End If
    End If
    
    TempPlayer(Index).Distance = 1 + Player(Index).Char(TempPlayer(Index).CharNum).BonusDistance
    TempPlayer(Index).Bombs = 1 + Player(Index).Char(TempPlayer(Index).CharNum).BonusLimit
    TempPlayer(Index).Throw = False
    TempPlayer(Index).Speed = False
    OldMap = GetPlayerMap(Index)
    RefreshWall (GetPlayerMap(Index))
    SendDataToMap OldMap, SWatching & SEP_CHAR & Index & SEP_CHAR & "8" & END_CHAR
    TempPlayer(Index).Watching = False
    SendDataTo Index, SLeaveRoom & END_CHAR
    GlobalMsg GetPlayerName(Index) & " saiu da sala " & GetPlayerMap(Index) & "!", Green
    PlayerWarp Index, 100, START_X, START_Y
    SendWhosOnline Index
    
    If Not TempMap(OldMap).InProgress And Map(OldMap).Moral <> MAP_MORAL_FFA Then
        If GetTotalMapPlayers(OldMap) > 0 Then
            ReDim Players(1 To Map(OldMap).PlayerLimit) As String
            
            ii = 1
            
            For i = 1 To MAX_PLAYERS
                If IsPlaying(i) Then
                    If GetPlayerMap(i) = OldMap Then
                        Players(ii) = GetPlayerName(i)
                        ii = ii + 1
                    End If
                End If
            Next
            
            packet = SStartMatch & SEP_CHAR & Map(OldMap).PlayerLimit & SEP_CHAR
            
            For i = 1 To Map(OldMap).PlayerLimit
                If Players(i) = vbNullString Then Players(i) = "Esperando..."
                packet = packet & Players(i) & SEP_CHAR
            Next
            
            packet = packet & END_CHAR
            
            SendDataToMap OldMap, packet
        End If
    End If

End Sub

Sub HandleRequestHighScores(ByVal Index As Long)
Dim i As Long
Dim packet As String

    packet = SHighScoreList
    
    For i = 1 To 100
        packet = packet & SEP_CHAR & Trim$(HighScore(i).Name) & SEP_CHAR & HighScore(i).Kills & SEP_CHAR & HighScore(i).Deaths & SEP_CHAR & HighScore(i).Matchs
    Next i
    
    packet = packet & END_CHAR
    
    SendDataTo Index, packet

End Sub

Sub HandleRequestStats(ByVal Index As Long)

    PlayerMsg Index, GetPlayerName(Index) & "> Matches: " & GetMatchsWon(Index) & "; Kills: " & GetPlayerKills(Index) & "; Deaths: " & GetPlayerDeaths(Index), Green

End Sub
Sub HandleAddFriend(ByVal Index As Long, ByRef Parse() As String)
Dim FriendName As String
Dim i As Long
Dim i2 As Long

    
    'See if character exsists
    If FindChar(Parse(1)) = False Then
        Call PlayerMsg(Index, "Jogador não existe", Red)
        Exit Sub
    Else
        'Add Friend to List
        For i = 1 To MAX_FRIENDS
            If Player(Index).Friends(i).FriendName = Parse(1) Then
            Player(Index).Friends(i).FriendName = vbNullString
            End If
            If Player(Index).Friends(i).FriendName = vbNullString Then
                Player(Index).Friends(i).FriendName = Parse(1)
                Player(Index).AmountofFriends = Player(Index).AmountofFriends + 1
                Exit For
            End If
            
        Next
    End If
    
    'Update Friend List
    Call UpdateFriendsList(Index)
End Sub

Sub HandleRemoveFriend(ByVal Index As Long, ByRef Parse() As String)
Dim i As Long
    
    If Parse(1) = vbNullString Then Exit Sub
    
    For i = 1 To MAX_FRIENDS
        If Player(Index).Friends(i).FriendName = (Parse(1)) Then
            Player(Index).Friends(i).FriendName = vbNullString
            Player(Index).AmountofFriends = Player(Index).AmountofFriends - 1
            Exit For
        End If
    Next
    
    'Update Friend List
    Call UpdateFriendsList(Index)
End Sub


'Friends List
Sub UpdateFriendsList(Index)
Dim FriendName As String
Dim tempName As String
Dim i As Long
Dim o As Long
Dim i2 As Long


    'Check to see if they are Online
    For i = 1 To MAX_FRIENDS
    For o = 1 To MAX_PLAYERS

    If GetPlayerName(o) = Player(Index).Friends(i).FriendName And IsPlaying(o) = True Then
    Call SendDataTo(Index, SUpdateFriendList & SEP_CHAR & Player(Index).Friends(i).FriendName & SEP_CHAR & Player(Index).AmountofFriends & SEP_CHAR & "(Online)" & END_CHAR)
    End If
    Next
    Next
    
    For i = 1 To MAX_FRIENDS
        Call SendDataTo(Index, SUpdateFriendList & SEP_CHAR & Player(Index).Friends(i).FriendName & SEP_CHAR & Player(Index).AmountofFriends & SEP_CHAR & "-" & END_CHAR)
    Next
        
End Sub
Sub HandleAddStatus(ByVal Index As Long, ByRef Parse() As String)
If Player(Index).Char(TempPlayer(Index).CharNum).BonusDistance = 0 And Player(Index).Char(TempPlayer(Index).CharNum).BonusLimit = 0 Then
Player(Index).Char(TempPlayer(Index).CharNum).BonusDistance = Parse(1)
Player(Index).Char(TempPlayer(Index).CharNum).BonusLimit = Parse(2)
End If

If Not Player(Index).Char(TempPlayer(Index).CharNum).BonusDistance = 0 And Player(Index).Char(TempPlayer(Index).CharNum).BonusDistance <= Parse(1) Then
Player(Index).Char(TempPlayer(Index).CharNum).BonusDistance = Parse(1)
End If

If Player(Index).Char(TempPlayer(Index).CharNum).BonusDistance = 0 Then
Player(Index).Char(TempPlayer(Index).CharNum).BonusDistance = Parse(1)
End If

If Not Player(Index).Char(TempPlayer(Index).CharNum).BonusLimit = 0 And Player(Index).Char(TempPlayer(Index).CharNum).BonusDistance <= Parse(1) Then
Player(Index).Char(TempPlayer(Index).CharNum).BonusLimit = Parse(2)
End If

If Player(Index).Char(TempPlayer(Index).CharNum).BonusLimit = 0 Then
Player(Index).Char(TempPlayer(Index).CharNum).BonusLimit = Parse(2)
End If


End Sub
Sub HandleDealBP(ByVal Index As Long, ByRef Parse() As String)

Call SetPlayerBPoints(Index, GetPlayerBPoints(Index) - Parse(1))

End Sub
Sub HandleDealBC(ByVal Index As Long, ByRef Parse() As String)

Call SetPlayerBCash(Index, GetPlayerBCash(Index) - Parse(1))

End Sub
Sub HandleSaveSprite(ByVal Index As Long, ByRef Parse() As String)

Call SetPlayerSprite(Index, Parse(1))

End Sub
