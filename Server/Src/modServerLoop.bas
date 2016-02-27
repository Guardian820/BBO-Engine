Attribute VB_Name = "modServerLoop"
Option Explicit

' ******************************************
' **              BMO Source              **
' ******************************************

Sub ServerLoop()
Dim i As Long
Dim X As Long
Dim Y As Long
Dim Tick As Long

Dim tmr500 As Long
Dim tmr1000 As Long

Dim LastUpdateSavePlayers As Long
Dim LastUpdateMapSpawnItems As Long
Dim LastUpdatePlayerVitals As Long

    ServerOnline = 1
    
    Do While ServerOnline
        Tick = GetTickCount
        
        '/////////////////////////////////////////////
        '// Checks if it's time to update something //
        '/////////////////////////////////////////////
        
        For i = 1 To MAX_MAPS
        
            If Map(i).Moral <> MAP_MORAL_FFA Then
                If Not TempMap(i).InProgress Then
                    If i <> 100 Then
                        If GetTotalMapPlayers(i) = Map(i).PlayerLimit Then
                            If TempMap(i).Timer < GetTickCount Then
                                GlobalMsg "Sala " & i & " iniciou uma batalha !", Green
                                'SendDataToMap i, SStartMatch & END_CHAR
                                SendDataToMap i, SInGame & END_CHAR
                                TempMap(i).InProgress = True
                            End If
                        End If
                    End If
                Else
                    If i <> 100 Then
                        If TempMap(i).Winner Then
                            If TempMap(i).WinnerTimer < GetTickCount Then
                                HandleJump FindPlayer(TempMap(i).WinnerName)
                                TempMap(i).JumpTimes = TempMap(i).JumpTimes + 1
                                If TempMap(i).JumpTimes = 5 Then
                                    For Y = 1 To Player_HighIndex
                                        If GetPlayerMap(Y) = i Then
                                            HandleLeaveRoom (Y)
                                        End If
                                    Next
                                    GlobalMsg TempMap(i).WinnerName & " venceu a batalha na sala " & i & " !", Green
                                    TempMap(i).InProgress = False
                                    TempMap(i).Winner = False
                                    TempMap(i).WinnerName = vbNullString
                                    SetMatchsWon FindPlayer(TempMap(i).WinnerName), GetMatchsWon(FindPlayer(TempMap(i).WinnerName)) + 1
                                    GlobalMsg "Sala " & i & " esta aberta!", Green
                                End If
                                TempMap(i).WinnerTimer = GetTickCount + 700
                            End If
                        Else
                            If GetTotalMapPlayers(i) <= 1 Then
                                TempMap(i).InProgress = False
                                If GetTotalMapPlayers(i) > 0 Then
                                    For Y = 1 To Player_HighIndex
                                        If GetPlayerMap(Y) = i Then 'And GetPlayerAccess(PlayersOnline(Y)) < 1 Then
                                            HandleLeaveRoom (Y)
                                            PlayerMsg Y, "Você foi o único a sair da sala " & Y & "!", Red
                                        End If
                                    Next
                                End If
                                GlobalMsg "Sala " & i & " esta aberta!", Green
                            Else
                                X = 0
                                For Y = 1 To Player_HighIndex
                                    If GetPlayerMap(Y) = i Then
                                        If TempPlayer(Y).Watching Then
                                            X = X + 1
                                        End If
                                    End If
                                Next
                                If X = GetTotalMapPlayers(i) - 1 And Not TempMap(i).Winner Then
                                    For Y = 1 To Player_HighIndex
                                        If GetPlayerMap(Y) = i Then
                                            If Not TempPlayer(Y).Watching Then
                                                TempMap(i).WinnerName = GetPlayerName(Y)
                                            End If
                                        End If
                                    Next
                                    TempMap(i).WinnerTimer = GetTickCount + 500
                                    TempMap(i).JumpTimes = 0
                                    TempMap(i).Winner = True
                                    SetMatchsWon FindPlayer(TempMap(i).WinnerName), GetMatchsWon(FindPlayer(TempMap(i).WinnerName) + 1)
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        
        Next
        
        For i = 1 To Player_HighIndex
        
            If TempPlayer(i).Dieing Then
                
                If TempPlayer(i).DeathTimer < Tick Then
                    OnDeath i, True
                End If
                
            End If
        
        Next
        
        For X = 0 To MAX_MAPX
            For Y = 0 To MAX_MAPY

                If Bomb(X, Y).Here Then
                    If Bomb(X, Y).MakerIndex > 0 Then
                        If LenB(Bomb(X, Y).Maker) > 0 Then
                            If Bomb(X, Y).Map > 0 Then
                                If Bomb(X, Y).Timer < Tick Then

                                    'bomb blow up
                                    Set_Bomb Bomb(X, Y).Map, X, Y
                                    
                                End If
                            End If
                        End If
                    End If
                End If
            
            Next
        Next
        
        ' Check for disconnections every half second
        If tmr500 < Tick Then
        
            For i = 1 To MAX_PLAYERS
                If frmServer.Socket(i).State > sckConnected Then
                    Call CloseSocket(i)
                End If
            Next
            
            tmr500 = GetTickCount + 500
        End If
        
        ' Checks to save players every 10 minutes - Can be tweaked
        If LastUpdateSavePlayers < Tick Then
            UpdateSavePlayers
            LastUpdateSavePlayers = GetTickCount + 600000
        End If
        
        Sleep 1
        DoEvents
        
    Loop
    
End Sub

Private Sub UpdateMapSpawnItems()
Dim X As Long, Y As Long
    
    ' ///////////////////////////////////////////
    ' // This is used for respawning map items //
    ' ///////////////////////////////////////////
    For Y = 1 To MAX_MAPS
        ' Make sure no one is on the map when it respawns
        If Not PlayersOnMap(Y) Then
            ' Clear out unnecessary junk
            For X = 1 To MAX_MAP_ITEMS
                Call ClearMapItem(X, Y)
            Next
                
            ' Spawn the items
            Call SpawnMapItems(Y)
            Call SendMapItemsToAll(Y)
        End If
        DoEvents
    Next
    
End Sub

Private Sub UpdateNpcAI()
Dim i As Long, X As Long, Y As Long, n As Long, x1 As Long, y1 As Long, TickCount As Long
Dim Damage As Long, DistanceX As Long, DistanceY As Long, NpcNum As Long, Target As Long
Dim DidWalk As Boolean
            
    For Y = 1 To MAX_MAPS
        If PlayersOnMap(Y) = YES Then
            TickCount = GetTickCount
            
            ' ////////////////////////////////////
            ' // This is used for closing doors //
            ' ////////////////////////////////////
            If TickCount > TempTile(Y).DoorTimer + 5000 Then
                For x1 = 0 To MAX_MAPX
                    For y1 = 0 To MAX_MAPY
                        If Map(Y).Tile(x1, y1).Type = TILE_TYPE_KEY And TempTile(Y).DoorOpen(x1, y1) = YES Then
                            TempTile(Y).DoorOpen(x1, y1) = NO
                            Call SendDataToMap(Y, SMapKey & SEP_CHAR & x1 & SEP_CHAR & y1 & SEP_CHAR & 0 & END_CHAR)
                        End If
                    Next
                Next
            End If
            
            For X = 1 To MAX_MAP_NPCS
                NpcNum = MapNpc(Y, X).Num
                
                ' /////////////////////////////////////////
                ' // This is used for ATTACKING ON SIGHT //
                ' /////////////////////////////////////////
                ' Make sure theres a npc with the map
                If Map(Y).Npc(X) > 0 And MapNpc(Y, X).Num > 0 Then
                    ' If the npc is a attack on sight, search for a player on the map
                    If Npc(NpcNum).Behavior = NPC_BEHAVIOR_ATTACKONSIGHT Or Npc(NpcNum).Behavior = NPC_BEHAVIOR_GUARD Then
                        For i = 1 To MAX_PLAYERS
                            If IsPlaying(i) Then
                                If GetPlayerMap(i) = Y And MapNpc(Y, X).Target = 0 And GetPlayerAccess(i) <= ADMIN_MONITOR Then
                                    n = Npc(NpcNum).Range
                                    
                                    DistanceX = MapNpc(Y, X).X - GetPlayerX(i)
                                    DistanceY = MapNpc(Y, X).Y - GetPlayerY(i)
                                    
                                    ' Make sure we get a positive value
                                    If DistanceX < 0 Then DistanceX = DistanceX * -1
                                    If DistanceY < 0 Then DistanceY = DistanceY * -1
                                    
                                    ' Are they in range?  if so GET'M!
                                    If DistanceX <= n And DistanceY <= n Then
                                        If Npc(NpcNum).Behavior = NPC_BEHAVIOR_ATTACKONSIGHT Or GetPlayerPK(i) = YES Then
                                            If LenB(Trim$(Npc(NpcNum).AttackSay)) > 0 Then
                                                Call PlayerMsg(i, "A " & Trim$(Npc(NpcNum).Name) & " says, '" & Trim$(Npc(NpcNum).AttackSay) & "' to you.", SayColor)
                                            End If
                                            
                                            MapNpc(Y, X).Target = i
                                        End If
                                    End If
                                End If
                            End If
                        Next
                    End If
                End If
                                                                        
                ' /////////////////////////////////////////////
                ' // This is used for NPC walking/targetting //
                ' /////////////////////////////////////////////
                ' Make sure theres a npc with the map
                If Map(Y).Npc(X) > 0 And MapNpc(Y, X).Num > 0 Then
                    Target = MapNpc(Y, X).Target
                    
                    ' Check to see if its time for the npc to walk
                    If Npc(NpcNum).Behavior <> NPC_BEHAVIOR_SHOPKEEPER Then
                        ' Check to see if we are following a player or not
                        If Target > 0 Then
                            ' Check if the player is even playing, if so follow'm
                            If IsPlaying(Target) And GetPlayerMap(Target) = Y Then
                                DidWalk = False
                                
                                i = Int(Rnd * 5)
                                
                                ' Lets move the npc
                                Select Case i
                                    Case 0
                                        ' Up
                                        If MapNpc(Y, X).Y > GetPlayerY(Target) And DidWalk = False Then
                                            If CanNpcMove(Y, X, DIR_UP) Then
                                                Call NpcMove(Y, X, DIR_UP, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
                                        ' Down
                                        If MapNpc(Y, X).Y < GetPlayerY(Target) And DidWalk = False Then
                                            If CanNpcMove(Y, X, DIR_DOWN) Then
                                                Call NpcMove(Y, X, DIR_DOWN, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
                                        ' Left
                                        If MapNpc(Y, X).X > GetPlayerX(Target) And DidWalk = False Then
                                            If CanNpcMove(Y, X, DIR_LEFT) Then
                                                Call NpcMove(Y, X, DIR_LEFT, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
                                        ' Right
                                        If MapNpc(Y, X).X < GetPlayerX(Target) And DidWalk = False Then
                                            If CanNpcMove(Y, X, DIR_RIGHT) Then
                                                Call NpcMove(Y, X, DIR_RIGHT, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
                                    
                                    Case 1
                                        ' Right
                                        If MapNpc(Y, X).X < GetPlayerX(Target) And DidWalk = False Then
                                            If CanNpcMove(Y, X, DIR_RIGHT) Then
                                                Call NpcMove(Y, X, DIR_RIGHT, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
                                        ' Left
                                        If MapNpc(Y, X).X > GetPlayerX(Target) And DidWalk = False Then
                                            If CanNpcMove(Y, X, DIR_LEFT) Then
                                                Call NpcMove(Y, X, DIR_LEFT, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
                                        ' Down
                                        If MapNpc(Y, X).Y < GetPlayerY(Target) And DidWalk = False Then
                                            If CanNpcMove(Y, X, DIR_DOWN) Then
                                                Call NpcMove(Y, X, DIR_DOWN, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
                                        ' Up
                                        If MapNpc(Y, X).Y > GetPlayerY(Target) And DidWalk = False Then
                                            If CanNpcMove(Y, X, DIR_UP) Then
                                                Call NpcMove(Y, X, DIR_UP, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
                                        
                                    Case 2
                                        ' Down
                                        If MapNpc(Y, X).Y < GetPlayerY(Target) And DidWalk = False Then
                                            If CanNpcMove(Y, X, DIR_DOWN) Then
                                                Call NpcMove(Y, X, DIR_DOWN, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
                                        ' Up
                                        If MapNpc(Y, X).Y > GetPlayerY(Target) And DidWalk = False Then
                                            If CanNpcMove(Y, X, DIR_UP) Then
                                                Call NpcMove(Y, X, DIR_UP, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
                                        ' Right
                                        If MapNpc(Y, X).X < GetPlayerX(Target) And DidWalk = False Then
                                            If CanNpcMove(Y, X, DIR_RIGHT) Then
                                                Call NpcMove(Y, X, DIR_RIGHT, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
                                        ' Left
                                        If MapNpc(Y, X).X > GetPlayerX(Target) And DidWalk = False Then
                                            If CanNpcMove(Y, X, DIR_LEFT) Then
                                                Call NpcMove(Y, X, DIR_LEFT, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
                                    
                                    Case 3
                                        ' Left
                                        If MapNpc(Y, X).X > GetPlayerX(Target) And DidWalk = False Then
                                            If CanNpcMove(Y, X, DIR_LEFT) Then
                                                Call NpcMove(Y, X, DIR_LEFT, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
                                        ' Right
                                        If MapNpc(Y, X).X < GetPlayerX(Target) And DidWalk = False Then
                                            If CanNpcMove(Y, X, DIR_RIGHT) Then
                                                Call NpcMove(Y, X, DIR_RIGHT, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
                                        ' Up
                                        If MapNpc(Y, X).Y > GetPlayerY(Target) And DidWalk = False Then
                                            If CanNpcMove(Y, X, DIR_UP) Then
                                                Call NpcMove(Y, X, DIR_UP, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
                                        ' Down
                                        If MapNpc(Y, X).Y < GetPlayerY(Target) And DidWalk = False Then
                                            If CanNpcMove(Y, X, DIR_DOWN) Then
                                                Call NpcMove(Y, X, DIR_DOWN, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
                                End Select
                                
                                
                            
                                ' Check if we can't move and if player is behind something and if we can just switch dirs
                                If Not DidWalk Then
                                    If MapNpc(Y, X).X - 1 = GetPlayerX(Target) And MapNpc(Y, X).Y = GetPlayerY(Target) Then
                                        If MapNpc(Y, X).Dir <> DIR_LEFT Then
                                            Call NpcDir(Y, X, DIR_LEFT)
                                        End If
                                        DidWalk = True
                                    End If
                                    If MapNpc(Y, X).X + 1 = GetPlayerX(Target) And MapNpc(Y, X).Y = GetPlayerY(Target) Then
                                        If MapNpc(Y, X).Dir <> DIR_RIGHT Then
                                            Call NpcDir(Y, X, DIR_RIGHT)
                                        End If
                                        DidWalk = True
                                    End If
                                    If MapNpc(Y, X).X = GetPlayerX(Target) And MapNpc(Y, X).Y - 1 = GetPlayerY(Target) Then
                                        If MapNpc(Y, X).Dir <> DIR_UP Then
                                            Call NpcDir(Y, X, DIR_UP)
                                        End If
                                        DidWalk = True
                                    End If
                                    If MapNpc(Y, X).X = GetPlayerX(Target) And MapNpc(Y, X).Y + 1 = GetPlayerY(Target) Then
                                        If MapNpc(Y, X).Dir <> DIR_DOWN Then
                                            Call NpcDir(Y, X, DIR_DOWN)
                                        End If
                                        DidWalk = True
                                    End If
                                    
                                    ' We could not move so player must be behind something, walk randomly.
                                    If Not DidWalk Then
                                        i = Int(Rnd * 2)
                                        If i = 1 Then
                                            i = Int(Rnd * 4)
                                            If CanNpcMove(Y, X, i) Then
                                                Call NpcMove(Y, X, i, MOVING_WALKING)
                                            End If
                                        End If
                                    End If
                                End If
                            Else
                                MapNpc(Y, X).Target = 0
                            End If
                        Else
                            i = Int(Rnd * 4)
                            If i = 1 Then
                                i = Int(Rnd * 4)
                                If CanNpcMove(Y, X, i) Then
                                    Call NpcMove(Y, X, i, MOVING_WALKING)
                                End If
                            End If
                        End If
                    End If
                End If
                
                ' /////////////////////////////////////////////
                ' // This is used for npcs to attack players //
                ' /////////////////////////////////////////////
                ' Make sure theres a npc with the map
                If Map(Y).Npc(X) > 0 And MapNpc(Y, X).Num > 0 Then
                    Target = MapNpc(Y, X).Target
                    
                    ' Check if the npc can attack the targeted player player
                    If Target > 0 Then
                        ' Is the target playing and on the same map?
                        If IsPlaying(Target) And GetPlayerMap(Target) = Y Then
                            ' Can the npc attack the player?
                            If CanNpcAttackPlayer(X, Target) Then
                                'If Not CanPlayerBlockHit(Target) Then
                                    Damage = Npc(NpcNum).Stat(Stats.Strength) '- GetPlayerProtection(Target)
                                    If Damage > 0 Then
                                        Call NpcAttackPlayer(X, Target, Damage)
                                    Else
                                        Call PlayerMsg(Target, "The " & Trim$(Npc(NpcNum).Name) & "'s hit didn't even phase you!", BrightBlue)
                                    End If
                                'Else
                                '    Call PlayerMsg(Target, "Your " & Trim$(Item(GetPlayerInvItemNum(Target, GetPlayerEquipmentSlot(Target, Shield))).Name) & " blocks the " & Trim$(Npc(NpcNum).Name) & "'s hit!", BrightCyan)
                                'End If
                            End If
                        Else
                            ' Player left map or game, set target to 0
                            MapNpc(Y, X).Target = 0
                        End If
                    End If
                End If
                
                ' ////////////////////////////////////////////
                ' // This is used for regenerating NPC's HP //
                ' ////////////////////////////////////////////
                ' Check to see if we want to regen some of the npc's hp
                If MapNpc(Y, X).Num > 0 And TickCount > GiveNPCHPTimer + 10000 Then
                    If MapNpc(Y, X).Vital(Vitals.HP) > 0 Then
                        MapNpc(Y, X).Vital(Vitals.HP) = MapNpc(Y, X).Vital(Vitals.HP) + GetNpcVitalRegen(NpcNum, Vitals.HP)
                    
                        ' Check if they have more then they should and if so just set it to max
                        If MapNpc(Y, X).Vital(Vitals.HP) > GetNpcMaxVital(NpcNum, Vitals.HP) Then
                            MapNpc(Y, X).Vital(Vitals.HP) = GetNpcMaxVital(NpcNum, Vitals.HP)
                        End If
                    End If
                End If
                
                ' //////////////////////////////////////
                ' // This is used for spawning an NPC //
                ' //////////////////////////////////////
                ' Check if we are supposed to spawn an npc or not
                If MapNpc(Y, X).Num = 0 And Map(Y).Npc(X) > 0 Then
                    If TickCount > MapNpc(Y, X).SpawnWait + (Npc(Map(Y).Npc(X)).SpawnSecs * 1000) Then
                        Call SpawnNpc(X, Y)
                    End If
                End If
            Next
        End If
        DoEvents
    Next
    
    ' Make sure we reset the timer for npc hp regeneration
    If GetTickCount > GiveNPCHPTimer + 10000 Then
        GiveNPCHPTimer = GetTickCount
    End If

    ' Make sure we reset the timer for door closing
    If GetTickCount > KeyTimer + 15000 Then
        KeyTimer = GetTickCount
    End If
End Sub

Private Sub UpdatePlayerVitals()
Dim i As Long

    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            If GetPlayerVital(i, Vitals.HP) <> GetPlayerMaxVital(i, Vitals.HP) Then
                Call SetPlayerVital(i, Vitals.HP, GetPlayerVital(i, Vitals.HP) + GetPlayerVitalRegen(i, Vitals.HP))
                Call SendVital(i, Vitals.HP)
            End If
            If GetPlayerVital(i, Vitals.MP) <> GetPlayerMaxVital(i, Vitals.MP) Then
                Call SetPlayerVital(i, Vitals.MP, GetPlayerVital(i, Vitals.MP) + GetPlayerVitalRegen(i, Vitals.MP))
                Call SendVital(i, Vitals.MP)
            End If
            If GetPlayerVital(i, Vitals.SP) <> GetPlayerMaxVital(i, Vitals.SP) Then
                Call SetPlayerVital(i, Vitals.SP, GetPlayerVital(i, Vitals.SP) + GetPlayerVitalRegen(i, Vitals.SP))
                Call SendVital(i, Vitals.SP)
            End If
        End If
    Next
End Sub

Private Sub UpdateSavePlayers()
Dim i As Long

    If TotalOnlinePlayers > 0 Then
        Call TextAdd(frmServer.txtText, "Salvando jogadores online...")
        Call GlobalMsg("Salvando todos os jogadores online...", Pink)
        For i = 1 To MAX_PLAYERS
            If IsPlaying(i) Then
                Call SavePlayer(i)
            End If
            DoEvents
        Next
    End If
End Sub

Private Sub HandleShutdown()

    If Secs <= 0 Then Secs = 30
    If Secs Mod 5 = 0 Or Secs <= 5 Then
        Call GlobalMsg("Server Shutdown in " & Secs & " seconds.", BrightBlue)
        Call TextAdd(frmServer.txtText, "Automated Server Shutdown in " & Secs & " seconds.")
    End If
    
    Secs = Secs - 1
    
    If Secs <= 0 Then
        Call GlobalMsg("Server Shutdown.", BrightRed)
        Call DestroyServer
    End If

End Sub

