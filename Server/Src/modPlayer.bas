Attribute VB_Name = "modPlayer"
Option Explicit

' ******************************************
' **              BMO Source              **
' ******************************************

Sub JoinGame(ByVal Index As Long)
Dim i As Long

    ' Set the flag so we know the person is in the game
    TempPlayer(Index).Bombs = 1 + Player(Index).Char(TempPlayer(Index).CharNum).BonusLimit
    TempPlayer(Index).Distance = 1 + Player(Index).Char(TempPlayer(Index).CharNum).BonusDistance
    TempPlayer(Index).Throw = False
    TempPlayer(Index).Speed = False
    'Update the log
    
    ' Send some more little goodies, no need to explain these
    'Call CheckEquippedItems(Index)
    Call SendClasses(Index)
    Call SendNpcs(Index)
    Call SendShops(Index)
    'Call SendSpells(Index)
    'Call SendInventory(Index)
    'Call SendWornEquipment(Index)
    For i = 1 To Vitals.Vital_Count - 1
        Call SendVital(Index, i)
    Next
    'Call SendStats(Index)
    
    ' Warp the player to his saved location
    'Call PlayerWarp(Index, GetPlayerMap(Index), GetPlayerX(Index), GetPlayerY(Index))
    
    ' Send the flag so they know they can start doing stuff
    Call SendDataTo(Index, SInLobby & END_CHAR)
End Sub

Sub LeftGame(ByVal Index As Long)
Dim n As Long

    If TempPlayer(Index).InGame = True Then
        TempPlayer(Index).InGame = False
        
        ' Check if player was the only player on the map and stop npc processing if so
        If GetTotalMapPlayers(GetPlayerMap(Index)) < 1 Then
            PlayersOnMap(GetPlayerMap(Index)) = NO
        End If
        
        Call SavePlayer(Index)
        
        ' Send a global message that he/she left
        If GetPlayerAccess(Index) <= ADMIN_MONITOR Then
            Call GlobalMsg(GetPlayerName(Index) & " saiu do " & GAME_NAME & "!", JoinLeftColor)
        Else
            Call GlobalMsg(GetPlayerName(Index) & " saiu do " & GAME_NAME & "!", Red)
        End If
        Call TextAdd(frmServer.txtText, GetPlayerName(Index) & " foi desconectado do " & GAME_NAME & ".")
        Call SendLeftGame(Index)
        TotalPlayersOnline = TotalPlayersOnline - 1
        Call UpdateHighIndex
    End If
    
    Call ClearPlayer(Index)
    
End Sub

Sub AttackNpc(ByVal Attacker As Long, ByVal MapNpcNum As Long, ByVal Damage As Long)
Dim Name As String
Dim Exp As Long
Dim n As Long, i As Long
Dim STR As Long, DEF As Long, MapNum As Long, NpcNum As Long

    ' Check for subscript out of range
    If IsPlaying(Attacker) = False Or MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Or Damage < 0 Then
        Exit Sub
    End If
    
    MapNum = GetPlayerMap(Attacker)
    NpcNum = MapNpc(MapNum, MapNpcNum).Num
    Name = Trim$(Npc(NpcNum).Name)
    
    ' Send this packet so they can see the person attacking
    Call SendDataToMapBut(Attacker, MapNum, SAttack & SEP_CHAR & Attacker & END_CHAR)
     
    ' Check for weapon
    n = 0
    If GetPlayerEquipmentSlot(Attacker, Weapon) > 0 Then
        n = GetPlayerInvItemNum(Attacker, GetPlayerEquipmentSlot(Attacker, Weapon))
    End If
    
    If Damage >= MapNpc(MapNum, MapNpcNum).Vital(Vitals.HP) Then
        ' Check for a weapon and say damage
        If n = 0 Then
            Call PlayerMsg(Attacker, "Você hitou " & Name & " causando dano de " & Damage & " pontos de vida.", BrightRed)
        Else
            Call PlayerMsg(Attacker, "Você hitou " & Name & " usando " & Trim$(Item(n).Name) & " causando dano de " & Damage & " pontos de vida.", BrightRed)
        End If
                        
        ' Calculate exp to give attacker
        STR = Npc(NpcNum).Stat(Stats.Strength)
        DEF = Npc(NpcNum).Stat(Stats.Defense)
        Exp = STR * DEF * 2
        
        ' Make sure we dont get less then 0
        If Exp < 0 Then
            Exp = 1
        End If
        
        ' Check if in party, if so divide the exp up by 2
        If TempPlayer(Attacker).InParty = NO Then
            Call SetPlayerExp(Attacker, GetPlayerExp(Attacker) + Exp)
            Call PlayerMsg(Attacker, "Você ganhou " & Exp & " de exp.", BrightBlue)
        Else
            Exp = Exp / 2
            
            If Exp < 0 Then
                Exp = 1
            End If
            
            Call SetPlayerExp(Attacker, GetPlayerExp(Attacker) + Exp)
            Call PlayerMsg(Attacker, "Você ganhou " & Exp & " de exp do grupo.", BrightBlue)
            
            n = TempPlayer(Attacker).PartyPlayer
            If n > 0 Then
                Call SetPlayerExp(n, GetPlayerExp(n) + Exp)
                Call PlayerMsg(n, "Você ganhou " & Exp & " de exp do grupo.", BrightBlue)
            End If
        End If
                                
        ' Drop the goods if they get it
        n = Int(Rnd * Npc(NpcNum).DropChance) + 1
        If n = 1 Then
            Call SpawnItem(Npc(NpcNum).DropItem, Npc(NpcNum).DropItemValue, MapNum, MapNpc(MapNum, MapNpcNum).x, MapNpc(MapNum, MapNpcNum).y)
        End If
        
        ' Now set HP to 0 so we know to actually kill them in the server loop (this prevents subscript out of range)
        MapNpc(MapNum, MapNpcNum).Num = 0
        MapNpc(MapNum, MapNpcNum).SpawnWait = GetTickCount
        MapNpc(MapNum, MapNpcNum).Vital(Vitals.HP) = 0
        Call SendDataToMap(MapNum, SNpcDead & SEP_CHAR & MapNpcNum & END_CHAR)
        
        ' Check for level up
        Call CheckPlayerLevelUp(Attacker)
        
        ' Check for level up party member
        If TempPlayer(Attacker).InParty = YES Then
            Call CheckPlayerLevelUp(TempPlayer(Attacker).PartyPlayer)
        End If
    
        ' Check if target is npc that died and if so set target to 0
        If TempPlayer(Attacker).TargetType = TARGET_TYPE_NPC Then
            If TempPlayer(Attacker).Target = MapNpcNum Then
                TempPlayer(Attacker).Target = 0
                TempPlayer(Attacker).TargetType = TARGET_TYPE_NONE
            End If
        End If
    Else
        ' NPC not dead, just do the damage
        MapNpc(MapNum, MapNpcNum).Vital(Vitals.HP) = MapNpc(MapNum, MapNpcNum).Vital(Vitals.HP) - Damage
        
        ' Check for a weapon and say damage
        If n = 0 Then
            Call PlayerMsg(Attacker, "Você hitou " & Name & " causando o dano de " & Damage & " pontos de vida.", Red)
        Else
            Call PlayerMsg(Attacker, "Você hitou " & Name & " com um " & Trim$(Item(n).Name) & " causando dano de " & Damage & " pontos de vida.", Red)
        End If
        
        ' Check if we should send a message
        If MapNpc(MapNum, MapNpcNum).Target = 0 Then
            If LenB(Trim$(Npc(NpcNum).AttackSay)) > 0 Then
                Call PlayerMsg(Attacker, "A " & Trim$(Npc(NpcNum).Name) & ": " & Trim$(Npc(NpcNum).AttackSay), SayColor)
            End If
        End If
        
        ' Set the NPC target to the player
        MapNpc(MapNum, MapNpcNum).Target = Attacker
        
        ' Now check for guard ai and if so have all onmap guards come after'm
        If Npc(MapNpc(MapNum, MapNpcNum).Num).Behavior = NPC_BEHAVIOR_GUARD Then
            For i = 1 To MAX_MAP_NPCS
                If MapNpc(MapNum, i).Num = MapNpc(MapNum, MapNpcNum).Num Then
                    MapNpc(MapNum, i).Target = Attacker
                End If
            Next
        End If
    End If
    
    ' Reduce durability of weapon
    Call DamageEquipment(Attacker, Weapon)
    
    ' Reset attack timer
    TempPlayer(Attacker).AttackTimer = GetTickCount
    
End Sub

Sub AttackPlayer(ByVal Attacker As Long, ByVal MapNum As Long, ByVal x As Long, ByVal y As Long, Optional ByVal Ended As Byte = 0, Optional ByVal Vert As Byte = 0, Optional ByVal BombMaker As String, Optional OnlyItems As Boolean = False)
Dim Exp As Long
Dim n As Long, i As Long

    If x < 0 Or x > MAX_MAPX Or y < 0 Or y > MAX_MAPY Then Exit Sub
    
    For i = 1 To MAX_MAP_ITEMS
        If (MapItem(MapNum, i).Num > 0) Then
            If (MapItem(MapNum, i).Num <= MAX_ITEMS) Then
                If (MapItem(MapNum, i).x = x) Then
                    If (MapItem(MapNum, i).y = y) Then
                    
                        MapItem(MapNum, i).Num = 0
                        MapItem(MapNum, i).Value = 0
                        MapItem(MapNum, i).Dur = 0
                        MapItem(MapNum, i).x = 0
                        MapItem(MapNum, i).y = 0
    
                        Call SpawnItemSlot(i, 0, 0, 0, MapNum, x, y)
                        Exit For
                        
                    End If
                End If
            End If
        End If
    Next
    
    If Bomb(x, y).MakerIndex > 0 Then Bomb(x, y).Timer = GetTickCount - 2000
    
    If Not OnlyItems Then Call SendDataToMap(GetPlayerMap(Attacker), SFire & SEP_CHAR & x & SEP_CHAR & y & SEP_CHAR & Ended & SEP_CHAR & Vert & SEP_CHAR & Attacker & END_CHAR)
    
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            If Not TempPlayer(i).Dieing Then
                If GetPlayerMap(i) = MapNum Then
                    If GetPlayerX(i) = x Then
                        If GetPlayerY(i) = y Then
    
                            If Attacker > 0 And IsPlaying(Attacker) And TempPlayer(Attacker).Team <> TempPlayer(i).Team Then
                                If Attacker <> i Then
                                    Call GlobalMsg(GetPlayerName(i) & " foi morto por " & GetPlayerName(Attacker) & "!", BrightRed)
                                Else
                                    Call GlobalMsg(GetPlayerName(Attacker) & " se matou!", BrightRed)
                                End If
                            Else
                                Call GlobalMsg(GetPlayerName(i) & " foi morto por " & BombMaker & "!", BrightRed)
                            End If
                            
                            Call OnDeath(i)
                            
                            If Attacker <> i Then
                                SetPlayerKills Attacker, GetPlayerKills(Attacker) + 1
                                SetPlayerBPoints Attacker, GetPlayerBPoints(Attacker) + 1
                                SetPlayerDeaths i, GetPlayerDeaths(i) + 1
                            Else
                                SetPlayerDeaths Attacker, GetPlayerDeaths(Attacker) + 1
                            End If
                            
                            ' Reset attack timer
                            TempPlayer(Attacker).AttackTimer = GetTickCount
    
                        End If
                    End If
                End If
            End If
        End If
    Next

End Sub

Sub PlayerWarp(ByVal Index As Long, ByVal MapNum As Long, ByVal x As Long, ByVal y As Long)
Dim ShopNum As Long, OldMap As Long

    ' Check for subscript out of range
    If IsPlaying(Index) = False Or MapNum <= 0 Or MapNum > MAX_MAPS Then Exit Sub
    
    SetPlayerKills Index, GetPlayerKills(Index)
    
    SendDataToMap GetPlayerMap(Index), SDieing & SEP_CHAR & Index & SEP_CHAR & 0 & END_CHAR
    
    TempPlayer(Index).Target = 0
    TempPlayer(Index).TargetType = TARGET_TYPE_NONE
    
    ' Save old map to send erase player data to
    OldMap = GetPlayerMap(Index)
    
    If OldMap <> MapNum Then
        Call SendLeaveMap(Index, OldMap)
    End If
    
    Call SetPlayerMap(Index, MapNum)
    Call SetPlayerX(Index, x)
    Call SetPlayerY(Index, y)
    
    ' Now we check if there were any players left on the map the player just left, and if not stop processing npcs
    If GetTotalMapPlayers(OldMap) = 0 Then
        PlayersOnMap(OldMap) = NO
    End If
    
    ' Sets it so we know to process npcs on the map
    PlayersOnMap(MapNum) = YES
    
    TempPlayer(Index).GettingMap = YES
    Call SendDataTo(Index, SCheckForMap & SEP_CHAR & MapNum & SEP_CHAR & Map(MapNum).Revision & END_CHAR)
    
    For OldMap = 1 To MAX_PLAYERS
        If IsPlaying(OldMap) Then HandleRequestRoomList OldMap
    Next
    
    SendPlayerData Index
    
End Sub

Sub PlayerMove(ByVal Index As Long, ByVal Dir As Long, ByVal Movement As Long)
Dim packet As String
Dim MapNum As Long
Dim x As Long
Dim y As Long
Dim Moved As Byte

    ' Check for subscript out of range
    If IsPlaying(Index) = False Or Dir < DIR_UP Or Dir > DIR_RIGHT Or Movement < 1 Or Movement > 2 Then Exit Sub
    
    Call SetPlayerDir(Index, Dir)
    
    Moved = NO
    
    Select Case Dir
        Case DIR_UP
            ' Check to make sure not outside of boundries
            If GetPlayerY(Index) > 0 Then
                ' Check to make sure that the tile is walkable
                If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) - 1).Type <> TILE_TYPE_BLOCKED Then
                    ' Check to see if the tile is a key and if it is check if its opened
                    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) - 1).Type <> TILE_TYPE_KEY Or (Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) - 1).Type = TILE_TYPE_KEY And TempTile(GetPlayerMap(Index)).DoorOpen(GetPlayerX(Index), GetPlayerY(Index) - 1) = YES) Then
                        Call SetPlayerY(Index, GetPlayerY(Index) - 1)
                        
                        packet = SPlayerMove & SEP_CHAR & Index & SEP_CHAR & GetPlayerX(Index) & SEP_CHAR & GetPlayerY(Index) & SEP_CHAR & GetPlayerDir(Index) & SEP_CHAR & Movement & END_CHAR
                        Call SendDataToMapBut(Index, GetPlayerMap(Index), packet)
                        Moved = YES
                    End If
                End If
            Else
                ' Check to see if we can move them to the another map
                If Map(GetPlayerMap(Index)).Up > 0 Then
                    Call PlayerWarp(Index, Map(GetPlayerMap(Index)).Up, GetPlayerX(Index), MAX_MAPY)
                    Moved = YES
                End If
            End If
                    
        Case DIR_DOWN
            ' Check to make sure not outside of boundries
            If GetPlayerY(Index) < MAX_MAPY Then
                ' Check to make sure that the tile is walkable
                If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) + 1).Type <> TILE_TYPE_BLOCKED Then
                    ' Check to see if the tile is a key and if it is check if its opened
                    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) + 1).Type <> TILE_TYPE_KEY Or (Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) + 1).Type = TILE_TYPE_KEY And TempTile(GetPlayerMap(Index)).DoorOpen(GetPlayerX(Index), GetPlayerY(Index) + 1) = YES) Then
                        Call SetPlayerY(Index, GetPlayerY(Index) + 1)
                        
                        packet = SPlayerMove & SEP_CHAR & Index & SEP_CHAR & GetPlayerX(Index) & SEP_CHAR & GetPlayerY(Index) & SEP_CHAR & GetPlayerDir(Index) & SEP_CHAR & Movement & END_CHAR
                        Call SendDataToMapBut(Index, GetPlayerMap(Index), packet)
                        Moved = YES
                    End If
                End If
            Else
                ' Check to see if we can move them to the another map
                If Map(GetPlayerMap(Index)).Down > 0 Then
                    Call PlayerWarp(Index, Map(GetPlayerMap(Index)).Down, GetPlayerX(Index), 0)
                    Moved = YES
                End If
            End If
        
        Case DIR_LEFT
            ' Check to make sure not outside of boundries
            If GetPlayerX(Index) > 0 Then
                ' Check to make sure that the tile is walkable
                If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) - 1, GetPlayerY(Index)).Type <> TILE_TYPE_BLOCKED Then
                    ' Check to see if the tile is a key and if it is check if its opened
                    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) - 1, GetPlayerY(Index)).Type <> TILE_TYPE_KEY Or (Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) - 1, GetPlayerY(Index)).Type = TILE_TYPE_KEY And TempTile(GetPlayerMap(Index)).DoorOpen(GetPlayerX(Index) - 1, GetPlayerY(Index)) = YES) Then
                        Call SetPlayerX(Index, GetPlayerX(Index) - 1)
                        
                        packet = SPlayerMove & SEP_CHAR & Index & SEP_CHAR & GetPlayerX(Index) & SEP_CHAR & GetPlayerY(Index) & SEP_CHAR & GetPlayerDir(Index) & SEP_CHAR & Movement & END_CHAR
                        Call SendDataToMapBut(Index, GetPlayerMap(Index), packet)
                        Moved = YES
                    End If
                End If
            Else
                ' Check to see if we can move them to the another map
                If Map(GetPlayerMap(Index)).Left > 0 Then
                    Call PlayerWarp(Index, Map(GetPlayerMap(Index)).Left, MAX_MAPX, GetPlayerY(Index))
                    Moved = YES
                End If
            End If
        
        Case DIR_RIGHT
            ' Check to make sure not outside of boundries
            If GetPlayerX(Index) < MAX_MAPX Then
                ' Check to make sure that the tile is walkable
                If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) + 1, GetPlayerY(Index)).Type <> TILE_TYPE_BLOCKED Then
                    ' Check to see if the tile is a key and if it is check if its opened
                    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) + 1, GetPlayerY(Index)).Type <> TILE_TYPE_KEY Or (Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) + 1, GetPlayerY(Index)).Type = TILE_TYPE_KEY And TempTile(GetPlayerMap(Index)).DoorOpen(GetPlayerX(Index) + 1, GetPlayerY(Index)) = YES) Then
                        Call SetPlayerX(Index, GetPlayerX(Index) + 1)
                        
                        packet = SPlayerMove & SEP_CHAR & Index & SEP_CHAR & GetPlayerX(Index) & SEP_CHAR & GetPlayerY(Index) & SEP_CHAR & GetPlayerDir(Index) & SEP_CHAR & Movement & END_CHAR
                        Call SendDataToMapBut(Index, GetPlayerMap(Index), packet)
                        Moved = YES
                    End If
                End If
            Else
                ' Check to see if we can move them to the another map
                If Map(GetPlayerMap(Index)).Right > 0 Then
                    Call PlayerWarp(Index, Map(GetPlayerMap(Index)).Right, 0, GetPlayerY(Index))
                    Moved = YES
                End If
            End If
    End Select
        
    ' Check to see if the tile is a warp tile, and if so warp them
    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Type = TILE_TYPE_WARP Then
        MapNum = Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data1
        x = Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data2
        y = Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data3
                        
        Call PlayerWarp(Index, MapNum, x, y)
        Moved = YES
    End If
    
    ' Check for key trigger open
    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Type = TILE_TYPE_KEYOPEN Then
        x = Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data1
        y = Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data2
        
        If Map(GetPlayerMap(Index)).Tile(x, y).Type = TILE_TYPE_KEY And TempTile(GetPlayerMap(Index)).DoorOpen(x, y) = NO Then
            TempTile(GetPlayerMap(Index)).DoorOpen(x, y) = YES
            TempTile(GetPlayerMap(Index)).DoorTimer = GetTickCount
                            
            Call SendDataToMap(GetPlayerMap(Index), SMapKey & SEP_CHAR & x & SEP_CHAR & y & SEP_CHAR & 1 & END_CHAR)
            Call MapMsg(GetPlayerMap(Index), "A porta foi destrancada.", Red)
        End If
    End If
    
    ' They tried to hack
    'If Moved = NO Then
    '    Call HackingAttempt(Index, "Position Modification")
    'End If
End Sub

Sub CheckEquippedItems(ByVal Index As Long)
Dim Slot As Long
Dim ItemNum As Long
Dim i As Long

    ' We want to check incase an admin takes away an object but they had it equipped
    For i = 1 To Equipment.Equipment_Count - 1
        Slot = GetPlayerEquipmentSlot(Index, i)
        If Slot > 0 Then
            ItemNum = GetPlayerInvItemNum(Index, Slot)
            
            If ItemNum > 0 Then
                Select Case i
                    Case Equipment.Weapon
                        If Item(ItemNum).Type <> ITEM_TYPE_WEAPON Then SetPlayerEquipmentSlot Index, 0, i
                    Case Equipment.Armor
                        If Item(ItemNum).Type <> ITEM_TYPE_ARMOR Then SetPlayerEquipmentSlot Index, 0, i
                    Case Equipment.Helmet
                        If Item(ItemNum).Type <> ITEM_TYPE_HELMET Then SetPlayerEquipmentSlot Index, 0, i
                    Case Equipment.Shield
                        If Item(ItemNum).Type <> ITEM_TYPE_SHIELD Then SetPlayerEquipmentSlot Index, 0, i
                End Select
            Else
               SetPlayerEquipmentSlot Index, 0, i
            End If
        End If
    Next
End Sub

Function FindOpenInvSlot(ByVal Index As Long, ByVal ItemNum As Long) As Long
Dim i As Long
    
    FindOpenInvSlot = 0
    
    ' Check for subscript out of range
    If IsPlaying(Index) = False Or ItemNum <= 0 Or ItemNum > MAX_ITEMS Then
        Exit Function
    End If
    
    If Item(ItemNum).Type = ITEM_TYPE_CURRENCY Then
        ' If currency then check to see if they already have an instance of the item and add it to that
        For i = 1 To MAX_INV
            If GetPlayerInvItemNum(Index, i) = ItemNum Then
                FindOpenInvSlot = i
                Exit Function
            End If
        Next
    End If
    
    For i = 1 To MAX_INV
        ' Try to find an open free slot
        If GetPlayerInvItemNum(Index, i) = 0 Then
            FindOpenInvSlot = i
            Exit Function
        End If
    Next
End Function

Function HasItem(ByVal Index As Long, ByVal ItemNum As Long) As Long
Dim i As Long
    
    HasItem = 0
    
    ' Check for subscript out of range
    If IsPlaying(Index) = False Or ItemNum <= 0 Or ItemNum > MAX_ITEMS Then
        Exit Function
    End If
    
    For i = 1 To MAX_INV
        ' Check to see if the player has the item
        If GetPlayerInvItemNum(Index, i) = ItemNum Then
            If Item(ItemNum).Type = ITEM_TYPE_CURRENCY Then
                HasItem = GetPlayerInvItemValue(Index, i)
            Else
                HasItem = 1
            End If
            Exit Function
        End If
    Next
End Function

Sub TakeItem(ByVal Index As Long, ByVal ItemNum As Long, ByVal ItemVal As Long)
Dim i As Long, n As Long
Dim TakeItem As Boolean

    TakeItem = False
    
    ' Check for subscript out of range
    If IsPlaying(Index) = False Or ItemNum <= 0 Or ItemNum > MAX_ITEMS Then
        Exit Sub
    End If
    
    For i = 1 To MAX_INV
        ' Check to see if the player has the item
        If GetPlayerInvItemNum(Index, i) = ItemNum Then
            If Item(ItemNum).Type = ITEM_TYPE_CURRENCY Then
                ' Is what we are trying to take away more then what they have?  If so just set it to zero
                If ItemVal >= GetPlayerInvItemValue(Index, i) Then
                    TakeItem = True
                Else
                    Call SetPlayerInvItemValue(Index, i, GetPlayerInvItemValue(Index, i) - ItemVal)
                    Call SendInventoryUpdate(Index, i)
                End If
            Else
                ' Check to see if its any sort of ArmorSlot/WeaponSlot
                Select Case Item(GetPlayerInvItemNum(Index, i)).Type
                    Case ITEM_TYPE_WEAPON
                        If GetPlayerEquipmentSlot(Index, Weapon) > 0 Then
                            If i = GetPlayerEquipmentSlot(Index, Weapon) Then
                                Call SetPlayerEquipmentSlot(Index, 0, Weapon)
                                Call SendWornEquipment(Index)
                                TakeItem = True
                            Else
                                ' Check if the item we are taking isn't already equipped
                                If ItemNum <> GetPlayerInvItemNum(Index, GetPlayerEquipmentSlot(Index, Weapon)) Then
                                    TakeItem = True
                                End If
                            End If
                        Else
                            TakeItem = True
                        End If
                
                    Case ITEM_TYPE_ARMOR
                        If GetPlayerEquipmentSlot(Index, Armor) > 0 Then
                            If i = GetPlayerEquipmentSlot(Index, Armor) Then
                                Call SetPlayerEquipmentSlot(Index, 0, Armor)
                                Call SendWornEquipment(Index)
                                TakeItem = True
                            Else
                                ' Check if the item we are taking isn't already equipped
                                If ItemNum <> GetPlayerInvItemNum(Index, GetPlayerEquipmentSlot(Index, Armor)) Then
                                    TakeItem = True
                                End If
                            End If
                        Else
                            TakeItem = True
                        End If
                    
                    Case ITEM_TYPE_HELMET
                        If GetPlayerEquipmentSlot(Index, Helmet) > 0 Then
                            If i = GetPlayerEquipmentSlot(Index, Helmet) Then
                                Call SetPlayerEquipmentSlot(Index, 0, Helmet)
                                Call SendWornEquipment(Index)
                                TakeItem = True
                            Else
                                ' Check if the item we are taking isn't already equipped
                                If ItemNum <> GetPlayerInvItemNum(Index, GetPlayerEquipmentSlot(Index, Helmet)) Then
                                    TakeItem = True
                                End If
                            End If
                        Else
                            TakeItem = True
                        End If
                    
                    Case ITEM_TYPE_SHIELD
                        If GetPlayerEquipmentSlot(Index, Shield) > 0 Then
                            If i = GetPlayerEquipmentSlot(Index, Shield) Then
                                Call SetPlayerEquipmentSlot(Index, 0, Shield)
                                Call SendWornEquipment(Index)
                                TakeItem = True
                            Else
                                ' Check if the item we are taking isn't already equipped
                                If ItemNum <> GetPlayerInvItemNum(Index, GetPlayerEquipmentSlot(Index, Shield)) Then
                                    TakeItem = True
                                End If
                            End If
                        Else
                            TakeItem = True
                        End If
                End Select

                
                n = Item(GetPlayerInvItemNum(Index, i)).Type
                ' Check if its not an equipable weapon, and if it isn't then take it away
                If (n <> ITEM_TYPE_WEAPON) And (n <> ITEM_TYPE_ARMOR) And (n <> ITEM_TYPE_HELMET) And (n <> ITEM_TYPE_SHIELD) Then
                    TakeItem = True
                End If
            End If
                            
            If TakeItem = True Then
                Call SetPlayerInvItemNum(Index, i, 0)
                Call SetPlayerInvItemValue(Index, i, 0)
                Call SetPlayerInvItemDur(Index, i, 0)
                
                ' Send the inventory update
                Call SendInventoryUpdate(Index, i)
                Exit Sub
            End If
        End If
    Next
End Sub

Sub GiveItem(ByVal Index As Long, ByVal ItemNum As Long, ByVal ItemVal As Long)
Dim i As Long

    ' Check for subscript out of range
    If IsPlaying(Index) = False Or ItemNum <= 0 Or ItemNum > MAX_ITEMS Then
        Exit Sub
    End If
    
    i = FindOpenInvSlot(Index, ItemNum)
    
    ' Check to see if inventory is full
    If i <> 0 Then
        Call SetPlayerInvItemNum(Index, i, ItemNum)
        Call SetPlayerInvItemValue(Index, i, GetPlayerInvItemValue(Index, i) + ItemVal)
        
        If (Item(ItemNum).Type = ITEM_TYPE_ARMOR) Or (Item(ItemNum).Type = ITEM_TYPE_WEAPON) Or (Item(ItemNum).Type = ITEM_TYPE_HELMET) Or (Item(ItemNum).Type = ITEM_TYPE_SHIELD) Then
            Call SetPlayerInvItemDur(Index, i, Item(ItemNum).Data1)
        End If
        
        Call SendInventoryUpdate(Index, i)
    Else
        Call PlayerMsg(Index, "Seu inventário está cheio.", BrightRed)
    End If
End Sub

Function HasSpell(ByVal Index As Long, ByVal SpellNum As Long) As Boolean
Dim i As Long

    HasSpell = False
    
    For i = 1 To MAX_PLAYER_SPELLS
        If GetPlayerSpell(Index, i) = SpellNum Then
            HasSpell = True
            Exit Function
        End If
    Next
End Function

Function FindOpenSpellSlot(ByVal Index As Long) As Long
Dim i As Long

    FindOpenSpellSlot = 0
    
    For i = 1 To MAX_PLAYER_SPELLS
        If GetPlayerSpell(Index, i) = 0 Then
            FindOpenSpellSlot = i
            Exit Function
        End If
    Next
End Function
Sub PlayerMapGetItem(ByVal Index As Long)
Dim i As Long
Dim n As Long
Dim MapNum As Long
Dim Msg As String

    If Not IsPlaying(Index) Then
        Exit Sub
    End If
    
    MapNum = GetPlayerMap(Index)
    
    For i = 1 To MAX_MAP_ITEMS
        ' See if theres even an item here
        If (MapItem(MapNum, i).Num > 0) Then
            If (MapItem(MapNum, i).Num <= MAX_ITEMS) Then
            
                ' Check if item is at the same location as the player
                If (MapItem(MapNum, i).x = GetPlayerX(Index)) Then
                
                    If (MapItem(MapNum, i).y = GetPlayerY(Index)) Then
                    
                        ' Find open slot
'                        n = FindOpenInvSlot(Index, MapItem(MapNum, i).Num)
'
'                        ' Open slot available?
'                        If n <> 0 Then
'                            ' Set item in players inventor
'                            Call SetPlayerInvItemNum(Index, n, MapItem(MapNum, i).Num)
'                            If Item(GetPlayerInvItemNum(Index, n)).Type = ITEM_TYPE_CURRENCY Then
'                                Call SetPlayerInvItemValue(Index, n, GetPlayerInvItemValue(Index, n) + MapItem(MapNum, i).Value)
'                                Msg = "You picked up " & MapItem(MapNum, i).Value & " " & Trim$(Item(GetPlayerInvItemNum(Index, n)).Name) & "."
'                            Else
'                                Call SetPlayerInvItemValue(Index, n, 0)
'                                Msg = "You picked up a " & Trim$(Item(GetPlayerInvItemNum(Index, n)).Name) & "."
'                            End If
'                            Call SetPlayerInvItemDur(Index, n, MapItem(MapNum, i).Dur)
                            
                            Select Case MapItem(MapNum, i).Num
                            
                                Case 1
                                If Not TempPlayer(Index).Distance = 4 Then
                                    TempPlayer(Index).Distance = TempPlayer(Index).Distance + 1
                                    End If
                                Case 2
                                    TempPlayer(Index).Bombs = TempPlayer(Index).Bombs + 1
                                Case 3
                                    TempPlayer(Index).Throw = True
                                Case 4
                                    TempPlayer(Index).Speed = True
                                    SendDataToMap GetPlayerMap(Index), SSpeed & SEP_CHAR & "1" & END_CHAR
                                    
                            End Select
                            
                            ' Erase item from the map
                            MapItem(MapNum, i).Num = 0
                            MapItem(MapNum, i).Value = 0
                            MapItem(MapNum, i).Dur = 0
                            MapItem(MapNum, i).x = 0
                            MapItem(MapNum, i).y = 0
                                
                            'Call SendInventoryUpdate(Index, n)
                            Call SpawnItemSlot(i, 0, 0, 0, GetPlayerMap(Index), GetPlayerX(Index), GetPlayerY(Index))
                            'Call PlayerMsg(Index, Msg, Yellow)
                            Exit For
                        'Else
                        '    Call PlayerMsg(Index, "Your inventory is full.", BrightRed)
                        '    Exit For
                        'End If
                        
                    End If
                    
                End If
            
            End If
            
        End If
    Next
End Sub

Sub PlayerMapDropItem(ByVal Index As Long, ByVal InvNum As Long, ByVal Ammount As Long)
Dim i As Long

    ' Check for subscript out of range
    If IsPlaying(Index) = False Or InvNum <= 0 Or InvNum > MAX_INV Then
        Exit Sub
    End If
    
    If (GetPlayerInvItemNum(Index, InvNum) > 0) Then
        If (GetPlayerInvItemNum(Index, InvNum) <= MAX_ITEMS) Then
        
            i = FindOpenMapItemSlot(GetPlayerMap(Index))
            
            If i <> 0 Then
                MapItem(GetPlayerMap(Index), i).Dur = 0
                
                ' Check to see if its any sort of ArmorSlot/WeaponSlot
                Select Case Item(GetPlayerInvItemNum(Index, InvNum)).Type
                    Case ITEM_TYPE_ARMOR
                        If InvNum = GetPlayerEquipmentSlot(Index, Armor) Then
                            Call SetPlayerEquipmentSlot(Index, 0, Armor)
                            Call SendWornEquipment(Index)
                        End If
                        MapItem(GetPlayerMap(Index), i).Dur = GetPlayerInvItemDur(Index, InvNum)
                    
                    Case ITEM_TYPE_WEAPON
                        If InvNum = GetPlayerEquipmentSlot(Index, Weapon) Then
                            Call SetPlayerEquipmentSlot(Index, 0, Weapon)
                            Call SendWornEquipment(Index)
                        End If
                        MapItem(GetPlayerMap(Index), i).Dur = GetPlayerInvItemDur(Index, InvNum)
                        
                    Case ITEM_TYPE_HELMET
                        If InvNum = GetPlayerEquipmentSlot(Index, Helmet) Then
                            Call SetPlayerEquipmentSlot(Index, 0, Helmet)
                            Call SendWornEquipment(Index)
                        End If
                        MapItem(GetPlayerMap(Index), i).Dur = GetPlayerInvItemDur(Index, InvNum)
                                        
                    Case ITEM_TYPE_SHIELD
                        If InvNum = GetPlayerEquipmentSlot(Index, Shield) Then
                            Call SetPlayerEquipmentSlot(Index, 0, Shield)
                            Call SendWornEquipment(Index)
                        End If
                        MapItem(GetPlayerMap(Index), i).Dur = GetPlayerInvItemDur(Index, InvNum)
                End Select
                                    
                MapItem(GetPlayerMap(Index), i).Num = GetPlayerInvItemNum(Index, InvNum)
                MapItem(GetPlayerMap(Index), i).x = GetPlayerX(Index)
                MapItem(GetPlayerMap(Index), i).y = GetPlayerY(Index)
                            
                If Item(GetPlayerInvItemNum(Index, InvNum)).Type = ITEM_TYPE_CURRENCY Then
                    ' Check if its more then they have and if so drop it all
                    If Ammount >= GetPlayerInvItemValue(Index, InvNum) Then
                        MapItem(GetPlayerMap(Index), i).Value = GetPlayerInvItemValue(Index, InvNum)
                        Call MapMsg(GetPlayerMap(Index), GetPlayerName(Index) & " drops " & GetPlayerInvItemValue(Index, InvNum) & " " & Trim$(Item(GetPlayerInvItemNum(Index, InvNum)).Name) & ".", Yellow)
                        Call SetPlayerInvItemNum(Index, InvNum, 0)
                        Call SetPlayerInvItemValue(Index, InvNum, 0)
                        Call SetPlayerInvItemDur(Index, InvNum, 0)
                    Else
                        MapItem(GetPlayerMap(Index), i).Value = Ammount
                        Call MapMsg(GetPlayerMap(Index), GetPlayerName(Index) & " drops " & Ammount & " " & Trim$(Item(GetPlayerInvItemNum(Index, InvNum)).Name) & ".", Yellow)
                        Call SetPlayerInvItemValue(Index, InvNum, GetPlayerInvItemValue(Index, InvNum) - Ammount)
                    End If
                Else
                    ' Its not a currency object so this is easy
                    MapItem(GetPlayerMap(Index), i).Value = 0
                    If Item(GetPlayerInvItemNum(Index, InvNum)).Type >= ITEM_TYPE_WEAPON And Item(GetPlayerInvItemNum(Index, InvNum)).Type <= ITEM_TYPE_SHIELD Then
                        Call MapMsg(GetPlayerMap(Index), GetPlayerName(Index) & " drops a " & Trim$(Item(GetPlayerInvItemNum(Index, InvNum)).Name) & " " & GetPlayerInvItemDur(Index, InvNum) & "/" & Item(GetPlayerInvItemNum(Index, InvNum)).Data1 & ".", Yellow)
                    Else
                        Call MapMsg(GetPlayerMap(Index), GetPlayerName(Index) & " drops a " & Trim$(Item(GetPlayerInvItemNum(Index, InvNum)).Name) & ".", Yellow)
                    End If
                    
                    Call SetPlayerInvItemNum(Index, InvNum, 0)
                    Call SetPlayerInvItemValue(Index, InvNum, 0)
                    Call SetPlayerInvItemDur(Index, InvNum, 0)
                End If
                                            
                ' Send inventory update
                Call SendInventoryUpdate(Index, InvNum)
                ' Spawn the item before we set the num or we'll get a different free map item slot
                Call SpawnItemSlot(i, MapItem(GetPlayerMap(Index), i).Num, Ammount, MapItem(GetPlayerMap(Index), i).Dur, GetPlayerMap(Index), GetPlayerX(Index), GetPlayerY(Index))
            Else
                'Call PlayerMsg(Index, "Já existem muitos itens no chão.", BrightRed)
            End If
        End If
    End If
End Sub

Sub CheckPlayerLevelUp(ByVal Index As Long)
Dim i As Long
Dim expRollover As Long

    ' Check if attacker got a level up
    If GetPlayerExp(Index) >= GetPlayerNextLevel(Index) Then
        expRollover = CLng(GetPlayerExp(Index) - GetPlayerNextLevel(Index))
        Call SetPlayerLevel(Index, GetPlayerLevel(Index) + 1)
                   
        ' Get the ammount of skill points to add
        i = Int(GetPlayerStat(Index, Stats.Speed) / 10)
        If i < 1 Then i = 1
        If i > 3 Then i = 3
           
        Call SetPlayerPOINTS(Index, GetPlayerPOINTS(Index) + i)
        Call SetPlayerExp(Index, expRollover)
        Call GlobalMsg(GetPlayerName(Index) & " ganhou um nível!", Brown)
        Call PlayerMsg(Index, "Você subiu de nível!  Agora você possúi " & GetPlayerPOINTS(Index) & " pontos.", BrightBlue)
    End If
   
End Sub

Function GetPlayerVitalRegen(ByVal Index As Long, ByVal Vital As Vitals) As Long
Dim i As Long

    ' Prevent subscript out of range
    If IsPlaying(Index) = False Or Index <= 0 Or Index > MAX_PLAYERS Then
        GetPlayerVitalRegen = 0
        Exit Function
    End If
    
    Select Case Vital
        Case HP
            i = Int(GetPlayerStat(Index, Stats.Defense) / 2)
        Case MP
            i = Int(GetPlayerStat(Index, Stats.Magic) / 2)
        Case SP
            i = Int(GetPlayerStat(Index, Stats.Speed) / 2)
    End Select
        
    If i < 2 Then i = 2

    GetPlayerVitalRegen = i
End Function

' //////////////////////
' // PLAYER FUNCTIONS //
' //////////////////////

Function GetPlayerLogin(ByVal Index As Long) As String
    GetPlayerLogin = Trim$(Player(Index).Login)
End Function

Sub SetPlayerLogin(ByVal Index As Long, ByVal Login As String)
    Player(Index).Login = Login
End Sub

Function GetPlayerPassword(ByVal Index As Long) As String
    GetPlayerPassword = Trim$(Player(Index).Password)
End Function

Sub SetPlayerPassword(ByVal Index As Long, ByVal Password As String)
    Player(Index).Password = Password
End Sub

Function GetPlayerName(ByVal Index As Long) As String
    GetPlayerName = Trim$(Player(Index).Char(TempPlayer(Index).CharNum).Name)
End Function

Sub SetPlayerName(ByVal Index As Long, ByVal Name As String)
    Player(Index).Char(TempPlayer(Index).CharNum).Name = Name
End Sub

Function GetPlayerClass(ByVal Index As Long) As Long
    GetPlayerClass = Player(Index).Char(TempPlayer(Index).CharNum).Class
End Function

Sub SetPlayerClass(ByVal Index As Long, ByVal ClassNum As Long)
    Player(Index).Char(TempPlayer(Index).CharNum).Class = ClassNum
End Sub

Function GetPlayerSprite(ByVal Index As Long) As Long
    GetPlayerSprite = Player(Index).Char(TempPlayer(Index).CharNum).Sprite
End Function

Sub SetPlayerSprite(ByVal Index As Long, ByVal Sprite As Long)
    Player(Index).Char(TempPlayer(Index).CharNum).Sprite = Sprite
End Sub

Function GetPlayerLevel(ByVal Index As Long) As Long
    GetPlayerLevel = Player(Index).Char(TempPlayer(Index).CharNum).Level
End Function

Sub SetPlayerLevel(ByVal Index As Long, ByVal Level As Long)
    If Level > MAX_LEVELS Then Exit Sub
    Player(Index).Char(TempPlayer(Index).CharNum).Level = Level
End Sub

Function GetPlayerNextLevel(ByVal Index As Long) As Long
    GetPlayerNextLevel = (GetPlayerLevel(Index) + 1) * (GetPlayerStat(Index, Stats.Strength) + GetPlayerStat(Index, Stats.Defense) + GetPlayerStat(Index, Stats.Magic) + GetPlayerStat(Index, Stats.Speed) + GetPlayerPOINTS(Index)) * 25
End Function

Function GetPlayerExp(ByVal Index As Long) As Long
    GetPlayerExp = Player(Index).Char(TempPlayer(Index).CharNum).Exp
End Function

Sub SetPlayerExp(ByVal Index As Long, ByVal Exp As Long)
    Player(Index).Char(TempPlayer(Index).CharNum).Exp = Exp
End Sub

Function GetPlayerAccess(ByVal Index As Long) As Long
    GetPlayerAccess = Player(Index).Char(TempPlayer(Index).CharNum).Access
End Function

Sub SetPlayerAccess(ByVal Index As Long, ByVal Access As Long)
    Player(Index).Char(TempPlayer(Index).CharNum).Access = Access
    If IsPlaying(Index) Then HandleRequestRoomList Index
End Sub

Function GetPlayerPK(ByVal Index As Long) As Long
    GetPlayerPK = Player(Index).Char(TempPlayer(Index).CharNum).PK
End Function

Sub SetPlayerPK(ByVal Index As Long, ByVal PK As Long)
    Player(Index).Char(TempPlayer(Index).CharNum).PK = PK
End Sub

Function GetPlayerVital(ByVal Index As Long, ByVal Vital As Vitals) As Long
    GetPlayerVital = Player(Index).Char(TempPlayer(Index).CharNum).Vital(Vital)
End Function

Sub SetPlayerVital(ByVal Index As Long, ByVal Vital As Vitals, ByVal Value As Long)
    Player(Index).Char(TempPlayer(Index).CharNum).Vital(Vital) = Value
    
    If GetPlayerVital(Index, Vital) > GetPlayerMaxVital(Index, Vital) Then
        Player(Index).Char(TempPlayer(Index).CharNum).Vital(Vital) = GetPlayerMaxVital(Index, Vital)
    End If
    If GetPlayerVital(Index, Vital) < 0 Then
        Player(Index).Char(TempPlayer(Index).CharNum).Vital(Vital) = 0
    End If
End Sub

Function GetPlayerMaxVital(ByVal Index As Long, ByVal Vital As Vitals) As Long
Dim CharNum As Long

    Select Case Vital
        Case HP
            CharNum = TempPlayer(Index).CharNum
            GetPlayerMaxVital = (Player(Index).Char(CharNum).Level + Int(GetPlayerStat(Index, Stats.Strength) / 2) + Class(Player(Index).Char(CharNum).Class).Stat(Stats.Strength)) * 2
        Case MP
            CharNum = TempPlayer(Index).CharNum
            GetPlayerMaxVital = (Player(Index).Char(CharNum).Level + Int(GetPlayerStat(Index, Stats.Magic) / 2) + Class(Player(Index).Char(CharNum).Class).Stat(Stats.Magic)) * 2
        Case SP
            CharNum = TempPlayer(Index).CharNum
            GetPlayerMaxVital = (Player(Index).Char(CharNum).Level + Int(GetPlayerStat(Index, Stats.Speed) / 2) + Class(Player(Index).Char(CharNum).Class).Stat(Stats.Speed)) * 2
    End Select
End Function

Public Function GetPlayerStat(ByVal Index As Long, ByVal Stat As Stats) As Long
    GetPlayerStat = Player(Index).Char(TempPlayer(Index).CharNum).Stat(Stat)
End Function

Public Sub SetPlayerStat(ByVal Index As Long, ByVal Stat As Stats, ByVal Value As Long)
    Player(Index).Char(TempPlayer(Index).CharNum).Stat(Stat) = Value
End Sub

Function GetPlayerPOINTS(ByVal Index As Long) As Long
    GetPlayerPOINTS = Player(Index).Char(TempPlayer(Index).CharNum).POINTS
End Function

Sub SetPlayerPOINTS(ByVal Index As Long, ByVal POINTS As Long)
    Player(Index).Char(TempPlayer(Index).CharNum).POINTS = POINTS
End Sub

Function GetPlayerMap(ByVal Index As Long) As Long
    GetPlayerMap = Player(Index).Char(TempPlayer(Index).CharNum).Map
End Function

Sub SetPlayerMap(ByVal Index As Long, ByVal MapNum As Long)
    If MapNum > 0 And MapNum <= MAX_MAPS Then
        Player(Index).Char(TempPlayer(Index).CharNum).Map = MapNum
    End If
End Sub

Function GetPlayerX(ByVal Index As Long) As Long
    GetPlayerX = Player(Index).Char(TempPlayer(Index).CharNum).x
End Function

Sub SetPlayerX(ByVal Index As Long, ByVal x As Long)
    Player(Index).Char(TempPlayer(Index).CharNum).x = x
End Sub

Function GetPlayerY(ByVal Index As Long) As Long
    GetPlayerY = Player(Index).Char(TempPlayer(Index).CharNum).y
End Function

Sub SetPlayerY(ByVal Index As Long, ByVal y As Long)
    Player(Index).Char(TempPlayer(Index).CharNum).y = y
End Sub

Function GetPlayerDir(ByVal Index As Long) As Long
    GetPlayerDir = Player(Index).Char(TempPlayer(Index).CharNum).Dir
End Function

Sub SetPlayerDir(ByVal Index As Long, ByVal Dir As Long)
    Player(Index).Char(TempPlayer(Index).CharNum).Dir = Dir
    Call SendDataToMapBut(Index, GetPlayerMap(Index), SPlayerDir & SEP_CHAR & Index & SEP_CHAR & GetPlayerDir(Index) & END_CHAR)
End Sub

Function GetPlayerIP(ByVal Index As Long) As String
    GetPlayerIP = frmServer.Socket(Index).RemoteHostIP
End Function

Function GetPlayerInvItemNum(ByVal Index As Long, ByVal InvSlot As Long) As Long
    GetPlayerInvItemNum = Player(Index).Char(TempPlayer(Index).CharNum).Inv(InvSlot).Num
End Function

Sub SetPlayerInvItemNum(ByVal Index As Long, ByVal InvSlot As Long, ByVal ItemNum As Long)
    Player(Index).Char(TempPlayer(Index).CharNum).Inv(InvSlot).Num = ItemNum
End Sub

Function GetPlayerInvItemValue(ByVal Index As Long, ByVal InvSlot As Long) As Long
    GetPlayerInvItemValue = Player(Index).Char(TempPlayer(Index).CharNum).Inv(InvSlot).Value
End Function

Sub SetPlayerInvItemValue(ByVal Index As Long, ByVal InvSlot As Long, ByVal ItemValue As Long)
    Player(Index).Char(TempPlayer(Index).CharNum).Inv(InvSlot).Value = ItemValue
End Sub

Function GetPlayerInvItemDur(ByVal Index As Long, ByVal InvSlot As Long) As Long
    GetPlayerInvItemDur = Player(Index).Char(TempPlayer(Index).CharNum).Inv(InvSlot).Dur
End Function

Sub SetPlayerInvItemDur(ByVal Index As Long, ByVal InvSlot As Long, ByVal ItemDur As Long)
    Player(Index).Char(TempPlayer(Index).CharNum).Inv(InvSlot).Dur = ItemDur
End Sub

Function GetPlayerSpell(ByVal Index As Long, ByVal SpellSlot As Long) As Long
    GetPlayerSpell = Player(Index).Char(TempPlayer(Index).CharNum).Spell(SpellSlot)
End Function

Sub SetPlayerSpell(ByVal Index As Long, ByVal SpellSlot As Long, ByVal SpellNum As Long)
    Player(Index).Char(TempPlayer(Index).CharNum).Spell(SpellSlot) = SpellNum
End Sub

Function GetPlayerEquipmentSlot(ByVal Index As Long, ByVal EquipmentSlot As Equipment) As Byte
    GetPlayerEquipmentSlot = Player(Index).Char(TempPlayer(Index).CharNum).Equipment(EquipmentSlot)
End Function

Sub SetPlayerEquipmentSlot(ByVal Index As Long, ByVal InvNum As Long, ByVal EquipmentSlot As Equipment)
    Player(Index).Char(TempPlayer(Index).CharNum).Equipment(EquipmentSlot) = InvNum
End Sub
Sub SetPlayerBCash(ByVal Index As Long, ByVal SetBCash As Long)

    Player(Index).Char(TempPlayer(Index).CharNum).BCash = SetBCash
    SendDataToAll SHighScore & SEP_CHAR & Index & SEP_CHAR & GetPlayerKills(Index) & SEP_CHAR & GetPlayerDeaths(Index) & SEP_CHAR & GetMatchsWon(Index) & SEP_CHAR & GetPlayerBPoints(Index) & SEP_CHAR & GetPlayerBCash(Index) & END_CHAR

End Sub
Function GetPlayerBCash(ByVal Index As Long) As Long
    GetPlayerBCash = Player(Index).Char(TempPlayer(Index).CharNum).BCash
End Function
Sub SetPlayerBPoints(ByVal Index As Long, ByVal SetBPoints As Long)

    Player(Index).Char(TempPlayer(Index).CharNum).BPoints = SetBPoints
    SendDataToAll SHighScore & SEP_CHAR & Index & SEP_CHAR & GetPlayerKills(Index) & SEP_CHAR & GetPlayerDeaths(Index) & SEP_CHAR & GetMatchsWon(Index) & SEP_CHAR & GetPlayerBPoints(Index) & SEP_CHAR & GetPlayerBCash(Index) & END_CHAR

End Sub
Function GetPlayerBPoints(ByVal Index As Long) As Long
    GetPlayerBPoints = Player(Index).Char(TempPlayer(Index).CharNum).BPoints
End Function
Sub SetPlayerKills(ByVal Index As Long, ByVal SetKills As Long)

    Player(Index).Char(TempPlayer(Index).CharNum).Kills = SetKills
    HandleHighScore Index
    SendDataToAll SHighScore & SEP_CHAR & Index & SEP_CHAR & GetPlayerKills(Index) & SEP_CHAR & GetPlayerDeaths(Index) & SEP_CHAR & GetMatchsWon(Index) & SEP_CHAR & GetPlayerBPoints(Index) & SEP_CHAR & GetPlayerBCash(Index) & END_CHAR

End Sub

Function GetPlayerKills(ByVal Index As Long) As Long
    GetPlayerKills = Player(Index).Char(TempPlayer(Index).CharNum).Kills
End Function

Sub SetPlayerDeaths(ByVal Index As Long, ByVal SetDeaths As Long)

    Player(Index).Char(TempPlayer(Index).CharNum).Deaths = SetDeaths
    HandleHighScore Index
    SendDataToAll SHighScore & SEP_CHAR & Index & SEP_CHAR & GetPlayerKills(Index) & SEP_CHAR & GetPlayerDeaths(Index) & SEP_CHAR & GetMatchsWon(Index) & SEP_CHAR & GetPlayerBPoints(Index) & SEP_CHAR & GetPlayerBCash(Index) & END_CHAR

End Sub

Function GetPlayerDeaths(ByVal Index As Long) As Long
    GetPlayerDeaths = Player(Index).Char(TempPlayer(Index).CharNum).Deaths
End Function

Sub SetMatchsWon(ByVal Index As Long, ByVal SetMatches As Long)

    Player(Index).Char(TempPlayer(Index).CharNum).MatchsWon = SetMatches
    HandleHighScore Index
    SendDataToAll SHighScore & SEP_CHAR & Index & SEP_CHAR & GetPlayerKills(Index) & SEP_CHAR & GetPlayerDeaths(Index) & SEP_CHAR & GetMatchsWon(Index) & SEP_CHAR & GetPlayerBPoints(Index) & SEP_CHAR & GetPlayerBCash(Index) & END_CHAR

End Sub

Function GetMatchsWon(ByVal Index As Long) As Long
    GetMatchsWon = Player(Index).Char(TempPlayer(Index).CharNum).MatchsWon
End Function

Sub HandleHighScore(ByVal Index As Long)
Dim i As Long
Dim id As Long

    For i = 1 To 100
        If Trim$(HighScore(i).Name) = GetPlayerName(Index) Then
            id = i
        End If
    Next
    
    If id < 1 Then id = FindOpenHighScoreSlot
    
    HighScore(id).Name = GetPlayerName(Index)
    HighScore(id).Kills = GetPlayerKills(Index)
    HighScore(id).Deaths = GetPlayerDeaths(Index)
    HighScore(id).Matchs = GetMatchsWon(Index)

End Sub

' ToDo
Sub OnDeath(ByVal Index As Long, Optional ByVal Kill_For_Real As Boolean = False)

    If Not Kill_For_Real Then
    
        TempPlayer(Index).DeathTimer = GetTickCount + 1000
        TempPlayer(Index).Dieing = True
        SendDataToMap GetPlayerMap(Index), SDieing & SEP_CHAR & Index & SEP_CHAR & 1 & END_CHAR
    
    Else
    
        ' Set HP to nothing
        Call SetPlayerVital(Index, Vitals.HP, 0)
        
        TempPlayer(Index).Distance = 1 + Player(Index).Char(TempPlayer(Index).CharNum).BonusDistance
        TempPlayer(Index).Bombs = 1 + Player(Index).Char(TempPlayer(Index).CharNum).BonusLimit
        TempPlayer(Index).Throw = False
        TempPlayer(Index).Speed = False
        
        ' Warp player away
        'Call PlayerWarp(Index, START_MAP, START_X, START_Y)
        If Not Map(GetPlayerMap(Index)).Moral = MAP_MORAL_FFA Then
            SetPlayerX Index, MAX_MAPX
            SetPlayerY Index, MAX_MAPY
            SendPlayerData Index
        Else
            SetPlayerX Index, START_X
            SetPlayerY Index, START_Y
            SendPlayerData Index
        End If
        
        ' Restore vitals
        Call SetPlayerVital(Index, Vitals.HP, GetPlayerMaxVital(Index, Vitals.HP))
        Call SetPlayerVital(Index, Vitals.MP, GetPlayerMaxVital(Index, Vitals.MP))
        Call SetPlayerVital(Index, Vitals.SP, GetPlayerMaxVital(Index, Vitals.SP))
        Call SendVital(Index, Vitals.HP)
        Call SendVital(Index, Vitals.MP)
        Call SendVital(Index, Vitals.SP)
        
        TempPlayer(Index).Dieing = False
        TempPlayer(Index).DeathTimer = 0
        SendDataToMap GetPlayerMap(Index), SDieing & SEP_CHAR & Index & SEP_CHAR & 0 & END_CHAR
        
        If Map(GetPlayerMap(Index)).Moral <> MAP_MORAL_FFA Then
             SendDataToMap GetPlayerMap(Index), SWatching & SEP_CHAR & Index & SEP_CHAR & "9" & END_CHAR
             TempPlayer(Index).Watching = True
        End If
        
    End If

End Sub

Sub DamageEquipment(ByVal Index As Long, ByVal EquipmentSlot As Equipment)
Dim Slot As Long
    
    Slot = GetPlayerEquipmentSlot(Index, EquipmentSlot)
    
    If Slot > 0 Then
        Call SetPlayerInvItemDur(Index, Slot, GetPlayerInvItemDur(Index, Slot) - 1)
            
        If GetPlayerInvItemDur(Index, Slot) <= 0 Then
            Call PlayerMsg(Index, "Seu " & Trim$(Item(GetPlayerInvItemNum(Index, Slot)).Name) & " quebrou.", Yellow)
            Call TakeItem(Index, GetPlayerInvItemNum(Index, Slot), 0)
        Else
            If GetPlayerInvItemDur(Index, Slot) <= 5 Then
                Call PlayerMsg(Index, "Seu " & Trim$(Item(GetPlayerInvItemNum(Index, Slot)).Name) & " está prestes a quebrar!", Yellow)
            End If
        End If
    End If
End Sub

