Attribute VB_Name = "modGameLogic"
Option Explicit

' ******************************************
' **              BMO Source              **
' ******************************************

Function FindOpenPlayerSlot() As Long
Dim i As Long

    FindOpenPlayerSlot = 0
    
    For i = 1 To MAX_PLAYERS
        If Not IsConnected(i) Then
            FindOpenPlayerSlot = i
            Exit Function
        End If
    Next
    
End Function

Function FindOpenHighScoreSlot() As Long
Dim i As Long

    For i = 1 To 100
        If HighScore(i).Kills = 0 And HighScore(i).Deaths = 0 And HighScore(i).Matchs = 0 Then
            FindOpenHighScoreSlot = i
            Exit Function
        End If
    Next

End Function

Function FindOpenMapItemSlot(ByVal MapNum As Long) As Long
Dim i As Long

    FindOpenMapItemSlot = 0
    
    ' Check for subscript out of range
    If MapNum <= 0 Or MapNum > MAX_MAPS Then
        Exit Function
    End If
    
    For i = 1 To MAX_MAP_ITEMS
        If MapItem(MapNum, i).Num = 0 Then
            FindOpenMapItemSlot = i
            Exit Function
        End If
    Next
End Function

Function TotalOnlinePlayers() As Long
Dim i As Long

    TotalOnlinePlayers = 0
    
    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            TotalOnlinePlayers = TotalOnlinePlayers + 1
        End If
    Next
End Function

Function FindPlayer(ByVal Name As String) As Long
Dim i As Long

    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            ' Make sure we dont try to check a name thats to small
            If Len(GetPlayerName(i)) >= Len(Trim$(Name)) Then
                If UCase$(Mid$(GetPlayerName(i), 1, Len(Trim$(Name)))) = UCase$(Trim$(Name)) Then
                    FindPlayer = i
                    Exit Function
                End If
            End If
        End If
    Next
    
    FindPlayer = 0
End Function

Sub SpawnItem(ByVal ItemNum As Long, ByVal ItemVal As Long, ByVal MapNum As Long, ByVal x As Long, ByVal y As Long)
Dim i As Long

    ' Check for subscript out of range
    If ItemNum < 1 Or ItemNum > MAX_ITEMS Or MapNum <= 0 Or MapNum > MAX_MAPS Then
        Exit Sub
    End If
    
    ' Find open map item slot
    i = FindOpenMapItemSlot(MapNum)
    
    Call SpawnItemSlot(i, ItemNum, ItemVal, Item(ItemNum).Data1, MapNum, x, y)
End Sub

Sub SpawnItemSlot(ByVal MapItemSlot As Long, ByVal ItemNum As Long, ByVal ItemVal As Long, ByVal ItemDur As Long, ByVal MapNum As Long, ByVal x As Long, ByVal y As Long)
Dim packet As String
Dim i As Long
    
    ' Check for subscript out of range
    If MapItemSlot <= 0 Or MapItemSlot > MAX_MAP_ITEMS Or ItemNum < 0 Or ItemNum > MAX_ITEMS Or MapNum <= 0 Or MapNum > MAX_MAPS Then
        Exit Sub
    End If
    
    i = MapItemSlot
    
    If i <> 0 Then
        If ItemNum >= 0 Then
            If ItemNum <= MAX_ITEMS Then
    
                MapItem(MapNum, i).Num = ItemNum
                MapItem(MapNum, i).Value = ItemVal
                
                If ItemNum <> 0 Then
                    If (Item(ItemNum).Type >= ITEM_TYPE_WEAPON) And (Item(ItemNum).Type <= ITEM_TYPE_SHIELD) Then
                        MapItem(MapNum, i).Dur = ItemDur
                    Else
                        MapItem(MapNum, i).Dur = 0
                    End If
                Else
                    MapItem(MapNum, i).Dur = 0
                End If
                
                MapItem(MapNum, i).x = x
                MapItem(MapNum, i).y = y
                    
                packet = SSpawnItem & SEP_CHAR & i & SEP_CHAR & ItemNum & SEP_CHAR & ItemVal & SEP_CHAR & MapItem(MapNum, i).Dur & SEP_CHAR & x & SEP_CHAR & y & END_CHAR
                Call SendDataToMap(MapNum, packet)
                
            End If
        End If
    End If

End Sub

Sub SpawnAllMapsItems()
Dim i As Long
    
    For i = 1 To MAX_MAPS
        Call SpawnMapItems(i)
    Next
End Sub

Sub SpawnMapItems(ByVal MapNum As Long)
Dim x As Long
Dim y As Long

    ' Check for subscript out of range
    If MapNum <= 0 Or MapNum > MAX_MAPS Then
        Exit Sub
    End If
    
    ' Spawn what we have
    For x = 0 To MAX_MAPX
        For y = 0 To MAX_MAPY
            ' Check if the tile type is an item or a saved tile incase someone drops something
            If (Map(MapNum).Tile(x, y).Type = TILE_TYPE_ITEM) Then
                ' Check to see if its a currency and if they set the value to 0 set it to 1 automatically
                If Item(Map(MapNum).Tile(x, y).Data1).Type = ITEM_TYPE_CURRENCY And Map(MapNum).Tile(x, y).Data2 <= 0 Then
                    Call SpawnItem(Map(MapNum).Tile(x, y).Data1, 1, MapNum, x, y)
                Else
                    Call SpawnItem(Map(MapNum).Tile(x, y).Data1, Map(MapNum).Tile(x, y).Data2, MapNum, x, y)
                End If
            End If
        Next
    Next
End Sub

Sub SpawnNpc(ByVal MapNpcNum As Long, ByVal MapNum As Long)
Dim packet As String
Dim NpcNum As Long
Dim i As Long, x As Long, y As Long
Dim Spawned As Boolean

    ' Check for subscript out of range
    If MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Or MapNum <= 0 Or MapNum > MAX_MAPS Then
        Exit Sub
    End If
    
    Spawned = False
    
    NpcNum = Map(MapNum).Npc(MapNpcNum)
    If NpcNum > 0 Then
        MapNpc(MapNum, MapNpcNum).Num = NpcNum
        MapNpc(MapNum, MapNpcNum).Target = 0
        
        MapNpc(MapNum, MapNpcNum).Vital(Vitals.HP) = GetNpcMaxVital(NpcNum, Vitals.HP)
        MapNpc(MapNum, MapNpcNum).Vital(Vitals.MP) = GetNpcMaxVital(NpcNum, Vitals.MP)
        MapNpc(MapNum, MapNpcNum).Vital(Vitals.SP) = GetNpcMaxVital(NpcNum, Vitals.SP)
        
        MapNpc(MapNum, MapNpcNum).Dir = Int(Rnd * 4)
        
        ' Well try 100 times to randomly place the sprite
        For i = 1 To 100
            x = Int(Rnd * MAX_MAPX)
            y = Int(Rnd * MAX_MAPY)
            
            ' Check if the tile is walkable
            If Map(MapNum).Tile(x, y).Type = TILE_TYPE_WALKABLE Then
                MapNpc(MapNum, MapNpcNum).x = x
                MapNpc(MapNum, MapNpcNum).y = y
                Spawned = True
                Exit For
            End If
        Next
        
        ' Didn't spawn, so now we'll just try to find a free tile
        If Not Spawned Then
            For x = 0 To MAX_MAPX
                For y = 0 To MAX_MAPY
                    If Map(MapNum).Tile(x, y).Type = TILE_TYPE_WALKABLE Then
                        MapNpc(MapNum, MapNpcNum).x = x
                        MapNpc(MapNum, MapNpcNum).y = y
                        Spawned = True
                    End If
                Next
            Next
        End If
             
        ' If we suceeded in spawning then send it to everyone
        If Spawned Then
            packet = SSpawnNpc & SEP_CHAR & MapNpcNum & SEP_CHAR & MapNpc(MapNum, MapNpcNum).Num & SEP_CHAR & MapNpc(MapNum, MapNpcNum).x & SEP_CHAR & MapNpc(MapNum, MapNpcNum).y & SEP_CHAR & MapNpc(MapNum, MapNpcNum).Dir & END_CHAR
            Call SendDataToMap(MapNum, packet)
        End If
    End If
End Sub

Sub SpawnMapNpcs(ByVal MapNum As Long)
Dim i As Long

    For i = 1 To MAX_MAP_NPCS
        Call SpawnNpc(i, MapNum)
    Next
End Sub

Sub SpawnAllMapNpcs()
Dim i As Long

    For i = 1 To MAX_MAPS
        Call SpawnMapNpcs(i)
    Next
End Sub

Function CanAttackPlayer(ByVal Attacker As Long, ByVal Victim As Long) As Boolean

    CanAttackPlayer = False

    ' Check attack timer
    If GetTickCount < TempPlayer(Attacker).AttackTimer + 1000 Then Exit Function
    
    ' Check for subscript out of range
    If Not IsPlaying(Victim) Then Exit Function

    ' Make sure they are on the same map
    If Not GetPlayerMap(Attacker) = GetPlayerMap(Victim) Then Exit Function
       
    ' Make sure we dont attack the player if they are switching maps
    If TempPlayer(Victim).GettingMap = YES Then Exit Function
   
    ' Check if at same coordinates
    Select Case GetPlayerDir(Attacker)
        Case DIR_UP
            If Not ((GetPlayerY(Victim) + 1 = GetPlayerY(Attacker)) And (GetPlayerX(Victim) = GetPlayerX(Attacker))) Then Exit Function
        Case DIR_DOWN
            If Not ((GetPlayerY(Victim) - 1 = GetPlayerY(Attacker)) And (GetPlayerX(Victim) = GetPlayerX(Attacker))) Then Exit Function
        Case DIR_LEFT
            If Not ((GetPlayerY(Victim) = GetPlayerY(Attacker)) And (GetPlayerX(Victim) + 1 = GetPlayerX(Attacker))) Then Exit Function
        Case DIR_RIGHT
            If Not ((GetPlayerY(Victim) = GetPlayerY(Attacker)) And (GetPlayerX(Victim) - 1 = GetPlayerX(Attacker))) Then Exit Function
        Case Else
            Exit Function
    End Select
    
    ' Check if map is attackable
    If Not Map(GetPlayerMap(Attacker)).Moral = MAP_MORAL_NONE Then
        If GetPlayerPK(Victim) = NO Then
            Call PlayerMsg(Attacker, "Essa é uma zona segura!", BrightRed)
            Exit Function
        End If
    End If
   
    ' Make sure they have more then 0 hp
    If GetPlayerVital(Victim, Vitals.HP) <= 0 Then Exit Function
    
    ' Check to make sure that they dont have access
    If GetPlayerAccess(Attacker) > ADMIN_MONITOR Then
        Call PlayerMsg(Attacker, "????", BrightBlue)
        Exit Function
    End If

    ' Check to make sure the victim isn't an admin
    If GetPlayerAccess(Victim) > ADMIN_MONITOR Then
        Call PlayerMsg(Attacker, "???? " & GetPlayerName(Victim) & " ????", BrightRed)
        Exit Function
    End If

    ' Make sure attacker is high enough level
    If GetPlayerLevel(Attacker) < 10 Then
        Call PlayerMsg(Attacker, "????", BrightRed)
        Exit Function
    End If
   
    ' Make sure victim is high enough level
    If GetPlayerLevel(Victim) < 10 Then
        Call PlayerMsg(Attacker, GetPlayerName(Victim) & " ????", BrightRed)
        Exit Function
    End If
    
    CanAttackPlayer = True

End Function

Function CanAttackNpc(ByVal Attacker As Long, ByVal MapNpcNum As Long) As Boolean
Dim MapNum As Long, NpcNum As Long
Dim NpcX As Long, NpcY As Long

    CanAttackNpc = False
    
    ' Check for subscript out of range
    If IsPlaying(Attacker) = False Or MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Then
        Exit Function
    End If
    
    ' Check for subscript out of range
    If MapNpc(GetPlayerMap(Attacker), MapNpcNum).Num <= 0 Then
        Exit Function
    End If
    
    MapNum = GetPlayerMap(Attacker)
    NpcNum = MapNpc(MapNum, MapNpcNum).Num
    
    ' Make sure the npc isn't already dead
    If MapNpc(MapNum, MapNpcNum).Vital(Vitals.HP) <= 0 Then
        Exit Function
    End If
    
    ' Make sure they are on the same map
    If IsPlaying(Attacker) Then
        If NpcNum > 0 And GetTickCount > TempPlayer(Attacker).AttackTimer + 1000 Then
            ' Check if at same coordinates
            Select Case GetPlayerDir(Attacker)
                Case DIR_UP
                    NpcX = MapNpc(MapNum, MapNpcNum).x
                    NpcY = MapNpc(MapNum, MapNpcNum).y + 1
                Case DIR_DOWN
                    NpcX = MapNpc(MapNum, MapNpcNum).x
                    NpcY = MapNpc(MapNum, MapNpcNum).y - 1
                Case DIR_LEFT
                    NpcX = MapNpc(MapNum, MapNpcNum).x + 1
                    NpcY = MapNpc(MapNum, MapNpcNum).y
                Case DIR_RIGHT
                    NpcX = MapNpc(MapNum, MapNpcNum).x - 1
                    NpcY = MapNpc(MapNum, MapNpcNum).y
            End Select
            
            If NpcX = GetPlayerX(Attacker) Then
                If NpcY = GetPlayerY(Attacker) Then
                    If Npc(NpcNum).Behavior <> NPC_BEHAVIOR_FRIENDLY And Npc(NpcNum).Behavior <> NPC_BEHAVIOR_SHOPKEEPER Then
                        CanAttackNpc = True
                    Else
                        Call PlayerMsg(Attacker, "???? " & Trim$(Npc(NpcNum).Name) & " ????", BrightBlue)
                    End If
                End If
            End If
        End If
    End If
End Function

Function CanNpcAttackPlayer(ByVal MapNpcNum As Long, ByVal Index As Long) As Boolean
Dim MapNum As Long, NpcNum As Long
    
    CanNpcAttackPlayer = False
    
    ' Check for subscript out of range
    If MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Or IsPlaying(Index) = False Then
        Exit Function
    End If
    
    ' Check for subscript out of range
    If MapNpc(GetPlayerMap(Index), MapNpcNum).Num <= 0 Then
        Exit Function
    End If
    
    MapNum = GetPlayerMap(Index)
    NpcNum = MapNpc(MapNum, MapNpcNum).Num
    
    ' Make sure the npc isn't already dead
    If MapNpc(MapNum, MapNpcNum).Vital(Vitals.HP) <= 0 Then
        Exit Function
    End If
    
    ' Make sure npcs dont attack more then once a second
    If GetTickCount < MapNpc(MapNum, MapNpcNum).AttackTimer + 1000 Then
        Exit Function
    End If
    
    ' Make sure we dont attack the player if they are switching maps
    If TempPlayer(Index).GettingMap = YES Then
        Exit Function
    End If
    
    MapNpc(MapNum, MapNpcNum).AttackTimer = GetTickCount
    
    ' Make sure they are on the same map
    If IsPlaying(Index) Then
        If NpcNum > 0 Then
            ' Check if at same coordinates
            If (GetPlayerY(Index) + 1 = MapNpc(MapNum, MapNpcNum).y) And (GetPlayerX(Index) = MapNpc(MapNum, MapNpcNum).x) Then
                CanNpcAttackPlayer = True
            Else
                If (GetPlayerY(Index) - 1 = MapNpc(MapNum, MapNpcNum).y) And (GetPlayerX(Index) = MapNpc(MapNum, MapNpcNum).x) Then
                    CanNpcAttackPlayer = True
                Else
                    If (GetPlayerY(Index) = MapNpc(MapNum, MapNpcNum).y) And (GetPlayerX(Index) + 1 = MapNpc(MapNum, MapNpcNum).x) Then
                        CanNpcAttackPlayer = True
                    Else
                        If (GetPlayerY(Index) = MapNpc(MapNum, MapNpcNum).y) And (GetPlayerX(Index) - 1 = MapNpc(MapNum, MapNpcNum).x) Then
                            CanNpcAttackPlayer = True
                        End If
                    End If
                End If
            End If

'            Select Case MapNpc(MapNum, MapNpcNum).Dir
'                Case DIR_UP
'                    If (GetPlayerY(Index) + 1 = MapNpc(MapNum, MapNpcNum).y) And (GetPlayerX(Index) = MapNpc(MapNum, MapNpcNum).x) Then
'                        CanNpcAttackPlayer = True
'                    End If
'
'                Case DIR_DOWN
'                    If (GetPlayerY(Index) - 1 = MapNpc(MapNum, MapNpcNum).y) And (GetPlayerX(Index) = MapNpc(MapNum, MapNpcNum).x) Then
'                        CanNpcAttackPlayer = True
'                    End If
'
'                Case DIR_LEFT
'                    If (GetPlayerY(Index) = MapNpc(MapNum, MapNpcNum).y) And (GetPlayerX(Index) + 1 = MapNpc(MapNum, MapNpcNum).x) Then
'                        CanNpcAttackPlayer = True
'                    End If
'
'                Case DIR_RIGHT
'                    If (GetPlayerY(Index) = MapNpc(MapNum, MapNpcNum).y) And (GetPlayerX(Index) - 1 = MapNpc(MapNum, MapNpcNum).x) Then
'                        CanNpcAttackPlayer = True
'                    End If
'            End Select
        End If
    End If
End Function

Sub NpcAttackPlayer(ByVal MapNpcNum As Long, ByVal Victim As Long, ByVal Damage As Long)

End Sub

Function CanNpcMove(ByVal MapNum As Long, ByVal MapNpcNum As Long, ByVal Dir As Byte) As Boolean
Dim i As Long, n As Long
Dim x As Long, y As Long

    CanNpcMove = False
    
    ' Check for subscript out of range
    If MapNum <= 0 Or MapNum > MAX_MAPS Or MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Or Dir < DIR_UP Or Dir > DIR_RIGHT Then
        Exit Function
    End If
    
    x = MapNpc(MapNum, MapNpcNum).x
    y = MapNpc(MapNum, MapNpcNum).y
    
    CanNpcMove = True
    
    Select Case Dir
        Case DIR_UP
            ' Check to make sure not outside of boundries
            If y > 0 Then
                n = Map(MapNum).Tile(x, y - 1).Type
                
                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM Then
                    CanNpcMove = False
                    Exit Function
                End If
                
                ' Check to make sure that there is not a player in the way
                For i = 1 To MAX_PLAYERS
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = MapNum) And (GetPlayerX(i) = MapNpc(MapNum, MapNpcNum).x) And (GetPlayerY(i) = MapNpc(MapNum, MapNpcNum).y - 1) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    End If
                Next
                
                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (i <> MapNpcNum) And (MapNpc(MapNum, i).Num > 0) And (MapNpc(MapNum, i).x = MapNpc(MapNum, MapNpcNum).x) And (MapNpc(MapNum, i).y = MapNpc(MapNum, MapNpcNum).y - 1) Then
                        CanNpcMove = False
                        Exit Function
                    End If
                Next
            Else
                CanNpcMove = False
            End If
                
        Case DIR_DOWN
            ' Check to make sure not outside of boundries
            If y < MAX_MAPY Then
                n = Map(MapNum).Tile(x, y + 1).Type
                
                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM Then
                    CanNpcMove = False
                    Exit Function
                End If
                
                ' Check to make sure that there is not a player in the way
                For i = 1 To MAX_PLAYERS
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = MapNum) And (GetPlayerX(i) = MapNpc(MapNum, MapNpcNum).x) And (GetPlayerY(i) = MapNpc(MapNum, MapNpcNum).y + 1) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    End If
                Next
                
                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (i <> MapNpcNum) And (MapNpc(MapNum, i).Num > 0) And (MapNpc(MapNum, i).x = MapNpc(MapNum, MapNpcNum).x) And (MapNpc(MapNum, i).y = MapNpc(MapNum, MapNpcNum).y + 1) Then
                        CanNpcMove = False
                        Exit Function
                    End If
                Next
            Else
                CanNpcMove = False
            End If
                
        Case DIR_LEFT
            ' Check to make sure not outside of boundries
            If x > 0 Then
                n = Map(MapNum).Tile(x - 1, y).Type
                
                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM Then
                    CanNpcMove = False
                    Exit Function
                End If
                
                ' Check to make sure that there is not a player in the way
                For i = 1 To MAX_PLAYERS
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = MapNum) And (GetPlayerX(i) = MapNpc(MapNum, MapNpcNum).x - 1) And (GetPlayerY(i) = MapNpc(MapNum, MapNpcNum).y) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    End If
                Next
                
                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (i <> MapNpcNum) And (MapNpc(MapNum, i).Num > 0) And (MapNpc(MapNum, i).x = MapNpc(MapNum, MapNpcNum).x - 1) And (MapNpc(MapNum, i).y = MapNpc(MapNum, MapNpcNum).y) Then
                        CanNpcMove = False
                        Exit Function
                    End If
                Next
            Else
                CanNpcMove = False
            End If
                
        Case DIR_RIGHT
            ' Check to make sure not outside of boundries
            If x < MAX_MAPX Then
                n = Map(MapNum).Tile(x + 1, y).Type
                
                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM Then
                    CanNpcMove = False
                    Exit Function
                End If
                
                ' Check to make sure that there is not a player in the way
                For i = 1 To MAX_PLAYERS
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = MapNum) And (GetPlayerX(i) = MapNpc(MapNum, MapNpcNum).x + 1) And (GetPlayerY(i) = MapNpc(MapNum, MapNpcNum).y) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    End If
                Next
                
                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (i <> MapNpcNum) And (MapNpc(MapNum, i).Num > 0) And (MapNpc(MapNum, i).x = MapNpc(MapNum, MapNpcNum).x + 1) And (MapNpc(MapNum, i).y = MapNpc(MapNum, MapNpcNum).y) Then
                        CanNpcMove = False
                        Exit Function
                    End If
                Next
            Else
                CanNpcMove = False
            End If
    End Select
End Function

Sub NpcMove(ByVal MapNum As Long, ByVal MapNpcNum As Long, ByVal Dir As Long, ByVal Movement As Long)
Dim packet As String

    ' Check for subscript out of range
    If MapNum <= 0 Or MapNum > MAX_MAPS Or MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Or Dir < DIR_UP Or Dir > DIR_RIGHT Or Movement < 1 Or Movement > 2 Then
        Exit Sub
    End If
    
    MapNpc(MapNum, MapNpcNum).Dir = Dir
    
    Select Case Dir
        Case DIR_UP
            MapNpc(MapNum, MapNpcNum).y = MapNpc(MapNum, MapNpcNum).y - 1
            packet = SNpcMove & SEP_CHAR & MapNpcNum & SEP_CHAR & MapNpc(MapNum, MapNpcNum).x & SEP_CHAR & MapNpc(MapNum, MapNpcNum).y & SEP_CHAR & MapNpc(MapNum, MapNpcNum).Dir & SEP_CHAR & Movement & END_CHAR
            Call SendDataToMap(MapNum, packet)
    
        Case DIR_DOWN
            MapNpc(MapNum, MapNpcNum).y = MapNpc(MapNum, MapNpcNum).y + 1
            packet = SNpcMove & SEP_CHAR & MapNpcNum & SEP_CHAR & MapNpc(MapNum, MapNpcNum).x & SEP_CHAR & MapNpc(MapNum, MapNpcNum).y & SEP_CHAR & MapNpc(MapNum, MapNpcNum).Dir & SEP_CHAR & Movement & END_CHAR
            Call SendDataToMap(MapNum, packet)
    
        Case DIR_LEFT
            MapNpc(MapNum, MapNpcNum).x = MapNpc(MapNum, MapNpcNum).x - 1
            packet = SNpcMove & SEP_CHAR & MapNpcNum & SEP_CHAR & MapNpc(MapNum, MapNpcNum).x & SEP_CHAR & MapNpc(MapNum, MapNpcNum).y & SEP_CHAR & MapNpc(MapNum, MapNpcNum).Dir & SEP_CHAR & Movement & END_CHAR
            Call SendDataToMap(MapNum, packet)
    
        Case DIR_RIGHT
            MapNpc(MapNum, MapNpcNum).x = MapNpc(MapNum, MapNpcNum).x + 1
            packet = SNpcMove & SEP_CHAR & MapNpcNum & SEP_CHAR & MapNpc(MapNum, MapNpcNum).x & SEP_CHAR & MapNpc(MapNum, MapNpcNum).y & SEP_CHAR & MapNpc(MapNum, MapNpcNum).Dir & SEP_CHAR & Movement & END_CHAR
            Call SendDataToMap(MapNum, packet)
    End Select
End Sub

Sub NpcDir(ByVal MapNum As Long, ByVal MapNpcNum As Long, ByVal Dir As Long)
Dim packet As String

    ' Check for subscript out of range
    If MapNum <= 0 Or MapNum > MAX_MAPS Or MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Or Dir < DIR_UP Or Dir > DIR_RIGHT Then
        Exit Sub
    End If
    
    MapNpc(MapNum, MapNpcNum).Dir = Dir
    packet = SNpcDir & SEP_CHAR & MapNpcNum & SEP_CHAR & Dir & END_CHAR
    Call SendDataToMap(MapNum, packet)
End Sub

Function GetTotalMapPlayers(ByVal MapNum As Long) As Long
Dim i As Long, n As Long

    n = 0
    
    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) And GetPlayerMap(i) = MapNum Then
            n = n + 1
        End If
    Next
    
    GetTotalMapPlayers = n
End Function

Function GetNpcMaxVital(ByVal NpcNum As Long, ByVal Vital As Vitals) As Long
Dim x As Long, y As Long

    ' Prevent subscript out of range
    If NpcNum <= 0 Or NpcNum > MAX_NPCS Then
        GetNpcMaxVital = 0
        Exit Function
    End If
    
    Select Case Vital
        Case HP
            x = Npc(NpcNum).Stat(Stats.Strength)
            y = Npc(NpcNum).Stat(Stats.Defense)
            GetNpcMaxVital = x * y
        Case MP
            GetNpcMaxVital = Npc(NpcNum).Stat(Stats.Magic) * 2
        Case SP
            GetNpcMaxVital = Npc(NpcNum).Stat(Stats.Speed) * 2
    End Select
End Function

Function GetNpcVitalRegen(ByVal NpcNum As Long, ByVal Vital As Vitals) As Long
Dim i As Long

    'Prevent subscript out of range
    If NpcNum <= 0 Or NpcNum > MAX_NPCS Then
        GetNpcVitalRegen = 0
        Exit Function
    End If
    
    Select Case Vital
        Case HP
            i = Int(Npc(NpcNum).Stat(Stats.Defense) / 3)
            If i < 1 Then i = 1
                GetNpcVitalRegen = i
        'Case MP
        
        'Case SP
    
    End Select
End Function

Function GetSpellReqLevel(ByVal SpellNum As Long) As Byte
    GetSpellReqLevel = Spell(SpellNum).LevelReq
End Function

Public Sub UpdateHighIndex()
    Call SendDataToAll(SHighIndex & SEP_CHAR & Player_HighIndex & END_CHAR)
End Sub

Public Sub Set_Bomb(ByVal MapNum As Long, ByVal x As Long, ByVal y As Long)
Dim StartX As Long, StartY As Long, StartMap As Long, Distance As Long, i As Long, Index As Long, Maker As String
Dim distancia As String
    StartX = x
    StartY = y
    
    With Bomb(StartX, StartY)
        If .Distance < 1 Or .Maker = vbNullString Or .MakerIndex < 1 Or .Map < 1 Then Exit Sub
    End With
    
    SendDataToMap MapNum, SBomb & SEP_CHAR & 1 & SEP_CHAR & StartX & SEP_CHAR & StartY & END_CHAR
    
    StartMap = MapNum
    Distance = Bomb(StartX, StartY).Distance
    Index = Bomb(StartX, StartY).MakerIndex
    Maker = Bomb(StartX, StartY).Maker
            
    For x = StartX To (StartX - Distance) Step -1
        If x <> StartX Then
            If x < MAX_MAPX Then
                If x > -1 Then
                    If Map(StartMap).Tile(x, StartY).Type <> TILE_TYPE_BLOCKED Then
                        If Map(StartMap).Tile(StartX - 1, StartY).Type = TILE_TYPE_BLOCKED Or x = StartX - Distance Then
                            AttackPlayer Index, StartMap, x, StartY, 1, , Maker
                        Else
                            AttackPlayer Index, StartMap, x, StartY, , , Maker
                        End If
                    Else
                        AttackPlayer Index, StartMap, x, StartY, 1, , Maker, True
                        Exit For
                    End If
                End If
            End If
        End If
    Next

    For x = StartX To (StartX + Distance)
        If x <> StartX Then
            If x < MAX_MAPX Then
                If x > -1 Then
                    If Map(StartMap).Tile(x, StartY).Type <> TILE_TYPE_BLOCKED Then
                        If Map(StartMap).Tile(StartX, StartY).Type = TILE_TYPE_BLOCKED Or x = StartX + Distance Then
                           AttackPlayer Index, StartMap, x, StartY, 2, , Maker
                        Else
                            AttackPlayer Index, StartMap, x, StartY, , , Maker
                        End If
                    Else
                        AttackPlayer Index, StartMap, x, StartY, 2, , Maker, True
                        Exit For
                    End If
                End If
            End If
        End If
    Next
    
    For y = StartY To (StartY - Distance) Step -1
        If y <> StartY Then
            If y < MAX_MAPY Then
                If y > -1 Then
                    If Map(StartMap).Tile(StartX, y).Type <> TILE_TYPE_BLOCKED Then
                        If Map(StartMap).Tile(StartX, StartY).Type = TILE_TYPE_BLOCKED Or y = StartY - Distance Then
                            AttackPlayer Index, StartMap, StartX, y, 1, 1, Maker
                        Else
                            AttackPlayer Index, StartMap, StartX, y, , 1, Maker
                        End If
                    Else
                        AttackPlayer Index, StartMap, StartX, y, 1, 1, Maker, True
                        Exit For
                    End If
                End If
            End If
        End If
    Next
    
    For y = StartY To (StartY + Distance)
        If y <> StartY Then
            If y < MAX_MAPY Then
                If y > -1 Then
                    If Map(StartMap).Tile(StartX, y).Type <> TILE_TYPE_BLOCKED Then
                        If Map(StartMap).Tile(StartX, y + 1).Type = TILE_TYPE_BLOCKED Or y = StartY + Distance Then
                            AttackPlayer Index, StartMap, StartX, y, 2, 1, Maker
                        Else
                            AttackPlayer Index, StartMap, StartX, y, , 1, Maker
                        End If
                    Else
                        AttackPlayer Index, StartMap, StartX, y, 2, 1, Maker, True
                        Exit For
                    End If
                End If
            End If
        End If
    Next
    
    AttackPlayer Index, StartMap, StartX, StartY, 3, , Maker

    If IsPlaying(Bomb(StartX, StartY).MakerIndex) Then
        TempPlayer(Bomb(StartX, StartY).MakerIndex).TotalBombs = TempPlayer(Bomb(StartX, StartY).MakerIndex).TotalBombs - 1
        If TempPlayer(Bomb(StartX, StartY).MakerIndex).TotalBombs < 0 Then TempPlayer(Bomb(StartX, StartY).MakerIndex).TotalBombs = 0
    End If
    
    distancia = TempPlayer(Index).Distance
    
    If Not StartX - 1 < 1 Then
    If Map(StartMap).Tile(StartX - 1, StartY).Type = TILE_TYPE_WALL Then
    Call DestroyWall(MapNum, StartX - 1, StartY)
    End If
    End If
    If Not StartX + 1 > MAX_MAPX Then
    If Map(StartMap).Tile(StartX + 1, StartY).Type = TILE_TYPE_WALL Then
    Call DestroyWall(MapNum, StartX + 1, StartY)
    End If
    End If
    If Not StartY - 1 < 1 Then
    If Map(StartMap).Tile(StartX, StartY - 1).Type = TILE_TYPE_WALL Then
    Call DestroyWall(MapNum, StartX, StartY - 1)
    End If
    End If
    If Not StartY + 1 > MAX_MAPY Then
    If Map(StartMap).Tile(StartX, StartY + 1).Type = TILE_TYPE_WALL Then
    Call DestroyWall(MapNum, StartX, StartY + 1)
    End If
    End If
        
    If distancia >= 2 Then
    If Not StartX - 2 < 1 Then
    If Map(StartMap).Tile(StartX - 2, StartY).Type = TILE_TYPE_WALL Then
    Call DestroyWall(MapNum, StartX - 2, StartY)
    End If
    End If
    If Not StartX + 2 > MAX_MAPX Then
    If Map(StartMap).Tile(StartX + 2, StartY).Type = TILE_TYPE_WALL Then
    Call DestroyWall(MapNum, StartX + 2, StartY)
    End If
    End If
    If Not StartY - 2 < 1 Then
    If Map(StartMap).Tile(StartX, StartY - 2).Type = TILE_TYPE_WALL Then
    Call DestroyWall(MapNum, StartX, StartY - 2)
    End If
    End If
    If Not StartY + 2 > MAX_MAPY Then
    If Map(StartMap).Tile(StartX, StartY + 2).Type = TILE_TYPE_WALL Then
    Call DestroyWall(MapNum, StartX, StartY + 2)
    End If
    End If
    End If
    
    If distancia >= 3 Then
    If Not StartX - 3 < 1 Then
    If Map(StartMap).Tile(StartX - 3, StartY).Type = TILE_TYPE_WALL Then
    Call DestroyWall(MapNum, StartX - 3, StartY)
    End If
    End If
    If Not StartX + 3 > MAX_MAPX Then
    If Map(StartMap).Tile(StartX + 3, StartY).Type = TILE_TYPE_WALL Then
    Call DestroyWall(MapNum, StartX + 3, StartY)
    End If
    End If
    If Not StartY - 3 < 1 Then
    If Map(StartMap).Tile(StartX, StartY - 3).Type = TILE_TYPE_WALL Then
    Call DestroyWall(MapNum, StartX, StartY - 3)
    End If
    End If
    If Not StartY + 3 > MAX_MAPY Then
    If Map(StartMap).Tile(StartX, StartY + 3).Type = TILE_TYPE_WALL Then
    Call DestroyWall(MapNum, StartX, StartY + 3)
    End If
    End If
    End If
    
    If distancia = 4 Then
    If Not StartX - 4 < 1 Then
    If Map(StartMap).Tile(StartX - 4, StartY).Type = TILE_TYPE_WALL Then
    Call DestroyWall(MapNum, StartX - 4, StartY)
    End If
    End If
    If Not StartX + 4 > MAX_MAPX Then
    If Map(StartMap).Tile(StartX + 4, StartY).Type = TILE_TYPE_WALL Then
    Call DestroyWall(MapNum, StartX + 4, StartY)
    End If
    End If
    If Not StartY - 4 < 1 Then
    If Map(StartMap).Tile(StartX, StartY - 4).Type = TILE_TYPE_WALL Then
    Call DestroyWall(MapNum, StartX, StartY - 4)
    End If
    End If
    If Not StartY + 4 > MAX_MAPY Then
    If Map(StartMap).Tile(StartX, StartY + 4).Type = TILE_TYPE_WALL Then
    Call DestroyWall(MapNum, StartX, StartY + 4)
    End If
    End If
    End If
    Bomb(StartX, StartY).Here = False
    Bomb(StartX, StartY).MakerIndex = 0
    Bomb(StartX, StartY).Maker = vbNullString
    Bomb(StartX, StartY).Timer = GetTickCount
    Bomb(StartX, StartY).Distance = 0
    Bomb(StartX, StartY).Map = 0

End Sub

Public Sub DestroyWall(ByVal MapNum As Long, ByVal x As Long, ByVal y As Long)
    If Map(MapNum).Tile(x, y).Type = TILE_TYPE_WALL And Not Wall(x, y).Gone Then
        SendDataToMap MapNum, SWall & SEP_CHAR & x & SEP_CHAR & y & END_CHAR
        Wall(x, y).Gone = True
    End If

End Sub
Public Sub RefreshWall(ByVal MapNum As Long)

        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 1 & SEP_CHAR & 1 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 1 & SEP_CHAR & 2 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 1 & SEP_CHAR & 3 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 1 & SEP_CHAR & 4 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 1 & SEP_CHAR & 5 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 1 & SEP_CHAR & 6 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 1 & SEP_CHAR & 7 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 1 & SEP_CHAR & 8 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 1 & SEP_CHAR & 9 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 1 & SEP_CHAR & 10 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 1 & SEP_CHAR & 11 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 1 & SEP_CHAR & 12 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 2 & SEP_CHAR & 1 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 2 & SEP_CHAR & 2 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 2 & SEP_CHAR & 3 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 2 & SEP_CHAR & 4 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 2 & SEP_CHAR & 5 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 2 & SEP_CHAR & 6 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 2 & SEP_CHAR & 7 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 2 & SEP_CHAR & 8 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 2 & SEP_CHAR & 9 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 2 & SEP_CHAR & 10 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 2 & SEP_CHAR & 11 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 2 & SEP_CHAR & 12 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 3 & SEP_CHAR & 1 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 3 & SEP_CHAR & 2 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 3 & SEP_CHAR & 3 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 3 & SEP_CHAR & 4 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 3 & SEP_CHAR & 5 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 3 & SEP_CHAR & 6 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 3 & SEP_CHAR & 7 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 3 & SEP_CHAR & 8 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 3 & SEP_CHAR & 9 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 3 & SEP_CHAR & 10 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 3 & SEP_CHAR & 11 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 3 & SEP_CHAR & 12 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 4 & SEP_CHAR & 1 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 4 & SEP_CHAR & 2 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 4 & SEP_CHAR & 3 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 4 & SEP_CHAR & 4 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 4 & SEP_CHAR & 5 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 4 & SEP_CHAR & 6 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 4 & SEP_CHAR & 7 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 4 & SEP_CHAR & 8 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 4 & SEP_CHAR & 9 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 4 & SEP_CHAR & 10 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 4 & SEP_CHAR & 11 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 4 & SEP_CHAR & 12 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 5 & SEP_CHAR & 1 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 5 & SEP_CHAR & 2 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 5 & SEP_CHAR & 3 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 5 & SEP_CHAR & 4 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 5 & SEP_CHAR & 5 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 5 & SEP_CHAR & 6 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 5 & SEP_CHAR & 7 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 5 & SEP_CHAR & 8 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 5 & SEP_CHAR & 9 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 5 & SEP_CHAR & 10 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 5 & SEP_CHAR & 11 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 5 & SEP_CHAR & 12 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 6 & SEP_CHAR & 1 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 6 & SEP_CHAR & 2 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 6 & SEP_CHAR & 3 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 6 & SEP_CHAR & 4 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 6 & SEP_CHAR & 5 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 6 & SEP_CHAR & 6 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 6 & SEP_CHAR & 7 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 6 & SEP_CHAR & 8 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 6 & SEP_CHAR & 9 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 6 & SEP_CHAR & 10 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 6 & SEP_CHAR & 11 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 6 & SEP_CHAR & 12 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 7 & SEP_CHAR & 1 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 7 & SEP_CHAR & 2 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 7 & SEP_CHAR & 3 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 7 & SEP_CHAR & 4 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 7 & SEP_CHAR & 5 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 7 & SEP_CHAR & 6 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 7 & SEP_CHAR & 7 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 7 & SEP_CHAR & 8 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 7 & SEP_CHAR & 9 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 7 & SEP_CHAR & 10 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 7 & SEP_CHAR & 11 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 7 & SEP_CHAR & 12 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 8 & SEP_CHAR & 1 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 8 & SEP_CHAR & 2 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 8 & SEP_CHAR & 3 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 8 & SEP_CHAR & 4 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 8 & SEP_CHAR & 5 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 8 & SEP_CHAR & 6 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 8 & SEP_CHAR & 7 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 8 & SEP_CHAR & 8 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 8 & SEP_CHAR & 9 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 8 & SEP_CHAR & 10 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 8 & SEP_CHAR & 11 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 8 & SEP_CHAR & 12 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 9 & SEP_CHAR & 1 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 9 & SEP_CHAR & 2 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 9 & SEP_CHAR & 3 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 9 & SEP_CHAR & 4 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 9 & SEP_CHAR & 5 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 9 & SEP_CHAR & 6 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 9 & SEP_CHAR & 7 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 9 & SEP_CHAR & 8 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 9 & SEP_CHAR & 9 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 9 & SEP_CHAR & 10 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 9 & SEP_CHAR & 11 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 9 & SEP_CHAR & 12 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 10 & SEP_CHAR & 1 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 10 & SEP_CHAR & 2 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 10 & SEP_CHAR & 3 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 10 & SEP_CHAR & 4 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 10 & SEP_CHAR & 5 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 10 & SEP_CHAR & 6 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 10 & SEP_CHAR & 7 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 10 & SEP_CHAR & 8 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 10 & SEP_CHAR & 9 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 10 & SEP_CHAR & 10 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 10 & SEP_CHAR & 11 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 10 & SEP_CHAR & 12 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 11 & SEP_CHAR & 1 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 11 & SEP_CHAR & 2 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 11 & SEP_CHAR & 3 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 11 & SEP_CHAR & 4 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 11 & SEP_CHAR & 5 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 11 & SEP_CHAR & 6 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 11 & SEP_CHAR & 7 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 11 & SEP_CHAR & 8 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 11 & SEP_CHAR & 9 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 11 & SEP_CHAR & 10 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 11 & SEP_CHAR & 11 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 11 & SEP_CHAR & 12 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 12 & SEP_CHAR & 1 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 12 & SEP_CHAR & 2 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 12 & SEP_CHAR & 3 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 12 & SEP_CHAR & 4 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 12 & SEP_CHAR & 5 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 12 & SEP_CHAR & 6 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 12 & SEP_CHAR & 7 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 12 & SEP_CHAR & 8 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 12 & SEP_CHAR & 9 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 12 & SEP_CHAR & 10 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 12 & SEP_CHAR & 11 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 12 & SEP_CHAR & 12 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 13 & SEP_CHAR & 1 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 13 & SEP_CHAR & 2 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 13 & SEP_CHAR & 3 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 13 & SEP_CHAR & 4 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 13 & SEP_CHAR & 5 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 13 & SEP_CHAR & 6 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 13 & SEP_CHAR & 7 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 13 & SEP_CHAR & 8 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 13 & SEP_CHAR & 9 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 13 & SEP_CHAR & 10 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 13 & SEP_CHAR & 11 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 13 & SEP_CHAR & 12 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 14 & SEP_CHAR & 1 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 14 & SEP_CHAR & 2 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 14 & SEP_CHAR & 3 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 14 & SEP_CHAR & 4 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 14 & SEP_CHAR & 5 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 14 & SEP_CHAR & 6 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 14 & SEP_CHAR & 7 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 14 & SEP_CHAR & 8 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 14 & SEP_CHAR & 9 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 14 & SEP_CHAR & 10 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 14 & SEP_CHAR & 11 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 14 & SEP_CHAR & 12 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 15 & SEP_CHAR & 1 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 15 & SEP_CHAR & 2 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 15 & SEP_CHAR & 3 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 15 & SEP_CHAR & 4 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 15 & SEP_CHAR & 5 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 15 & SEP_CHAR & 6 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 15 & SEP_CHAR & 7 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 15 & SEP_CHAR & 8 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 15 & SEP_CHAR & 9 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 15 & SEP_CHAR & 10 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 15 & SEP_CHAR & 11 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 15 & SEP_CHAR & 12 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 16 & SEP_CHAR & 1 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 16 & SEP_CHAR & 2 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 16 & SEP_CHAR & 3 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 16 & SEP_CHAR & 4 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 16 & SEP_CHAR & 5 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 16 & SEP_CHAR & 6 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 16 & SEP_CHAR & 7 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 16 & SEP_CHAR & 8 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 16 & SEP_CHAR & 9 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 16 & SEP_CHAR & 10 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 16 & SEP_CHAR & 11 & END_CHAR
        SendDataToMap MapNum, SRefreshWall & SEP_CHAR & 16 & SEP_CHAR & 12 & END_CHAR

        
        Wall(1, 1).Gone = False
        Wall(1, 2).Gone = False
        Wall(1, 3).Gone = False
        Wall(1, 4).Gone = False
        Wall(1, 5).Gone = False
        Wall(1, 6).Gone = False
        Wall(1, 7).Gone = False
        Wall(1, 8).Gone = False
        Wall(1, 9).Gone = False
        Wall(1, 10).Gone = False
        Wall(1, 11).Gone = False
        Wall(1, 12).Gone = False
        Wall(2, 1).Gone = False
        Wall(2, 2).Gone = False
        Wall(2, 3).Gone = False
        Wall(2, 4).Gone = False
        Wall(2, 5).Gone = False
        Wall(2, 6).Gone = False
        Wall(2, 7).Gone = False
        Wall(2, 8).Gone = False
        Wall(2, 9).Gone = False
        Wall(2, 10).Gone = False
        Wall(2, 11).Gone = False
        Wall(2, 12).Gone = False
        Wall(3, 1).Gone = False
        Wall(3, 2).Gone = False
        Wall(3, 3).Gone = False
        Wall(3, 4).Gone = False
        Wall(3, 5).Gone = False
        Wall(3, 6).Gone = False
        Wall(3, 7).Gone = False
        Wall(3, 8).Gone = False
        Wall(3, 9).Gone = False
        Wall(3, 10).Gone = False
        Wall(3, 11).Gone = False
        Wall(3, 12).Gone = False
        Wall(4, 1).Gone = False
        Wall(4, 2).Gone = False
        Wall(4, 3).Gone = False
        Wall(4, 4).Gone = False
        Wall(4, 5).Gone = False
        Wall(4, 6).Gone = False
        Wall(4, 7).Gone = False
        Wall(4, 8).Gone = False
        Wall(4, 9).Gone = False
        Wall(4, 10).Gone = False
        Wall(4, 11).Gone = False
        Wall(4, 12).Gone = False
        Wall(5, 1).Gone = False
        Wall(5, 2).Gone = False
        Wall(5, 3).Gone = False
        Wall(5, 4).Gone = False
        Wall(5, 5).Gone = False
        Wall(5, 6).Gone = False
        Wall(5, 7).Gone = False
        Wall(5, 8).Gone = False
        Wall(5, 9).Gone = False
        Wall(5, 10).Gone = False
        Wall(5, 11).Gone = False
        Wall(5, 12).Gone = False
        Wall(6, 1).Gone = False
        Wall(6, 2).Gone = False
        Wall(6, 3).Gone = False
        Wall(6, 4).Gone = False
        Wall(6, 5).Gone = False
        Wall(6, 6).Gone = False
        Wall(6, 7).Gone = False
        Wall(6, 8).Gone = False
        Wall(6, 9).Gone = False
        Wall(6, 10).Gone = False
        Wall(6, 11).Gone = False
        Wall(6, 12).Gone = False
        Wall(7, 1).Gone = False
        Wall(7, 2).Gone = False
        Wall(7, 3).Gone = False
        Wall(7, 4).Gone = False
        Wall(7, 5).Gone = False
        Wall(7, 6).Gone = False
        Wall(7, 7).Gone = False
        Wall(7, 8).Gone = False
        Wall(7, 9).Gone = False
        Wall(7, 10).Gone = False
        Wall(7, 11).Gone = False
        Wall(7, 12).Gone = False
        Wall(8, 1).Gone = False
        Wall(8, 2).Gone = False
        Wall(8, 3).Gone = False
        Wall(8, 4).Gone = False
        Wall(8, 5).Gone = False
        Wall(8, 6).Gone = False
        Wall(8, 7).Gone = False
        Wall(8, 8).Gone = False
        Wall(8, 9).Gone = False
        Wall(8, 10).Gone = False
        Wall(8, 11).Gone = False
        Wall(8, 12).Gone = False
        Wall(9, 1).Gone = False
        Wall(9, 2).Gone = False
        Wall(9, 3).Gone = False
        Wall(9, 4).Gone = False
        Wall(9, 5).Gone = False
        Wall(9, 6).Gone = False
        Wall(9, 7).Gone = False
        Wall(9, 8).Gone = False
        Wall(9, 9).Gone = False
        Wall(9, 10).Gone = False
        Wall(9, 11).Gone = False
        Wall(9, 12).Gone = False
        Wall(10, 1).Gone = False
        Wall(10, 2).Gone = False
        Wall(10, 3).Gone = False
        Wall(10, 4).Gone = False
        Wall(10, 5).Gone = False
        Wall(10, 6).Gone = False
        Wall(10, 7).Gone = False
        Wall(10, 8).Gone = False
        Wall(10, 9).Gone = False
        Wall(10, 10).Gone = False
        Wall(10, 11).Gone = False
        Wall(10, 12).Gone = False
        Wall(11, 12).Gone = False
        Wall(11, 2).Gone = False
        Wall(11, 3).Gone = False
        Wall(11, 4).Gone = False
        Wall(11, 5).Gone = False
        Wall(11, 6).Gone = False
        Wall(11, 7).Gone = False
        Wall(11, 8).Gone = False
        Wall(11, 9).Gone = False
        Wall(11, 10).Gone = False
        Wall(11, 11).Gone = False
        Wall(11, 12).Gone = False
        Wall(12, 1).Gone = False
        Wall(12, 2).Gone = False
        Wall(12, 3).Gone = False
        Wall(12, 4).Gone = False
        Wall(12, 5).Gone = False
        Wall(12, 6).Gone = False
        Wall(12, 7).Gone = False
        Wall(12, 8).Gone = False
        Wall(12, 9).Gone = False
        Wall(12, 10).Gone = False
        Wall(12, 11).Gone = False
        Wall(12, 12).Gone = False
        Wall(13, 1).Gone = False
        Wall(13, 2).Gone = False
        Wall(13, 3).Gone = False
        Wall(13, 4).Gone = False
        Wall(13, 5).Gone = False
        Wall(13, 6).Gone = False
        Wall(13, 7).Gone = False
        Wall(13, 8).Gone = False
        Wall(13, 9).Gone = False
        Wall(13, 10).Gone = False
        Wall(13, 11).Gone = False
        Wall(13, 12).Gone = False
        Wall(14, 1).Gone = False
        Wall(14, 2).Gone = False
        Wall(14, 3).Gone = False
        Wall(14, 4).Gone = False
        Wall(14, 5).Gone = False
        Wall(14, 6).Gone = False
        Wall(14, 7).Gone = False
        Wall(14, 8).Gone = False
        Wall(14, 9).Gone = False
        Wall(14, 10).Gone = False
        Wall(14, 11).Gone = False
        Wall(14, 12).Gone = False
        Wall(15, 1).Gone = False
        Wall(15, 2).Gone = False
        Wall(15, 3).Gone = False
        Wall(15, 4).Gone = False
        Wall(15, 5).Gone = False
        Wall(15, 6).Gone = False
        Wall(15, 7).Gone = False
        Wall(15, 8).Gone = False
        Wall(15, 9).Gone = False
        Wall(15, 10).Gone = False
        Wall(15, 11).Gone = False
        Wall(15, 12).Gone = False
        Wall(16, 1).Gone = False
        Wall(16, 2).Gone = False
        Wall(16, 3).Gone = False
        Wall(16, 4).Gone = False
        Wall(16, 5).Gone = False
        Wall(16, 6).Gone = False
        Wall(16, 7).Gone = False
        Wall(16, 8).Gone = False
        Wall(16, 9).Gone = False
        Wall(16, 10).Gone = False
        Wall(16, 11).Gone = False
        Wall(16, 12).Gone = False
        
        

End Sub

Function Random(Lowerbound As Integer, Upperbound As Integer) As Integer
    Random = Int((Upperbound - Lowerbound + 1) * Rnd) + Lowerbound
End Function
