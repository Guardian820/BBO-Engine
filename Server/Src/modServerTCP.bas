Attribute VB_Name = "modServerTCP"
Option Explicit

' ******************************************
' **              BMO Source              **
' ******************************************

Sub UpdateCaption()
    frmServer.Caption = "Bomberman Online Server <IP " & frmServer.Socket(0).LocalIP & "" & CStr(frmServer.Socket(0).LocalPort) & "> (" & TotalOnlinePlayers & ")"
End Sub

Function IsConnected(ByVal Index As Long) As Boolean
    If frmServer.Socket(Index).State = sckConnected Then
        IsConnected = True
    Else
        IsConnected = False
    End If
End Function

Function IsPlaying(ByVal Index As Long) As Boolean

    If Index <= 0 Then Exit Function
    
    If IsConnected(Index) And TempPlayer(Index).InGame = True Then
        IsPlaying = True
    Else
        IsPlaying = False
    End If
End Function

Function IsLoggedIn(ByVal Index As Long) As Boolean
    If IsConnected(Index) And LenB(Trim$(Player(Index).Login)) > 0 Then
        IsLoggedIn = True
    Else
        IsLoggedIn = False
    End If
End Function

Function IsMultiAccounts(ByVal Login As String) As Boolean
Dim i As Long

    IsMultiAccounts = False
    For i = 1 To MAX_PLAYERS
        If IsConnected(i) And LCase$(Trim$(Player(i).Login)) = LCase$(Login) Then
            IsMultiAccounts = True
            Exit Function
        End If
    Next
End Function

Function IsMultiIPOnline(ByVal IP As String) As Boolean
Dim i As Long
Dim n As Long

    n = 0
    IsMultiIPOnline = False
    For i = 1 To MAX_PLAYERS
        If IsConnected(i) And Trim$(GetPlayerIP(i)) = IP Then
            n = n + 1
            
            If (n > 1) Then
                IsMultiIPOnline = True
                Exit Function
            End If
        End If
    Next
End Function

Function IsBanned(ByVal IP As String) As Boolean
Dim FileName As String, fIP As String, fName As String
Dim f As Long

    IsBanned = False
    
    FileName = App.Path & "\data\banlist.txt"
    
    ' Check if file exists
    If Not FileExist("data\banlist.txt") Then
        f = FreeFile
        Open FileName For Output As #f
        Close #f
    End If
    
    f = FreeFile
    Open FileName For Input As #f
    
    Do While Not EOF(f)
        Input #f, fIP
        Input #f, fName
    
        ' Is banned?
        If Trim$(LCase$(fIP)) = Trim$(LCase$(Mid$(IP, 1, Len(fIP)))) Then
            IsBanned = True
            Close #f
            Exit Function
        End If
    Loop
    
    Close #f
End Function

Sub SendDataTo(ByVal Index As Long, ByVal Data As String)
    If IsConnected(Index) Then
        frmServer.Socket(Index).SendData Data
        DoEvents
    End If
End Sub

Sub SendDataToAll(ByVal Data As String)
Dim i As Long

    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            Call SendDataTo(i, Data)
        End If
    Next
End Sub

Sub SendDataToAllBut(ByVal Index As Long, ByVal Data As String)
Dim i As Long

    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) And i <> Index Then
            Call SendDataTo(i, Data)
        End If
    Next
End Sub

Sub SendDataToMap(ByVal MapNum As Long, ByVal Data As String)
Dim i As Long

    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            If GetPlayerMap(i) = MapNum Then
                Call SendDataTo(i, Data)
            End If
        End If
    Next
End Sub

Sub SendDataToMapBut(ByVal Index As Long, ByVal MapNum As Long, ByVal Data As String)
Dim i As Long

    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            If GetPlayerMap(i) = MapNum Then
                If i <> Index Then
                    Call SendDataTo(i, Data)
                End If
            End If
        End If
    Next
End Sub

Sub GlobalMsg(ByVal Msg As String, ByVal Color As Byte)
Dim packet As String

    packet = SGlobalMsg & SEP_CHAR & Msg & SEP_CHAR & Color & END_CHAR
    
    Call SendDataToAll(packet)
End Sub

Sub AdminMsg(ByVal Msg As String, ByVal Color As Byte)
Dim packet As String
Dim i As Long

    packet = SAdminMsg & SEP_CHAR & Msg & SEP_CHAR & Color & END_CHAR
    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) And GetPlayerAccess(i) > 0 Then
            Call SendDataTo(i, packet)
        End If
    Next
End Sub

Sub PlayerMsg(ByVal Index As Long, ByVal Msg As String, ByVal Color As Byte)
Dim packet As String

    packet = SPlayerMsg & SEP_CHAR & Msg & SEP_CHAR & Color & END_CHAR
    
    Call SendDataTo(Index, packet)
End Sub

Sub MapMsg(ByVal MapNum As Long, ByVal Msg As String, ByVal Color As Byte)
Dim packet As String
    
    packet = SMapMsg & SEP_CHAR & Msg & SEP_CHAR & Color & END_CHAR
    
    Call SendDataToMap(MapNum, packet)
End Sub

Sub AlertMsg(ByVal Index As Long, ByVal Msg As String)
Dim packet As String

    packet = SAlertMsg & SEP_CHAR & Msg & END_CHAR
    
    Call SendDataTo(Index, packet)
    Call CloseSocket(Index)
End Sub

Sub HackingAttempt(ByVal Index As Long, ByVal Reason As String)
    If Index > 0 Then
        If IsPlaying(Index) Then
            Call GlobalMsg(GetPlayerLogin(Index) & "/" & GetPlayerName(Index) & " foi retirado pelo seguinte motivo: (" & Reason & ")", Red)
        End If
    
        Call AlertMsg(Index, "Você perdeu a conexão com " & GAME_NAME & ".")
    End If
End Sub

Sub AcceptConnection(ByVal Index As Long, ByVal SocketId As Long)
Dim i As Long

    If (Index = 0) Then
        i = FindOpenPlayerSlot
        
        If i <> 0 Then
            ' we can connect them
            frmServer.Socket(i).Close
            frmServer.Socket(i).Accept SocketId
            Call SocketConnected(i)
        End If
    End If
End Sub

Sub SocketConnected(ByVal Index As Long)
Dim i As Long
    If Index <> 0 Then
        ' make sure they're not banned
            If Not IsBanned(GetPlayerIP(Index)) Then
                Call TextAdd(frmServer.txtText, "Recebeu conexão de " & GetPlayerIP(Index) & ".")
            Else
                Call AlertMsg(Index, "Você foi banido do " & GAME_NAME & ", e não pode jogar.")
            End If
        ' re-set the high index
        Player_HighIndex = 0
        For i = MAX_PLAYERS To 1 Step -1
            If IsConnected(i) Then
                Player_HighIndex = i
                Exit For
            End If
        Next
        ' send the new highindex to all logged in players
        UpdateHighIndex
    End If
End Sub

Sub IncomingData(ByVal Index As Long, ByVal DataLength As Long)
Dim Buffer As String
Dim packet As String
Dim Start As Integer

    If Index > 0 Then
    
        frmServer.Socket(Index).GetData Buffer, vbString, DataLength
            
        TempPlayer(Index).Buffer = TempPlayer(Index).Buffer & Buffer
        
        Start = InStr(TempPlayer(Index).Buffer, END_CHAR)
        Do While Start > 0
            packet = Mid$(TempPlayer(Index).Buffer, 1, Start - 1)
            TempPlayer(Index).Buffer = Mid$(TempPlayer(Index).Buffer, Start + 1, Len(TempPlayer(Index).Buffer))
            TempPlayer(Index).DataPackets = TempPlayer(Index).DataPackets + 1
            Start = InStr(TempPlayer(Index).Buffer, END_CHAR)
            
            If Len(packet) > 0 Then
                Call HandleData(Index, packet)
            End If
        Loop

        ' Check if elapsed time has passed
        TempPlayer(Index).DataBytes = TempPlayer(Index).DataBytes + DataLength
        
        If GetTickCount >= TempPlayer(Index).DataTimer + 1000 Then
            TempPlayer(Index).DataTimer = GetTickCount
            TempPlayer(Index).DataBytes = 0
            TempPlayer(Index).DataPackets = 0
            Exit Sub
        End If
        
        If GetPlayerAccess(Index) <= ADMIN_MONITOR Then
            ' Check for data flooding
            If TempPlayer(Index).DataBytes > 1000 Then
                Call HackingAttempt(Index, "Data Flooding")
                Exit Sub
            End If
            
            ' Check for packet flooding
            'If TempPlayer(Index).DataPackets > 25 Then
            '    Call HackingAttempt(Index, "Packet Flooding")
            '    Exit Sub
            'End If
        End If
        
    End If
End Sub

Sub CloseSocket(ByVal Index As Long)

    If Index > 0 Then
    
        If GetPlayerMap(Index) > 0 Then
            If Map(GetPlayerMap(Index)).Moral = MAP_MORAL_FFA Then
                If TempMap(GetPlayerMap(Index)).InProgress Then
                    If TempPlayer(Index).Watching Then
                        SetPlayerDeaths Index, GetPlayerDeaths(Index) + 1
                        GlobalMsg GetPlayerName(Index) & " ganhou um death por sair antes.", BrightRed
                    End If
                End If
            End If
        End If
        
        Call TextAdd(frmServer.txtText, "Connection from " & GetPlayerIP(Index) & " has been terminated.")
        
        frmServer.Socket(Index).Close
        
        Call LeftGame(Index)
        Call UpdateCaption
        'Call ClearPlayer(Index)
    End If
End Sub

Public Sub MapCache_Create(ByVal MapNum As Long)
    Dim MapData As String
    Dim X As Long
    Dim Y As Long

    MapData = SMapData & SEP_CHAR & MapNum & SEP_CHAR & Trim$(Map(MapNum).Name) & SEP_CHAR & Map(MapNum).Revision & SEP_CHAR & Map(MapNum).PlayerLimit & SEP_CHAR & Map(MapNum).Moral & SEP_CHAR & Map(MapNum).TileSet & SEP_CHAR & Map(MapNum).Up & SEP_CHAR & Map(MapNum).Down & SEP_CHAR & Map(MapNum).Left & SEP_CHAR & Map(MapNum).Right & SEP_CHAR & Map(MapNum).Music & SEP_CHAR & Map(MapNum).BootMap & SEP_CHAR & Map(MapNum).BootX & SEP_CHAR & Map(MapNum).BootY & SEP_CHAR & Map(MapNum).Shop
    
    For X = 0 To MAX_MAPX
        For Y = 0 To MAX_MAPY
            With Map(MapNum).Tile(X, Y)
                MapData = MapData & SEP_CHAR & .Ground & SEP_CHAR & .Mask & SEP_CHAR & .Anim & SEP_CHAR & .Fringe & SEP_CHAR & .Type & SEP_CHAR & .Data1 & SEP_CHAR & .Data2 & SEP_CHAR & .Data3
            End With
        Next
    Next
    
    For X = 1 To MAX_MAP_NPCS
        MapData = MapData & SEP_CHAR & Map(MapNum).Npc(X)
    Next
    
    MapData = MapData & END_CHAR

    MapCache(MapNum) = MapData
End Sub

' ******************************
' ** Outcoming Server Packets **
' ******************************

Sub SendWhosOnline(ByVal Index As Long)
Dim s As String
Dim n As Long, i As Long
Dim RoomName(1 To MAX_MAPS) As String

    For i = 1 To MAX_MAPS
        If i = 100 Then
            RoomName(i) = "Lobby"
        Else
            RoomName(i) = "Sala " & i
        End If
    Next

    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            If i <> Index Then
                s = s & "[" & RoomName(GetPlayerMap(i)) & "] " & GetPlayerName(i) & ", "
                n = n + 1
            End If
        End If
    Next
    
    If n = 0 Then
        s = "Status do servidor: Vazio."
    Else
        s = Mid$(s, 1, Len(s) - 2)
        s = "Há " & n & " jogadores online,entre eles: " & s & "."
    End If
        
    Call PlayerMsg(Index, s, WhoColor)
End Sub

Sub SendChars(ByVal Index As Long)
Dim packet As String
Dim i As Long
    
    packet = SAllChars
    For i = 1 To MAX_CHARS
        packet = packet & SEP_CHAR & Trim$(Player(Index).Char(i).Name) & SEP_CHAR & Trim$(Class(Player(Index).Char(i).Class).Name) & SEP_CHAR & Player(Index).Char(i).Level
    Next
    
    packet = packet & END_CHAR
    
    Call SendDataTo(Index, packet)
End Sub

Sub SendJoinMap(ByVal Index As Long)
Dim packet As String
Dim i As Long
    
    ' Send all players on current map to index
    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            If i <> Index Then
                If GetPlayerMap(i) = GetPlayerMap(Index) Then
                    packet = SPlayerData & SEP_CHAR & i & SEP_CHAR & GetPlayerName(i) & SEP_CHAR & GetPlayerSprite(i) & SEP_CHAR & GetPlayerMap(i) & SEP_CHAR & GetPlayerX(i) & SEP_CHAR & GetPlayerY(i) & SEP_CHAR & GetPlayerDir(i) & SEP_CHAR & GetPlayerAccess(i) & SEP_CHAR & GetPlayerPK(i) & SEP_CHAR & GetPlayerKills(i) & SEP_CHAR & GetPlayerDeaths(i) & SEP_CHAR & GetPlayerBPoints(i) & SEP_CHAR & GetPlayerBCash(i) & SEP_CHAR & vbNullString & END_CHAR
                    Call SendDataTo(Index, packet)
                End If
            End If
        End If
    Next
    
    ' Send index's player data to everyone on the map including himself
    packet = SPlayerData & SEP_CHAR & Index & SEP_CHAR & GetPlayerName(Index) & SEP_CHAR & GetPlayerSprite(Index) & SEP_CHAR & GetPlayerMap(Index) & SEP_CHAR & GetPlayerX(Index) & SEP_CHAR & GetPlayerY(Index) & SEP_CHAR & GetPlayerDir(Index) & SEP_CHAR & GetPlayerAccess(Index) & SEP_CHAR & GetPlayerPK(Index) & SEP_CHAR & GetPlayerKills(Index) & SEP_CHAR & GetPlayerDeaths(Index) & SEP_CHAR & GetPlayerBPoints(Index) & SEP_CHAR & GetPlayerBCash(Index) & SEP_CHAR & vbNullString & END_CHAR
    Call SendDataToMap(GetPlayerMap(Index), packet)
End Sub

Sub SendLeaveMap(ByVal Index As Long, ByVal MapNum As Long)
Dim packet As String

    'Packet = SPlayerData & SEP_CHAR & Index & SEP_CHAR & GetPlayerName(Index) & SEP_CHAR & GetPlayerSprite(Index) & SEP_CHAR & 0 & SEP_CHAR & GetPlayerX(Index) & SEP_CHAR & GetPlayerY(Index) & SEP_CHAR & GetPlayerDir(Index) & SEP_CHAR & GetPlayerAccess(Index) & SEP_CHAR & GetPlayerPK(Index) & END_CHAR
    packet = SLeft & SEP_CHAR & Index & END_CHAR
    Call SendDataToMapBut(Index, MapNum, packet)
End Sub

Sub SendPlayerData(ByVal Index As Long)
Dim packet As String

    SendDataToMap GetPlayerMap(Index), SDieing & SEP_CHAR & Index & SEP_CHAR & 0 & END_CHAR

    ' Send index's player data to everyone including himself on th emap
    packet = SPlayerData & SEP_CHAR & Index & SEP_CHAR & GetPlayerName(Index) & SEP_CHAR & GetPlayerSprite(Index) & SEP_CHAR & GetPlayerMap(Index) & SEP_CHAR & GetPlayerX(Index) & SEP_CHAR & GetPlayerY(Index) & SEP_CHAR & GetPlayerDir(Index) & SEP_CHAR & GetPlayerAccess(Index) & SEP_CHAR & GetPlayerPK(Index) & SEP_CHAR & GetPlayerKills(Index) & SEP_CHAR & GetPlayerDeaths(Index) & SEP_CHAR & GetPlayerBPoints(Index) & SEP_CHAR & GetPlayerBCash(Index) & SEP_CHAR & vbNullString & END_CHAR
    Call SendDataToMap(GetPlayerMap(Index), packet)
End Sub

Public Sub SendMap(ByVal Index As Long, ByVal MapNum As Long)
    Call SendDataTo(Index, MapCache(MapNum))
End Sub

Sub SendMapItemsTo(ByVal Index As Long, ByVal MapNum As Long)
Dim packet As String
Dim i As Long

    packet = SMapItemData
    For i = 1 To MAX_MAP_ITEMS
        packet = packet & SEP_CHAR & MapItem(MapNum, i).Num & SEP_CHAR & MapItem(MapNum, i).Value & SEP_CHAR & MapItem(MapNum, i).Dur & SEP_CHAR & MapItem(MapNum, i).X & SEP_CHAR & MapItem(MapNum, i).Y
    Next
    packet = packet & END_CHAR
    
    Call SendDataTo(Index, packet)
End Sub

Sub SendMapItemsToAll(ByVal MapNum As Long)
Dim packet As String
Dim i As Long

    packet = SMapItemData
    For i = 1 To MAX_MAP_ITEMS
        packet = packet & SEP_CHAR & MapItem(MapNum, i).Num & SEP_CHAR & MapItem(MapNum, i).Value & SEP_CHAR & MapItem(MapNum, i).Dur & SEP_CHAR & MapItem(MapNum, i).X & SEP_CHAR & MapItem(MapNum, i).Y
    Next
    packet = packet & END_CHAR
    
    Call SendDataToMap(MapNum, packet)
End Sub

Sub SendMapNpcsTo(ByVal Index As Long, ByVal MapNum As Long)
Dim packet As String
Dim i As Long

    packet = SMapNpcData
    For i = 1 To MAX_MAP_NPCS
        packet = packet & SEP_CHAR & MapNpc(MapNum, i).Num & SEP_CHAR & MapNpc(MapNum, i).X & SEP_CHAR & MapNpc(MapNum, i).Y & SEP_CHAR & MapNpc(MapNum, i).Dir
    Next
    packet = packet & END_CHAR
    
    Call SendDataTo(Index, packet)
End Sub

Sub SendMapNpcsToMap(ByVal MapNum As Long)
Dim packet As String
Dim i As Long

    packet = SMapNpcData
    For i = 1 To MAX_MAP_NPCS
        packet = packet & SEP_CHAR & MapNpc(MapNum, i).Num & SEP_CHAR & MapNpc(MapNum, i).X & SEP_CHAR & MapNpc(MapNum, i).Y & SEP_CHAR & MapNpc(MapNum, i).Dir
    Next
    packet = packet & END_CHAR
    
    Call SendDataToMap(MapNum, packet)
End Sub

Sub SendItems(ByVal Index As Long)
Dim i As Long
Dim p As Long

    For i = 1 To MAX_ITEMS
    For p = 1 To MAX_PLAYERS
        If LenB(Trim$(Item(i).Name)) > 0 Then
            Call SendUpdateItemTo(p, i)
        End If
    Next
    Next
End Sub

Sub SendNpcs(ByVal Index As Long)
Dim i As Long

    For i = 1 To MAX_NPCS
        If LenB(Trim$(Npc(i).Name)) > 0 Then
            Call SendUpdateNpcTo(Index, i)
        End If
    Next
End Sub

Sub SendInventory(ByVal Index As Long)
Dim packet As String
Dim i As Long

    packet = SPlayerInv & SEP_CHAR
    For i = 1 To MAX_INV
        packet = packet & GetPlayerInvItemNum(Index, i) & SEP_CHAR & GetPlayerInvItemValue(Index, i) & SEP_CHAR & GetPlayerInvItemDur(Index, i) & SEP_CHAR
    Next
    packet = packet & END_CHAR
    
    Call SendDataTo(Index, packet)
End Sub

Sub SendInventoryUpdate(ByVal Index As Long, ByVal InvSlot As Long)
Dim packet As String
    
    packet = SPlayerInvUpdate & SEP_CHAR & InvSlot & SEP_CHAR & GetPlayerInvItemNum(Index, InvSlot) & SEP_CHAR & GetPlayerInvItemValue(Index, InvSlot) & SEP_CHAR & GetPlayerInvItemDur(Index, InvSlot) & END_CHAR
    Call SendDataTo(Index, packet)
End Sub

Sub SendWornEquipment(ByVal Index As Long)
Dim packet As String
    
    packet = SPlayerWornEq & SEP_CHAR & GetPlayerEquipmentSlot(Index, Armor) & SEP_CHAR & GetPlayerEquipmentSlot(Index, Weapon) & SEP_CHAR & GetPlayerEquipmentSlot(Index, Helmet) & SEP_CHAR & GetPlayerEquipmentSlot(Index, Shield) & END_CHAR
    Call SendDataTo(Index, packet)
End Sub

Sub SendVital(ByVal Index As Long, ByVal Vital As Vitals)
Dim packet As String
    
    Select Case Vital
        Case HP
            packet = SPlayerHp & SEP_CHAR & GetPlayerMaxVital(Index, Vitals.HP) & SEP_CHAR & GetPlayerVital(Index, Vitals.HP) & END_CHAR
        Case MP
            packet = SPlayerMp & SEP_CHAR & GetPlayerMaxVital(Index, Vitals.MP) & SEP_CHAR & GetPlayerVital(Index, Vitals.MP) & END_CHAR
        Case SP
            packet = SPlayerSp & SEP_CHAR & GetPlayerMaxVital(Index, Vitals.SP) & SEP_CHAR & GetPlayerVital(Index, Vitals.SP) & END_CHAR
    End Select
    
    Call SendDataTo(Index, packet)
End Sub

Sub SendStats(ByVal Index As Long)
Dim packet As String
    
    packet = SPlayerStats & SEP_CHAR & GetPlayerStat(Index, Stats.Strength) & SEP_CHAR & GetPlayerStat(Index, Stats.Defense) & SEP_CHAR & GetPlayerStat(Index, Stats.Speed) & SEP_CHAR & GetPlayerStat(Index, Stats.Magic) & END_CHAR
    Call SendDataTo(Index, packet)
End Sub

Sub SendWelcome(ByVal Index As Long)

    ' Send them welcome
    Call PlayerMsg(Index, "Evento: Não há eventos atualmente!", Cyan)
    Call PlayerMsg(Index, "Compre cash e aumente o poder de suas bombas,visite o site **************!", Cyan)
    
    ' Send them MOTD
    If LenB(MOTD) > 0 Then
        Call PlayerMsg(Index, "MDD: " & MOTD, Cyan)
    End If
    
    ' Send whos online
    Call SendWhosOnline(Index)
    
End Sub

Sub SendClasses(ByVal Index As Long)
Dim packet As String
Dim i As Long

    packet = SClassesData & SEP_CHAR & Max_Classes
    For i = 1 To Max_Classes
        packet = packet & SEP_CHAR & GetClassName(i) & SEP_CHAR & GetClassMaxVital(i, Vitals.HP) & SEP_CHAR & GetClassMaxVital(i, Vitals.MP) & SEP_CHAR & GetClassMaxVital(i, Vitals.SP) & SEP_CHAR & Class(i).Stat(Stats.Strength) & SEP_CHAR & Class(i).Stat(Stats.Defense) & SEP_CHAR & Class(i).Stat(Stats.Speed) & SEP_CHAR & Class(i).Stat(Stats.Magic)
    Next
    packet = packet & END_CHAR
    
    Call SendDataTo(Index, packet)
End Sub

Sub SendNewCharClasses(ByVal Index As Long)
Dim packet As String
Dim i As Long

    packet = SNewCharClasses & SEP_CHAR & Max_Classes
    For i = 1 To Max_Classes
        packet = packet & SEP_CHAR & GetClassName(i) & SEP_CHAR & GetClassMaxVital(i, Vitals.HP) & SEP_CHAR & GetClassMaxVital(i, Vitals.MP) & SEP_CHAR & GetClassMaxVital(i, Vitals.SP) & SEP_CHAR & Class(i).Stat(Stats.Strength) & SEP_CHAR & Class(i).Stat(Stats.Defense) & SEP_CHAR & Class(i).Stat(Stats.Speed) & SEP_CHAR & Class(i).Stat(Stats.Magic)
    Next
    packet = packet & END_CHAR
    
    Call SendDataTo(Index, packet)
End Sub

Sub SendLeftGame(ByVal Index As Long)
'Dim packet As String

    'packet = SPlayerData & SEP_CHAR & Index & SEP_CHAR & vbNullString & SEP_CHAR & 0 & SEP_CHAR & 0 & SEP_CHAR & 0 & SEP_CHAR & 0 & SEP_CHAR & 0 & SEP_CHAR & 0 & SEP_CHAR & 0 & SEP_CHAR & 0 & SEP_CHAR & 0 & SEP_CHAR & 0 & SEP_CHAR & 0 & SEP_CHAR & GetPlayerName(Index) & END_CHAR
    'Packet = SLeft & SEP_CHAR & Index & END_CHAR
    'Call SendDataToAllBut(Index, packet)
End Sub

Sub SendPlayerXY(ByVal Index As Long)
Dim packet As String

    packet = SPlayerXY & SEP_CHAR & GetPlayerX(Index) & SEP_CHAR & GetPlayerY(Index) & END_CHAR
    Call SendDataTo(Index, packet)
End Sub

Sub SendUpdateItemToAll(ByVal ItemNum As Long)
Dim packet As String

    packet = SUpdateItem & SEP_CHAR & ItemNum & SEP_CHAR & Trim$(Item(ItemNum).Name) & SEP_CHAR & Item(ItemNum).Pic & SEP_CHAR & Item(ItemNum).Type & SEP_CHAR & Item(ItemNum).Data1 & SEP_CHAR & Item(ItemNum).Data2 & SEP_CHAR & Item(ItemNum).Data3 & END_CHAR
    Call SendDataToAll(packet)
End Sub

Sub SendUpdateItemTo(ByVal Index As Long, ByVal ItemNum As Long)
Dim packet As String

    packet = SUpdateItem & SEP_CHAR & ItemNum & SEP_CHAR & Trim$(Item(ItemNum).Name) & SEP_CHAR & Item(ItemNum).Pic & SEP_CHAR & Item(ItemNum).Type & SEP_CHAR & Item(ItemNum).Data1 & SEP_CHAR & Item(ItemNum).Data2 & SEP_CHAR & Item(ItemNum).Data3 & END_CHAR
    Call SendDataTo(Index, packet)
End Sub

Sub SendEditItemTo(ByVal Index As Long, ByVal ItemNum As Long)
Dim packet As String

    packet = SEditItem & SEP_CHAR & ItemNum & SEP_CHAR & Trim$(Item(ItemNum).Name) & SEP_CHAR & Item(ItemNum).Pic & SEP_CHAR & Item(ItemNum).Type & SEP_CHAR & Item(ItemNum).Data1 & SEP_CHAR & Item(ItemNum).Data2 & SEP_CHAR & Item(ItemNum).Data3 & END_CHAR
    Call SendDataTo(Index, packet)
End Sub

Sub SendUpdateNpcToAll(ByVal NpcNum As Long)
Dim packet As String

    'Packet = SUpdateNpc & SEP_CHAR & NpcNum & SEP_CHAR & Trim$(Npc(NpcNum).Name) & SEP_CHAR & Npc(NpcNum).Sprite & END_CHAR
    Call SendDataToAll(packet)
End Sub

Sub SendUpdateNpcTo(ByVal Index As Long, ByVal NpcNum As Long)
Dim packet As String

    'Packet = SUpdateNpc & SEP_CHAR & NpcNum & SEP_CHAR & Trim$(Npc(NpcNum).Name) & SEP_CHAR & Npc(NpcNum).Sprite & END_CHAR
    Call SendDataTo(Index, packet)
End Sub

Sub SendEditNpcTo(ByVal Index As Long, ByVal NpcNum As Long)
Dim packet As String

    packet = SEditNpc & SEP_CHAR & NpcNum & SEP_CHAR & Trim$(Npc(NpcNum).Name) & SEP_CHAR & Trim$(Npc(NpcNum).AttackSay) & SEP_CHAR & Npc(NpcNum).Sprite & SEP_CHAR & Npc(NpcNum).SpawnSecs & SEP_CHAR & Npc(NpcNum).Behavior & SEP_CHAR & Npc(NpcNum).Range & SEP_CHAR & Npc(NpcNum).DropChance & SEP_CHAR & Npc(NpcNum).DropItem & SEP_CHAR & Npc(NpcNum).DropItemValue & SEP_CHAR & Npc(NpcNum).Stat(Stats.Strength) & SEP_CHAR & Npc(NpcNum).Stat(Stats.Defense) & SEP_CHAR & Npc(NpcNum).Stat(Stats.Speed) & SEP_CHAR & Npc(NpcNum).Stat(Stats.Magic) & END_CHAR
    Call SendDataTo(Index, packet)
End Sub

Sub SendShops(ByVal Index As Long)
Dim i As Long

    For i = 1 To MAX_SHOPS
        If LenB(Trim$(Shop(i).Name)) > 0 Then
            Call SendUpdateShopTo(Index, i)
        End If
    Next
End Sub

Sub SendUpdateShopToAll(ByVal ShopNum As Long)
Dim packet As String

    packet = SUpdateShop & SEP_CHAR & ShopNum & SEP_CHAR & Trim$(Shop(ShopNum).Name) & END_CHAR
    Call SendDataToAll(packet)
End Sub

Sub SendUpdateShopTo(ByVal Index As Long, ByVal ShopNum As Long)
Dim packet As String

    packet = SUpdateShop & SEP_CHAR & ShopNum & SEP_CHAR & Trim$(Shop(ShopNum).Name) & END_CHAR
    Call SendDataTo(Index, packet)
End Sub

Sub SendEditShopTo(ByVal Index As Long, ByVal ShopNum As Long)
Dim packet As String
Dim i As Long

    packet = SEditShop & SEP_CHAR & ShopNum & SEP_CHAR & Trim$(Shop(ShopNum).Name) & SEP_CHAR & Trim$(Shop(ShopNum).JoinSay) & SEP_CHAR & Trim$(Shop(ShopNum).LeaveSay) & SEP_CHAR & Shop(ShopNum).FixesItems
    For i = 1 To MAX_TRADES
        packet = packet & SEP_CHAR & Shop(ShopNum).TradeItem(i).GiveItem & SEP_CHAR & Shop(ShopNum).TradeItem(i).GiveValue & SEP_CHAR & Shop(ShopNum).TradeItem(i).GetItem & SEP_CHAR & Shop(ShopNum).TradeItem(i).GetValue
    Next
    packet = packet & END_CHAR

    Call SendDataTo(Index, packet)
End Sub

Sub SendSpells(ByVal Index As Long)
Dim i As Long

    For i = 1 To MAX_SPELLS
        If LenB(Trim$(Spell(i).Name)) > 0 Then
            Call SendUpdateSpellTo(Index, i)
        End If
    Next
End Sub

Sub SendUpdateSpellToAll(ByVal SpellNum As Long)
Dim packet As String

    packet = SUpdateSpell & SEP_CHAR & SpellNum & SEP_CHAR & Trim$(Spell(SpellNum).Name) & END_CHAR
    Call SendDataToAll(packet)
End Sub

Sub SendUpdateSpellTo(ByVal Index As Long, ByVal SpellNum As Long)
Dim packet As String

    packet = SUpdateSpell & SEP_CHAR & SpellNum & SEP_CHAR & Trim$(Spell(SpellNum).Name) & END_CHAR
    Call SendDataTo(Index, packet)
End Sub

Sub SendEditSpellTo(ByVal Index As Long, ByVal SpellNum As Long)
Dim packet As String

    packet = SEditSpell & SEP_CHAR & SpellNum & SEP_CHAR & Trim$(Spell(SpellNum).Name) & SEP_CHAR & Spell(SpellNum).ClassReq & SEP_CHAR & Spell(SpellNum).LevelReq & SEP_CHAR & Spell(SpellNum).Type & SEP_CHAR & Spell(SpellNum).Data1 & SEP_CHAR & Spell(SpellNum).Data2 & SEP_CHAR & Spell(SpellNum).Data3 & END_CHAR
    Call SendDataTo(Index, packet)
End Sub

Sub SendTrade(ByVal Index As Long, ByVal ShopNum As Long)
Dim packet As String
Dim i As Long, X As Long, Y As Long

    packet = STrade & SEP_CHAR & ShopNum & SEP_CHAR & Shop(ShopNum).FixesItems
    For i = 1 To MAX_TRADES
        packet = packet & SEP_CHAR & Shop(ShopNum).TradeItem(i).GiveItem & SEP_CHAR & Shop(ShopNum).TradeItem(i).GiveValue & SEP_CHAR & Shop(ShopNum).TradeItem(i).GetItem & SEP_CHAR & Shop(ShopNum).TradeItem(i).GetValue
        
        ' Item #
        X = Shop(ShopNum).TradeItem(i).GetItem
        
        If Item(X).Type = ITEM_TYPE_SPELL Then
            ' Spell class requirement
            Y = Spell(Item(X).Data1).ClassReq
            
            If Y = 0 Then
                Call PlayerMsg(Index, Trim$(Item(X).Name) & " can be used by all classes.", Yellow)
            Else
                Call PlayerMsg(Index, Trim$(Item(X).Name) & " can only be used by a " & GetClassName(Y - 1) & ".", Yellow)
            End If
        End If
    Next
    packet = packet & END_CHAR
    
    Call SendDataTo(Index, packet)
End Sub

Sub SendPlayerSpells(ByVal Index As Long)
Dim packet As String
Dim i As Long

    packet = SSpells
    For i = 1 To MAX_PLAYER_SPELLS
        packet = packet & SEP_CHAR & GetPlayerSpell(Index, i)
    Next
    packet = packet & END_CHAR
    
    Call SendDataTo(Index, packet)
End Sub

