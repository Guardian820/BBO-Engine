Attribute VB_Name = "modHandleData"
Option Explicit

' ******************************************
' **              BMO Source              **
' ** Parses and handles String packets    **
' ******************************************

Public Sub HandleData(ByVal Data As String)
Dim parse() As String

    ' Handle Data
    parse = Split(Data, SEP_CHAR)
    
    ' Add the data to the debug window if we are in debug mode
    If DEBUG_MODE Then
        If frmDebug.Visible = False Then frmDebug.Visible = True
        Call TextAdd(frmDebug.txtDebug, "((( Processed Packet " & parse(0) & " )))", True)
    End If
    
    Select Case parse(0)
    
        Case SAlertMsg
            HandleAlertMsg parse
        Case SAllChars
            HandleAllChars parse
        Case SLoginOk
            HandleLoginOk parse
        Case SNewCharClasses
            HandleNewCharClasses parse
        Case SClassesData
            HandleClassesData parse
        Case SInGame
            HandleInGame parse
        Case SPlayerInv
            HandlePlayerInv parse
        Case SPlayerInvUpdate
            HandlePlayerInvUpdate parse
        Case SPlayerWornEq
            HandlePlayerWornEq parse
        Case SPlayerHp
            HandlePlayerHp parse
        Case SPlayerMp
            HandlePlayerMp parse
        Case SPlayerSp
            HandlePlayerSp parse
        Case SPlayerStats
            HandlePlayerStats parse
        Case SPlayerData
            HandlePlayerData parse
        Case SPlayerMove
            HandlePlayerMove parse
        Case SNpcMove
            HandleNpcMove parse
        Case SPlayerDir
            HandlePlayerDir parse
        Case SNpcDir
            HandleNpcDir parse
        Case SPlayerXY
            HandlePlayerXY parse
        Case SAttack
            HandleAttack parse
        Case SNpcAttack
            HandleNpcAttack parse
        Case SCheckForMap
            HandleCheckForMap parse
        Case SMapData
            HandleMapData parse
        Case SMapItemData
            HandleMapItemData parse
        Case SMapNpcData
            HandleMapNpcData parse
        Case SMapDone
            HandleMapDone
        Case SSayMsg
            HandleSayMsg parse
        Case SGlobalMsg
            HandleGlobalMsg parse
        Case SAdminMsg
            HandleAdminMsg parse
        Case SPlayerMsg
            HandlePlayerMsg parse
        Case SMapMsg
            HandleMapMsg parse
        Case SSpawnItem
            HandleSpawnItem parse
        Case SItemEditor
            HandleItemEditor
        Case SUpdateItem
            HandleUpdateItem parse
        Case SEditItem
            HandleEditItem parse
        Case SSpawnNpc
            HandleSpawnNpc parse
        Case SNpcDead
            HandleNpcDead parse
        Case SNpcEditor
            HandleNpcEditor
        Case SUpdateNpc
            HandleUpdateNpc parse
        Case SEditNpc
            HandleEditNpc parse
        Case SMapKey
            HandleMapKey parse
        Case SEditMap
            HandleEditMap
        Case SShopEditor
            HandleShopEditor
        Case SUpdateShop
            HandleUpdateShop parse
        Case SEditShop
            HandleEditShop parse
        Case SREditor
            HandleRefresh
        Case SSpellEditor
            HandleSpellEditor
        Case SUpdateSpell
            HandleUpdateSpell parse
        Case SEditSpell
            HandleEditSpell parse
        Case STrade
            HandleTrade parse
        Case SSpells
            HandleSpells parse
        Case SLeft
            HandleLeft parse
        Case SHighIndex
            HandleHighIndex parse
        Case SFire
            HandleFire parse
        Case SBomb
            HandleBomb parse
        Case SDieing
            HandleDieing parse
        Case SWall
            HandleWall parse
        Case SJump
            HandleJump parse
        Case SSendRoomList
            HandleRoomList parse
        Case SInLobby
            HandleInLobby
        Case SStartMatch
            HandleStartMatch parse
        Case SStartTimer
            HandleStartTimer
        Case SWatching
            HandleWatching parse
        Case SLeaveRoom
            HandleLeaveRoom
        Case SHighScore
            HandleHighScore parse
        Case SHighScoreList
            HandleHighScoreList parse
        Case SRefreshWall
            HandleRefreshWall parse
        Case SUpdateFriendList
            HandleUpdateFriendList parse
        Case SSpeed
            HandleSpeed
    End Select
        
End Sub

' ::::::::::::::::::::::::::
 ' :: Alert message packet ::
 ' ::::::::::::::::::::::::::
Sub HandleAlertMsg(ByRef parse() As String)
Dim Msg As String

     frmMain.picLoading.Visible = False
     frmMain.picMainMenu.Visible = True
     
     Msg = parse(1)
        frmMsg.Visible = True
        frmMsg.lblAtenção.Caption = "Atenção"
        frmMsg.lblMsg.Caption = Msg
End Sub
 ' :::::::::::::::::::::::::::
 ' :: All characters packet ::
 ' :::::::::::::::::::::::::::
Sub HandleAllChars(ByRef parse() As String)
Dim n As Long, i As Long, Level As Long
Dim Name As String, Msg As String

     n = 1
     
     'frmChars.Visible = True
     'frmSendGetData.Visible = False
     
     'frmChars.lstChars.Clear
     
     'For i = 1 To MAX_CHARS
     '    Name = Parse(n)
     '    Msg = Parse(n + 1)
     '    Level = CLng(Parse(n + 2))
         
     '    If LenB(Trim$(Name)) = 0 Then
     '        frmChars.lstChars.AddItem "Free Character Slot"
     '    Else
     '        frmChars.lstChars.AddItem Name & " a level " & Level & " " & Msg
     '    End If
         
     '    n = n + 3
     'Next
     
     'frmChars.lstChars.ListIndex = 0
End Sub
 ' :::::::::::::::::::::::::::::::::
 ' :: Login was successful packet ::
 ' :::::::::::::::::::::::::::::::::
Sub HandleLoginOk(ByRef parse() As String)
     ' Now we can receive game data
     MyIndex = CLng(parse(1))
     Player_HighIndex = parse(2)
     'frmSendGetData.Visible = True
     frmMain.picLogin.Visible = False
     
End Sub
 ' :::::::::::::::::::::::::::::::::::::::
 ' :: New character classes data packet ::
 ' :::::::::::::::::::::::::::::::::::::::
Sub HandleNewCharClasses(ByRef parse() As String)
Dim n As Long, i As Long
     
     n = 1
     
     ' Max classes
     Max_Classes = CByte(parse(n))
     ReDim Class(1 To Max_Classes) As ClassRec
     
     n = n + 1
     
     For i = 1 To Max_Classes
         With Class(i)
             .Name = parse(n)
             
             .Vital(Vitals.HP) = CLng(parse(n + 1))
             .Vital(Vitals.MP) = CLng(parse(n + 2))
             .Vital(Vitals.SP) = CLng(parse(n + 3))
             
             .Stat(Stats.Strength) = CLng(parse(n + 4))
             .Stat(Stats.Defense) = CLng(parse(n + 5))
             .Stat(Stats.SPEED) = CLng(parse(n + 6))
             .Stat(Stats.Magic) = CLng(parse(n + 7))
         End With
         
         n = n + 8
     Next
     
     frmMain.height = Main_MoreHeight
     
     ' Used for if the player is creating a new character
     frmMain.picRegister.Visible = True
     frmMain.picLoading.Visible = False
     
     RegisterBlt
     
     frmMain.txtRegisterName.SetFocus

     'frmMain.cmbClass.Clear

     'For i = 1 To Max_Classes
     '    frmMain.cmbClass.AddItem Trim$(Class(i).Name)
     'Next

     'frmMain.cmbClass.ListIndex = 0
     
     'n = frmMain.cmbClass.ListIndex + 1
     
End Sub
 ' :::::::::::::::::::::::::
 ' :: Classes data packet ::
 ' :::::::::::::::::::::::::
Sub HandleClassesData(ByRef parse() As String)
Dim n As Long, i As Long
     
     n = 1
     
     ' Max classes
     Max_Classes = CByte(parse(n))
     ReDim Class(1 To Max_Classes) As ClassRec
     
     n = n + 1
     
     For i = 1 To Max_Classes
         With Class(i)
             .Name = parse(n)
             
             .Vital(Vitals.HP) = CLng(parse(n + 1))
             .Vital(Vitals.MP) = CLng(parse(n + 2))
             .Vital(Vitals.SP) = CLng(parse(n + 3))
             
             .Stat(Stats.Strength) = CLng(parse(n + 4))
             .Stat(Stats.Defense) = CLng(parse(n + 5))
             .Stat(Stats.SPEED) = CLng(parse(n + 6))
             .Stat(Stats.Magic) = CLng(parse(n + 7))
         End With
         
         n = n + 8
     Next
End Sub
 ' ::::::::::::::::::::
 ' :: In game packet ::
 ' ::::::::::::::::::::
Sub HandleInGame(ByRef parse() As String)

     InGame = True
     'Call GameInit
     'frmLobby.Hide
     frmWaiting.Hide
     frmMainGame.Show
     frmMainGame.lblRoom.Caption = GetPlayerMap(MyIndex)
     frmWaiting.tmrGameStart.Enabled = False
     Call GameLoop
     
End Sub

Sub HandleInLobby()

    Call GameInit

End Sub

 ' :::::::::::::::::::::::::::::
 ' :: Player inventory packet ::
 ' :::::::::::::::::::::::::::::
Sub HandlePlayerInv(ByRef parse() As String)
Dim n As Long, i As Long

     n = 1
     For i = 1 To MAX_INV
         Call SetPlayerInvItemNum(MyIndex, i, CLng(parse(n)))
         Call SetPlayerInvItemValue(MyIndex, i, CLng(parse(n + 1)))
         Call SetPlayerInvItemDur(MyIndex, i, CLng(parse(n + 2)))
         
         n = n + 3
     Next
     Call UpdateInventory
 End Sub
 ' ::::::::::::::::::::::::::::::::::::
 ' :: Player inventory update packet ::
 ' ::::::::::::::::::::::::::::::::::::
Sub HandlePlayerInvUpdate(ByRef parse() As String)
Dim n As Long

     n = CLng(parse(1))
     
     Call SetPlayerInvItemNum(MyIndex, n, CLng(parse(2)))
     Call SetPlayerInvItemValue(MyIndex, n, CLng(parse(3)))
     Call SetPlayerInvItemDur(MyIndex, n, CLng(parse(4)))
     Call UpdateInventory
End Sub
 ' ::::::::::::::::::::::::::::::::::
 ' :: Player worn equipment packet ::
 ' ::::::::::::::::::::::::::::::::::
Sub HandlePlayerWornEq(ByRef parse() As String)
     Call SetPlayerEquipmentSlot(MyIndex, CLng(parse(1)), Armor)
     Call SetPlayerEquipmentSlot(MyIndex, CLng(parse(2)), Weapon)
     Call SetPlayerEquipmentSlot(MyIndex, CLng(parse(3)), Helmet)
     Call SetPlayerEquipmentSlot(MyIndex, CLng(parse(4)), Shield)
     Call UpdateInventory
End Sub
 ' ::::::::::::::::::::::
 ' :: Player hp packet ::
 ' ::::::::::::::::::::::
Sub HandlePlayerHp(ByRef parse() As String)
     Player(MyIndex).MaxHP = CLng(parse(1))
     Call SetPlayerVital(MyIndex, Vitals.HP, CLng(parse(2)))
     If GetPlayerMaxVital(MyIndex, Vitals.HP) > 0 Then
         'frmMainGame.lblHP.Caption = Int(GetPlayerVital(MyIndex, Vitals.HP) / GetPlayerMaxVital(MyIndex, Vitals.HP) * 100) & "%"
     End If
End Sub
 ' ::::::::::::::::::::::
 ' :: Player mp packet ::
 ' ::::::::::::::::::::::
Sub HandlePlayerMp(ByRef parse() As String)
     Player(MyIndex).MaxMP = CLng(parse(1))
     Call SetPlayerVital(MyIndex, Vitals.MP, CLng(parse(2)))
     If GetPlayerMaxVital(MyIndex, Vitals.MP) > 0 Then
         'frmMainGame.lblMP.Caption = Int(GetPlayerVital(MyIndex, Vitals.MP) / GetPlayerMaxVital(MyIndex, Vitals.MP) * 100) & "%"
     End If
End Sub
 ' ::::::::::::::::::::::
 ' :: Player sp packet ::
 ' ::::::::::::::::::::::
Sub HandlePlayerSp(ByRef parse() As String)
     Player(MyIndex).MaxSP = CLng(parse(1))
     Call SetPlayerVital(MyIndex, Vitals.SP, CLng(parse(2)))
     If GetPlayerMaxVital(MyIndex, Vitals.SP) > 0 Then
         'frmMainGame.lblSP.Caption = Int(GetPlayerVital(MyIndex, Vitals.SP) / GetPlayerMaxVital(MyIndex, Vitals.SP) * 100) & "%"
     End If
End Sub
 ' :::::::::::::::::::::::::
 ' :: Player stats packet ::
 ' :::::::::::::::::::::::::
Sub HandlePlayerStats(ByRef parse() As String)
     Call SetPlayerStat(MyIndex, Stats.Strength, CLng(parse(1)))
     Call SetPlayerStat(MyIndex, Stats.Defense, CLng(parse(2)))
     Call SetPlayerStat(MyIndex, Stats.SPEED, CLng(parse(3)))
     Call SetPlayerStat(MyIndex, Stats.Magic, CLng(parse(4)))
End Sub
 ' ::::::::::::::::::::::::
 ' :: Player data packet ::
 ' ::::::::::::::::::::::::
Sub HandlePlayerData(ByRef parse() As String)
Dim i As Long

     i = CLng(parse(1))
     
     Call SetPlayerName(i, parse(2))
     Call SetPlayerSprite(i, CLng(parse(3)))
     Call SetPlayerMap(i, CLng(parse(4)))
     Call SetPlayerX(i, CLng(parse(5)))
     Call SetPlayerY(i, CLng(parse(6)))
     Call SetPlayerDir(i, CLng(parse(7)))
     Call SetPlayerAccess(i, CLng(parse(8)))
     Call SetPlayerPK(i, CLng(parse(9)))
     Player(i).Kills = (parse(10))
     Player(i).Deaths = (parse(11))
     Player(i).BPoints = CLng(parse(12))
     Player(i).BCash = CLng(parse(13))
     
     ' Check if the player is the client player, and if so reset directions
     If i = MyIndex Then
         DirUp = False
         DirDown = False
         DirLeft = False
         DirRight = False
     End If
     
     ' Make sure they aren't walking
     Player(i).Moving = 0
     Player(i).XOffset = 0
     Player(i).YOffset = 0
     
     Call GetPlayersOnMap
End Sub
 ' ::::::::::::::::::::::::::::
 ' :: Player movement packet ::
 ' ::::::::::::::::::::::::::::
Sub HandlePlayerMove(ByRef parse() As String)
Dim i As Long, x As Long, y As Long, Dir As Long
Dim n As Byte

     i = CLng(parse(1))
     x = CLng(parse(2))
     y = CLng(parse(3))
     Dir = CLng(parse(4))
     n = CByte(parse(5))

     Call SetPlayerX(i, x)
     Call SetPlayerY(i, y)
     Call SetPlayerDir(i, Dir)
             
     Player(i).XOffset = 0
     Player(i).YOffset = 0
     Player(i).Moving = n
     
     Select Case GetPlayerDir(i)
         Case DIR_UP
             Player(i).YOffset = PIC_Y
         Case DIR_DOWN
             Player(i).YOffset = PIC_Y * -1
         Case DIR_LEFT
             Player(i).XOffset = PIC_X
         Case DIR_RIGHT
             Player(i).XOffset = PIC_X * -1
     End Select
End Sub
 ' :::::::::::::::::::::::::
 ' :: Npc movement packet ::
 ' :::::::::::::::::::::::::
Sub HandleNpcMove(ByRef parse() As String)
Dim i As Long, x As Long, y As Long, Dir As Long
Dim n As Byte

     i = CLng(parse(1))
     x = CLng(parse(2))
     y = CLng(parse(3))
     Dir = CLng(parse(4))
     n = CByte(parse(5))

     MapNpc(i).x = x
     MapNpc(i).y = y
     MapNpc(i).Dir = Dir
     MapNpc(i).XOffset = 0
     MapNpc(i).YOffset = 0
     MapNpc(i).Moving = n
     
     Select Case MapNpc(i).Dir
         Case DIR_UP
             MapNpc(i).YOffset = PIC_Y
         Case DIR_DOWN
             MapNpc(i).YOffset = PIC_Y * -1
         Case DIR_LEFT
             MapNpc(i).XOffset = PIC_X
         Case DIR_RIGHT
             MapNpc(i).XOffset = PIC_X * -1
     End Select
End Sub
 ' :::::::::::::::::::::::::::::
 ' :: Player direction packet ::
 ' :::::::::::::::::::::::::::::
Sub HandlePlayerDir(ByRef parse() As String)
Dim i As Long
Dim Dir As Byte

     i = CLng(parse(1))
     Dir = CByte(parse(2))
     Call SetPlayerDir(i, Dir)
     
     Player(i).XOffset = 0
     Player(i).YOffset = 0
     Player(i).Moving = 0
End Sub
 ' ::::::::::::::::::::::::::
 ' :: NPC direction packet ::
 ' ::::::::::::::::::::::::::
Sub HandleNpcDir(ByRef parse() As String)
Dim i As Long
Dim Dir As Byte

     i = CLng(parse(1))
     Dir = CByte(parse(2))
     MapNpc(i).Dir = Dir
     
     MapNpc(i).XOffset = 0
     MapNpc(i).YOffset = 0
     MapNpc(i).Moving = 0
End Sub
 ' :::::::::::::::::::::::::::::::
 ' :: Player XY location packet ::
 ' :::::::::::::::::::::::::::::::
Sub HandlePlayerXY(ByRef parse() As String)
Dim x As Long, y As Long

     x = CLng(parse(1))
     y = CLng(parse(2))
     
     Call SetPlayerX(MyIndex, x)
     Call SetPlayerY(MyIndex, y)
     
     ' Make sure they aren't walking
     Player(MyIndex).Moving = 0
     Player(MyIndex).XOffset = 0
     Player(MyIndex).YOffset = 0
End Sub
 ' ::::::::::::::::::::::::::
 ' :: Player attack packet ::
 ' ::::::::::::::::::::::::::
Sub HandleAttack(ByRef parse() As String)
Dim i As Long

     i = CLng(parse(1))
     
     ' Set player to attacking
     Player(i).Attacking = 1
     Player(i).AttackTimer = GetTickCount

End Sub

Sub HandleFire(ByRef parse() As String)

    Map.Fire(Val(parse(1)), Val(parse(2))).Here = True
    
    If Val(parse(3)) > 0 Then
        Map.Fire((parse(1)), Val(parse(2))).Ended = Val(parse(3))
    Else
        Map.Fire((parse(1)), Val(parse(2))).Ended = 0
    End If
    
    Map.Fire((parse(1)), Val(parse(2))).Direction = Val(parse(4))
    
    Map.Fire(Val(parse(1)), Val(parse(2))).Timer = GetTickCount ' + 1000
    
    Map.Fire(Val(parse(1)), Val(parse(2))).Owner = Val(parse(5))

End Sub

Sub HandleBomb(ByRef parse() As String)

    If Val(parse(1)) = 0 Then
        Map.Tile(Val(parse(2)), Val(parse(3))).Type = TILE_TYPE_BOMB
    Else
        Map.Tile(Val(parse(2)), Val(parse(3))).Type = 0
    End If
    
End Sub

 ' :::::::::::::::::::::::
 ' :: NPC attack packet ::
 ' :::::::::::::::::::::::
Sub HandleNpcAttack(ByRef parse() As String)
Dim i As Long

     i = CLng(parse(1))
     
     ' Set player to attacking
     MapNpc(i).Attacking = 1
     MapNpc(i).AttackTimer = GetTickCount
End Sub
 ' ::::::::::::::::::::::::::
 ' :: Check for map packet ::
 ' ::::::::::::::::::::::::::
Sub HandleCheckForMap(ByRef parse() As String)
Dim x As Long, y As Long, i As Long

     ' Erase all players except self
     For i = 1 To Player_HighIndex
         If i <> MyIndex Then
             Call SetPlayerMap(i, 0)
         End If
     Next

     ' Erase all temporary tile values
     Call ClearTempTile
     
     ' Get map num
     x = CLng(parse(1))
     
     ' Get revision
     y = CLng(parse(2))
     
     If FileExist(MAP_PATH & "map" & x & MAP_EXT, False) Then
         Call LoadMap(x)
     
         ' Check to see if the revisions match
         If Map.Revision = y Then
             ' We do so we dont need the map
             Call SendData(CNeedMap & SEP_CHAR & "no" & END_CHAR)
             Exit Sub
         End If
     End If

     ' Either the revisions didn't match or we dont have the map, so we need it
     Call SendData(CNeedMap & SEP_CHAR & "yes" & END_CHAR)

End Sub
 ' :::::::::::::::::::::
 ' :: Map data packet ::
 ' :::::::::::::::::::::
Sub HandleMapData(ByRef parse() As String)
Dim n As Long, x As Long, y As Long

     n = 1
     
     Map.Name = parse(n + 1)
     Map.Revision = CLng(parse(n + 2))
     Map.PlayerLimit = CByte(parse(n + 3))
     Map.Moral = CByte(parse(n + 4))
     Map.TileSet = CInt(parse(n + 5))
     Map.Up = CInt(parse(n + 6))
     Map.Down = CInt(parse(n + 7))
     Map.Left = CInt(parse(n + 8))
     Map.Right = CInt(parse(n + 9))
     Map.Music = CByte(parse(n + 10))
     Map.BootMap = CInt(parse(n + 11))
     Map.BootX = CByte(parse(n + 12))
     Map.BootY = CByte(parse(n + 13))
     Map.Shop = CByte(parse(n + 14))
     
     n = n + 15
     
     For x = 0 To MAX_MAPX
         For y = 0 To MAX_MAPY
             Map.Tile(x, y).Ground = CInt(parse(n))
             Map.Tile(x, y).Mask = CInt(parse(n + 1))
             Map.Tile(x, y).Anim = CInt(parse(n + 2))
             Map.Tile(x, y).Fringe = CInt(parse(n + 3))
             Map.Tile(x, y).Type = CByte(parse(n + 4))
             Map.Tile(x, y).Data1 = CInt(parse(n + 5))
             Map.Tile(x, y).Data2 = CInt(parse(n + 6))
             Map.Tile(x, y).Data3 = CInt(parse(n + 7))
             
             n = n + 8
         Next
     Next
     
     For x = 1 To MAX_MAP_NPCS
         Map.Npc(x) = CByte(parse(n))
         n = n + 1
     Next
             
     ' Save the map
     Call SaveMap(CLng(parse(1)))
     
     ' Check if we get a map from someone else and if we were editing a map cancel it out
     If InEditor Then
         InEditor = False
         frmMainGame.picMapEditor.Visible = False
         
         'If frmMapWarp.Visible Then
         '    Unload frmMapWarp
         'End If
         
         If frmMapProperties.Visible Then
             Unload frmMapProperties
         End If
     End If
End Sub
 ' :::::::::::::::::::::::::::
 ' :: Map items data packet ::
 ' :::::::::::::::::::::::::::
Sub HandleMapItemData(ByRef parse() As String)
Dim n As Long, i As Long

     n = 1
     
     For i = 1 To MAX_MAP_ITEMS
         MapItem(i).Num = CByte(parse(n))
         MapItem(i).Value = CLng(parse(n + 1))
         MapItem(i).Dur = CInt(parse(n + 2))
         MapItem(i).x = CByte(parse(n + 3))
         MapItem(i).y = CByte(parse(n + 4))
         
         n = n + 5
     Next
End Sub
 ' :::::::::::::::::::::::::
 ' :: Map npc data packet ::
 ' :::::::::::::::::::::::::
Sub HandleMapNpcData(ByRef parse() As String)
Dim n As Long, i As Long

     n = 1
     
     For i = 1 To MAX_MAP_NPCS
         MapNpc(i).Num = CByte(parse(n))
         MapNpc(i).x = CByte(parse(n + 1))
         MapNpc(i).y = CByte(parse(n + 2))
         MapNpc(i).Dir = CByte(parse(n + 3))
         
         n = n + 4
     Next
End Sub
 ' :::::::::::::::::::::::::::::::
 ' :: Map send completed packet ::
 ' :::::::::::::::::::::::::::::::
Sub HandleMapDone()
Dim i As Long

    GettingMap = False
    
    ' get high NPC index
    High_Npc_Index = 0
    
    For i = 1 To MAX_MAP_NPCS
    
        If MapNpc(i).Num > 0 Then
            High_Npc_Index = High_Npc_Index + 1
        Else
            Exit For
        End If
    Next
    
    ' Play music
    If Map.Music > 0 Then
        DirectMusic_StopMidi
        Call DirectMusic_PlayMidi("music" & Trim$(CStr(Map.Music)) & ".mid")
    Else
        DirectMusic_StopMidi
    End If
    
    Call UpdateDrawMapName
    
    Call InitTileSurf(Map.TileSet)
    
    CanMoveNow = True
End Sub
 ' ::::::::::::::::::::
 ' :: Social packets ::
 ' ::::::::::::::::::::
Sub HandleSayMsg(ByRef parse() As String)
     Call AddText(parse(1), CInt(parse(2)))
End Sub

 Sub HandleBroadcastMsg(ByRef parse() As String)
     Call AddText(parse(1), CInt(parse(2)))
End Sub
 
Sub HandleGlobalMsg(ByRef parse() As String)
     Call AddText(parse(1), CInt(parse(2)))
End Sub
     
Sub HandlePlayerMsg(ByRef parse() As String)
     Call AddText(parse(1), CInt(parse(2)))
End Sub
     
Sub HandleMapMsg(ByRef parse() As String)
     Call AddText(parse(1), CInt(parse(2)))
End Sub
 
Sub HandleAdminMsg(ByRef parse() As String)
     Call AddText(parse(1), CInt(parse(2)))
End Sub

 ' ::::::::::::::::::::::::
 ' :: Refresh editor packet ::
 ' ::::::::::::::::::::::::
Sub HandleRefresh()
    Dim i As Long
    
    frmIndex.lstIndex.Clear
    
    Select Case Editor
        Case EDITOR_ITEM
            For i = 1 To MAX_ITEMS
                frmIndex.lstIndex.AddItem i & ": " & Trim$(Item(i).Name)
            Next
        Case EDITOR_NPC
            For i = 1 To MAX_NPCS
                frmIndex.lstIndex.AddItem i & ": " & Trim$(Npc(i).Name)
            Next
        Case EDITOR_SHOP
            For i = 1 To MAX_SHOPS
                frmIndex.lstIndex.AddItem i & ": " & Trim$(Shop(i).Name)
            Next
        Case EDITOR_SPELL
            For i = 1 To MAX_SPELLS
                frmIndex.lstIndex.AddItem i & ": " & Trim$(Spell(i).Name)
            Next
    End Select
     
    frmIndex.lstIndex.ListIndex = 0
     
End Sub

 ' :::::::::::::::::::::::
 ' :: Item spawn packet ::
 ' :::::::::::::::::::::::
Sub HandleSpawnItem(ByRef parse() As String)
Dim n As Long

     n = CLng(parse(1))
     
     MapItem(n).Num = CByte(parse(2))
     MapItem(n).Value = CLng(parse(3))
     MapItem(n).Dur = CInt(parse(4))
     MapItem(n).x = CByte(parse(5))
     MapItem(n).y = CByte(parse(6))
End Sub
 
 ' ::::::::::::::::::::::::
 ' :: Item editor packet ::
 ' ::::::::::::::::::::::::
Sub HandleItemEditor()
Dim i As Long

     Editor = 1

     frmIndex.lstIndex.Clear
     
     ' Add the names
     For i = 1 To MAX_ITEMS
         frmIndex.lstIndex.AddItem i & ": " & Trim$(Item(i).Name)
     Next
     
     frmIndex.lstIndex.ListIndex = 0
     frmIndex.Show vbModal
End Sub
 ' ::::::::::::::::::::::::
 ' :: Update item packet ::
 ' ::::::::::::::::::::::::
Sub HandleUpdateItem(ByRef parse() As String)
Dim n As Long

     n = CLng(parse(1))
     
     ' Update the item
     Item(n).Name = parse(2)
     Item(n).Pic = CInt(parse(3))
     Item(n).Type = CByte(parse(4))
     Item(n).Data1 = 0
     Item(n).Data2 = 0
     Item(n).Data3 = 0
End Sub
 ' ::::::::::::::::::::::
 ' :: Edit item packet ::
 ' ::::::::::::::::::::::
Sub HandleEditItem(ByRef parse() As String)
Dim n As Long

     n = CLng(parse(1))
     
     ' Update the item
     Item(n).Name = parse(2)
     Item(n).Pic = CInt(parse(3))
     Item(n).Type = CByte(parse(4))
     Item(n).Data1 = CInt(parse(5))
     Item(n).Data2 = CInt(parse(6))
     Item(n).Data3 = CInt(parse(7))
     
     ' Initialize the item editor
     Call ItemEditorInit
End Sub
 ' ::::::::::::::::::::::
 ' :: Npc spawn packet ::
 ' ::::::::::::::::::::::
Sub HandleSpawnNpc(ByRef parse() As String)
Dim n As Long

     n = CLng(parse(1))
     
     MapNpc(n).Num = CByte(parse(2))
     MapNpc(n).x = CByte(parse(3))
     MapNpc(n).y = CByte(parse(4))
     MapNpc(n).Dir = CByte(parse(5))
     
     ' Client use only
     MapNpc(n).XOffset = 0
     MapNpc(n).YOffset = 0
     MapNpc(n).Moving = 0
End Sub
 ' :::::::::::::::::::::
 ' :: Npc dead packet ::
 ' :::::::::::::::::::::
 Sub HandleNpcDead(ByRef parse() As String)
 Dim n As Long
 
     n = CLng(parse(1))
     
     MapNpc(n).Num = 0
     MapNpc(n).x = 0
     MapNpc(n).y = 0
     MapNpc(n).Dir = 0
     
     ' Client use only
     MapNpc(n).XOffset = 0
     MapNpc(n).YOffset = 0
     MapNpc(n).Moving = 0
End Sub
 ' :::::::::::::::::::::::
 ' :: Npc editor packet ::
 ' :::::::::::::::::::::::
Sub HandleNpcEditor()
Dim i As Long

    Editor = 2
     
     frmIndex.lstIndex.Clear
     
     ' Add the names
     For i = 1 To MAX_NPCS
         frmIndex.lstIndex.AddItem i & ": " & Trim$(Npc(i).Name)
     Next
     
     frmIndex.lstIndex.ListIndex = 0
     frmIndex.Show vbModal
End Sub
 ' :::::::::::::::::::::::
 ' :: Update npc packet ::
 ' :::::::::::::::::::::::
Sub HandleUpdateNpc(ByRef parse() As String)
Dim n As Long

     n = CLng(parse(1))
     
     ' Update the item
     Npc(n).Name = parse(2)
     Npc(n).AttackSay = vbNullString
     Npc(n).sprite = CInt(parse(3))
     Npc(n).SpawnSecs = 0
     Npc(n).Behavior = 0
     Npc(n).Range = 0
     Npc(n).DropChance = 0
     Npc(n).DropItem = 0
     Npc(n).DropItemValue = 0
     Npc(n).Stat(Stats.Strength) = 0
     Npc(n).Stat(Stats.Defense) = 0
     Npc(n).Stat(Stats.SPEED) = 0
     Npc(n).Stat(Stats.Magic) = 0
End Sub
 ' :::::::::::::::::::::
 ' :: Edit npc packet ::
 ' :::::::::::::::::::::
Sub HandleEditNpc(ByRef parse() As String)
Dim n As Long

     n = CLng(parse(1))
     
     ' Update the npc
     Npc(n).Name = parse(2)
     Npc(n).AttackSay = parse(3)
     Npc(n).sprite = CInt(parse(4))
     Npc(n).SpawnSecs = CLng(parse(5))
     Npc(n).Behavior = CByte(parse(6))
     Npc(n).Range = CByte(parse(7))
     Npc(n).DropChance = CInt(parse(8))
     Npc(n).DropItem = CByte(parse(9))
     Npc(n).DropItemValue = CInt(parse(10))
     Npc(n).Stat(Stats.Strength) = CByte(parse(11))
     Npc(n).Stat(Stats.Defense) = CByte(parse(12))
     Npc(n).Stat(Stats.SPEED) = CByte(parse(13))
     Npc(n).Stat(Stats.Magic) = CByte(parse(14))
     
     ' Initialize the npc editor
     Call NpcEditorInit
End Sub
 ' ::::::::::::::::::::
 ' :: Map key packet ::
 ' ::::::::::::::::::::
Sub HandleMapKey(ByRef parse() As String)
Dim n As Long, x As Long, y As Long

     x = CLng(parse(1))
     y = CLng(parse(2))
     n = CLng(parse(3))
     
     TempTile(x, y).DoorOpen = n
End Sub
 ' :::::::::::::::::::::
 ' :: Edit map packet ::
 ' :::::::::::::::::::::
Sub HandleEditMap()
     Call MapEditorInit
End Sub
 ' ::::::::::::::::::::::::
 ' :: Shop editor packet ::
 ' ::::::::::::::::::::::::
Sub HandleShopEditor()
Dim i As Long

     Editor = 4
     
     frmIndex.lstIndex.Clear
     
     ' Add the names
     For i = 1 To MAX_SHOPS
         frmIndex.lstIndex.AddItem i & ": " & Trim$(Shop(i).Name)
     Next
     
     frmIndex.lstIndex.ListIndex = 0
     frmIndex.Show vbModal
End Sub
 ' ::::::::::::::::::::::::
 ' :: Update shop packet ::
 ' ::::::::::::::::::::::::
Sub HandleUpdateShop(ByRef parse() As String)
Dim n As Long

     n = CLng(parse(1))
     
     ' Update the shop name
     Shop(n).Name = parse(2)
End Sub
 ' ::::::::::::::::::::::
 ' :: Edit shop packet ::
 ' ::::::::::::::::::::::
Sub HandleEditShop(ByRef parse() As String)
Dim n As Long, i As Long, ShopNum As Long
Dim GiveItem As Long, GiveValue As Long, GetItem As Long, GetValue As Long

     ShopNum = CLng(parse(1))
     
     ' Update the shop
     Shop(ShopNum).Name = parse(2)
     Shop(ShopNum).JoinSay = parse(3)
     Shop(ShopNum).LeaveSay = parse(4)
     Shop(ShopNum).FixesItems = CByte(parse(5))
     
     n = 6
     For i = 1 To MAX_TRADES
         
         GiveItem = CLng(parse(n))
         GiveValue = CLng(parse(n + 1))
         GetItem = CLng(parse(n + 2))
         GetValue = CLng(parse(n + 3))
         
         Shop(ShopNum).TradeItem(i).GiveItem = GiveItem
         Shop(ShopNum).TradeItem(i).GiveValue = GiveValue
         Shop(ShopNum).TradeItem(i).GetItem = GetItem
         Shop(ShopNum).TradeItem(i).GetValue = GetValue
         
         n = n + 4
     Next
     
     ' Initialize the shop editor
     Call ShopEditorInit
     
End Sub
 ' :::::::::::::::::::::::::
 ' :: Spell editor packet ::
 ' :::::::::::::::::::::::::
Sub HandleSpellEditor()
Dim i As Long

     Editor = 3
     
     frmIndex.lstIndex.Clear
     
     ' Add the names
     For i = 1 To MAX_SPELLS
         frmIndex.lstIndex.AddItem i & ": " & Trim$(Spell(i).Name)
     Next
     
     frmIndex.lstIndex.ListIndex = 0
     frmIndex.Show vbModal
End Sub
 ' ::::::::::::::::::::::::
 ' :: Update spell packet ::
 ' ::::::::::::::::::::::::
Sub HandleUpdateSpell(ByRef parse() As String)
Dim n As Long

     n = CLng(parse(1))
     
     ' Update the spell name
     Spell(n).Name = parse(2)
End Sub
 ' :::::::::::::::::::::::
 ' :: Edit spell packet ::
Sub HandleEditSpell(ByRef parse() As String)
Dim n As Long

     n = CLng(parse(1))
     
     ' Update the spell
     Spell(n).Name = parse(2)
     Spell(n).ClassReq = CByte(parse(3))
     Spell(n).LevelReq = CByte(parse(4))
     Spell(n).Type = CByte(parse(5))
     Spell(n).Data1 = CInt(parse(6))
     Spell(n).Data2 = CInt(parse(7))
     Spell(n).Data3 = CInt(parse(8))
                     
     ' Initialize the spell editor
     Call SpellEditorInit
End Sub
 ' ::::::::::::::::::
 ' :: Trade packet ::
 ' ::::::::::::::::::
 Sub HandleTrade(ByRef parse() As String)
' Dim n As Long, i As Long, ShopNum As Long
' Dim GiveItem As Long, GiveValue As Long, GetItem As Long, GetValue As Long
'
'     ShopNum = CLng(Parse(1))
'
'     If CByte(Parse(2)) = 1 Then
'         frmTrade.lblFixItem.Visible = True
'     Else
'         frmTrade.lblFixItem.Visible = False
'     End If
'
'     n = 3
'     For i = 1 To MAX_TRADES
'         GiveItem = CLng(Parse(n))
'         GiveValue = CLng(Parse(n + 1))
'         GetItem = CLng(Parse(n + 2))
'         GetValue = CLng(Parse(n + 3))
'
'         If GiveItem > 0 Then
'             If GetItem > 0 Then
'                 frmTrade.lstTrade.AddItem "Give " & Trim$(Shop(ShopNum).Name) & " " & GiveValue & " " & Trim$(Item(GiveItem).Name) & " for " & GetValue & " " & Trim$(Item(GetItem).Name)
'             End If
'         End If
'         n = n + 4
'     Next
'
'     If frmTrade.lstTrade.ListCount > 0 Then
'         frmTrade.lstTrade.ListIndex = 0
'     End If
'     frmTrade.Show vbModal
End Sub
 ' :::::::::::::::::::
 ' :: Spells packet ::
 ' :::::::::::::::::::
Sub HandleSpells(ByRef parse() As String)
'Dim i As Long
'
'     'frmMainGame.picPlayerSpells.Visible = True
'     frmMainGame.lstSpells.Clear
'
'     ' Put spells known in player record
'     For i = 1 To MAX_PLAYER_SPELLS
'         Player(MyIndex).Spell(i) = CByte(Parse(i))
'         If Player(MyIndex).Spell(i) <> 0 Then
'             frmMainGame.lstSpells.AddItem i & ": " & Trim$(Spell(Player(MyIndex).Spell(i)).Name)
'         Else
'             frmMainGame.lstSpells.AddItem "<free spells slot>"
'         End If
'     Next
'
'     frmMainGame.lstSpells.ListIndex = 0
End Sub

Sub HandleLeft(ByRef parse() As String)
    Call ClearPlayer(parse(1))
    Call GetPlayersOnMap
End Sub

Sub HandleHighIndex(ByRef parse() As String)
    Player_HighIndex = CLng(parse(1))
End Sub

Sub HandleDieing(ByRef parse() As String)

    If Val(parse(2)) = 1 Then
        Player(Val(parse(1))).Dieing = True
        Player(Val(parse(1))).DeathTimer = GetTickCount
    Else
        Player(Val(parse(1))).Dieing = False
        Player(Val(parse(1))).DeathTimer = 0
    End If
    
End Sub

Sub HandleWall(ByRef parse() As String)

    If Not Wall(Val(parse(1)), Val(parse(2))).Here Then
        'Map.Tile(Val(Parse(1)), Val(Parse(2))).Type = 0
        
        Wall(Val(parse(1)), Val(parse(2))).Dieing = True
        Wall(Val(parse(1)), Val(parse(2))).Timer = GetTickCount
        Wall(Val(parse(1)), Val(parse(2))).Here = True
        SendData CRequestItemSpawn & SEP_CHAR & Val(parse(1)) & SEP_CHAR & Val(parse(2)) & END_CHAR
    End If

End Sub
Sub HandleRefreshWall(ByRef parse() As String)
        'Map.Tile(Val(Parse(1)), Val(Parse(2))).Type = 0
    On Error Resume Next
        Wall(Val(parse(1)), Val(parse(2))).Here = False
    
End Sub

Sub HandleJump(ByRef parse() As String)

    Player(Val(parse(1))).Jumping = True
    Player(Val(parse(1))).JumpTimer = GetTickCount

End Sub

Sub HandleRoomList(ByRef parse() As String)
Dim RoomCount As Long
Dim i As Long
Dim n As Long

    frmLobby.lstRooms.Clear

    RoomCount = CLng(parse(1))
    
    n = 2
    
    If RoomCount > 0 Then
        For i = 1 To RoomCount
            frmLobby.lstRooms.AddItem parse(n) & " (" & parse(n + 1) & "/" & parse(n + 2) & ")" & " " & parse(n + 3)
            n = n + 4
        Next
    Else
        frmLobby.lstRooms.AddItem "Não há salas disponíveis."
    End If

End Sub

Sub HandleStartMatch(ByRef parse() As String)
Dim i As Long
Dim p As Long
Dim l As Long
Dim q As Long
Dim Players() As String

For l = 1 To PlayersOnMapHighIndex
If Player(PlayersOnMap(l)).Kills * 5 >= 0 Then
Player(PlayersOnMap(l)).Rank = "Bombinha de aguá"
End If

If Player(PlayersOnMap(l)).Kills * 5 > 20 Then
Player(PlayersOnMap(l)).Rank = "Bomba de aguá"
End If

If Player(PlayersOnMap(l)).Kills * 5 > 50 Then
Player(PlayersOnMap(l)).Rank = "Super bomba de aguá"
End If

If Player(PlayersOnMap(l)).Kills * 5 > 120 Then
Player(PlayersOnMap(l)).Rank = "Bombinha de terra"
End If

If Player(PlayersOnMap(l)).Kills * 5 > 240 Then
Player(PlayersOnMap(l)).Rank = "Bomba de terra"
End If

If Player(PlayersOnMap(l)).Kills * 5 > 500 Then
Player(PlayersOnMap(l)).Rank = "Super bomba de terra"
End If

If Player(PlayersOnMap(l)).Kills * 5 > 1200 Then
Player(PlayersOnMap(l)).Rank = "Bombinha de vento"
End If

If Player(PlayersOnMap(l)).Kills * 5 > 2600 Then
Player(PlayersOnMap(l)).Rank = "Bomba de vento"
End If

If Player(PlayersOnMap(l)).Kills * 5 > 5400 Then
Player(PlayersOnMap(l)).Rank = "Super bomba de vento"
End If

If Player(PlayersOnMap(l)).Kills * 5 > 11000 Then
Player(PlayersOnMap(l)).Rank = "Bombinha de fogo"
End If

If Player(PlayersOnMap(l)).Kills * 5 > 23000 Then
Player(PlayersOnMap(l)).Rank = "Bomba de fogo"
End If

If Player(PlayersOnMap(l)).Kills * 5 > 50000 Then
Player(PlayersOnMap(l)).Rank = "Super bomba de fogo"
End If
Next
    If Not frmWaiting.Visible Then
        frmWaiting.Visible = True
        frmLobby.Visible = False
    End If
    
    ReDim Players(1 To Val(parse(1))) As String
    
    For i = 1 To Val(parse(1))
    On Error GoTo Duelo
        Players(i) = parse(i + 1)
    Next
    
    frmWaiting.lstPlayers.Clear
    
    For i = 1 To Val(parse(1))
        frmWaiting.lstPlayers.AddItem parse(i + 1)
    Next
    
    frmWaiting.lstSprites.Clear
    frmWaiting.lstInfos.Clear
     
    For p = 1 To PlayersOnMapHighIndex
    frmWaiting.lstSprites.AddItem Player(PlayersOnMap(p)).sprite
    frmWaiting.lstInfos.AddItem Player(PlayersOnMap(p)).Kills & "/" & Player(PlayersOnMap(p)).Deaths
    
    Next

    frmWaiting.lblInfoPlayer1.Caption = frmWaiting.lstInfos.List(0)
    frmWaiting.lblInfoPlayer2.Caption = frmWaiting.lstInfos.List(1)
    frmWaiting.lblInfoPlayer3.Caption = frmWaiting.lstInfos.List(2)
    frmWaiting.lblInfoPlayer4.Caption = frmWaiting.lstInfos.List(3)
    frmWaiting.lblRoomPName1.Caption = frmWaiting.lstPlayers.List(0)
    frmWaiting.lblRoomPName2.Caption = frmWaiting.lstPlayers.List(1)
    frmWaiting.lblRoomPName3.Caption = frmWaiting.lstPlayers.List(2)
    frmWaiting.lblRoomPName4.Caption = frmWaiting.lstPlayers.List(3)
    
    For q = 1 To PlayersOnMapHighIndex
    
    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(0) And Player(PlayersOnMap(q)).Rank = "Bombinha de aguá" Then
    frmWaiting.imgRoomRank1.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank1" & ".gif")
    End If
    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(0) And Player(PlayersOnMap(q)).Rank = "Bomba de aguá" Then
    frmWaiting.imgRoomRank1.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank2" & ".gif")
    End If
    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(0) And Player(PlayersOnMap(q)).Rank = "Super bomba de aguá" Then
    frmWaiting.imgRoomRank1.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank3" & ".gif")
    End If
    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(0) And Player(PlayersOnMap(q)).Rank = "Bombinha de terra" Then
    frmWaiting.imgRoomRank1.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank4" & ".gif")
    End If
    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(0) And Player(PlayersOnMap(q)).Rank = "Bomba de terra" Then
    frmWaiting.imgRoomRank1.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank5" & ".gif")
    End If
    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(0) And Player(PlayersOnMap(q)).Rank = "Super bomba de terra" Then
    frmWaiting.imgRoomRank1.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank6" & ".gif")
    End If
    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(0) And Player(PlayersOnMap(q)).Rank = "Bombinha de vento" Then
    frmWaiting.imgRoomRank1.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank7" & ".gif")
    End If
    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(0) And Player(PlayersOnMap(q)).Rank = "Bomba de vento" Then
    frmWaiting.imgRoomRank1.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank8" & ".gif")
    End If
    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(0) And Player(PlayersOnMap(q)).Rank = "Super bomba de vento" Then
    frmWaiting.imgRoomRank1.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank9" & ".gif")
    End If
    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(0) And Player(PlayersOnMap(q)).Rank = "Bombinha de fogo" Then
    frmWaiting.imgRoomRank1.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank10" & ".gif")
    End If
    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(0) And Player(PlayersOnMap(q)).Rank = "Bomba de fogo" Then
    frmWaiting.imgRoomRank1.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank11" & ".gif")
    End If
    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(0) And Player(PlayersOnMap(q)).Rank = "Super bomba de fogo" Then
    frmWaiting.imgRoomRank1.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank12" & ".gif")
    End If


    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(1) And Player(PlayersOnMap(q)).Rank = "Bombinha de aguá" Then
    frmWaiting.imgRoomRank2.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank1" & ".gif")
    End If
    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(1) And Player(PlayersOnMap(q)).Rank = "Bomba de aguá" Then
    frmWaiting.imgRoomRank2.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank2" & ".gif")
    End If
    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(1) And Player(PlayersOnMap(q)).Rank = "Super bomba de aguá" Then
    frmWaiting.imgRoomRank2.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank3" & ".gif")
    End If
    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(1) And Player(PlayersOnMap(q)).Rank = "Bombinha de terra" Then
    frmWaiting.imgRoomRank2.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank4" & ".gif")
    End If
    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(1) And Player(PlayersOnMap(q)).Rank = "Bomba de terra" Then
    frmWaiting.imgRoomRank2.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank5" & ".gif")
    End If
    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(1) And Player(PlayersOnMap(q)).Rank = "Super bomba de terra" Then
    frmWaiting.imgRoomRank2.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank6" & ".gif")
    End If
    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(1) And Player(PlayersOnMap(q)).Rank = "Bombinha de vento" Then
    frmWaiting.imgRoomRank2.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank7" & ".gif")
    End If
    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(1) And Player(PlayersOnMap(q)).Rank = "Bomba de vento" Then
    frmWaiting.imgRoomRank2.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank8" & ".gif")
    End If
    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(1) And Player(PlayersOnMap(q)).Rank = "Super bomba de vento" Then
    frmWaiting.imgRoomRank2.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank9" & ".gif")
    End If
    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(1) And Player(PlayersOnMap(q)).Rank = "Bombinha de fogo" Then
    frmWaiting.imgRoomRank2.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank10" & ".gif")
    End If
    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(1) And Player(PlayersOnMap(q)).Rank = "Bomba de fogo" Then
    frmWaiting.imgRoomRank2.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank11" & ".gif")
    End If
    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(1) And Player(PlayersOnMap(q)).Rank = "Super bomba de fogo" Then
    frmWaiting.imgRoomRank2.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank12" & ".gif")
    End If


    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(2) And Player(PlayersOnMap(q)).Rank = "Bombinha de aguá" Then
    frmWaiting.imgRoomRank3.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank1" & ".gif")
    End If
    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(2) And Player(PlayersOnMap(q)).Rank = "Bomba de aguá" Then
    frmWaiting.imgRoomRank3.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank2" & ".gif")
    End If
    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(2) And Player(PlayersOnMap(q)).Rank = "Super bomba de aguá" Then
    frmWaiting.imgRoomRank3.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank3" & ".gif")
    End If
    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(2) And Player(PlayersOnMap(q)).Rank = "Bombinha de terra" Then
    frmWaiting.imgRoomRank3.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank4" & ".gif")
    End If
    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(2) And Player(PlayersOnMap(q)).Rank = "Bomba de terra" Then
    frmWaiting.imgRoomRank3.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank5" & ".gif")
    End If
    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(2) And Player(PlayersOnMap(q)).Rank = "Super bomba de terra" Then
    frmWaiting.imgRoomRank3.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank6" & ".gif")
    End If
    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(2) And Player(PlayersOnMap(q)).Rank = "Bombinha de vento" Then
    frmWaiting.imgRoomRank3.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank7" & ".gif")
    End If
    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(2) And Player(PlayersOnMap(q)).Rank = "Bomba de vento" Then
    frmWaiting.imgRoomRank3.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank8" & ".gif")
    End If
    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(2) And Player(PlayersOnMap(q)).Rank = "Super bomba de vento" Then
    frmWaiting.imgRoomRank3.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank9" & ".gif")
    End If
    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(2) And Player(PlayersOnMap(q)).Rank = "Bombinha de fogo" Then
    frmWaiting.imgRoomRank3.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank10" & ".gif")
    End If
    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(2) And Player(PlayersOnMap(q)).Rank = "Bomba de fogo" Then
    frmWaiting.imgRoomRank3.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank11" & ".gif")
    End If
    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(2) And Player(PlayersOnMap(q)).Rank = "Super bomba de fogo" Then
    frmWaiting.imgRoomRank3.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank12" & ".gif")
    End If


    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(3) And Player(PlayersOnMap(q)).Rank = "Bombinha de aguá" Then
    frmWaiting.imgRoomRank4.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank1" & ".gif")
    End If
    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(3) And Player(PlayersOnMap(q)).Rank = "Bomba de aguá" Then
    frmWaiting.imgRoomRank4.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank2" & ".gif")
    End If
    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(3) And Player(PlayersOnMap(q)).Rank = "Super bomba de aguá" Then
    frmWaiting.imgRoomRank4.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank3" & ".gif")
    End If
    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(3) And Player(PlayersOnMap(q)).Rank = "Bombinha de terra" Then
    frmWaiting.imgRoomRank4.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank4" & ".gif")
    End If
    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(3) And Player(PlayersOnMap(q)).Rank = "Bomba de terra" Then
    frmWaiting.imgRoomRank4.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank5" & ".gif")
    End If
    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(3) And Player(PlayersOnMap(q)).Rank = "Super bomba de terra" Then
    frmWaiting.imgRoomRank4.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank6" & ".gif")
    End If
    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(3) And Player(PlayersOnMap(q)).Rank = "Bombinha de vento" Then
    frmWaiting.imgRoomRank4.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank7" & ".gif")
    End If
    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(3) And Player(PlayersOnMap(q)).Rank = "Bomba de vento" Then
    frmWaiting.imgRoomRank4.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank8" & ".gif")
    End If
    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(3) And Player(PlayersOnMap(q)).Rank = "Super bomba de vento" Then
    frmWaiting.imgRoomRank4.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank9" & ".gif")
    End If
    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(3) And Player(PlayersOnMap(q)).Rank = "Bombinha de fogo" Then
    frmWaiting.imgRoomRank4.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank10" & ".gif")
    End If
    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(3) And Player(PlayersOnMap(q)).Rank = "Bomba de fogo" Then
    frmWaiting.imgRoomRank4.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank11" & ".gif")
    End If
    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(3) And Player(PlayersOnMap(q)).Rank = "Super bomba de fogo" Then
    frmWaiting.imgRoomRank4.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank12" & ".gif")
    End If
    

    Next
Duelo:
        For i = 1 To 2
        Players(i) = parse(i + 1)
    Next
    
    frmWaiting.lstPlayers.Clear
    
    For i = 1 To 2
        frmWaiting.lstPlayers.AddItem parse(i + 1)
    Next
    
    frmWaiting.lstSprites.Clear
    frmWaiting.lstInfos.Clear
     
    For p = 1 To PlayersOnMapHighIndex
    frmWaiting.lstSprites.AddItem Player(PlayersOnMap(p)).sprite
    frmWaiting.lstInfos.AddItem Player(PlayersOnMap(p)).Kills & "/" & Player(PlayersOnMap(p)).Deaths
    
    Next

    frmWaiting.lblInfoPlayer1.Caption = frmWaiting.lstInfos.List(0)
    frmWaiting.lblInfoPlayer2.Caption = frmWaiting.lstInfos.List(1)
    frmWaiting.lblInfoPlayer3.Caption = frmWaiting.lstInfos.List(2)
    frmWaiting.lblInfoPlayer4.Caption = frmWaiting.lstInfos.List(3)
    frmWaiting.lblRoomPName1.Caption = frmWaiting.lstPlayers.List(0)
    frmWaiting.lblRoomPName2.Caption = frmWaiting.lstPlayers.List(1)
    frmWaiting.lblRoomPName3.Caption = frmWaiting.lstPlayers.List(2)
    frmWaiting.lblRoomPName4.Caption = frmWaiting.lstPlayers.List(3)
    
    For q = 1 To PlayersOnMapHighIndex
    
    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(0) And Player(PlayersOnMap(q)).Rank = "Bombinha de aguá" Then
    frmWaiting.imgRoomRank1.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank1" & ".gif")
    End If
    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(0) And Player(PlayersOnMap(q)).Rank = "Bomba de aguá" Then
    frmWaiting.imgRoomRank1.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank2" & ".gif")
    End If
    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(0) And Player(PlayersOnMap(q)).Rank = "Super bomba de aguá" Then
    frmWaiting.imgRoomRank1.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank3" & ".gif")
    End If
    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(0) And Player(PlayersOnMap(q)).Rank = "Bombinha de terra" Then
    frmWaiting.imgRoomRank1.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank4" & ".gif")
    End If
    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(0) And Player(PlayersOnMap(q)).Rank = "Bomba de terra" Then
    frmWaiting.imgRoomRank1.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank5" & ".gif")
    End If
    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(0) And Player(PlayersOnMap(q)).Rank = "Super bomba de terra" Then
    frmWaiting.imgRoomRank1.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank6" & ".gif")
    End If
    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(0) And Player(PlayersOnMap(q)).Rank = "Bombinha de vento" Then
    frmWaiting.imgRoomRank1.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank7" & ".gif")
    End If
    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(0) And Player(PlayersOnMap(q)).Rank = "Bomba de vento" Then
    frmWaiting.imgRoomRank1.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank8" & ".gif")
    End If
    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(0) And Player(PlayersOnMap(q)).Rank = "Super bomba de vento" Then
    frmWaiting.imgRoomRank1.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank9" & ".gif")
    End If
    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(0) And Player(PlayersOnMap(q)).Rank = "Bombinha de fogo" Then
    frmWaiting.imgRoomRank1.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank10" & ".gif")
    End If
    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(0) And Player(PlayersOnMap(q)).Rank = "Bomba de fogo" Then
    frmWaiting.imgRoomRank1.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank11" & ".gif")
    End If
    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(0) And Player(PlayersOnMap(q)).Rank = "Super bomba de fogo" Then
    frmWaiting.imgRoomRank1.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank12" & ".gif")
    End If


    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(1) And Player(PlayersOnMap(q)).Rank = "Bombinha de aguá" Then
    frmWaiting.imgRoomRank2.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank1" & ".gif")
    End If
    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(1) And Player(PlayersOnMap(q)).Rank = "Bomba de aguá" Then
    frmWaiting.imgRoomRank2.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank2" & ".gif")
    End If
    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(1) And Player(PlayersOnMap(q)).Rank = "Super bomba de aguá" Then
    frmWaiting.imgRoomRank2.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank3" & ".gif")
    End If
    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(1) And Player(PlayersOnMap(q)).Rank = "Bombinha de terra" Then
    frmWaiting.imgRoomRank2.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank4" & ".gif")
    End If
    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(1) And Player(PlayersOnMap(q)).Rank = "Bomba de terra" Then
    frmWaiting.imgRoomRank2.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank5" & ".gif")
    End If
    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(1) And Player(PlayersOnMap(q)).Rank = "Super bomba de terra" Then
    frmWaiting.imgRoomRank2.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank6" & ".gif")
    End If
    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(1) And Player(PlayersOnMap(q)).Rank = "Bombinha de vento" Then
    frmWaiting.imgRoomRank2.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank7" & ".gif")
    End If
    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(1) And Player(PlayersOnMap(q)).Rank = "Bomba de vento" Then
    frmWaiting.imgRoomRank2.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank8" & ".gif")
    End If
    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(1) And Player(PlayersOnMap(q)).Rank = "Super bomba de vento" Then
    frmWaiting.imgRoomRank2.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank9" & ".gif")
    End If
    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(1) And Player(PlayersOnMap(q)).Rank = "Bombinha de fogo" Then
    frmWaiting.imgRoomRank2.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank10" & ".gif")
    End If
    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(1) And Player(PlayersOnMap(q)).Rank = "Bomba de fogo" Then
    frmWaiting.imgRoomRank2.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank11" & ".gif")
    End If
    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(1) And Player(PlayersOnMap(q)).Rank = "Super bomba de fogo" Then
    frmWaiting.imgRoomRank2.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank12" & ".gif")
    End If


    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(2) And Player(PlayersOnMap(q)).Rank = "Bombinha de aguá" Then
    frmWaiting.imgRoomRank3.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank1" & ".gif")
    End If
    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(2) And Player(PlayersOnMap(q)).Rank = "Bomba de aguá" Then
    frmWaiting.imgRoomRank3.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank2" & ".gif")
    End If
    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(2) And Player(PlayersOnMap(q)).Rank = "Super bomba de aguá" Then
    frmWaiting.imgRoomRank3.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank3" & ".gif")
    End If
    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(2) And Player(PlayersOnMap(q)).Rank = "Bombinha de terra" Then
    frmWaiting.imgRoomRank3.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank4" & ".gif")
    End If
    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(2) And Player(PlayersOnMap(q)).Rank = "Bomba de terra" Then
    frmWaiting.imgRoomRank3.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank5" & ".gif")
    End If
    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(2) And Player(PlayersOnMap(q)).Rank = "Super bomba de terra" Then
    frmWaiting.imgRoomRank3.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank6" & ".gif")
    End If
    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(2) And Player(PlayersOnMap(q)).Rank = "Bombinha de vento" Then
    frmWaiting.imgRoomRank3.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank7" & ".gif")
    End If
    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(2) And Player(PlayersOnMap(q)).Rank = "Bomba de vento" Then
    frmWaiting.imgRoomRank3.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank8" & ".gif")
    End If
    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(2) And Player(PlayersOnMap(q)).Rank = "Super bomba de vento" Then
    frmWaiting.imgRoomRank3.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank9" & ".gif")
    End If
    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(2) And Player(PlayersOnMap(q)).Rank = "Bombinha de fogo" Then
    frmWaiting.imgRoomRank3.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank10" & ".gif")
    End If
    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(2) And Player(PlayersOnMap(q)).Rank = "Bomba de fogo" Then
    frmWaiting.imgRoomRank3.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank11" & ".gif")
    End If
    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(2) And Player(PlayersOnMap(q)).Rank = "Super bomba de fogo" Then
    frmWaiting.imgRoomRank3.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank12" & ".gif")
    End If


    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(3) And Player(PlayersOnMap(q)).Rank = "Bombinha de aguá" Then
    frmWaiting.imgRoomRank4.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank1" & ".gif")
    End If
    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(3) And Player(PlayersOnMap(q)).Rank = "Bomba de aguá" Then
    frmWaiting.imgRoomRank4.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank2" & ".gif")
    End If
    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(3) And Player(PlayersOnMap(q)).Rank = "Super bomba de aguá" Then
    frmWaiting.imgRoomRank4.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank3" & ".gif")
    End If
    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(3) And Player(PlayersOnMap(q)).Rank = "Bombinha de terra" Then
    frmWaiting.imgRoomRank4.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank4" & ".gif")
    End If
    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(3) And Player(PlayersOnMap(q)).Rank = "Bomba de terra" Then
    frmWaiting.imgRoomRank4.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank5" & ".gif")
    End If
    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(3) And Player(PlayersOnMap(q)).Rank = "Super bomba de terra" Then
    frmWaiting.imgRoomRank4.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank6" & ".gif")
    End If
    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(3) And Player(PlayersOnMap(q)).Rank = "Bombinha de vento" Then
    frmWaiting.imgRoomRank4.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank7" & ".gif")
    End If
    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(3) And Player(PlayersOnMap(q)).Rank = "Bomba de vento" Then
    frmWaiting.imgRoomRank4.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank8" & ".gif")
    End If
    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(3) And Player(PlayersOnMap(q)).Rank = "Super bomba de vento" Then
    frmWaiting.imgRoomRank4.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank9" & ".gif")
    End If
    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(3) And Player(PlayersOnMap(q)).Rank = "Bombinha de fogo" Then
    frmWaiting.imgRoomRank4.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank10" & ".gif")
    End If
    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(3) And Player(PlayersOnMap(q)).Rank = "Bomba de fogo" Then
    frmWaiting.imgRoomRank4.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank11" & ".gif")
    End If
    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(3) And Player(PlayersOnMap(q)).Rank = "Super bomba de fogo" Then
    frmWaiting.imgRoomRank4.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank12" & ".gif")
    End If
    

    Next
End Sub

Sub HandleStartTimer()
Dim i As Long

    frmWaiting.tmrGameStart.Enabled = True
    For i = 1 To PlayersOnMapHighIndex
    If Player(PlayersOnMap(i)).Watching = True Then
    Player(PlayersOnMap(i)).Watching = False
    End If
    Next

End Sub

Sub HandleWatching(ByRef parse() As String)
Dim i As Long

If Val(parse(2)) = "8" Then
Player(parse(1)).Watching = False
End If

If Val(parse(2)) = "9" Then
Player(Val(parse(1))).Watching = True
End If

End Sub

Sub HandleLeaveRoom()

    frmLobby.txtChat.Text = vbNullString
    frmWaiting.Hide
    frmMainGame.Hide
    frmLobby.Show
    isLogging = True
    InGame = False
    Player(MyIndex).Watching = False
    
End Sub

Sub HandleHighScore(ByRef parse() As String)
Dim Index As Long

    Index = Val(parse(1))
    
    Player(Index).Kills = Val(parse(2))
    Player(Index).Deaths = Val(parse(3))
    Player(Index).MatchsWon = Val(parse(4))
    Player(Index).BPoints = Val(parse(5))
    Player(Index).BCash = Val(parse(6))
    
    If Index = MyIndex Then
        frmMainGame.lblKills = Player(Index).Kills
        frmMainGame.lblPlayerExp.Caption = Player(MyIndex).Kills * 5
        frmMainGame.lblDeaths = Player(Index).Deaths
        frmLobby.lblKillDeath = Player(MyIndex).Kills & "/" & Player(MyIndex).Deaths
        frmLobby.lblExp = Player(Index).Kills * 5
        frmLobby.lblBP = Player(Index).BPoints
        frmLobby.lblBC = Player(Index).BCash
    End If

End Sub

Sub HandleHighScoreList(ByRef parse() As String)
Dim i As Long
Dim n As Long
    
    n = 1
    
    For i = 1 To 100
        frmHighScores.lstHighScores.ListItems.Add i
        If parse(n) <> vbNullString Then
            frmHighScores.lstHighScores.ListItems(i).Text = i
            frmHighScores.lstHighScores.ListItems(i).SubItems(1) = parse(n)
            frmHighScores.lstHighScores.ListItems(i).SubItems(2) = parse(n + 3)
            frmHighScores.lstHighScores.ListItems(i).SubItems(3) = parse(n + 1)
            frmHighScores.lstHighScores.ListItems(i).SubItems(4) = parse(n + 2)
        Else
            frmHighScores.lstHighScores.ListItems(i).Text = i
            frmHighScores.lstHighScores.ListItems(i).SubItems(1) = "Nada"
            frmHighScores.lstHighScores.ListItems(i).SubItems(2) = "0"
            frmHighScores.lstHighScores.ListItems(i).SubItems(3) = "0"
            frmHighScores.lstHighScores.ListItems(i).SubItems(4) = "0"
        End If
        n = n + 4
    Next
    
    SortColumn frmHighScores.lstHighScores, 1, sortNumeric, sortAscending
    
    frmHighScores.Show
    frmLobby.Hide

End Sub
Sub HandleUpdateFriendList(ByRef parse() As String)
Dim i As Long
       'Prevents error and clears your friends list when you have no friends
        If parse(2) = 0 Then
            frmLobby.lstFriend.Clear
            frmLobby.lstFriend.AddItem "Não há amigos online"
            Exit Sub
        End If
    
    'clear lstbox so it can be updated correctly.

    
    'Adds Friends Name to the List
    'For i = 1 To MAX_FRIENDS
           ' If parse(3) = " (Offline)" Then
              '  GoTo Continue
           ' Else
                frmLobby.lstFriend.AddItem parse(1) & parse(3)
           ' End If
        
'Continue:
  '  Next
    
    'If frmLobby.lstFriend.ListCount = 0 Then
   '     frmLobby.lstFriend.AddItem "Não há amigos online"
   ' End If
End Sub
Sub HandleSpeed()
Walk_Speed = 8
Run_Speed = 8
BonusSpeed = True
End Sub
