Attribute VB_Name = "modClientTCP"
Option Explicit

' ******************************************
' **              BMO Source              **
' ******************************************
' ** Communcation to server, TCP          **
' ** Winsock Control (mswinsck.ocx)       **
' ** String packets (slow and big)        **
' ******************************************

Sub TcpInit()

    ' used for parsing packets
    SEP_CHAR = vbNullChar ' ChrW$(0)
    END_CHAR = ChrW$(237)
    
    GAME_IP = GetVar(App.Path & "\config.ini", "Opções", "IP")
    GAME_PORT = GetVar(App.Path & "\config.ini", "Opções", "Port")
    
    frmMainGame.Socket.RemoteHost = GAME_IP
    frmMainGame.Socket.RemotePort = GAME_PORT
    
End Sub

Sub DestroyTCP()
    Call DestroyDirectDraw
    frmMainGame.Socket.Close
End Sub

Sub IncomingData(ByVal DataLength As Long)
Dim Buffer As String
Dim packet As String
Dim Start As Integer

    frmMainGame.Socket.GetData Buffer, vbString, DataLength
    PlayerBuffer = PlayerBuffer & Buffer
        
    Start = InStr(PlayerBuffer, END_CHAR)
    Do While Start > 0
        packet = Mid$(PlayerBuffer, 1, Start - 1)
        PlayerBuffer = Mid$(PlayerBuffer, Start + 1, Len(PlayerBuffer))
        Start = InStr(PlayerBuffer, END_CHAR)
        
        If Len(packet) > 0 Then
            Call HandleData(packet)
        End If
    Loop
End Sub

Public Function ConnectToServer() As Boolean
Dim Wait As Long

    ' Check to see if we are already connected, if so just exit
    If IsConnected Then
        ConnectToServer = True
        Exit Function
    End If
    
    Wait = GetTickCount
    With frmMainGame.Socket
        .Close
        .Connect
    End With
    
    ' Wait until connected or 4 seconds have passed and report the server being down
    Do While (Not IsConnected) And (GetTickCount <= Wait + 2000)
        DoEvents
        Sleep 20
    Loop
    
    ' return value
    If IsConnected Then
        ConnectToServer = True
    Else
        ConnectToServer = False
    End If
End Function

Private Function IsIP(ByVal IPAddress As String) As Boolean
Dim s() As String
Dim i As Long

    ' Check if connecting to localhost or URL
    If IPAddress = "localhost" Or InStr(1, IPAddress, "http://", vbTextCompare) = 1 Then
        IsIP = True
        Exit Function
    End If

    'If there are no periods, I have no idea what we have...
    If InStr(1, IPAddress, ".") = 0 Then Exit Function
    
    'Split up the string by the periods
    s = Split(IPAddress, ".")
    
    'Confirm we have ubound = 3, since xxx.xxx.xxx.xxx has 4 elements and we start at index 0
    If UBound(s) <> 3 Then Exit Function
    
    'Check that the values are numeric and in a valid range
    For i = 0 To 3
        If Val(s(i)) < 0 Then Exit Function
        If Val(s(i)) > 255 Then Exit Function
    Next
    
    'Looks like we were passed a valid IP!
    IsIP = True
    
End Function

Function IsConnected() As Boolean
    If frmMainGame.Socket.State = sckConnected Then
        IsConnected = True
    End If
End Function

Function IsPlaying(ByVal Index As Long) As Boolean
    ' if the player doesn't exist, the name will equal 0
    If LenB(GetPlayerName(Index)) > 0 Then
        IsPlaying = True
    End If
End Function

Sub SendData(ByVal Data As String)
    ' check if connection exist, otherwise will error
    If IsConnected Then
        frmMainGame.Socket.SendData Data
        DoEvents
    End If
End Sub

' ******************************
' ** Outcoming Client Packets **
' ******************************

Sub SendNewAccount(ByVal Name As String, ByVal Password As String, ByVal Gender As Byte, ByVal Class As Byte)
Dim packet As String

    packet = CNewAccount & SEP_CHAR & Trim$(Name) & SEP_CHAR & Trim$(Password) & SEP_CHAR & Gender & SEP_CHAR & Class & END_CHAR
    Call SendData(packet)
End Sub

Sub SendDelAccount(ByVal Name As String, ByVal Password As String)
Dim packet As String
    
    packet = CDelAccount & SEP_CHAR & Trim$(Name) & SEP_CHAR & Trim$(Password) & END_CHAR
    Call SendData(packet)
End Sub

Sub SendLogin(ByVal Name As String, ByVal Password As String)
Dim packet As String

    packet = CLogin & SEP_CHAR & Trim$(Name) & SEP_CHAR & Trim$(Password) & SEP_CHAR & App.Major & SEP_CHAR & App.Minor & SEP_CHAR & App.Revision & END_CHAR
    Call SendData(packet)
End Sub

Sub SendAddChar(ByVal Name As String, ByVal Sex As Long, ByVal ClassNum As Long, ByVal Slot As Long)
Dim packet As String

    packet = CAddChar & SEP_CHAR & Trim$(Name) & SEP_CHAR & Sex & SEP_CHAR & ClassNum & SEP_CHAR & Slot & END_CHAR
    Call SendData(packet)
End Sub

Sub SendDelChar(ByVal Slot As Long)
Dim packet As String
    
    packet = CDelChar & SEP_CHAR & Slot & END_CHAR
    Call SendData(packet)
End Sub

Sub SendGetClasses()
Dim packet As String

    packet = CGetClasses & END_CHAR
    Call SendData(packet)
    
End Sub

Sub SendUseChar(ByVal CharSlot As Long)
Dim packet As String

    packet = CUseChar & SEP_CHAR & 1 & END_CHAR
    Call SendData(packet)
    
End Sub

Sub SayMsg(ByVal Text As String)
Dim packet As String

    packet = CSayMsg & SEP_CHAR & Text & END_CHAR
    Call SendData(packet)
End Sub

Sub GlobalMsg(ByVal Text As String)
Dim packet As String

    packet = CGlobalMsg & SEP_CHAR & Text & END_CHAR
    Call SendData(packet)
End Sub

Sub BroadcastMsg(ByVal Text As String)
Dim packet As String

    packet = CBroadcastMsg & SEP_CHAR & Text & END_CHAR
    Call SendData(packet)
End Sub

Sub EmoteMsg(ByVal Text As String)
Dim packet As String

    packet = CEmoteMsg & SEP_CHAR & Text & END_CHAR
    Call SendData(packet)
End Sub

Sub PlayerMsg(ByVal Text As String, ByVal MsgTo As String)
Dim packet As String

    packet = CPlayerMsg & SEP_CHAR & MsgTo & SEP_CHAR & Text & END_CHAR
    Call SendData(packet)
End Sub

Sub AdminMsg(ByVal Text As String)
Dim packet As String

    packet = CAdminMsg & SEP_CHAR & Text & END_CHAR
    Call SendData(packet)
End Sub

Sub SendPlayerMove()
Dim packet As String

    packet = CPlayerMove & SEP_CHAR & GetPlayerDir(MyIndex) & SEP_CHAR & Player(MyIndex).Moving & END_CHAR
    Call SendData(packet)
End Sub

Sub SendPlayerDir()
Dim packet As String

    packet = CPlayerDir & SEP_CHAR & GetPlayerDir(MyIndex) & END_CHAR
    Call SendData(packet)
End Sub

Sub SendPlayerRequestNewMap()
Dim packet As String
    
    packet = CRequestNewMap & SEP_CHAR & GetPlayerDir(MyIndex) & END_CHAR
    Call SendData(packet)
End Sub

Public Sub SendMap()
Dim packet As String
Dim x As Long
Dim y As Long

    CanMoveNow = False
    
    With Map
        packet = CMapData & SEP_CHAR & Trim$(.Name) & SEP_CHAR & .PlayerLimit & SEP_CHAR & .Moral & SEP_CHAR & .TileSet & SEP_CHAR & .Up & SEP_CHAR & .Down & SEP_CHAR & .Left & SEP_CHAR & .Right & SEP_CHAR & .Music & SEP_CHAR & .BootMap & SEP_CHAR & .BootX & SEP_CHAR & .BootY & SEP_CHAR & .Shop
    End With
    
    For x = 0 To MAX_MAPX
        For y = 0 To MAX_MAPY
            With Map.Tile(x, y)
                packet = packet & SEP_CHAR & .Ground & SEP_CHAR & .Mask & SEP_CHAR & .Anim & SEP_CHAR & .Fringe & SEP_CHAR & .Type & SEP_CHAR & .Data1 & SEP_CHAR & .Data2 & SEP_CHAR & .Data3
            End With
        Next
    Next
    
    With Map
        For x = 1 To MAX_MAP_NPCS
            packet = packet & SEP_CHAR & .Npc(x)
        Next
    End With
    
    packet = packet & END_CHAR
    
    Call SendData(packet)
End Sub

Sub WarpMeTo(ByVal Name As String)
Dim packet As String

    packet = CWarpMeTo & SEP_CHAR & Name & END_CHAR
    Call SendData(packet)
End Sub

Sub WarpToMe(ByVal Name As String)
Dim packet As String

    packet = CWarpToMe & SEP_CHAR & Name & END_CHAR
    Call SendData(packet)
End Sub

Sub WarpTo(ByVal MapNum As Long)
Dim packet As String
    
    packet = CWarpTo & SEP_CHAR & MapNum & END_CHAR
    Call SendData(packet)
End Sub

Sub SendSetAccess(ByVal Name As String, ByVal Access As Byte)
Dim packet As String

    packet = CSetAccess & SEP_CHAR & Name & SEP_CHAR & Access & END_CHAR
    Call SendData(packet)
End Sub

Sub SendSetSprite(ByVal SpriteNum As Long)
Dim packet As String

    packet = CSetSprite & SEP_CHAR & SpriteNum & END_CHAR
    Call SendData(packet)
End Sub

Sub SendKick(ByVal Name As String)
Dim packet As String

    packet = CKickPlayer & SEP_CHAR & Name & END_CHAR
    Call SendData(packet)
End Sub

Sub SendBan(ByVal Name As String)
Dim packet As String

    packet = CBanPlayer & SEP_CHAR & Name & END_CHAR
    Call SendData(packet)
End Sub

Sub SendBanList()
Dim packet As String

    packet = CBanList & END_CHAR
    Call SendData(packet)
End Sub

Sub SendRequestEditItem()
Dim packet As String

    packet = CRequestEditItem & END_CHAR
    Call SendData(packet)
End Sub

Public Sub SendSaveItem(ByVal ItemNum As Long)
Dim packet As String
    
    With Item(ItemNum)
        packet = CSaveItem & SEP_CHAR & ItemNum & SEP_CHAR & Trim$(.Name) & SEP_CHAR & .Pic & SEP_CHAR & .Type & SEP_CHAR & .Data1 & SEP_CHAR & .Data2 & SEP_CHAR & .Data3 & END_CHAR
    End With
    
    Call SendData(packet)
End Sub
                
Sub SendRequestEditNpc()
Dim packet As String

    packet = CRequestEditNpc & END_CHAR
    Call SendData(packet)
End Sub

Public Sub SendSaveNpc(ByVal NpcNum As Long)
Dim packet As String
    
    With Npc(NpcNum)
        packet = CSaveNpc & SEP_CHAR & NpcNum & SEP_CHAR & Trim$(.Name) & SEP_CHAR & Trim$(.AttackSay) & SEP_CHAR & .sprite & SEP_CHAR & .SpawnSecs & SEP_CHAR & .Behavior & SEP_CHAR & .Range & SEP_CHAR & .DropChance & SEP_CHAR & .DropItem & SEP_CHAR & .DropItemValue & SEP_CHAR & .Stat(Stats.Strength) & SEP_CHAR & .Stat(Stats.Defense) & SEP_CHAR & .Stat(Stats.SPEED) & SEP_CHAR & .Stat(Stats.Magic) & END_CHAR
    End With
    
    Call SendData(packet)
End Sub

Sub SendMapRespawn()
Dim packet As String

    packet = CMapRespawn & END_CHAR
    Call SendData(packet)
End Sub

Sub SendUseItem(ByVal InvNum As Long)
Dim packet As String

    packet = CUseItem & SEP_CHAR & InvNum & END_CHAR
    Call SendData(packet)
End Sub

Sub SendDropItem(ByVal InvNum As Long, ByVal Amount As Long)
Dim packet As String

    packet = CMapDropItem & SEP_CHAR & InvNum & SEP_CHAR & Amount & END_CHAR
    Call SendData(packet)
End Sub

Sub SendWhosOnline()
Dim packet As String

    packet = CWhosOnline & END_CHAR
    Call SendData(packet)
End Sub
            
Sub SendMOTDChange(ByVal MOTD As String)
Dim packet As String

    packet = CSetMotd & SEP_CHAR & MOTD & END_CHAR
    Call SendData(packet)
End Sub

Sub SendRequestEditShop()
Dim packet As String

    packet = CRequestEditShop & END_CHAR
    Call SendData(packet)
End Sub

Public Sub SendSaveShop(ByVal ShopNum As Long)
Dim packet As String
Dim i As Long

    With Shop(ShopNum)
        packet = CSaveShop & SEP_CHAR & ShopNum & SEP_CHAR & Trim$(.Name) & SEP_CHAR & Trim$(.JoinSay) & SEP_CHAR & Trim$(.LeaveSay) & SEP_CHAR & .FixesItems
    End With
    
    For i = 1 To MAX_TRADES
        With Shop(ShopNum).TradeItem(i)
            packet = packet & SEP_CHAR & .GiveItem & SEP_CHAR & .GiveValue & SEP_CHAR & .GetItem & SEP_CHAR & .GetValue
        End With
    Next
    
    packet = packet & END_CHAR
    Call SendData(packet)
End Sub

Sub SendRequestEditSpell()
Dim packet As String

    packet = CRequestEditSpell & END_CHAR
    Call SendData(packet)
End Sub

Public Sub SendSaveSpell(ByVal SpellNum As Long)
Dim packet As String

    With Spell(SpellNum)
        packet = CSaveSpell & SEP_CHAR & SpellNum & SEP_CHAR & Trim$(.Name) & SEP_CHAR & .ClassReq & SEP_CHAR & .LevelReq & SEP_CHAR & .Type & SEP_CHAR & .Data1 & SEP_CHAR & .Data2 & SEP_CHAR & .Data3 & END_CHAR
    End With
    
    Call SendData(packet)
End Sub

Sub SendRequestEditMap()
Dim packet As String

    packet = CRequestEditMap & END_CHAR
    Call SendData(packet)
End Sub

Sub SendPartyRequest(ByVal Name As String)
Dim packet As String

    packet = CParty & SEP_CHAR & Name & END_CHAR
    Call SendData(packet)
End Sub

Sub SendJoinParty()
Dim packet As String

    packet = CJoinParty & END_CHAR
    Call SendData(packet)
End Sub

Sub SendLeaveParty()
Dim packet As String

    packet = CLeaveParty & END_CHAR
    Call SendData(packet)
End Sub

Sub SendBanDestroy()
Dim packet As String
    
    packet = CBanDestroy & END_CHAR
    Call SendData(packet)
End Sub
Sub addfriend(ByVal addfriend As String)
Dim packet As String
    
    packet = CAddFriend & SEP_CHAR & addfriend & END_CHAR
    Call SendData(packet)
End Sub
Sub RemoveFriend(ByVal FriendRemove As String)
Dim packet As String
    
    packet = CRemoveFriend & SEP_CHAR & FriendRemove & END_CHAR
    Call SendData(packet)
End Sub
Sub UpdateFriendList()
Dim packet As String
    
    packet = CUpdateFriendList & END_CHAR
    Call SendData(packet)
End Sub
Sub AddStatus(ByVal alcance As Long, ByVal limite As Long)
Dim packet As String

packet = CAddStatus & SEP_CHAR & alcance & SEP_CHAR & limite & END_CHAR
Call SendData(packet)
End Sub
Sub DealBP(ByVal valor As Long)
Dim packet As String

packet = CDealBP & SEP_CHAR & valor & END_CHAR
Call SendData(packet)
End Sub
Sub DealBC(ByVal valor As Long)
Dim packet As String

packet = CDealBC & SEP_CHAR & valor & END_CHAR
Call SendData(packet)
End Sub
Sub SendSprite(ByVal sprite As Long)
Dim packet As String
    packet = CSaveSprite & SEP_CHAR & Player(MyIndex).sprite & END_CHAR
    Call SendData(packet)
End Sub
