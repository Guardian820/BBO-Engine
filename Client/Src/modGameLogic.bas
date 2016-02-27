Attribute VB_Name = "modGameLogic"
Option Explicit

' ******************************************
' **              BMO Source              **
' ******************************************

Public Sub GameLoop()
    Dim Tick As Long
    Dim TickFPS As Long
    Dim FPS As Long
    Dim i As Long
    Dim WalkTimer As Long
    Dim tmr25 As Long

    ' Used for calculating fps
    TickFPS = GetTickCount
    FPS = 0

    ' ** Start GameLoop **
    Do While InGame
    
        Tick = GetTickCount
        
        InGame = IsConnected
        
        If frmLobby.Visible Then frmLobby.Hide
        
        If tmr25 < Tick Then
            
            ' Check to make sure they aren't trying to auto do anything
            Call CheckKeys
            
            ' Change map animation every 250 milliseconds
            If MapAnimTimer < Tick Then
                MapAnim = Not MapAnim
                MapAnimTimer = Tick + 250
            End If
            
            tmr25 = Tick + 25
        End If
        
        ' Process input before rendering, otherwise input will be behind by 1 frame
        If WalkTimer < Tick Then
            ' Process player movements (actually move them)
            For i = 1 To PlayersOnMapHighIndex
                If Player(PlayersOnMap(i)).Moving > 0 Then
                    Call ProcessMovement(PlayersOnMap(i))
                End If
            Next
            WalkTimer = Tick + 30 ' edit this value to change WalkTimer
        End If
        
        If CanMoveNow Then
            ' Check if player is trying to move
            Call CheckMovement
            
            ' Check to see if player is trying to attack
            Call CheckAttack
            Call CheckThrow
        End If
        
        CheckMapGetItem
        
        ' *********************
        ' ** Render Graphics **
        ' *********************
        Render_Graphics
        
        ' Calculate fps
        If TickFPS < Tick Then
            GameFPS = FPS
            TickFPS = Tick + 1000
            FPS = 0
        Else
            FPS = FPS + 1
        End If
        
        DoEvents
        Sleep 1
        
        ' FPS cap
        If (GetTickCount - Tick) < 31 Then
            Sleep 31 - (GetTickCount - Tick)
        End If
        
    Loop
    
    frmMainGame.Visible = False
    
    If isLogging Then
        frmLobby.Show
        frmMainGame.Hide
        frmWaiting.Hide
        GettingMap = True
    Else
         ' Shutdown the game
        frmMain.picLoading.Visible = True
        Call SetStatus("Apagando informações do jogo...")
        Call DestroyGame
    End If
    
End Sub

Public Sub CheaterLoop()

    Do
        ' WPE Pro? ---------------------
        TargetName = "WPE PRO"
        TargetHwnd = 0
        
        ' Examine the window names.
        EnumWindows AddressOf WindowEnumerator, 0
        
        ' See if we got an hwnd.
        If TargetHwnd > 0 Then
            MsgBox "Cheat detectado,fechando o jogo..."
            DestroyGame
        End If
        
        ' Cheat Engine? ---------------------
        TargetName = "Cheat Engine"
        TargetHwnd = 0
        
        ' Examine the window names.
        EnumWindows AddressOf WindowEnumerator, 0
        
        ' See if we got an hwnd.
        If TargetHwnd > 0 Then
            MsgBox "Cheat detectado,fechando o jogo..."
            DestroyGame
        End If
        
        DoEvents
    Loop
    
End Sub

Sub ProcessMovement(ByVal Index As Long)
    
    Select Case GetPlayerDir(Index)
        Case DIR_UP
            Player(Index).YOffset = Player(Index).YOffset - WALK_SPEED
        Case DIR_DOWN
            Player(Index).YOffset = Player(Index).YOffset + WALK_SPEED
        Case DIR_LEFT
            Player(Index).XOffset = Player(Index).XOffset - WALK_SPEED
        Case DIR_RIGHT
            Player(Index).XOffset = Player(Index).XOffset + WALK_SPEED
    End Select

    ' Check if completed walking over to the next tile
    If Player(Index).XOffset = 0 Then
        If Player(Index).YOffset = 0 Then
            Player(Index).Moving = 0
        End If
    End If

End Sub

Sub ProcessNpcMovement(ByVal MapNpcNum As Long)

End Sub
Sub HandleKeypresses(ByVal KeyAscii As Integer)
Dim ChatText As String
Dim Name As String
Dim i As Long
Dim n As Long
Dim Command() As String

    ChatText = Trim$(MyText)
    
    If LenB(ChatText) = 0 Then Exit Sub
    
    MyText = LCase$(ChatText)
    
    ' Handle when the player presses the return key
    If KeyAscii = vbKeyReturn Then
        If frmMainGame.Visible Then
            ' Broadcast message
            If Left$(ChatText, 1) = "'" Then
                ChatText = Mid$(ChatText, 2, Len(ChatText) - 1)
                If Len(ChatText) > 0 Then
                    Call BroadcastMsg(ChatText)
                End If
                MyText = vbNullString
                frmMainGame.txtMyChat.Text = vbNullString
                Exit Sub
            End If
            
            ' Emote message
            If Left$(ChatText, 1) = "-" Then
                MyText = Mid$(ChatText, 2, Len(ChatText) - 1)
                If Len(ChatText) > 0 Then
                    Call EmoteMsg(ChatText)
                End If
                MyText = vbNullString
                frmMainGame.txtMyChat.Text = vbNullString
                Exit Sub
            End If
            
            ' Player message
            If Left$(ChatText, 1) = "!" Then
                Command = Split(ChatText, " ")
                        
                ' Make sure they are actually sending something
                If UBound(Command) > 0 Then
                    For i = 1 To UBound(Command)
                        Name = Name & " " & Command(i)
                    Next
                    
                    ' Send the message to the player
                    Call PlayerMsg(Right$(Name, Len(Name) - 1), Right$(Command(0), Len(Command(0)) - 1))
                Else
                    Call AddText("Usando: !nomedojogador (mensagem)", AlertColor)
                End If
                MyText = vbNullString
                frmMainGame.txtMyChat.Text = vbNullString
                Exit Sub
            End If
            
            ' Global Message
            If Left$(ChatText, 1) = vbQuote Then
                If GetPlayerAccess(MyIndex) >= ADMIN_MAPPER Then
                    ChatText = Mid$(ChatText, 2, Len(ChatText) - 1)
                    If Len(ChatText) > 0 Then
                        Call GlobalMsg(ChatText)
                    End If
                    MyText = vbNullString
                    frmMainGame.txtMyChat.Text = vbNullString
                    Exit Sub
                End If
            End If
            
            ' Admin Message
            If Left$(ChatText, 1) = "=" Then
                If GetPlayerAccess(MyIndex) >= ADMIN_MAPPER Then
                    ChatText = Mid$(ChatText, 2, Len(ChatText) - 1)
                    If Len(ChatText) > 0 Then
                        Call AdminMsg(MyText)
                    End If
                    MyText = vbNullString
                    frmMainGame.txtMyChat.Text = vbNullString
                    Exit Sub
                End If
            End If
            
            If Left$(MyText, 1) = "/" Then
                Command = Split(MyText, " ")
                
                Select Case Command(0)
                
                    Case "/pular"
                        If Not Player(MyIndex).Dieing Then Do_Jump
                        
                    ' Whos Online
                    Case "/online"
                        SendWhosOnline
                        
                    ' Checking fps
                    Case "/fps"
                        BFPS = Not BFPS
                        
                    Case "/sair"
                        SendData CLeaveRoom & END_CHAR
                        
                    ' // Monitor Admin Commands //
                    ' Admin Help
                    Case "/admin"
                        If GetPlayerAccess(MyIndex) < ADMIN_MONITOR Then
                            AddText "Comando não existe!", AlertColor
                            GoTo Continue
                        End If
                        
                    ' Kicking a player
                    Case "/retirar"
                        If GetPlayerAccess(MyIndex) < ADMIN_MONITOR Then
                            AddText "Comando não existe", AlertColor
                            GoTo Continue
                        End If
                        
                        If UBound(Command) < 1 Then
                            GoTo Continue
                        End If
                        
                        If IsNumeric(Command(1)) Then
                            GoTo Continue
                        End If
                        
                        SendKick Command(1)
                        
                    ' // Mapper Admin Commands //
                    ' Location
                    Case "/local"
                    
                        If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
                            AddText "Comando não existe.", AlertColor
                            GoTo Continue
                        End If
                        
                        BLoc = Not BLoc
                        
                    ' Map Editor
                    Case "/editarmapa"
                    
                        If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
                            AddText "Comando não existe.", AlertColor
                            GoTo Continue
                        End If
                        
                        SendRequestEditMap
                    
                    ' Warping to a player
                    Case "/seguir"
                    
                        If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
                            AddText "Comando não existe.", AlertColor
                            GoTo Continue
                        End If
                        
                        If UBound(Command) < 1 Then
                            GoTo Continue
                        End If
                        
                        If IsNumeric(Command(1)) Then
                            GoTo Continue
                        End If
                        
                        WarpMeTo Command(1)
                                
                    ' Warping a player to you
                    Case "/invocar"
                        If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
                            AddText "Comando não existe.", AlertColor
                            GoTo Continue
                        End If
                        
                        If UBound(Command) < 1 Then
                            GoTo Continue
                        End If
                        
                        If IsNumeric(Command(1)) Then
                            GoTo Continue
                        End If
                        
                        WarpToMe Command(1)
                                
                    ' Warping to a map
                    Case "/teleportar"
                        If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
                            AddText "Comando não existe.", AlertColor
                            GoTo Continue
                        End If
                        
                        If UBound(Command) < 1 Then
                            GoTo Continue
                        End If
                        
                        If Not IsNumeric(Command(1)) Then
                            GoTo Continue
                        End If
                        
                        n = CLng(Command(1))
                    
                        ' Check to make sure its a valid map #
                        If n > 0 And n <= MAX_MAPS Then
                            Call WarpTo(n)
                        Else
                            Call AddText("Numero do mapa invalido.", Red)
                        End If
                    
                    ' Setting sprite
                    Case "/aparencia"
                        If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
                            AddText "Comando não existe.", AlertColor
                            GoTo Continue
                        End If
                        
                        If UBound(Command) < 1 Then
                            GoTo Continue
                        End If
                        
                        If Not IsNumeric(Command(1)) Then
                            GoTo Continue
                        End If
                        
                        SendSetSprite CLng(Command(1))
                    
                    ' Map report
                    Case "/reportar"
                        If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
                            AddText "Comando não existe.", AlertColor
                            GoTo Continue
                        End If
                            
                        SendData CMapReport & END_CHAR
                
                    ' Respawn request
                    Case "/atualizar"
                        If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
                            AddText "Comando não existe.", AlertColor
                            GoTo Continue
                        End If
                        
                        SendMapRespawn
                
                    ' MOTD change
                    Case "/mdd"
                        If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
                            AddText "Comando não existe.", AlertColor
                            GoTo Continue
                        End If
                        
                        If UBound(Command) < 1 Then
                            GoTo Continue
                        End If
    
                        SendMOTDChange Right$(ChatText, Len(ChatText) - 5)
                        
                    ' Check the ban list
                    Case "/banlist"
                        If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
                            AddText "Comando não existe.", AlertColor
                            GoTo Continue
                        End If
                        
                        SendBanList
                        
                    ' Banning a player
                    Case "/ban"
                        If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
                            AddText "Comando não existe.", AlertColor
                            GoTo Continue
                        End If
                        
                        If UBound(Command) < 1 Then
                            GoTo Continue
                        End If
                        
                        SendBan Command(1)
                        
                    ' // Developer Admin Commands //
                    ' Editing item request
                    Case "/item"
                        If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then
                            AddText "Comando não existe.", AlertColor
                            GoTo Continue
                        End If
    
                        SendRequestEditItem
                        
                    ' // Creator Admin Commands //
                    ' Giving another player access
                    Case "/setaraccesso"
                        If GetPlayerAccess(MyIndex) < ADMIN_CREATOR Then
                            AddText "Comando não existe.", AlertColor
                            GoTo Continue
                        End If
                        
                        If UBound(Command) < 2 Then
                            GoTo Continue
                        End If
                        
                        If IsNumeric(Command(1)) = True Or IsNumeric(Command(2)) = False Then
                            GoTo Continue
                        End If
                        
                        SendSetAccess Command(1), CLng(Command(2))
                        
                    ' Ban destroy
                    Case "/destruirbanlist"
                        If GetPlayerAccess(MyIndex) < ADMIN_CREATOR Then
                            AddText "Comando não existe.", AlertColor
                            GoTo Continue
                        End If
                        
                        SendBanDestroy
                        
                    Case "/status"
                        GoTo Continue
                        
                    Case Else
                        AddText "Comando não existe.", HelpColor
                        
                End Select
                
    'continue label where we go instead of exiting the sub
Continue:
                MyText = vbNullString
                frmMainGame.txtMyChat.Text = vbNullString
                frmLobby.txtEnterChat.Text = vbNullString
                Exit Sub
            End If
            
            ' Say message
            If Len(ChatText) > 0 Then
                Call SayMsg(ChatText)
            End If
            
            MyText = vbNullString
            frmMainGame.txtMyChat.Text = vbNullString
            Exit Sub
        ElseIf frmLobby.Visible Then
        
            If Left$(ChatText, 1) = "'" Then
                ChatText = Mid$(ChatText, 2, Len(ChatText) - 1)
                If Len(ChatText) > 0 Then
                    Call BroadcastMsg(ChatText)
                End If
                MyText = vbNullString
                frmLobby.txtEnterChat.Text = vbNullString
                Exit Sub
            End If
        
            ' Player message
            If Left$(ChatText, 1) = "!" Then
                Command = Split(ChatText, " ")
                        
                ' Make sure they are actually sending something
                If UBound(Command) > 0 Then
                    For i = 1 To UBound(Command)
                        Name = Name & " " & Command(i)
                    Next
                    
                    ' Send the message to the player
                    Call PlayerMsg(Right$(Name, Len(Name) - 1), Right$(Command(0), Len(Command(0)) - 1))
                Else
                    Call AddText("Usando: !nomedojogador (mensagem)", AlertColor)
                End If
                MyText = vbNullString
                frmLobby.txtEnterChat.Text = vbNullString
                Exit Sub
            End If
            
            ' Global Message
            If Left$(ChatText, 1) = vbQuote Then
                If GetPlayerAccess(MyIndex) >= ADMIN_MAPPER Then
                    ChatText = Mid$(ChatText, 2, Len(ChatText) - 1)
                    If Len(ChatText) > 0 Then
                        Call GlobalMsg(ChatText)
                    End If
                    MyText = vbNullString
                    frmLobby.txtEnterChat.Text = vbNullString
                    Exit Sub
                End If
            End If
            
            ' Admin Message
            If Left$(ChatText, 1) = "=" Then
                If GetPlayerAccess(MyIndex) >= ADMIN_MAPPER Then
                    ChatText = Mid$(ChatText, 2, Len(ChatText) - 1)
                    If Len(ChatText) > 0 Then
                        Call AdminMsg(ChatText)
                    End If
                    MyText = vbNullString
                    frmLobby.txtEnterChat.Text = vbNullString
                    Exit Sub
                End If
            End If
            
            If Left$(MyText, 1) = "/" Then
                Command = Split(MyText, " ")
                
                Select Case Command(0)
                
                    Case "/status"
                        SendData CRequestStats & END_CHAR
                        GoTo Continue
                        
                    Case "/sala"
                        If UBound(Command) < 1 Or UBound(Command) > 1 Then
                            AddText "Usando: /sala (número)", AlertColor
                            GoTo Continue
                        ElseIf Not IsNumeric(Command(1)) Then
                            AddText "Usando: /sala (número)", AlertColor
                            GoTo Continue
                        End If
                        
                        If GetPlayerAccess(MyIndex) < 1 Then
                            AddText "Comando não existe.", AlertColor
                            GoTo Continue
                        End If
                        
                        SendData CRequestJoinRoom & SEP_CHAR & Val(Command(1)) & END_CHAR
                    
                    Case "/online"
                        SendWhosOnline
                        
                    Case "/retirar"
                        If GetPlayerAccess(MyIndex) < ADMIN_MONITOR Then
                            AddText "Comando não existe.", AlertColor
                            GoTo Continue
                        End If
                        
                        If UBound(Command) < 1 Then
                            GoTo Continue
                        End If
                        
                        If IsNumeric(Command(1)) Then
                            GoTo Continue
                        End If
                        
                        SendKick Command(1)
                        
                    Case "/ban"
                        If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
                            AddText "Comando não existe.", AlertColor
                            GoTo Continue
                        End If
                        
                        If UBound(Command) < 1 Then
                            GoTo Continue
                        End If
                        
                        SendBan Command(1)
                    
                    Case "/acesso"
                        If GetPlayerAccess(MyIndex) < ADMIN_CREATOR Then
                            AddText "Comando não existe.", AlertColor
                            GoTo Continue
                        End If
                        
                        If UBound(Command) < 2 Then
                            GoTo Continue
                        End If
                        
                        If IsNumeric(Command(1)) = True Or IsNumeric(Command(2)) = False Then
                            GoTo Continue
                        End If
                        
                        SendSetAccess Command(1), CLng(Command(2))
                        
                    Case Else
                        AddText "Comando não existe.", HelpColor
                        frmLobby.txtEnterChat.Text = vbNullString
                        Exit Sub
                End Select
            End If
            
            SayMsg ChatText
            frmLobby.txtEnterChat.Text = vbNullString
        End If
    End If
    
    ' Handle when the user presses the backspace key
    If (KeyAscii = vbKeyBack) Then
        If LenB(MyText) > 0 Then MyText = Mid$(MyText, 1, Len(MyText) - 1)
    End If
    
    ' And if neither, then add the character to the user's text buffer
    If (KeyAscii <> vbKeyReturn) Then
        If (KeyAscii <> vbKeyBack) Then
            
            ' Make sure the character is on standard English keyboard
            If KeyAscii >= 32 Then
                If KeyAscii <= 126 Then
                    MyText = MyText & ChrW$(KeyAscii)
                End If
            End If
            
        End If
    End If
    
End Sub

Sub CheckMapGetItem()
Dim i As Long

    For i = 1 To MAX_MAP_ITEMS
        If MapItem(i).x = GetPlayerX(MyIndex) And MapItem(i).y = GetPlayerY(MyIndex) Then
            Call SendData(CMapGetItem & END_CHAR)
        End If
    Next
    
End Sub

Public Sub CheckAttack()

    If Player(MyIndex).Watching Then Exit Sub
    
    If SpaceDown Then
        With Player(MyIndex)
            .Attacking = 1
            .AttackTimer = GetTickCount
        End With
        Call SendData(CAttack & END_CHAR)
    End If
    
End Sub
Public Sub CheckThrow()

    If Player(MyIndex).Watching Then Exit Sub
    
    If ControlDown Then
        With Player(MyIndex)
            .Attacking = 1
            .AttackTimer = GetTickCount
        End With
        Call SendData(CThrow & END_CHAR)
    End If
    
End Sub

Sub CheckInput(ByVal KeyState As Byte, ByVal KeyCode As Integer, ByVal Shift As Integer)
    If Not GettingMap Then
        If KeyState = 1 Then
            Select Case KeyCode
            
                Case vbKeyReturn
                    CheckMapGetItem
                
                Case vbKeySpace
                    SpaceDown = True
                    
                Case vbKeyControl
                    ControlDown = True
                    
                Case vbKeyShift
                    ShiftDown = True
                    
                Case vbKeyUp
                    DirUp = True
                    DirDown = False
                    DirLeft = False
                    DirRight = False
                    
                Case vbKeyDown
                    DirUp = False
                    DirDown = True
                    DirLeft = False
                    DirRight = False
                    
                Case vbKeyLeft
                    DirUp = False
                    DirDown = False
                    DirLeft = True
                    DirRight = False
                    
                Case vbKeyRight
                    DirUp = False
                    DirDown = False
                    DirLeft = False
                    DirRight = True
                    
                Case vbKeyEnd
                    If Player(MyIndex).Moving < 1 Then
                        If Player(MyIndex).AttackTimer + 500 < GetTickCount Then
                            Select Case GetPlayerDir(MyIndex)
                                Case DIR_UP
                                    SetPlayerDir MyIndex, DIR_RIGHT
                                Case DIR_RIGHT
                                    SetPlayerDir MyIndex, DIR_DOWN
                                Case DIR_DOWN
                                    SetPlayerDir MyIndex, DIR_LEFT
                                Case DIR_LEFT
                                    SetPlayerDir MyIndex, DIR_UP
                            End Select
                            SendPlayerDir
                            Player(MyIndex).AttackTimer = GetTickCount
                        End If
                    End If
                    
            End Select
        Else
            Select Case KeyCode
            
                Case vbKeyUp
                    DirUp = False
                    
                Case vbKeyDown
                    DirDown = False
                    
                Case vbKeyLeft
                    DirLeft = False
                    
                Case vbKeyRight
                    DirRight = False
                    
                Case vbKeyShift
                    ShiftDown = False
                    
                Case vbKeyControl
                    ControlDown = False
                    
            End Select
        End If
    End If
End Sub

Function IsTryingToMove() As Boolean
    IsTryingToMove = DirUp Or DirDown Or DirLeft Or DirRight
End Function

Function CanMove() As Boolean
Dim d As Long

    CanMove = True
    
    ' Make sure they aren't trying to move when they are already moving
    If Player(MyIndex).Moving <> 0 Then
        CanMove = False
        Exit Function
    End If
    
    If Player(MyIndex).Watching Then
        CanMove = False
        Exit Function
    End If
    
    ' Make sure they haven't just casted a spell
    If Player(MyIndex).CastedSpell = YES Then
        If GetTickCount > Player(MyIndex).AttackTimer + 1000 Then
            Player(MyIndex).CastedSpell = NO
        Else
            CanMove = False
            Exit Function
        End If
    End If
   
    d = GetPlayerDir(MyIndex)
    If DirUp Then
        Call SetPlayerDir(MyIndex, DIR_UP)
       
        ' Check to see if they are trying to go out of bounds
        If GetPlayerY(MyIndex) > 0 Then
            If CheckDirection(DIR_UP) Then
                CanMove = False
                
                ' Set the new direction if they weren't facing that direction
                If d <> DIR_UP Then
                    Call SendPlayerDir
                End If
                Exit Function
            End If
        Else
            ' Check if they can warp to a new map
            If Map.Up > 0 Then
                Call MapEditorLeaveMap
                Call SendPlayerRequestNewMap
                GettingMap = True
                CanMoveNow = False
            End If
            CanMove = False
            Exit Function
        End If
    End If
           
    If DirDown Then
        Call SetPlayerDir(MyIndex, DIR_DOWN)
       
        ' Check to see if they are trying to go out of bounds
        If GetPlayerY(MyIndex) < MAX_MAPY Then
            If CheckDirection(DIR_DOWN) Then
                CanMove = False
               
                ' Set the new direction if they weren't facing that direction
                If d <> DIR_DOWN Then
                    Call SendPlayerDir
                End If
                Exit Function
            End If
        Else
            ' Check if they can warp to a new map
            If Map.Down > 0 Then
                Call MapEditorLeaveMap
                Call SendPlayerRequestNewMap
                GettingMap = True
                CanMoveNow = False
            End If
            CanMove = False
            Exit Function
        End If
    End If
               
    If DirLeft Then
        Call SetPlayerDir(MyIndex, DIR_LEFT)
       
        ' Check to see if they are trying to go out of bounds
        If GetPlayerX(MyIndex) > 0 Then
            If CheckDirection(DIR_LEFT) Then
                CanMove = False
               
                ' Set the new direction if they weren't facing that direction
                If d <> DIR_LEFT Then
                    Call SendPlayerDir
                End If
                Exit Function
            End If
        Else
            ' Check if they can warp to a new map
            If Map.Left > 0 Then
                Call MapEditorLeaveMap
                Call SendPlayerRequestNewMap
                GettingMap = True
                CanMoveNow = False
            End If
            CanMove = False
            Exit Function
        End If
    End If
       
    If DirRight Then
        Call SetPlayerDir(MyIndex, DIR_RIGHT)
       
        ' Check to see if they are trying to go out of bounds
        If GetPlayerX(MyIndex) < MAX_MAPX Then
            If CheckDirection(DIR_RIGHT) Then
                CanMove = False
                ' Set the new direction if they weren't facing that direction
                If d <> DIR_RIGHT Then
                    Call SendPlayerDir
                End If
                Exit Function
            End If
        Else
            ' Check if they can warp to a new map
            If Map.Right > 0 Then
                Call MapEditorLeaveMap
                Call SendPlayerRequestNewMap
                GettingMap = True
                CanMoveNow = False
            End If
            CanMove = False
            Exit Function
        End If
    End If
End Function

Function CheckDirection(ByVal Direction As Byte) As Boolean
Dim x As Long, y As Long, i As Long

    CheckDirection = False
    
    Select Case Direction
        Case DIR_UP
            x = GetPlayerX(MyIndex)
            y = GetPlayerY(MyIndex) - 1
        Case DIR_DOWN
            x = GetPlayerX(MyIndex)
            y = GetPlayerY(MyIndex) + 1
        Case DIR_LEFT
            x = GetPlayerX(MyIndex) - 1
            y = GetPlayerY(MyIndex)
        Case DIR_RIGHT
            x = GetPlayerX(MyIndex) + 1
            y = GetPlayerY(MyIndex)
    End Select
    
    ' Check to see if the map tile is blocked or not
    If Map.Tile(x, y).Type = TILE_TYPE_BLOCKED Or Map.Tile(x, y).Type = TILE_TYPE_BOMB Then
        CheckDirection = True
        Exit Function
    End If
    
    If Map.Tile(x, y).Type = TILE_TYPE_WALL Then
        If Not Wall(x, y).Here Then
            CheckDirection = True
            Exit Function
        End If
    End If
    
    ' Check to see if the key door is open or not
    If Map.Tile(x, y).Type = TILE_TYPE_KEY Then
        ' This actually checks if its open or not
        If TempTile(x, y).DoorOpen = NO Then
            CheckDirection = True
            Exit Function
        End If
    End If
    
    If Player(MyIndex).Dieing Then
        CheckDirection = True
        Exit Function
    End If
    
    If Player(MyIndex).Jumping Then
        CheckDirection = True
        Exit Function
    End If
    
End Function

Sub CheckMovement()
    If IsTryingToMove Then
        If CanMove Then
            
            Player(MyIndex).Moving = MOVING_WALKING
            
            Select Case GetPlayerDir(MyIndex)
                Case DIR_UP
                    Call SendPlayerMove
                    Player(MyIndex).YOffset = PIC_Y
                    Call SetPlayerY(MyIndex, GetPlayerY(MyIndex) - 1)
            
                Case DIR_DOWN
                    Call SendPlayerMove
                    Player(MyIndex).YOffset = PIC_Y * -1
                    Call SetPlayerY(MyIndex, GetPlayerY(MyIndex) + 1)
            
                Case DIR_LEFT
                    Call SendPlayerMove
                    Player(MyIndex).XOffset = PIC_X
                    Call SetPlayerX(MyIndex, GetPlayerX(MyIndex) - 1)
            
                Case DIR_RIGHT
                    Call SendPlayerMove
                    Player(MyIndex).XOffset = PIC_X * -1
                    Call SetPlayerX(MyIndex, GetPlayerX(MyIndex) + 1)
            End Select
            
            If tmrFireDeath < GetTickCount Then
                If Map.Fire(GetPlayerX(MyIndex), GetPlayerY(MyIndex)).Here Then
                    SendData CFireDeath & SEP_CHAR & Map.Fire(GetPlayerX(MyIndex), GetPlayerY(MyIndex)).Owner & END_CHAR
                    tmrFireDeath = GetTickCount + 500
                End If
            End If
            
            If Player(MyIndex).XOffset = 0 Then
                If Player(MyIndex).YOffset = 0 Then
                    If Map.Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex)).Type = TILE_TYPE_WARP Then
                        GettingMap = True
                    End If
                End If
            End If
            
        End If
    End If
End Sub

Public Sub UpdateInventory()

End Sub

Sub PlayerSearch()
    If isInBounds Then Call SendData(CSearch & SEP_CHAR & CurX & SEP_CHAR & CurY & END_CHAR)
End Sub

Public Function isInBounds()
    If (CurX >= 0) Then
        If (CurX <= MAX_MAPX) Then
            If (CurY >= 0) Then
                If (CurY <= MAX_MAPY) Then
                    isInBounds = True
                End If
            End If
        End If
    End If
End Function

Public Sub UpdateDrawMapName()
    
End Sub

Private Sub CheckKeys()
    If GetAsyncKeyState(VK_UP) >= 0 Then DirUp = False
    If GetAsyncKeyState(VK_DOWN) >= 0 Then DirDown = False
    If GetAsyncKeyState(VK_LEFT) >= 0 Then DirLeft = False
    If GetAsyncKeyState(VK_RIGHT) >= 0 Then DirRight = False
    If GetAsyncKeyState(VK_CONTROL) >= 0 Then ControlDown = False
    If GetAsyncKeyState(VK_SPACE) >= 0 Then SpaceDown = False
    If GetAsyncKeyState(VK_SHIFT) >= 0 Then ShiftDown = False
End Sub

Public Sub Do_Jump()
    SendData CJump & END_CHAR
End Sub
