Attribute VB_Name = "modGeneral"
Option Explicit

' ******************************************
' **              BMO Source              **
' ******************************************

Sub Main()
    Call InitServer
End Sub

Sub InitServer()
Dim i As Long
Dim f As Long

Dim time1 As Long
Dim time2 As Long

    time1 = GetTickCount
    
    frmServer.Show
    
    ' Initialize the random-number generator
    Randomize ', seed
    
    ' Check if the directory is there, if its not make it
    
    If LCase$(Dir(App.Path & "\Data\items", vbDirectory)) <> "items" Then
        Call MkDir(App.Path & "\Data\items")
    End If
    
    If LCase$(Dir(App.Path & "\Data\maps", vbDirectory)) <> "maps" Then
        Call MkDir(App.Path & "\Data\maps")
    End If
    
    If LCase$(Dir(App.Path & "\accounts", vbDirectory)) <> "accounts" Then
        Call MkDir(App.Path & "\accounts")
    End If
    
    ' used for parsing packets
    SEP_CHAR = vbNullChar ' ChrW$(0)
    END_CHAR = ChrW$(237)
    
    ' set quote character
    vbQuote = ChrW$(34) ' "
    
    ' set MOTD
    MOTD = GetVar(App.Path & "\data\motd.ini", "MOTD", "Msg")
    
    GAME_PORT = Val(GetVar(App.Path & "\config.ini", "IPCONFIG", "Port"))
    
    ' Get the listening socket ready to go
    frmServer.Socket(0).RemoteHost = frmServer.Socket(0).LocalIP
    frmServer.Socket(0).LocalPort = GAME_PORT
    
    ' Init all the player sockets
    Call SetStatus("Inicializando informações dos jogadores...")
    For i = 1 To MAX_PLAYERS
        Call ClearPlayer(i)
        Load frmServer.Socket(i)
    Next
    
    ' Serves as a constructor
    Call ClearGameData
    
    Call LoadGameData
    
    Call SetStatus("Liberando itens...")
    Call SpawnAllMapsItems
    
    Call SetStatus("Criando cachê dos mapas...")
    Call CreateFullMapCache
    
    Call SetStatus("Carregando proteção...")
    Call LoadSystemTray
    
    ' Check if the master charlist file exists for checking duplicate names, and if it doesnt make it
    If Not FileExist("accounts\charlist.txt") Then
        f = FreeFile
        Open App.Path & "\accounts\charlist.txt" For Output As #f
        Close #f
    End If
    
    ' Start listening
    frmServer.Socket(0).Listen
    
    UpdateCaption
    
    ' To allow first connection
    Player_HighIndex = 1
    
    time2 = GetTickCount
    
    Call SetStatus("Inicialização completa,servidor iniciado em " & time2 - time1 & "ms.")
    
    ' Starts the server loop
    ServerLoop
    
End Sub

Sub DestroyServer()
Dim i As Long
 
    ServerOnline = 0
    
    SaveHighScores
    
    Call SetStatus("Desligando proteção...")
    Call DestroySystemTray
    
    Call SetStatus("Salvando jogadores online...")
    Call SaveAllPlayersOnline
    
    Call ClearGameData
    
    Call SetStatus("Descarregando sockets...")
    
    For i = 1 To MAX_PLAYERS
        Unload frmServer.Socket(i)
    Next
    
    End
    
End Sub

Sub CreateFullMapCache()
Dim i As Long

    For i = 1 To MAX_MAPS
        Call MapCache_Create(i)
    Next
    
End Sub

Sub SetStatus(ByVal Status As String)
    Call TextAdd(frmServer.txtText, Status)
End Sub

Public Sub ClearGameData()
    Call SetStatus("Limpando tiles...")
    Call ClearTempTile
    Call SetStatus("Limpando mapas...")
    Call ClearMaps
    Call SetStatus("Limpando itens dropados...")
    Call ClearMapItems
    Call SetStatus("Limpando itens...")
    Call ClearItems
End Sub

Public Sub LoadGameData()
    Call SetStatus("Carregando dados dos jogadores ingame...")
    Call LoadHighScores
    Call SetStatus("Carregando classes...")
    Call LoadClasses
    Call SetStatus("Carregando mapas...")
    Call LoadMaps
    Call SetStatus("Carregando itens...")
    Call LoadItems
End Sub

Sub DestroySystemTray()
    nid.cbSize = Len(nid)
    nid.hWnd = frmServer.hWnd
    nid.uId = vbNull
    nid.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    nid.uCallBackMessage = WM_MOUSEMOVE
    nid.hIcon = frmServer.Icon
    nid.szTip = "Server" & vbNullChar
    Call Shell_NotifyIcon(NIM_DELETE, nid) ' Add to the sys tray
End Sub

Sub LoadSystemTray()
    nid.cbSize = Len(nid)
    nid.hWnd = frmServer.hWnd
    nid.uId = vbNull
    nid.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    nid.uCallBackMessage = WM_MOUSEMOVE
    nid.hIcon = frmServer.Icon
    nid.szTip = "Server" & vbNullChar   'You can add your game name or something.
    Call Shell_NotifyIcon(NIM_ADD, nid) 'Add to the sys tray
End Sub
