Attribute VB_Name = "modDatabase"
Option Explicit

' ******************************************
' **              BMO Source              **
' ******************************************

' Outputs string to text file
Sub AddLog(ByVal Text As String, ByVal FN As String)
Dim FileName As String
Dim f As Integer

    If ServerLog = True Then
        FileName = App.Path & "\logs\" & FN
    
        If Not FileExist(FileName, True) Then
            f = FreeFile
            Open FileName For Output As #f
            Close #f
        End If
    
        f = FreeFile
        Open FileName For Append As #f
            Print #f, Time & ": " & Text
        Close #f
    End If
End Sub

' gets a string from a text file
Public Function GetVar(File As String, Header As String, Var As String) As String
Dim sSpaces As String   ' Max string length
Dim szReturn As String  ' Return default value if not found
  
    szReturn = vbNullString
  
    sSpaces = Space$(5000)
  
    Call GetPrivateProfileString$(Header, Var, szReturn, sSpaces, Len(sSpaces), File)
  
    GetVar = RTrim$(sSpaces)
    GetVar = Left$(GetVar, Len(GetVar) - 1)
End Function

' writes a variable to a text file
Public Sub PutVar(File As String, Header As String, Var As String, Value As String)
    Call WritePrivateProfileString$(Header, Var, Value, File)
End Sub

Public Function FileExist(ByVal FileName As String, Optional RAW As Boolean = False) As Boolean
    
    If Not RAW Then
        If Dir(App.Path & "\" & FileName) = vbNullString Then
            FileExist = False
            Exit Function
        Else
            FileExist = True
            Exit Function
        End If
    Else
        If Dir(FileName) = vbNullString Then
            FileExist = False
            Exit Function
        Else
            FileExist = True
        End If
    End If
End Function

Sub BanIndex(ByVal BanPlayerIndex As Long, ByVal BannedByIndex As Long)
Dim FileName As String
Dim IP As String
Dim f As Long
Dim i As Long

    FileName = App.Path & "\data\banlist.txt"
    
    ' Make sure the file exists
    If Not FileExist("data\banlist.txt") Then
        f = FreeFile
        Open FileName For Output As #f
        Close #f
    End If
    
    ' Cut off last portion of ip
    IP = GetPlayerIP(BanPlayerIndex)
            
    For i = Len(IP) To 1 Step -1
        If Mid$(IP, i, 1) = "." Then
            Exit For
        End If
    Next
    IP = Mid$(IP, 1, i)
            
    f = FreeFile
    Open FileName For Append As #f
        Print #f, IP & "," & GetPlayerName(BannedByIndex)
    Close #f
    
    Call GlobalMsg(GetPlayerName(BanPlayerIndex) & " foi banido do " & GAME_NAME & " por " & GetPlayerName(BannedByIndex) & "!", Red)
    Call AddLog(GetPlayerName(BannedByIndex) & " baniu " & GetPlayerName(BanPlayerIndex) & ".", ADMIN_LOG)
    Call AlertMsg(BanPlayerIndex, "Você foi banido por " & GetPlayerName(BannedByIndex) & "!")
End Sub

Sub ServerBanIndex(ByVal BanPlayerIndex As Long)
Dim FileName As String, IP As String
Dim f As Long, i As Long

    FileName = App.Path & "data\banlist.txt"
    
    ' Make sure the file exists
    If Not FileExist("data\banlist.txt") Then
        f = FreeFile
        Open FileName For Output As #f
        Close #f
    End If
    
    ' Cut off last portion of ip
    IP = GetPlayerIP(BanPlayerIndex)
            
    For i = Len(IP) To 1 Step -1
        If Mid$(IP, i, 1) = "." Then
            Exit For
        End If
    Next
    IP = Mid$(IP, 1, i)
            
    f = FreeFile
    Open FileName For Append As #f
        Print #f, IP & "," & "Server"
    Close #f
    
    Call GlobalMsg(GetPlayerName(BanPlayerIndex) & " foi banido do " & GAME_NAME & " pelo " & "servidor" & "!", Red)
    Call AddLog("O servidor" & " baniu " & GetPlayerName(BanPlayerIndex) & ".", ADMIN_LOG)
    Call AlertMsg(BanPlayerIndex, "Você foi banido pelo " & "servidor" & "!")
End Sub

' **************
' ** Accounts **
' **************
Function AccountExist(ByVal Name As String) As Boolean
Dim FileName As String

    FileName = "accounts\" & Trim(Name) & ".bin"
   
    If FileExist(FileName) Then
        AccountExist = True
    Else
        AccountExist = False
    End If
End Function

Function PasswordOK(ByVal Name As String, ByVal Password As String) As Boolean
Dim FileName As String
Dim RightPassword As String * NAME_LENGTH
Dim nFileNum As Integer

    PasswordOK = False
   
    If AccountExist(Name) Then
        FileName = App.Path & "\Accounts\" & Trim$(Name) & ".bin"
        nFileNum = FreeFile
        Open FileName For Binary As #nFileNum
       
        Get #nFileNum, NAME_LENGTH, RightPassword
       
        Close #nFileNum
       
        If UCase$(Trim$(Password)) = UCase$(Trim$(RightPassword)) Then
            PasswordOK = True
        End If
    End If
    
End Function
    
Sub AddAccount(ByVal Index As Long, ByVal Name As String, ByVal Password As String)
Dim i As Long

    Player(Index).Login = Name
    Player(Index).Password = Password
    
    For i = 1 To MAX_CHARS
        Call ClearChar(Index, i)
    Next
    
    Call SavePlayer(Index)
End Sub

Sub DeleteName(ByVal Name As String)
Dim f1 As Long, f2 As Long
Dim s As String

    Call FileCopy(App.Path & "\accounts\charlist.txt", App.Path & "\accounts\chartemp.txt")
    
    ' Destroy name from charlist
    f1 = FreeFile
    Open App.Path & "\accounts\chartemp.txt" For Input As #f1
    
        f2 = FreeFile
        Open App.Path & "\accounts\charlist.txt" For Output As #f2
            
        Do While Not EOF(f1)
            Input #f1, s
            If Trim$(LCase$(s)) <> Trim$(LCase$(Name)) Then
                Print #f2, s
            End If
        Loop
        
        Close #f1
        
    Close #f2
    
    Call Kill(App.Path & "\accounts\chartemp.txt")
End Sub

' ****************
' ** Characters **
' ****************

Function CharExist(ByVal Index As Long, ByVal CharNum As Long) As Boolean
    If LenB(Trim$(Player(Index).Char(CharNum).Name)) > 0 Then
        CharExist = True
    Else
        CharExist = False
    End If
End Function

Sub AddChar(ByVal Index As Long, ByVal Name As String, ByVal Sex As Byte, ByVal ClassNum As Byte, ByVal CharNum As Long)
Dim f As Long
Dim n As Long

    If LenB(Trim$(Player(Index).Char(CharNum).Name)) = 0 Then
        TempPlayer(Index).CharNum = CharNum
        
        If ClassNum = 0 Then ClassNum = 1
        
        Player(Index).Char(CharNum).Name = Name
        Player(Index).Char(CharNum).Sex = Sex
        Player(Index).Char(CharNum).Class = 1
        
        Player(Index).Char(CharNum).Sprite = ClassNum
        
        Player(Index).Char(CharNum).Level = 1
        
        For n = 1 To Stats.Stat_Count - 1
            Player(Index).Char(CharNum).Stat(n) = Class(1).Stat(n)
        Next
        
        Player(Index).Char(CharNum).Map = START_MAP
        Player(Index).Char(CharNum).x = START_X
        Player(Index).Char(CharNum).y = START_Y
        
        Player(Index).Char(CharNum).Vital(Vitals.HP) = GetPlayerMaxVital(Index, Vitals.HP)
        Player(Index).Char(CharNum).Vital(Vitals.MP) = GetPlayerMaxVital(Index, Vitals.MP)
        Player(Index).Char(CharNum).Vital(Vitals.SP) = GetPlayerMaxVital(Index, Vitals.SP)
        
        ' Append name to file
        f = FreeFile
        Open App.Path & "\accounts\charlist.txt" For Append As #f
            Print #f, Name
        Close #f
        
        Call SavePlayer(Index)
        
    End If
End Sub

Sub DelChar(ByVal Index As Long, ByVal CharNum As Long)
    Call DeleteName(Player(Index).Char(CharNum).Name)
    Call ClearChar(Index, CharNum)
    Call SavePlayer(Index)
End Sub

Function FindChar(ByVal Name As String) As Boolean
Dim f As Long
Dim s As String

    FindChar = False
    
    f = FreeFile
    Open App.Path & "\Accounts\charlist.txt" For Input As #f
        Do While Not EOF(f)
            Input #f, s
            
            If Trim$(LCase$(s)) = Trim$(LCase$(Name)) Then
                FindChar = True
                Close #f
                Exit Function
            End If
        Loop
    Close #f
End Function

' *************
' ** Players **
' *************

Sub SaveAllPlayersOnline()
Dim i As Long

    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            Call SavePlayer(i)
        End If
    Next
End Sub

Sub SavePlayer(ByVal Index As Long)
Dim FileName As String
Dim f As Long

    FileName = App.Path & "\accounts\" & Trim$(Player(Index).Login) & ".bin"
       
    f = FreeFile
    Open FileName For Binary As #f
        Put #f, , Player(Index)
    Close #f
End Sub

Sub LoadPlayer(ByVal Index As Long, ByVal Name As String)
Dim FileName As String
Dim f As Long

    Call ClearPlayer(Index)
   
    FileName = App.Path & "\accounts\" & Trim(Name) & ".bin"

    f = FreeFile
    Open FileName For Binary As #f
        Get #f, , Player(Index)
    Close #f
End Sub

Sub ClearPlayer(ByVal Index As Long)
Dim i As Byte

    Call ZeroMemory(ByVal VarPtr(Player(Index)), LenB(Player(Index)))

    Player(Index).Login = vbNullString
    Player(Index).Password = vbNullString
    TempPlayer(Index).Buffer = vbNullString
    TempPlayer(Index).Distance = 1 + Player(Index).Char(TempPlayer(Index).CharNum).BonusDistance
    TempPlayer(Index).Bombs = 1 + Player(Index).Char(TempPlayer(Index).CharNum).BonusLimit
    TempPlayer(Index).Throw = False
    TempPlayer(Index).Speed = False
    
    For i = 0 To MAX_CHARS
        Player(Index).Char(i).Name = vbNullString
        Player(Index).Char(i).Class = 1
    Next
    
    frmServer.lvwInfo.ListItems(Index).SubItems(1) = vbNullString
    frmServer.lvwInfo.ListItems(Index).SubItems(2) = vbNullString
    frmServer.lvwInfo.ListItems(Index).SubItems(3) = vbNullString
End Sub

Sub ClearChar(ByVal Index As Long, ByVal CharNum As Long)
    Call ZeroMemory(ByVal VarPtr(Player(Index).Char(CharNum)), LenB(Player(Index).Char(CharNum)))
    Player(Index).Char(CharNum).Name = vbNullString
    Player(Index).Char(CharNum).Class = 1
End Sub

Sub LoadHighScores()
Dim FileName As String
Dim f As Long, i As Long

    'ClearHighScores
    
    For i = 1 To 100
        FileName = App.Path & "\Data\Highscores\Highscores" & i & ".bin"
    
        f = FreeFile
        Open FileName For Binary As #f
            Get #f, , HighScore(i)
        Close #f
    Next
    
End Sub

Sub ClearHighScores()
Dim i As Byte

    For i = 1 To 100
        Call ZeroMemory(ByVal VarPtr(HighScore(i)), LenB(HighScore(i)))
        HighScore(i).Name = vbNullString
    Next
    
End Sub

Sub SaveHighScores()
Dim FileName As String
Dim f As Long, i As Long

    For i = 1 To 100
        FileName = App.Path & "\Data\Highscores\Highscores" & i & ".bin"
           
        f = FreeFile
        Open FileName For Binary As #f
            Put #f, , HighScore(i)
        Close #f
    Next
    
End Sub

' *************
' ** Classes **
' *************

Public Sub CreateClassesINI()
Dim FileName As String
Dim File As String

    FileName = App.Path & "\data\classes.ini"
    
    Max_Classes = 2
    
    If Not FileExist(FileName, True) Then
        File = FreeFile
    
        Open FileName For Output As File
            Print #File, "[INIT]"
            Print #File, "MaxClasses=" & Max_Classes
        Close File
    End If

End Sub

Sub LoadClasses()
Dim FileName As String
Dim i As Long

    If CheckClasses Then
        ReDim Class(1 To Max_Classes) As ClassRec
        Call SaveClasses
    Else
        FileName = App.Path & "\data\classes.ini"
        Max_Classes = Val(GetVar(FileName, "INIT", "MaxClasses"))
        ReDim Class(1 To Max_Classes) As ClassRec
        
    End If
    
    Call ClearClasses
    
    For i = 1 To Max_Classes
        Class(i).Name = GetVar(FileName, "CLASS" & i, "Name")
        Class(i).Sprite = Val(GetVar(FileName, "CLASS" & i, "Sprite"))
        Class(i).Stat(Stats.Strength) = Val(GetVar(FileName, "CLASS" & i, "STR"))
        Class(i).Stat(Stats.Defense) = Val(GetVar(FileName, "CLASS" & i, "DEF"))
        Class(i).Stat(Stats.Speed) = Val(GetVar(FileName, "CLASS" & i, "Speed"))
        Class(i).Stat(Stats.Magic) = Val(GetVar(FileName, "CLASS" & i, "MAGI"))
    Next
End Sub

Sub SaveClasses()
Dim FileName As String
Dim i As Long

    FileName = App.Path & "\data\classes.ini"
    
    For i = 1 To Max_Classes
        Call PutVar(FileName, "CLASS" & i, "Name", Trim$(Class(i).Name))
        Call PutVar(FileName, "CLASS" & i, "Sprite", CStr(Class(i).Sprite))
        Call PutVar(FileName, "CLASS" & i, "STR", CStr(Class(i).Stat(Stats.Strength)))
        Call PutVar(FileName, "CLASS" & i, "DEF", CStr(Class(i).Stat(Stats.Defense)))
        Call PutVar(FileName, "CLASS" & i, "Speed", CStr(Class(i).Stat(Stats.Speed)))
        Call PutVar(FileName, "CLASS" & i, "MAGI", CStr(Class(i).Stat(Stats.Magic)))
    Next
End Sub

Function CheckClasses() As Boolean
Dim FileName As String

    FileName = App.Path & "\data\classes.ini"

    CheckClasses = False

    If Not FileExist(FileName, True) Then
        Call CreateClassesINI
        CheckClasses = True
    End If

End Function

Sub ClearClasses()
Dim i As Long

    For i = 1 To Max_Classes
        Call ZeroMemory(ByVal VarPtr(Class(i)), LenB(Class(i)))
        Class(i).Name = vbNullString
    Next
End Sub

' ***********
' ** Items **
' ***********

Sub SaveItems()
Dim i As Long
    
    For i = 1 To MAX_ITEMS
        Call SaveItem(i)
    Next
End Sub

Sub SaveItem(ByVal ItemNum As Long)
Dim FileName As String
Dim f  As Long
    
    FileName = App.Path & "\Data\items\item" & ItemNum & ".dat"
    f = FreeFile
    
    Open FileName For Binary As #f
        Put #f, , Item(ItemNum)
    Close #f
End Sub

Sub LoadItems()
Dim FileName As String
Dim i As Long
Dim f As Long

    Call CheckItems

    For i = 1 To MAX_ITEMS
        FileName = App.Path & "\Data\Items\Item" & i & ".dat"
        f = FreeFile
        
        Open FileName For Binary As #f
            Get #f, , Item(i)
        Close #f

    Next
End Sub

Sub CheckItems()
Dim i As Long

    For i = 1 To MAX_ITEMS
        If Not FileExist("\Data\Items\Item" & i & ".dat") Then
            Call SaveItem(i)
        End If
    Next
    
End Sub

Sub ClearItem(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(Item(Index)), LenB(Item(Index)))
    Item(Index).Name = vbNullString
End Sub

Sub ClearItems()
Dim i As Long

    For i = 1 To MAX_ITEMS
        Call ClearItem(i)
    Next
End Sub

' ***********
' ** Shops **
' ***********

Sub SaveShops()
Dim i As Long

    For i = 1 To MAX_SHOPS
        Call SaveShop(i)
    Next
End Sub

Sub SaveShop(ByVal ShopNum As Long)
    Dim FileName As String
    Dim f As Long

    FileName = App.Path & "\Data\shops\shop" & ShopNum & ".dat"

    f = FreeFile
    Open FileName For Binary As #f
        Put #f, , Shop(ShopNum)
    Close #f
End Sub

Sub LoadShops()
    Dim FileName As String
    Dim i As Long, f As Long

    Call CheckShops

    For i = 1 To MAX_SHOPS
        FileName = App.Path & "\Data\shops\shop" & i & ".dat"
        f = FreeFile
        
        Open FileName For Binary As #f
            Get #f, , Shop(i)
        Close #f

    Next
End Sub

Sub CheckShops()
Dim i As Long

    For i = 1 To MAX_SHOPS
        If Not FileExist("\Data\shops\shop" & i & ".dat") Then
            Call SaveShop(i)
        End If
    Next
End Sub

Sub ClearShop(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(Shop(Index)), LenB(Shop(Index)))
    Shop(Index).Name = vbNullString
    Shop(Index).JoinSay = vbNullString
    Shop(Index).LeaveSay = vbNullString
End Sub

Sub ClearShops()
Dim i As Long

    For i = 1 To MAX_SHOPS
        Call ClearShop(i)
    Next
End Sub

' ************
' ** Spells **
' ************

Sub SaveSpell(ByVal SpellNum As Long)
Dim FileName As String
Dim f As Long

    FileName = App.Path & "\Data\spells\spells" & SpellNum & ".dat"
    f = FreeFile
    
    Open FileName For Binary As #f
       Put #f, , Spell(SpellNum)
    Close #f
End Sub

Sub SaveSpells()
    Dim i As Long

    Call SetStatus("Salvando magias... ")
    For i = 1 To MAX_SPELLS
        Call SaveSpell(i)
    Next
End Sub

Sub LoadSpells()
    Dim FileName As String
    Dim i As Long
    Dim f As Long

    Call CheckSpells

    For i = 1 To MAX_SPELLS
        FileName = App.Path & "\Data\spells\spells" & i & ".dat"
        f = FreeFile
        
        Open FileName For Binary As #f
            Get #f, , Spell(i)
        Close #f

    Next
End Sub

Sub CheckSpells()
Dim i As Long

    For i = 1 To MAX_SPELLS
        If Not FileExist("\Data\spells\spell" & i & ".dat") Then
            Call SaveSpell(i)
        End If
    Next
End Sub

Sub ClearSpell(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(Spell(Index)), LenB(Spell(Index)))
    Spell(Index).Name = vbNullString
    Spell(Index).LevelReq = 1 'Needs to be 1 for the spell editor
End Sub

Sub ClearSpells()
Dim i As Long

    For i = 1 To MAX_SPELLS
        Call ClearSpell(i)
    Next
End Sub

' **********
' ** NPCs **
' **********

Sub SaveNpcs()
Dim i As Long
    
    For i = 1 To MAX_NPCS
        Call SaveNpc(i)
    Next
End Sub

Sub SaveNpc(ByVal NpcNum As Long)
Dim FileName As String
Dim f As Long

    FileName = App.Path & "\Data\npcs\npc" & NpcNum & ".dat"
    f = FreeFile
    
    Open FileName For Binary As #f
        Put #f, , Npc(NpcNum)
    Close #f
End Sub

Sub LoadNpcs()
Dim FileName As String
Dim i As Integer
Dim f As Long

    Call CheckNpcs

    For i = 1 To MAX_NPCS
        FileName = App.Path & "\Data\npcs\npc" & i & ".dat"
        f = FreeFile
        
        Open FileName For Binary As #f
            Get #f, , Npc(i)
        Close #f

    Next
End Sub

Sub CheckNpcs()
Dim i As Long

    For i = 1 To MAX_NPCS
        If Not FileExist("\Data\npcs\npc" & i & ".dat") Then
            Call SaveNpc(i)
        End If
    Next
End Sub

Sub ClearNpc(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(Npc(Index)), LenB(Npc(Index)))
    Npc(Index).Name = vbNullString
    Npc(Index).AttackSay = vbNullString
End Sub

Sub ClearNpcs()
Dim i As Long

    For i = 1 To MAX_NPCS
        Call ClearNpc(i)
    Next
End Sub

' **********
' ** Maps **
' **********

Sub SaveMap(ByVal MapNum As Long)
Dim FileName As String
Dim f As Long

    FileName = App.Path & "\Data\maps\map" & MapNum & ".dat"
        
    f = FreeFile
    Open FileName For Binary As #f
        Put #f, , Map(MapNum)
    Close #f
End Sub

Sub SaveMaps()
Dim i As Long

    For i = 1 To MAX_MAPS
        Call SaveMap(i)
    Next
End Sub

Sub LoadMaps()
Dim FileName As String
Dim i As Long
Dim f As Long

    Call CheckMaps
    
    For i = 1 To MAX_MAPS
        FileName = App.Path & "\Data\maps\map" & i & ".dat"
        
        f = FreeFile
        Open FileName For Binary As #f
            Get #f, , Map(i)
        Close #f
        If Map(i).PlayerLimit < 2 Then Map(i).PlayerLimit = 2
        If Trim$(Map(i).Name) = vbNullString Then Map(i).Moral = MAP_MORAL_FFA
    Next
End Sub
Sub CheckMaps()
Dim i As Long
        
    For i = 1 To MAX_MAPS
        
        If Not FileExist("\Data\maps\map" & i & ".dat") Then
            Call SaveMap(i)
        End If
    Next
End Sub

Sub ClearTempTile()
Dim i As Long, y As Long, x As Long

    For i = 1 To MAX_MAPS
        TempTile(i).DoorTimer = 0
        
        For x = 0 To MAX_MAPX
            For y = 0 To MAX_MAPY
                TempTile(i).DoorOpen(x, y) = NO
            Next
        Next
    Next
End Sub

Sub ClearMapItem(ByVal Index As Long, ByVal MapNum As Long)
    Call ZeroMemory(ByVal VarPtr(MapItem(MapNum, Index)), LenB(MapItem(MapNum, Index)))
End Sub

Sub ClearMapItems()
Dim x As Long
Dim y As Long

    For y = 1 To MAX_MAPS
        For x = 1 To MAX_MAP_ITEMS
            Call ClearMapItem(x, y)
        Next
    Next
End Sub

Sub ClearMapNpc(ByVal Index As Long, ByVal MapNum As Long)
    Call ZeroMemory(ByVal VarPtr(MapNpc(MapNum, Index)), LenB(MapNpc(MapNum, Index)))
End Sub

Sub ClearMapNpcs()
Dim x As Long
Dim y As Long

    For y = 1 To MAX_MAPS
        For x = 1 To MAX_MAP_NPCS
            Call ClearMapNpc(x, y)
        Next
    Next
End Sub

Sub ClearMap(ByVal MapNum As Long)
    Call ZeroMemory(ByVal VarPtr(Map(MapNum)), LenB(Map(MapNum)))
    
    Map(MapNum).TileSet = 0
    Map(MapNum).PlayerLimit = 4
    Map(MapNum).Moral = MAP_MORAL_FFA
    
    Map(MapNum).Name = vbNullString
    
    ' Reset the values for if a player is on the map or not
    PlayersOnMap(MapNum) = NO
    
    ' Reset the map cache array for this map.
    MapCache(MapNum) = vbNullString
    
End Sub

Sub ClearMaps()
Dim i As Long

    For i = 1 To MAX_MAPS
        Call ClearMap(i)
    Next
End Sub

Function GetClassName(ByVal ClassNum As Long) As String
    GetClassName = Trim$(Class(ClassNum).Name)
End Function

Function GetClassMaxVital(ByVal ClassNum As Long, ByVal Vital As Vitals) As Long
    Select Case Vital
        Case HP
            GetClassMaxVital = (1 + Int(Class(ClassNum).Stat(Stats.Strength) / 2) + Class(ClassNum).Stat(Stats.Strength)) * 2
        Case MP
            GetClassMaxVital = (1 + Int(Class(ClassNum).Stat(Stats.Magic) / 2) + Class(ClassNum).Stat(Stats.Magic)) * 2
        Case SP
            GetClassMaxVital = (1 + Int(Class(ClassNum).Stat(Stats.Speed) / 2) + Class(ClassNum).Stat(Stats.Speed)) * 2
    End Select
End Function

Function GetClassStat(ByVal ClassNum As Long, ByVal Stat As Stats) As Long
    GetClassStat = Class(ClassNum).Stat(Stat)
End Function
