Attribute VB_Name = "modTypes"
Option Explicit

' ******************************************
' **              BMO Source              **
' ******************************************

' Public data structures
Public Map(1 To MAX_MAPS) As MapRec
Public TempMap(1 To MAX_MAPS) As TempMapRec
Public MapCache(1 To MAX_MAPS) As String
Public TempTile(1 To MAX_MAPS) As TempTileRec
Public PlayersOnMap(1 To MAX_MAPS) As Long
Public Player(1 To MAX_PLAYERS) As AccountRec
Public TempPlayer(1 To MAX_PLAYERS) As TempPlayerRec
Public Class() As ClassRec
Public Item(1 To MAX_ITEMS) As ItemRec
Public Npc(1 To MAX_NPCS) As NpcRec
Public MapItem(1 To MAX_MAPS, 1 To MAX_MAP_ITEMS) As MapItemRec
Public MapNpc(1 To MAX_MAPS, 1 To MAX_MAP_NPCS) As MapNpcRec
Public Shop(1 To MAX_SHOPS) As ShopRec
Public Spell(1 To MAX_SPELLS) As SpellRec
Public Bomb(0 To MAX_MAPX, 0 To MAX_MAPY) As BombRec
Public Wall(0 To MAX_MAPX, 0 To MAX_MAPY) As WallRec
Public HighScore(1 To 100) As HighScoreRec

Type HighScoreRec
    Name As String * NAME_LENGTH
    Kills As Long
    Deaths As Long
    Matchs As Long
End Type

Type PlayerInvRec
    Num As Byte
    Value As Long
    Dur As Integer
End Type

Type WallRec
    Gone As Boolean
End Type

Type TempMapRec
    InProgress As Boolean
    Timer As Long
    JumpTimes As Long
    Winner As Boolean
    WinnerTimer As Long
    WinnerName As String
End Type

Type FriendsListUDT
    FriendName As String
End Type

Type PlayerRec
    ' General
    Name As String * NAME_LENGTH
    Sex As Byte
    Class As Byte
    Sprite As Integer
    Level As Byte
    Exp As Long
    Access As Byte
    PK As Byte
    Kills As Long
    Deaths As Long
    MatchsWon As Long
    BPoints As Long
    BCash As Long
    BonusDistance As Long
    BonusLimit As Long
    
    ' Vitals
    Vital(1 To Vitals.Vital_Count - 1) As Long
    
    ' Stats
    Stat(1 To Stats.Stat_Count - 1) As Byte
    POINTS As Byte
    
    ' Worn equipment
    Equipment(1 To Equipment.Equipment_Count - 1) As Byte
    
    ' Inventory
    Inv(1 To MAX_INV) As PlayerInvRec
    Spell(1 To MAX_PLAYER_SPELLS) As Byte
    
    ' Position
    Map As Integer
    x As Byte
    y As Byte
    Dir As Byte
End Type
    
Type AccountRec
    ' Account
    Login As String * NAME_LENGTH
    Password As String * NAME_LENGTH
       
    ' Characters (we use 0 to prevent a crash that still needs to be figured out)
    Char(0 To MAX_CHARS) As PlayerRec
    Friends(1 To MAX_FRIENDS) As FriendsListUDT
    AmountofFriends As Long
End Type

Type TempPlayerRec
    ' Non saved local vars
    Buffer As String
    CharNum As Byte
    InGame As Boolean
    AttackTimer As Long
    DataTimer As Long
    DataBytes As Long
    DataPackets As Long
    PartyPlayer As Long
    InParty As Byte
    TargetType As Byte
    Target As Byte
    CastedSpell As Byte
    PartyStarter As Byte
    GettingMap As Byte
    Distance As Byte
    DeathTimer As Long
    Dieing As Boolean
    Bombs As Byte
    TotalBombs As Integer
    Watching As Boolean
    Throw As Boolean
    Speed As Boolean
    Team As Long
End Type

Type BombRec
    Map As Long
    MakerIndex As Long
    Maker As String
    Timer As Long
    Distance As Long
    Here As Boolean
End Type

Type TileRec
    Ground As Integer
    Mask As Integer
    Anim As Integer
    Fringe As Integer
    Type As Byte
    Data1 As Integer
    Data2 As Integer
    Data3 As Integer
End Type

Type MapRec
    Name As String * NAME_LENGTH
    PlayerLimit As Byte
    Revision As Long
    Moral As Byte
    TileSet As Integer
    Up As Integer
    Down As Integer
    Left As Integer
    Right As Integer
    Music As Byte
    BootMap As Integer
    BootX As Byte
    BootY As Byte
    Shop As Byte
    Indoors As Byte
    Tile(0 To MAX_MAPX, 0 To MAX_MAPY) As TileRec
    Npc(1 To MAX_MAP_NPCS) As Byte
End Type

Type ClassRec
    Name As String * NAME_LENGTH
    
    Sprite As Integer
    
    Stat(1 To Stats.Stat_Count - 1) As Byte
End Type

Type ItemRec
    Name As String * NAME_LENGTH
    
    Pic As Integer
    Type As Byte
    Data1 As Integer
    Data2 As Integer
    Data3 As Integer
End Type

Type MapItemRec
    Num As Byte
    Value As Long
    Dur As Integer
    
    x As Byte
    y As Byte
End Type

Type NpcRec
    Name As String * NAME_LENGTH
    AttackSay As String * 100
    
    Sprite As Integer
    SpawnSecs As Long
    Behavior As Byte
    Range As Byte
    
    DropChance As Integer
    DropItem As Byte
    DropItemValue As Integer
    
    Stat(1 To Stats.Stat_Count - 1) As Byte
End Type

Type MapNpcRec
    Num As Integer
    
    Target As Integer
    
    Vital(1 To Vitals.Vital_Count - 1) As Long
        
    x As Byte
    y As Byte
    Dir As Integer
    
    ' For server use only
    SpawnWait As Long
    AttackTimer As Long
End Type

Type TradeItemRec
    GiveItem As Long
    GiveValue As Long
    GetItem As Long
    GetValue As Long
End Type

Type ShopRec
    Name As String * NAME_LENGTH
    JoinSay As String * 255
    LeaveSay As String * 255
    FixesItems As Byte
    TradeItem(1 To MAX_TRADES) As TradeItemRec
End Type
    
Type SpellRec
    Name As String * NAME_LENGTH
    ClassReq As Byte
    LevelReq As Byte
    Type As Byte
    Data1 As Integer
    Data2 As Integer
    Data3 As Integer
End Type

Type TempTileRec
    DoorOpen(0 To MAX_MAPX, 0 To MAX_MAPY)  As Byte
    DoorTimer As Long
End Type
