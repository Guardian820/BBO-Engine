Attribute VB_Name = "modTypes"
Option Explicit

' ******************************************
' **              BMO Source              **
' ******************************************

' Public data structures
Public Map As MapRec
Public TempTile(0 To MAX_MAPX, 0 To MAX_MAPY) As TempTileRec
Public Player(1 To MAX_PLAYERS) As PlayerRec
Public Class() As ClassRec
Public Item(1 To MAX_ITEMS) As ItemRec
Public Npc(1 To MAX_NPCS) As NpcRec
Public MapItem(1 To MAX_MAP_ITEMS) As MapItemRec
Public MapNpc(1 To MAX_MAP_NPCS) As MapNpcRec
Public Shop(1 To MAX_SHOPS) As ShopRec
Public Spell(1 To MAX_SPELLS) As SpellRec
Public Wall(0 To MAX_MAPX, 0 To MAX_MAPY) As WallRec

Type PlayerInvRec
    Num As Byte
    Value As Long
    Dur As Integer
End Type
Type FriendsListUDT
    FriendName As String
End Type

Type PlayerRec
    ' General
    Name As String * NAME_LENGTH
    Class As Byte
    sprite As Integer
    Level As Byte
    Exp As Long
    Access As Byte
    PK As Byte
    Kills As Long
    Deaths As Long
    MatchsWon As Long
    BPoints As Long
    BCash As Long
    Rank As String
    
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
    
    ' Client use only
    MaxHP As Long
    MaxMP As Long
    MaxSP As Long
    XOffset As Integer
    YOffset As Integer
    Moving As Byte
    Attacking As Byte
    AttackTimer As Long
    MapGetTimer As Long
    CastedSpell As Byte
    WalkTimer As Long
    Dieing As Boolean
    DeathTimer As Long
    Jumping As Boolean
    JumpTimer As Long
    Watching As Boolean
End Type
    
Type TileRec
    Ground As Long
    Mask As Long
    Anim As Long
    Fringe As Long
    Type As Byte
    Data1 As Integer
    Data2 As Integer
    Data3 As Integer
End Type

Type FireRec
    Here As Boolean
    Timer As Long
    Ended As Byte
    Direction As Byte
    Owner As Long
End Type

Type WallRec
    Here As Boolean
    Timer As Long
    Dieing As Boolean
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
    Tile(0 To MAX_MAPX, 0 To MAX_MAPY) As TileRec
    Npc(1 To MAX_MAP_NPCS) As Byte
    Fire(0 To MAX_MAPX, 0 To MAX_MAPY) As FireRec
End Type

Type ClassRec
    Name As String * NAME_LENGTH
    sprite As Integer
    
    Stat(1 To Stats.Stat_Count - 1) As Byte
    
    ' For client use
    Vital(1 To Vitals.Vital_Count - 1) As Long
End Type

Type ItemRec
    Name As String * NAME_LENGTH
    
    Pic As Integer
    Type As Byte
    Data1 As Integer
    Data2 As Integer
    Data3 As Integer
    
    Anim As Byte
    AnimTimer As Long
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
    
    sprite As Integer
    SpawnSecs As Long
    Behavior As Byte
    Range As Byte
    
    DropChance As Integer
    DropItem As Byte
    DropItemValue As Integer
    
    Stat(1 To Stats.Stat_Count - 1) As Byte
End Type

Type MapNpcRec
    Num As Byte
    
    Target As Byte
    
    Vital(1 To Vitals.Vital_Count - 1) As Long
        
    Map As Integer
    x As Byte
    y As Byte
    Dir As Byte
    Friends(1 To MAX_FRIENDS) As FriendsListUDT
    AmountofFriends As Long
    ' Client use only
    XOffset As Integer
    YOffset As Integer
    Moving As Byte
    Attacking As Byte
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
    JoinSay As String * 100
    LeaveSay As String * 100
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
    DoorOpen As Byte
End Type

