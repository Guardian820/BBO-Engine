Attribute VB_Name = "modEnumerations"
Option Explicit

' ******************************************
' **              BMO Source              **
' ******************************************

' *************
' ** Packets **
' *************

' The order of the packets must match with the client's packet enumeration

' Packets sent by the server
Public Enum ServerPackets
    SAlertMsg = 1
    SAllChars
    SLoginOk
    SNewCharClasses
    SClassesData
    SInGame
    SPlayerInv
    SPlayerInvUpdate
    SPlayerWornEq
    SPlayerHp
    SPlayerMp
    SPlayerSp
    SPlayerStats
    SPlayerData
    SPlayerMove
    SNpcMove
    SPlayerDir
    SNpcDir
    SPlayerXY
    SAttack
    SNpcAttack
    SCheckForMap
    SMapData
    SMapItemData
    SMapNpcData
    SMapDone
    SSayMsg
    SGlobalMsg
    SAdminMsg
    SPlayerMsg
    SMapMsg
    SSpawnItem
    SItemEditor
    SUpdateItem
    SEditItem
    SREditor
    SSpawnNpc
    SNpcDead
    SNpcEditor
    SUpdateNpc
    SEditNpc
    SMapKey
    SEditMap
    SShopEditor
    SUpdateShop
    SEditShop
    SSpellEditor
    SUpdateSpell
    SEditSpell
    STrade
    SSpells
    SLeft
    SHighIndex
    SFire
    SBomb
    SDieing
    SWall
    SJump
    SSendRoomList
    SInLobby
    SStartMatch
    SStartTimer
    SWatching
    SLeaveRoom
    SHighScore
    SHighScoreList
    SRefreshWall
    SUpdateFriendList
    SSpeed
End Enum

' Packets recieved by the server
Public Enum ClientPackets
    CGetClasses = 1
    CNewAccount
    CDelAccount
    CLogin
    CAddChar
    CDelChar
    CUseChar
    CSayMsg
    CEmoteMsg
    CBroadcastMsg
    CGlobalMsg
    CAdminMsg
    CPlayerMsg
    CPlayerMove
    CPlayerDir
    CUseItem
    CAttack
    CUseStatPoint
    CPlayerInfoRequest
    CWarpMeTo
    CWarpToMe
    CWarpTo
    CSetSprite
    CGetStats
    CRequestNewMap
    CMapData
    CNeedMap
    CMapGetItem
    CMapDropItem
    CMapRespawn
    CMapReport
    CKickPlayer
    CBanList
    CBanDestroy
    CBanPlayer
    CRequestEditMap
    CRequestEditItem
    CEditItem
    CSaveItem
    CRequestEditNpc
    CEditNpc
    CSaveNpc
    CRequestEditShop
    CEditShop
    CSaveShop
    CRequestEditSpell
    CEditSpell
    CSaveSpell
    CDelete
    CSetAccess
    CWhosOnline
    CSetMotd
    CTrade
    CTradeRequest
    CFixItem
    CSearch
    CParty
    CJoinParty
    CLeaveParty
    CSpells
    CCast
    CQuit
    CFireDeath
    CJump
    CRequestItemSpawn
    CRequestRoomList
    CRequestJoinRoom
    CLeaveRoom
    CRequestHighScores
    CRequestStats
    CRefresh
    CAddFriend
    CRemoveFriend
    CUpdateFriendList
    CAddStatus
    CDealBP
    CDealBC
    CSaveSprite
    CThrow
End Enum

' ****************
' ** Statistics **
' ****************

' Stats used by Players, Npcs and Classes
Public Enum Stats
    Strength = 1
    Defense
    Speed
    Magic
    ' Make sure Stat_Count is below everything else
    Stat_Count
End Enum

' Vitals used by Players, Npcs and Classes
Public Enum Vitals
    HP = 1
    MP
    SP
    ' Mak sure Vital_Count is below everything else
    Vital_Count
End Enum

' Equipment used by Players
Public Enum Equipment
    Weapon = 1
    Armor
    Helmet
    Shield
    ' Mak sure Equipment_Count is below everything else
    Equipment_Count
End Enum

