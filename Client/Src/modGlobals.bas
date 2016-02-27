Attribute VB_Name = "modGlobals"
Option Explicit

' ******************************************
' **              BMO Source              **
' ******************************************

' Number of tiles in width
Public Const TILESHEET_WIDTH As Long = 7 ' * PIC_X pixels

' TCP variables
Public ServerIP As String
Public PlayerBuffer As String

Public Current_SpriteNum As Long

Public FEMALE As Boolean

Public RegisterSprite As Long
Public RegisterAnim As Long
Public RegisterWhite As Boolean
Public RegisterBlack As Boolean
Public RegisterGreen As Boolean
Public RegisterRed As Boolean
Public RegisterBlue As Boolean
Public RegisterYellow As Boolean

Public TargetName As String
Public TargetHwnd As Long

Public tmrFireDeath As Long

' Controls main gameloop
Public InGame As Boolean
Public isLogging As Boolean

' Text variables
Public TexthDC As Long
Public GameFont As Long

' Draw map name location
Public DrawMapNameX As Single
Public DrawMapNameY As Single
Public DrawMapNameColor As Long

' Stops movement when updating a map
Public CanMoveNow As Boolean

' Game direction vars
Public DirUp As Boolean
Public DirDown As Boolean
Public DirLeft As Boolean
Public DirRight As Boolean
Public ShiftDown As Boolean
Public ControlDown As Boolean
Public SpaceDown As Boolean

' Game text buffer
Public MyText As String

' Index of actual player
Public MyIndex As Long

' Map animation #, used to keep track of what map animation is currently on
Public MapAnim As Byte
Public MapAnimTimer As Long

' Used to freeze controls when getting a new map
Public GettingMap As Boolean

' Used to check if FPS needs to be drawn
Public BFPS As Boolean
Public BLoc As Boolean

Public GameFPS As Long ' frames per second rendered

' Text vars
Public vbQuote As String

' Mouse cursor tile location
Public CurX As Integer
Public CurY As Integer

' Game editors
Public Editor As Byte
Public EditorIndex As Long

Public MAX_TILESETS As Long

' Used to check if in editor or not and variables for use in editor
Public InEditor As Boolean
Public EditorTileX As Long
Public EditorTileY As Long
Public EditorWarpMap As Long
Public EditorWarpX As Long
Public EditorWarpY As Long

' Used for map item editor
Public ItemEditorNum As Long
Public ItemEditorValue As Long

' Used for map key editor
Public KeyEditorNum As Long
Public KeyEditorTake As Long

Public WallPicture As Long

' Used for map key open editor
Public KeyOpenEditorX As Long
Public KeyOpenEditorY As Long

' Used for parsing String packets
Public SEP_CHAR As String * 1
Public END_CHAR As String * 1

' Maximum classes
Public Max_Classes As Byte

' Used for improved looping
Public Player_HighIndex As Long
Public High_Npc_Index As Long

Public PlayersOnMapHighIndex As Long
Public PlayersOnMap() As Long

'Lobby anim
Public LobbyAnim As Long
'Waiting Anim
Public WaitAnim1 As Long
Public WaitAnim2 As Long
Public WaitAnim3 As Long
Public Waitanim4 As Long
Public Walk_Speed As Long
Public Run_Speed As Long
Public BonusSpeed As Boolean
Public BonusSpeedTime As Long
Public Temporizer As Long
Public Foco As Boolean
