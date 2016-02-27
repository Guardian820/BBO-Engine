Attribute VB_Name = "modDirectDraw7"
Option Explicit

' ******************************************
' **              BMO Source              **
' ******************************************
' ** Renders graphics                     **
' ******************************************

' Master Object, leave one commented out
'Public DX7 As DirectX7 ' late binding
Public DX7 As New DirectX7 ' early binding
' ------------------------------------------

Public DD As DirectDraw7 ' DirectDraw7 Object

Public DD_Clip As DirectDrawClipper ' Clipper object

' primary surface
Public DDS_Primary As DirectDrawSurface7
Public DDSD_Primary As DDSURFACEDESC2

' back buffer
Public DDS_BackBuffer As DirectDrawSurface7
Public DDSD_BackBuffer As DDSURFACEDESC2

' gfx buffers
Public DDS_Item As DirectDrawSurface7
Public DDS_Sprite As DirectDrawSurface7
Public DDS_Tile As DirectDrawSurface7
Public DDS_Misc As DirectDrawSurface7
Public DDS_Bomb As DirectDrawSurface7
Public DDS_Fire As DirectDrawSurface7
Public DDS_Jump(1 To 3) As DirectDrawSurface7
Public DDS_Wall As DirectDrawSurface7

Public DDSD_Item As DDSURFACEDESC2
Public DDSD_Sprite As DDSURFACEDESC2
Public DDSD_Tile As DDSURFACEDESC2
Public DDSD_Misc As DDSURFACEDESC2
Public DDSD_Bomb As DDSURFACEDESC2
Public DDSD_Fire As DDSURFACEDESC2
Public DDSD_Jump(1 To 3) As DDSURFACEDESC2
Public DDSD_Wall As DDSURFACEDESC2

Private tmrBomb As Long
Private BombAnim As Byte
Public ItemAnim As Byte
Private tmrItemAnim As Long
Public WalkAnim(1 To MAX_PLAYERS) As Byte
Public tmrWalkAnim(1 To MAX_PLAYERS) As Byte

' ********************
' ** Initialization **
' ********************
Public Sub InitDirectDraw()
On Error GoTo ErrorHandle
    
    ' Initialize direct draw
    Set DD = DX7.DirectDrawCreate(vbNullString) ' empty string forces primary device

    ' dictates how we access thescreen and how other programs
    ' running at the same time will be allowed to access the screen as well.
    Call DD.SetCooperativeLevel(frmMainGame.hWnd, DDSCL_NORMAL)
        
    ' Init type and set the primary surface
    With DDSD_Primary
        .lFlags = DDSD_CAPS
        .ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE
    End With
        
    Set DDS_Primary = DD.CreateSurface(DDSD_Primary)
    
    ' Create the clipper
    Set DD_Clip = DD.CreateClipper(0)
    
    ' Associate the picture hwnd with the clipper
    Call DD_Clip.SetHWnd(frmMainGame.picScreen.hWnd)
        
    ' Have the blits to the screen clipped to the picture box
    Call DDS_Primary.SetClipper(DD_Clip) ' method attaches a clipper object to, or deletes one from, a surface.
 
    ' Initialize back buffer
    With DDSD_BackBuffer
        .lFlags = DDSD_CAPS Or DDSD_WIDTH Or DDSD_HEIGHT
        .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
        .lWidth = (MAX_MAPX + 1) * PIC_X
        .lHeight = (MAX_MAPY + 1) * PIC_Y
    End With
        
    '  sets the backbuffer dimensions to picScreen
    frmMainGame.picScreen.width = DDSD_BackBuffer.lWidth
    frmMainGame.picScreen.height = DDSD_BackBuffer.lHeight
    
    ' initialize the backbuffer
    Set DDS_BackBuffer = DD.CreateSurface(DDSD_BackBuffer)
    Exit Sub
    
ErrorHandle:
    DestroyGame
End Sub

' This sub gets the mask color from the surface loaded from a bitmap image
Public Sub SetMaskColorFromPixel(ByRef TheSurface As DirectDrawSurface7, ByVal x As Long, ByVal y As Long)
  Dim TmpR As RECT
  Dim TmpDDSD As DDSURFACEDESC2
  Dim TmpColorKey As DDCOLORKEY

    With TmpR
        .Left = x
        .Top = y
        .Right = x
        .Bottom = y
    End With

    TheSurface.Lock TmpR, TmpDDSD, DDLOCK_WAIT Or DDLOCK_READONLY, 0

    With TmpColorKey
        .Low = TheSurface.GetLockedPixel(x, y)
        .High = .Low
    End With

    TheSurface.SetColorKey DDCKEY_SRCBLT, TmpColorKey
    TheSurface.Unlock TmpR
End Sub

Public Sub InitSurfaces()
    Call InitDDSurf("sprites", DDSD_Sprite, DDS_Sprite)
    Call InitDDSurf("items", DDSD_Item, DDS_Item)
    Call InitDDSurf("bomb", DDSD_Bomb, DDS_Bomb)
    Call InitDDSurf("fire", DDSD_Fire, DDS_Fire)
    Call InitDDSurf("jump1", DDSD_Jump(1), DDS_Jump(1))
    Call InitDDSurf("jump2", DDSD_Jump(2), DDS_Jump(2))
    Call InitDDSurf("jump3", DDSD_Jump(3), DDS_Jump(3))
    Call InitDDSurf("wall", DDSD_Wall, DDS_Wall)
End Sub

' Initializing a surface, using a bitmap
Public Sub InitDDSurf(FileName As String, ByRef SurfDesc As DDSURFACEDESC2, ByRef Surf As DirectDrawSurface7)
    ' Set path
    FileName = App.Path & GFX_PATH & FileName & GFX_EXT
    
    ' check if file exists
    If Not FileExist(FileName, True) Then
        MsgBox "Arquivo não encontrado: " & FileName
        DestroyGame
    End If
    
    ' set flags
    SurfDesc.lFlags = DDSD_CAPS
    
    ' select one
    SurfDesc.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN ' auto determine best
    'SurfDesc.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY ' system memory
    'SurfDesc.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_VIDEOMEMORY ' video memory
    
    ' init object
    Set Surf = DD.CreateSurfaceFromFile(FileName, SurfDesc)

    Call SetMaskColorFromPixel(Surf, 0, 0)
End Sub

Public Sub InitTileSurf(ByVal TileSet As Integer)
    ' Destroy surface if it exist
    If Not DDS_Tile Is Nothing Then
        Set DDS_Tile = Nothing
        Call ZeroMemory(ByVal VarPtr(DDSD_Tile), LenB(DDSD_Tile))
    End If

    Call InitDDSurf("tiles" & TileSet, DDSD_Tile, DDS_Tile)
End Sub

Public Function CheckSurfaces() As Boolean
On Error GoTo ErrorHandle

    ' Check if we need to restore surfaces
    If NeedToRestoreSurfaces Then
        DD.RestoreAllSurfaces
        Call InitSurfaces
    End If
    
    CheckSurfaces = True
    Exit Function
    
ErrorHandle:
    ' re-initialize DirectDraw
    Call DestroyDirectDraw
    Call InitDirectDraw
    Call InitTileSurf(Map.TileSet)
    Call InitSurfaces
    
    CheckSurfaces = False
    
End Function

Private Function NeedToRestoreSurfaces() As Boolean
    
    NeedToRestoreSurfaces = True
    
    If DD.TestCooperativeLevel = DD_OK Then NeedToRestoreSurfaces = False
    
End Function

Public Sub DestroyDirectDraw()
' Unload DirectDraw
Set DDS_Tile = Nothing
Set DDS_Item = Nothing
Set DDS_Sprite = Nothing
Set DDS_Misc = Nothing

Set DDS_BackBuffer = Nothing
Set DDS_Primary = Nothing

Set DD_Clip = Nothing
Set DD = Nothing

End Sub

' **************
' ** Blitting **
' **************

Public Function Engine_BltToDC(ByRef Surface As DirectDrawSurface7, sRECT As DxVBLib.RECT, dRECT As DxVBLib.RECT, ByRef picBox As VB.PictureBox, Optional Clear As Boolean = True) As Boolean
On Error GoTo ErrorHandle
    
    If Clear Then
        picBox.Cls
    End If
    
    Call Surface.BltToDC(picBox.hdc, sRECT, dRECT)
    picBox.Refresh

    Engine_BltToDC = True
    Exit Function

ErrorHandle:
    ' returns false on error
    Engine_BltToDC = False

End Function

Public Sub BltMapTile(ByVal x As Long, ByVal y As Long)
Dim rec As DxVBLib.RECT

    With Map.Tile(x, y)
    
        rec.Top = (.Ground \ TILESHEET_WIDTH) * PIC_Y
        rec.Bottom = rec.Top + PIC_Y
        rec.Left = (.Ground Mod TILESHEET_WIDTH) * PIC_X
        rec.Right = rec.Left + PIC_X
        Call DDS_BackBuffer.BltFast(x * PIC_X, y * PIC_Y, DDS_Tile, rec, DDBLTFAST_WAIT)
    
        If MapAnim = 0 Or .Anim <= 0 Then
            If .Mask > 0 Then
                If TempTile(x, y).DoorOpen = NO Then
                    rec.Top = (.Mask \ TILESHEET_WIDTH) * PIC_Y
                    rec.Bottom = rec.Top + PIC_Y
                    rec.Left = (.Mask Mod TILESHEET_WIDTH) * PIC_X
                    rec.Right = rec.Left + PIC_X
                    Call DDS_BackBuffer.BltFast(x * PIC_X, y * PIC_Y, DDS_Tile, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
            End If
        Else
            ' Is there an animation tile to draw?
            If .Anim > 0 Then
                rec.Top = (.Anim \ TILESHEET_WIDTH) * PIC_Y
                rec.Bottom = rec.Top + PIC_Y
                rec.Left = (.Anim Mod TILESHEET_WIDTH) * PIC_X
                rec.Right = rec.Left + PIC_X
                Call DDS_BackBuffer.BltFast(x * PIC_X, y * PIC_Y, DDS_Tile, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
            End If
        End If
    
    End With

End Sub

Public Sub BltMapFringeTile(ByVal x As Long, ByVal y As Long)
Dim rec As DxVBLib.RECT

    With Map.Tile(x, y)
        If .Fringe > 0 Then
            rec.Top = (.Fringe \ TILESHEET_WIDTH) * PIC_Y
            rec.Bottom = rec.Top + PIC_Y
            rec.Left = Int(.Fringe Mod TILESHEET_WIDTH) * PIC_X
            rec.Right = rec.Left + PIC_X
            Call DDS_BackBuffer.BltFast(x * PIC_X, y * PIC_Y, DDS_Tile, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    End With
    
End Sub

Public Sub BltItem(ByVal ItemNum As Long)
Dim rec As DxVBLib.RECT

    With rec
        .Top = Item(MapItem(ItemNum).Num).Pic * PIC_Y
        .Bottom = .Top + PIC_Y
        .Left = Item(MapItem(ItemNum).Num).Anim * PIC_X
        .Right = .Left + PIC_X
    End With
    
    Call DDS_BackBuffer.BltFast(MapItem(ItemNum).x * PIC_X, MapItem(ItemNum).y * PIC_Y, DDS_Item, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
End Sub

Public Sub BltPlayer(ByVal Index As Long)
Dim Anim As Byte
Dim x As Long
Dim y As Long
Dim rec As DxVBLib.RECT

    If Player(Index).Moving > 0 Or (Index = MyIndex And IsTryingToMove) Or Player(Index).Dieing Then Anim = WalkAnim(Index) Else Anim = 0
    
    With rec
        .Top = GetPlayerSprite(Index) * 64 'SIZE_Y
        .Bottom = .Top + 64 'SIZE_Y
        If Not Player(Index).Dieing Then
            .Left = (GetPlayerDir(Index) * 8 + Anim) * 64 'SIZE_X
        Else
            .Left = (Anim) * 64 'SIZE_X
        End If
        .Right = .Left + 64 'SIZE_X
    End With
    
    x = GetPlayerX(Index) * PIC_X + Player(Index).XOffset - 16
    y = GetPlayerY(Index) * PIC_Y + Player(Index).YOffset - 32 ' - 4 ' to raise the sprite by 4 pixels
    
    ' Check if its out of bounds because of the offset
    If y < 0 Then
        y = 0
        With rec
            .Top = .Top + (y * -1)
        End With
    End If
    If x < 0 Then
        x = 0
        With rec
            .Left = .Left + (x * -1)
        End With
    End If
    If x > 480 Then
        x = 480
    End If
    Call DDS_BackBuffer.BltFast(x, y, DDS_Sprite, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
End Sub


Public Sub BltNpc(ByVal MapNpcNum As Long)
Dim Anim As Byte
Dim x As Long
Dim y As Long
Dim rec As DxVBLib.RECT

    ' Make sure that theres an npc there, and if not exit the sub
    If MapNpc(MapNpcNum).Num <= 0 Then
        Exit Sub
    End If
    
    ' Check for animation
    Anim = 0
    If MapNpc(MapNpcNum).Moving > 0 Then
        Select Case MapNpc(MapNpcNum).Dir
            Case DIR_UP
                If (MapNpc(MapNpcNum).YOffset < SIZE_Y / 2) Then Anim = 1 Else Anim = 2
            Case DIR_DOWN
                If (MapNpc(MapNpcNum).YOffset < SIZE_Y / 2 * -1) Then Anim = 1 Else Anim = 2
            Case DIR_LEFT
                If (MapNpc(MapNpcNum).XOffset < SIZE_Y / 2) Then Anim = 1 Else Anim = 2
            Case DIR_RIGHT
                If (MapNpc(MapNpcNum).XOffset < SIZE_Y / 2 * -1) Then Anim = 1 Else Anim = 2
        End Select
    End If
    'Else
    '    If MapNpc(MapNpcNum).AttackTimer + 500 > GetTickCount Then
    '        Anim = 2
    '    End If
    'End If
    
    ' Check to see if we want to stop making him attack
    With MapNpc(MapNpcNum)
        If .AttackTimer + 1000 < GetTickCount Then
            .Attacking = 0
            .AttackTimer = 0
        End If
    End With
    
    With rec
        .Top = Npc(MapNpc(MapNpcNum).Num).sprite * SIZE_Y
        .Bottom = .Top + SIZE_Y
        .Left = (MapNpc(MapNpcNum).Dir * 3 + Anim) * SIZE_X
        .Right = .Left + SIZE_X
    End With
    
    With MapNpc(MapNpcNum)
        x = .x * PIC_X + .XOffset
        y = .y * PIC_Y + .YOffset ' - 4 ' to raise the sprite by 4 pixels
    End With
    
    ' Check if its out of bounds because of the offset
    If y < 0 Then
        y = 0
        rec.Top = rec.Top + (y * -1)
    End If
        
    Call DDS_BackBuffer.BltFast(x, y, DDS_Sprite, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
End Sub

Public Sub BltWall(ByVal x As Long, ByVal y As Long)
Dim rec As DxVBLib.RECT
Dim Anim As Byte
    
    If Map.Tile(x, y).Type = TILE_TYPE_WALL Then
        If Wall(x, y).Dieing Then
            If Wall(x, y).Timer + 100 > GetTickCount Then
                Anim = 0
            ElseIf Wall(x, y).Timer + 200 > GetTickCount Then
                Anim = 1
            ElseIf Wall(x, y).Timer + 300 > GetTickCount Then
                Anim = 2
            ElseIf Wall(x, y).Timer + 400 > GetTickCount Then
                Anim = 3
            ElseIf Wall(x, y).Timer + 500 > GetTickCount Then
                Anim = 4
            Else
                Wall(x, y).Dieing = False
            End If
        Else
            Anim = 0
        End If
        
        If Not Wall(x, y).Dieing And Not Wall(x, y).Here Then
            Anim = 0
        Else
            If Not Wall(x, y).Dieing Then Anim = 10
        End If
        
        With rec
            .Top = Map.Tile(x, y).Data1 * 32
            .Bottom = .Top + 32
            .Left = Anim * 32
            .Right = .Left + 32
        End With
        
        Call DDS_BackBuffer.BltFast(x * PIC_X, y * PIC_Y, DDS_Wall, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    End If

End Sub

Public Sub BltJumping(ByVal Index As Long)
Dim rec As DxVBLib.RECT
Dim x As Long
Dim y As Long
Dim Anim As Byte
Dim JumpNum As Byte

    If GetPlayerSprite(Index) < 7 Then JumpNum = 1
    If GetPlayerSprite(Index) > 6 And GetPlayerSprite(Index) < 14 Then JumpNum = 3
    If GetPlayerSprite(Index) > 13 And GetPlayerSprite(Index) < 21 Then JumpNum = 2
    
    Select Case JumpNum
    
        Case 1
            If Player(Index).JumpTimer + 200 > GetTickCount Then
                Anim = 0
            ElseIf Player(Index).JumpTimer + 400 > GetTickCount Then
                Anim = 1
            ElseIf Player(Index).JumpTimer + 600 > GetTickCount Then
                Anim = 2
            ElseIf Player(Index).JumpTimer + 800 > GetTickCount Then
                Anim = 3
            ElseIf Player(Index).JumpTimer + 1000 > GetTickCount Then
                Anim = 4
            Else
                Player(Index).Jumping = False
            End If
            
            x = (GetPlayerX(Index) * PIC_X + Player(Index).XOffset - 16)
            y = (GetPlayerY(Index) * PIC_Y + Player(Index).YOffset - 32) - 2
            
            With rec
                .Top = GetPlayerSprite(Index) * 64
                .Bottom = .Top + 64
                .Left = Anim * 66
                .Right = .Left + 66
            End With
            
            Call DDS_BackBuffer.BltFast(x, y, DDS_Jump(1), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
            Exit Sub
            
        Case 2
            If Player(Index).JumpTimer + 50 > GetTickCount Then
                Anim = 0
            ElseIf Player(Index).JumpTimer + 150 > GetTickCount Then
                Anim = 1
            ElseIf Player(Index).JumpTimer + 300 > GetTickCount Then
                Anim = 2
            ElseIf Player(Index).JumpTimer + 450 > GetTickCount Then
                Anim = 3
            ElseIf Player(Index).JumpTimer + 600 > GetTickCount Then
                Anim = 4
            ElseIf Player(Index).JumpTimer + 850 > GetTickCount Then
                Anim = 5
            ElseIf Player(Index).JumpTimer + 1000 > GetTickCount Then
                Anim = 6
            Else
                Player(Index).Jumping = False
            End If
            
            x = (GetPlayerX(Index) * PIC_X + Player(Index).XOffset - 16)
            y = (GetPlayerY(Index) * PIC_Y + Player(Index).YOffset - 32) - 7
            
            With rec
                .Top = (GetPlayerSprite(Index) - 14) * 78
                .Bottom = .Top + 78
                .Left = Anim * 64
                .Right = .Left + 64
            End With
            
            Call DDS_BackBuffer.BltFast(x, y, DDS_Jump(2), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
            Exit Sub
            
        Case 3
            If Player(Index).JumpTimer + 100 > GetTickCount Then
                Anim = 0
            ElseIf Player(Index).JumpTimer + 200 > GetTickCount Then
                Anim = 1
            ElseIf Player(Index).JumpTimer + 400 > GetTickCount Then
                Anim = 2
            ElseIf Player(Index).JumpTimer + 600 > GetTickCount Then
                Anim = 3
            ElseIf Player(Index).JumpTimer + 800 > GetTickCount Then
                Anim = 4
            ElseIf Player(Index).JumpTimer + 1000 > GetTickCount Then
                Anim = 5
            Else
                Player(Index).Jumping = False
            End If
            
            x = (GetPlayerX(Index) * PIC_X + Player(Index).XOffset - 16)
            y = (GetPlayerY(Index) * PIC_Y + Player(Index).YOffset - 32) - 15
            
            With rec
                .Top = (GetPlayerSprite(Index) - 7) * 96
                .Bottom = .Top + 96
                .Left = Anim * 64
                .Right = .Left + 64
            End With
            
            Call DDS_BackBuffer.BltFast(x, y, DDS_Jump(3), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
            Exit Sub
            
        Case Else
            Player(Index).Jumping = False
            Exit Sub
        
    End Select

End Sub

Public Sub BltFire()
Dim x As Long
Dim y As Long
Dim rec As DxVBLib.RECT
Dim Anim As Byte

    For x = 0 To MAX_MAPX
        For y = 0 To MAX_MAPY
            With Map.Fire(x, y)
                If .Here Then
                    If .Timer + 50 > GetTickCount Then
                        Anim = 0
                    ElseIf .Timer + 200 > GetTickCount Then
                        Anim = 1
                    ElseIf .Timer + 350 > GetTickCount Then
                        Anim = 2
                    ElseIf .Timer + 400 > GetTickCount Then
                        Anim = 3
                    ElseIf .Timer + 550 > GetTickCount Then
                        Anim = 4
                    ElseIf .Timer + 700 > GetTickCount Then
                        Anim = 5
                    ElseIf .Timer + 850 > GetTickCount Then
                        Anim = 6
                    ElseIf .Timer + 1000 > GetTickCount Then
                        Anim = 7
                        .Here = False
                    End If
                    
                    If .Here Then
                    
                        With rec
                            .Top = Map.Fire(x, y).Ended * 32
                            .Bottom = .Top + 32
                            .Left = (Anim * 32) + (Map.Fire(x, y).Direction * 256)
                            .Right = .Left + PIC_X
                        End With
                        
                        Call DDS_BackBuffer.BltFast(x * PIC_X, y * PIC_Y, DDS_Fire, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    End If
                    
                End If
            End With
        Next
    Next

End Sub

Public Sub BltBombs()
Dim x As Long
Dim y As Long
Dim rec As DxVBLib.RECT

    For x = 0 To MAX_MAPX
        For y = 0 To MAX_MAPY
            With Map.Tile(x, y)
                If .Type = TILE_TYPE_BOMB Then
                    With rec
                        .Top = 0
                        .Bottom = 32
                        .Left = BombAnim * 32
                        .Right = .Left + PIC_X
                    End With
                    
                    Call DDS_BackBuffer.BltFast(x * PIC_X, y * PIC_Y, DDS_Bomb, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
            End With
        Next
    Next
    
End Sub

' ******************
' ** Game Editors **
' ******************

Public Sub BltMapEditor()
Dim height As Long
Dim width As Long
Dim sRECT As DxVBLib.RECT
Dim dRECT As DxVBLib.RECT

    height = DDSD_Tile.lHeight
    width = DDSD_Tile.lWidth
    
    dRECT.Top = 0
    dRECT.Bottom = height
    dRECT.Left = 0
    dRECT.Right = width
    
    frmMainGame.picBackSelect.height = height
    frmMainGame.picBackSelect.width = width
   
    Call Engine_BltToDC(DDS_Tile, sRECT, dRECT, frmMainGame.picBackSelect)
    
End Sub

Public Sub BltMapEditorTilePreview()
Dim sRECT As DxVBLib.RECT
Dim dRECT As DxVBLib.RECT

    sRECT.Top = EditorTileY * PIC_Y
    sRECT.Bottom = sRECT.Top + PIC_Y
    sRECT.Left = EditorTileX * PIC_X
    sRECT.Right = sRECT.Left + PIC_X
    
    dRECT.Top = 0
    dRECT.Bottom = PIC_Y
    dRECT.Left = 0
    dRECT.Right = PIC_X
    
    Call Engine_BltToDC(DDS_Tile, sRECT, dRECT, frmMainGame.picSelect)
End Sub

Public Sub BltTileOutline()
Dim rec As DxVBLib.RECT

    With rec
        .Top = 0
        .Bottom = .Top + PIC_Y
        .Left = 0
        .Right = .Left + PIC_X
    End With
        
    Call DDS_BackBuffer.BltFast(CurX * PIC_X, CurY * PIC_Y, DDS_Misc, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
End Sub

Public Sub MapItemEditorBltItem()
Dim sRECT As DxVBLib.RECT
Dim dRECT As DxVBLib.RECT

    sRECT.Top = frmMapItem.scrlItem.Value * PIC_Y
    sRECT.Bottom = sRECT.Top + PIC_Y
    sRECT.Left = 0
    sRECT.Right = sRECT.Left + PIC_X
    
    dRECT.Top = 0
    dRECT.Bottom = PIC_Y
    dRECT.Left = 0
    dRECT.Right = PIC_X

    Call Engine_BltToDC(DDS_Item, sRECT, dRECT, frmMapItem.picPreview)
    
End Sub

Public Sub KeyItemEditorBltItem()
'Dim sRECT As DxVBLib.RECT
'Dim dRECT As DxVBLib.RECT
'
'    sRECT.Top = frmMapKey.scrlItem.Value * PIC_Y
'    sRECT.Bottom = sRECT.Top + PIC_Y
'    sRECT.Left = 0
'    sRECT.Right = sRECT.Left + PIC_X
'
'    dRECT.Top = 0
'    dRECT.Bottom = PIC_Y
'    dRECT.Left = 0
'    dRECT.Right = PIC_X
'
'    Call Engine_BltToDC(DDS_Item, sRECT, dRECT, frmMapKey.picPreview)
'
End Sub

Public Sub ItemEditorBltItem()
Dim sRECT As DxVBLib.RECT
Dim dRECT As DxVBLib.RECT

    sRECT.Top = frmItemEditor.scrlPic.Value * PIC_Y
    sRECT.Bottom = sRECT.Top + PIC_Y
    sRECT.Left = 0
    sRECT.Right = sRECT.Left + PIC_X
    
    dRECT.Top = 0
    dRECT.Bottom = PIC_Y
    dRECT.Left = 0
    dRECT.Right = PIC_X

    Call Engine_BltToDC(DDS_Item, sRECT, dRECT, frmItemEditor.picPic)
    
End Sub

Public Sub WallEditorBltPic()
Dim sRECT As DxVBLib.RECT
Dim dRECT As DxVBLib.RECT

    sRECT.Top = frmWall.scrlPicture.Value * PIC_Y
    sRECT.Bottom = sRECT.Top + PIC_Y
    sRECT.Left = 0
    sRECT.Right = sRECT.Left + PIC_X
    
    dRECT.Top = 0
    dRECT.Bottom = PIC_Y
    dRECT.Left = 0
    dRECT.Right = PIC_X

    Call Engine_BltToDC(DDS_Wall, sRECT, dRECT, frmWall.picPreview)
    
End Sub

Public Sub NpcEditorBltSprite()

End Sub

Public Sub Render_Graphics()
Dim x As Long
Dim y As Long
Dim i As Long
Dim rec As DxVBLib.RECT
Dim rec_pos As DxVBLib.RECT

    If frmMainGame.WindowState = vbMinimized Then Exit Sub
    
    If Not CheckSurfaces Then Exit Sub
 
    If Not GettingMap Then
    
        Call DDS_BackBuffer.BltColorFill(rec, 0) ' clear backbuffer
        
        ' blit lower tiles
        For x = 0 To MAX_MAPX
            For y = 0 To MAX_MAPY
                Call BltMapTile(x, y)
            Next
        Next
        
        For i = 1 To MAX_MAP_ITEMS
            If Item(i).AnimTimer < GetTickCount Then
                If Item(i).Pic > 0 Then
                    Item(i).Anim = Item(i).Anim + 1
                    If Item(i).Anim > 4 Then Item(i).Anim = 0
                    Item(i).AnimTimer = GetTickCount + 250
                End If
            End If
        Next
        
        ' Blit out the items
        For i = 1 To MAX_MAP_ITEMS
            If MapItem(i).Num > 0 Then
                Call BltItem(i)
            End If
        Next
        
        For x = 0 To MAX_MAPX
            For y = 0 To MAX_MAPY
                'If Wall(x, y).Dieing Then
                    BltWall x, y
                'End If
            Next
        Next
        
        If tmrBomb < GetTickCount Then
            BombAnim = BombAnim + 1
            If BombAnim > 3 Then BombAnim = 0
            tmrBomb = GetTickCount + 250
        End If
        
        BltBombs
        
        BltFire
        
        For i = 1 To PlayersOnMapHighIndex
        
            If Player(PlayersOnMap(i)).Moving > 0 Or IsTryingToMove Then
            
                If WalkAnim(PlayersOnMap(i)) < 1 Then WalkAnim(PlayersOnMap(i)) = 1
                
                If Player(PlayersOnMap(i)).WalkTimer < GetTickCount Then
                    WalkAnim(PlayersOnMap(i)) = WalkAnim(PlayersOnMap(i)) + 1
                    If WalkAnim(PlayersOnMap(i)) = 3 Then WalkAnim(PlayersOnMap(i)) = 4
                    If WalkAnim(PlayersOnMap(i)) > 7 Then WalkAnim(PlayersOnMap(i)) = 1
                    Player(PlayersOnMap(i)).WalkTimer = GetTickCount + 100
                End If
            
            End If
            
            If Player(PlayersOnMap(i)).Dieing Then
                If Player(PlayersOnMap(i)).DeathTimer + 200 > GetTickCount Then
                    WalkAnim(PlayersOnMap(i)) = 32
                ElseIf Player(PlayersOnMap(i)).DeathTimer + 400 > GetTickCount Then
                    WalkAnim(PlayersOnMap(i)) = 33
                ElseIf Player(PlayersOnMap(i)).DeathTimer + 600 > GetTickCount Then
                    WalkAnim(PlayersOnMap(i)) = 34
                ElseIf Player(PlayersOnMap(i)).DeathTimer + 800 > GetTickCount Then
                    WalkAnim(PlayersOnMap(i)) = 35
                ElseIf Player(PlayersOnMap(i)).DeathTimer + 1000 > GetTickCount Then
                    WalkAnim(PlayersOnMap(i)) = 36
                End If
            End If
        
        Next
        
        ' Blit out players
        For y = 0 To MAX_MAPY
            For i = 1 To PlayersOnMapHighIndex
                If GetPlayerY(PlayersOnMap(i)) = y Then
                    If Not Player(PlayersOnMap(i)).Jumping Then
                        If Not Player(PlayersOnMap(i)).Watching Then Call BltPlayer(PlayersOnMap(i))
                    Else
                        If Not Player(PlayersOnMap(i)).Watching Then BltJumping PlayersOnMap(i)
                    End If
                End If
            Next
        Next
        
        ' blit out upper tiles
        For x = 0 To MAX_MAPX
            For y = 0 To MAX_MAPY
                Call BltMapFringeTile(x, y)
            Next
        Next
        
        ' blit out a square at mouse cursor
        If InEditor Then
            Call BltTileOutline
        End If
        
        ' Lock the backbuffer so we can draw text and names
        TexthDC = DDS_BackBuffer.GetDC
        
        ' draw FPS
        If BFPS Then
            Call DrawText(TexthDC, (((MAX_MAPX + 1) * PIC_X) - Len("FPS: " & GameFPS) * FONT_WIDTH), 1, Trim$("FPS: " & GameFPS), QBColor(Yellow))
        End If
        
        ' draw cursor, player X and Y locations
        If BLoc Then
            Call DrawText(TexthDC, 0, 1, Trim$("cur x: " & CurX & " y: " & CurY), QBColor(Yellow))
            Call DrawText(TexthDC, 0, 15, Trim$("loc x: " & GetPlayerX(MyIndex) & " y: " & GetPlayerY(MyIndex)), QBColor(Yellow))
            Call DrawText(TexthDC, 0, 27, Trim$(" (map #" & GetPlayerMap(MyIndex) & ")"), QBColor(Yellow))
        End If
        
        If Player(MyIndex).Watching = True Then
            DrawText TexthDC, (MAX_MAPX + 1) * PIC_X \ 2 - ((Len(Trim$("ASSISTINDO")) \ 2) * FONT_WIDTH), ((MAX_MAPY + 1) * PIC_Y \ 2) - FONT_HEIGHT, "ASSISTINDO", QBColor(BrightRed)
        End If
        
        ' draw player names
        For i = 1 To PlayersOnMapHighIndex
            If Not Player(PlayersOnMap(i)).Watching Then
            Call DrawPlayerName(PlayersOnMap(i))
            Call DrawPlayerLevel(PlayersOnMap(i))
            End If
        Next
        
        ' Blit out map attributes
        If InEditor Then Call BltMapAttributes
        
    Else
    
        TexthDC = DDS_BackBuffer.GetDC ' Lock the backbuffer so we can draw text and names
        
        ' Check if we are getting a map, and if we are tell them so
        Call DrawText(TexthDC, 50, 50, "Recebendo mapa...", QBColor(BrightCyan))
        
    End If
    
    ' Release DC
    Call DDS_BackBuffer.ReleaseDC(TexthDC)
    
    ' Get the rect to blit to
    Call DX7.GetWindowRect(frmMainGame.picScreen.hWnd, rec_pos)
    
    ' Blit the backbuffer
    Call DDS_Primary.Blt(rec_pos, DDS_BackBuffer, rec, DDBLT_WAIT)
    
End Sub

Private Function BltMapAttributes()
    Dim x As Long
    Dim y As Long

    If frmMainGame.optAttribs.Value Then
        For x = 0 To MAX_MAPX
            For y = 0 To MAX_MAPY
                With Map.Tile(x, y)
                    Select Case .Type
                    
                        Case TILE_TYPE_BLOCKED
                            DrawText TexthDC, ((x * PIC_X) - 4) + (PIC_X * 0.5), ((y * PIC_Y) - 7) + (PIC_Y * 0.5), "B", QBColor(BrightRed)
                        
                        Case TILE_TYPE_WARP
                            DrawText TexthDC, ((x * PIC_X) - 4) + (PIC_X * 0.5), ((y * PIC_Y) - 7) + (PIC_Y * 0.5), "W", QBColor(BrightBlue)
                    
                        Case TILE_TYPE_ITEM
                            DrawText TexthDC, ((x * PIC_X) - 4) + (PIC_X * 0.5), ((y * PIC_Y) - 7) + (PIC_Y * 0.5), "I", QBColor(White)
                    
                        Case TILE_TYPE_NPCAVOID
                            DrawText TexthDC, ((x * PIC_X) - 4) + (PIC_X * 0.5), ((y * PIC_Y) - 7) + (PIC_Y * 0.5), "N", QBColor(White)
                    
                        Case TILE_TYPE_KEY
                            DrawText TexthDC, ((x * PIC_X) - 4) + (PIC_X * 0.5), ((y * PIC_Y) - 7) + (PIC_Y * 0.5), "K", QBColor(White)
                    
                        Case TILE_TYPE_KEYOPEN
                            DrawText TexthDC, ((x * PIC_X) - 4) + (PIC_X * 0.5), ((y * PIC_Y) - 7) + (PIC_Y * 0.5), "O", QBColor(White)
                        
                        Case TILE_TYPE_WALL
                            DrawText TexthDC, ((x * PIC_X) - 4) + (PIC_X * 0.5), ((y * PIC_Y) - 7) + (PIC_Y * 0.5), "WA", QBColor(White)
                    
                    End Select
                End With
            Next
        Next
    End If
    
End Function

