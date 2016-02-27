Attribute VB_Name = "modGameEditors"
Option Explicit

' ******************************************
' **              BMO Source              **
' ******************************************

' ////////////////
' // Map Editor //
' ////////////////

Public Sub MapEditorInit()
    InEditor = True
    frmMainGame.picMapEditor.Visible = True
     
    Call InitDDSurf("misc", DDSD_Misc, DDS_Misc)
    frmMainGame.lblTileset = Map.TileSet
     
    frmMainGame.Width = 12600
     
    Call BltMapEditor
    Call BltMapEditorTilePreview
    
    frmMainGame.scrlPicture.Max = (frmMainGame.picBackSelect.Height \ PIC_Y) - (frmMainGame.picBack.Height \ PIC_Y)

End Sub

Public Sub MapEditorMouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Not isInBounds Then
        Exit Sub
    End If

    If Button = vbLeftButton Then
        If frmMainGame.optLayers.Value Then
            
            With Map.Tile(CurX, CurY)
                If frmMainGame.optGround.Value Then .Ground = EditorTileY * TILESHEET_WIDTH + EditorTileX
                If frmMainGame.optMask.Value Then .Mask = EditorTileY * TILESHEET_WIDTH + EditorTileX
                If frmMainGame.optAnim.Value Then .Anim = EditorTileY * TILESHEET_WIDTH + EditorTileX
                If frmMainGame.optFringe.Value Then .Fringe = EditorTileY * TILESHEET_WIDTH + EditorTileX
            End With
            
        Else
            With Map.Tile(CurX, CurY)
                If frmMainGame.optBlocked.Value Then .Type = TILE_TYPE_BLOCKED
                'If frmMainGame.optWarp.Value Then
                '    .Type = TILE_TYPE_WARP
                '    .Data1 = EditorWarpMap
                '    .Data2 = EditorWarpX
                '    .Data3 = EditorWarpY
                'End If
                If frmMainGame.optItem.Value Then
                    .Type = TILE_TYPE_ITEM
                    .Data1 = ItemEditorNum
                    .Data2 = ItemEditorValue
                    .Data3 = 0
                End If
                'If frmMainGame.optNpcAvoid.Value Then
                '    .Type = TILE_TYPE_NPCAVOID
                '    .Data1 = 0
                '    .Data2 = 0
                '    .Data3 = 0
                'End If
                If frmMainGame.OptWall.Value Then
                    .Type = TILE_TYPE_WALL
                    .Data1 = WallPicture
                    .Data2 = 0
                    .Data3 = 0
                End If
            End With
        End If
    End If
    
    If Button = vbRightButton Then
        If frmMainGame.optLayers.Value Then
            With Map.Tile(CurX, CurY)
                If frmMainGame.optGround.Value Then .Ground = 0
                If frmMainGame.optMask.Value Then .Mask = 0
                If frmMainGame.optAnim.Value Then .Anim = 0
                If frmMainGame.optFringe.Value Then .Fringe = 0
            End With
        Else
            With Map.Tile(CurX, CurY)
                .Type = 0
                .Data1 = 0
                .Data2 = 0
                .Data3 = 0
            End With
        End If
    End If

End Sub

Public Sub MapEditorChooseTile(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
        EditorTileX = x \ PIC_X
        EditorTileY = y \ PIC_Y
        
        frmMainGame.shpSelected.Top = EditorTileY * PIC_Y
        frmMainGame.shpSelected.Left = EditorTileX * PIC_Y
        
        Call BltMapEditorTilePreview
    End If
End Sub

Public Sub MapEditorTileScroll()
    frmMainGame.picBackSelect.Top = (frmMainGame.scrlPicture.Value * PIC_Y) * -1
End Sub

Public Sub MapEditorSend()
    Call SendMap
    Call MapEditorCancel
End Sub

Public Sub MapEditorCancel()
    Call LoadMap(GetPlayerMap(MyIndex))
    InEditor = False
    frmMainGame.picMapEditor.Visible = False
    Set DDS_Misc = Nothing
    
    frmMainGame.Width = 8565
End Sub

Public Sub MapEditorClearLayer()
Dim x As Long
Dim y As Long

    ' Ground layer
    If frmMainGame.optGround.Value Then
        If MsgBox("Você quer mesmo limpar o mapa?", vbYesNo, GAME_NAME) = vbYes Then
            For x = 0 To MAX_MAPX
                For y = 0 To MAX_MAPY
                    Map.Tile(x, y).Ground = 0
                Next
            Next
        End If
    End If

    ' Mask layer
    If frmMainGame.optMask.Value Then
        If MsgBox("Você quer mesmo limpar o mapa?", vbYesNo, GAME_NAME) = vbYes Then
            For x = 0 To MAX_MAPX
                For y = 0 To MAX_MAPY
                    Map.Tile(x, y).Mask = 0
                Next
            Next
        End If
    End If

    ' Animation layer
    If frmMainGame.optAnim.Value Then
        If MsgBox("Você quer mesmo limpar o mapa?", vbYesNo, GAME_NAME) = vbYes Then
            For x = 0 To MAX_MAPX
                For y = 0 To MAX_MAPY
                    Map.Tile(x, y).Anim = 0
                Next
            Next
        End If
    End If

    ' Fringe layer
    If frmMainGame.optFringe.Value Then
        If MsgBox("Você quer mesmo limpar o mapa?", vbYesNo, GAME_NAME) = vbYes Then
            For x = 0 To MAX_MAPX
                For y = 0 To MAX_MAPY
                    Map.Tile(x, y).Fringe = 0
                Next
            Next
        End If
    End If
    
End Sub

Public Sub MapEditorFillLayer()
Dim x As Long
Dim y As Long

    ' Ground layer
    If frmMainGame.optGround.Value Then
        If MsgBox("Você quer mesmo pintar o mapa?", vbYesNo, GAME_NAME) = vbYes Then
            For x = 0 To MAX_MAPX
                For y = 0 To MAX_MAPY
                    Map.Tile(x, y).Ground = EditorTileY * TILESHEET_WIDTH + EditorTileX
                Next
            Next
        End If
    End If

    ' Mask layer
    If frmMainGame.optMask.Value Then
        If MsgBox("Você quer mesmo pintar o mapa?", vbYesNo, GAME_NAME) = vbYes Then
            For x = 0 To MAX_MAPX
                For y = 0 To MAX_MAPY
                    Map.Tile(x, y).Mask = EditorTileY * TILESHEET_WIDTH + EditorTileX
                Next
            Next
        End If
    End If

    ' Animation layer
    If frmMainGame.optAnim.Value Then
        If MsgBox("Você quer mesmo limpar o mapa?", vbYesNo, GAME_NAME) = vbYes Then
            For x = 0 To MAX_MAPX
                For y = 0 To MAX_MAPY
                    Map.Tile(x, y).Anim = EditorTileY * TILESHEET_WIDTH + EditorTileX
                Next
            Next
        End If
    End If

    ' Fringe layer
    If frmMainGame.optFringe.Value Then
        If MsgBox("Você quer mesmo pintar o mapa?", vbYesNo, GAME_NAME) = vbYes Then
            For x = 0 To MAX_MAPX
                For y = 0 To MAX_MAPY
                    Map.Tile(x, y).Fringe = EditorTileY * TILESHEET_WIDTH + EditorTileX
                Next
            Next
        End If
    End If
    
End Sub

Public Sub MapEditorClearAttribs()
Dim x As Long
Dim y As Long
    
    If MsgBox("Você quer mesmo limpar o mapa?", vbYesNo, GAME_NAME) = vbYes Then
        For x = 0 To MAX_MAPX
            For y = 0 To MAX_MAPY
                Map.Tile(x, y).Type = 0
            Next
        Next
    End If
End Sub

Public Sub MapEditorLeaveMap()
     If InEditor Then
        If MsgBox("Salvar alterações no mapa?", vbYesNo) = vbYes Then
            Call MapEditorSend
        Else
            Call MapEditorCancel
        End If
    End If
End Sub

' /////////////////
' // Item Editor //
' /////////////////

Public Sub ItemEditorInit()
  
    frmItemEditor.txtName.Text = Trim$(Item(EditorIndex).Name)
    frmItemEditor.scrlPic.Value = Item(EditorIndex).Pic
    'frmItemEditor.cmbType.ListIndex = Item(EditorIndex).Type
    
    'If (frmItemEditor.cmbType.ListIndex >= ITEM_TYPE_WEAPON) And (frmItemEditor.cmbType.ListIndex <= ITEM_TYPE_SHIELD) Then
    '    frmItemEditor.fraEquipment.Visible = True
    '    frmItemEditor.scrlDurability.Value = Item(EditorIndex).Data1
    '    frmItemEditor.scrlStrength.Value = Item(EditorIndex).Data2
    'Else
    '    frmItemEditor.fraEquipment.Visible = False
    'End If
    
    'If (frmItemEditor.cmbType.ListIndex >= ITEM_TYPE_POTIONADDHP) And (frmItemEditor.cmbType.ListIndex <= ITEM_TYPE_POTIONSUBSP) Then
    '    frmItemEditor.fraVitals.Visible = True
    '    frmItemEditor.scrlVitalMod.Value = Item(EditorIndex).Data1
    'Else
    '    frmItemEditor.fraVitals.Visible = False
    'End If
    
    'If (frmItemEditor.cmbType.ListIndex = ITEM_TYPE_SPELL) Then
    '    frmItemEditor.fraSpell.Visible = True
    '    frmItemEditor.scrlSpell.Value = Item(EditorIndex).Data1
    'Else
    '    frmItemEditor.fraSpell.Visible = False
    'End If
    
    frmItemEditor.scrlPic.Max = (DDSD_Item.lHeight \ PIC_Y) - 1
    
    Call ItemEditorBltItem
    
    frmItemEditor.Show vbModal
    
End Sub

Public Sub ItemEditorOk()

    Item(EditorIndex).Name = frmItemEditor.txtName.Text
    Item(EditorIndex).Pic = frmItemEditor.scrlPic.Value
    Item(EditorIndex).Type = 0
    
    Call SendSaveItem(EditorIndex)
    
    Editor = 0
    Unload frmItemEditor
    
End Sub

Public Sub ItemEditorCancel()
    Editor = 0
    Unload frmItemEditor
End Sub

' ////////////////
' // Npc Editor //
' ////////////////

Public Sub NpcEditorInit()

End Sub

Public Sub NpcEditorOk()

End Sub

Public Sub NpcEditorCancel()

End Sub

' /////////////////
' // Shop Editor //
' /////////////////

Public Sub ShopEditorInit()

End Sub

Public Sub UpdateShopTrade()

End Sub

Public Sub ShopEditorOk()

End Sub

Public Sub ShopEditorCancel()

End Sub

' //////////////////
' // Spell Editor //
' //////////////////

Public Sub SpellEditorInit()

End Sub

Public Sub SpellEditorOk()

End Sub

Public Sub SpellEditorCancel()

End Sub

