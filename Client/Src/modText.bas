Attribute VB_Name = "modText"
Option Explicit

' ******************************************
' **              BMO Source              **
' ******************************************

' Used to set a font for GDI text drawing
Public Sub SetFont(ByVal Font As String, ByVal Size As Byte)
    GameFont = CreateFont(Size, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, Font)
End Sub

' GDI text drawing onto buffer
Public Sub DrawText(ByVal hdc As Long, ByVal x, ByVal y, ByVal Text As String, Color As Long, Optional ByVal NotGame As Boolean = False)
    
    ' Selects an object into the specified DC.
    ' The new object replaces the previous object of the same type.
    Call SelectObject(hdc, GameFont)
    
    ' Sets the background mix mode of the specified DC.
    Call SetBkMode(hdc, vbTransparent)
    
    ' color of text drop shadow
    If Not (Color = QBColor(Black)) Then
        Call SetTextColor(hdc, RGB(0, 0, 0))
    Else
        Call SetTextColor(hdc, RGB(255, 255, 255))
    End If
    
    If Not NotGame Then
        ' Draw name
        If x <= 1 Then
            x = 1
        End If
        
        If y <= 1 Then
            y = 1
        End If
        
        If x + ((Len(Text) * FONT_WIDTH)) >= ((MAX_MAPX + 1) * PIC_X) Then
            x = (frmMainGame.picScreen.width) - ((Len(Text) * FONT_WIDTH) + 1)
        End If
    End If
    
    ' draw with offset
    'Call TextOut(hdc, x - 2, y - 2, Text, Len(Text))
    'Call TextOut(hdc, x + 2, y + 2, Text, Len(Text))
    Call TextOut(hdc, x, y + 1, Text, Len(Text))
    Call TextOut(hdc, x + 1, y, Text, Len(Text))
    Call TextOut(hdc, x, y - 1, Text, Len(Text))
    Call TextOut(hdc, x - 1, y, Text, Len(Text))
    Call TextOut(hdc, x, y + 2, Text, Len(Text))
    Call TextOut(hdc, x + 2, y, Text, Len(Text))
    Call TextOut(hdc, x, y - 2, Text, Len(Text))
    Call TextOut(hdc, x - 2, y, Text, Len(Text))
    
    ' draw text with color
    Call SetTextColor(hdc, Color)
    Call TextOut(hdc, x, y, Text, Len(Text))
    
End Sub

Sub DrawPlayerName(ByVal Index As Long)
Dim TextX As Long
Dim TextY As Long

    ' Determine location for text
    TextX = ((GetPlayerX(Index) * PIC_X + Player(Index).XOffset) + 16) - ((Len(GetPlayerName(Index)) * FONT_WIDTH) / 2)
    TextY = GetPlayerY(Index) * PIC_Y + Player(Index).YOffset - (PIC_Y \ 2) - 22
    
    ' Draw name
    Call DrawText(TexthDC, TextX, TextY, GetPlayerName(Index), QBColor(Black))
End Sub
Sub DrawPlayerLevel(ByVal Index As Long)
Dim TextX As Long
Dim TextY As Long
Dim Rank As String

If Player(Index).Kills * 5 >= 0 Then
Rank = "Bombinha de aguá"
End If

If Player(Index).Kills * 5 > 20 Then
Rank = "Bomba de aguá"
End If

If Player(Index).Kills * 5 > 50 Then
Rank = "Super bomba de aguá"
End If

If Player(Index).Kills * 5 > 120 Then
Rank = "Bombinha de terra"
End If

If Player(Index).Kills * 5 > 240 Then
Rank = "Bomba de terra"
End If

If Player(Index).Kills * 5 > 500 Then
Rank = "Super bomba de terra"
End If

If Player(Index).Kills * 5 > 1200 Then
Rank = "Bombinha de vento"
End If

If Player(Index).Kills * 5 > 2600 Then
Rank = "Bomba de vento"
End If

If Player(Index).Kills * 5 > 5400 Then
Rank = "Super bomba de vento"
End If

If Player(Index).Kills * 5 > 11000 Then
Rank = "Bombinha de fogo"
End If

If Player(Index).Kills * 5 > 23000 Then
Rank = "Bomba de fogo"
End If

If Player(Index).Kills * 5 > 50000 Then
Rank = "Super bomba de fogo"
End If


    ' Determine location for text
    If Rank = "Bomba de aguá" Then
    TextX = ((GetPlayerX(Index) * PIC_X + Player(Index).XOffset) - 20) - ((Len(GetPlayerName(Index)) * FONT_WIDTH) / 2)
    TextY = GetPlayerY(Index) * PIC_Y + Player(Index).YOffset - (PIC_Y \ 2) - 37
    End If
    If Rank = "Bomba de terra" Then
    TextX = ((GetPlayerX(Index) * PIC_X + Player(Index).XOffset) - 20) - ((Len(GetPlayerName(Index)) * FONT_WIDTH) / 2)
    TextY = GetPlayerY(Index) * PIC_Y + Player(Index).YOffset - (PIC_Y \ 2) - 37
    End If
    If Rank = "Bomba de vento" Then
    TextX = ((GetPlayerX(Index) * PIC_X + Player(Index).XOffset) - 20) - ((Len(GetPlayerName(Index)) * FONT_WIDTH) / 2)
    TextY = GetPlayerY(Index) * PIC_Y + Player(Index).YOffset - (PIC_Y \ 2) - 37
    End If
    If Rank = "Bomba de fogo" Then
    TextX = ((GetPlayerX(Index) * PIC_X + Player(Index).XOffset) - 20) - ((Len(GetPlayerName(Index)) * FONT_WIDTH) / 2)
    TextY = GetPlayerY(Index) * PIC_Y + Player(Index).YOffset - (PIC_Y \ 2) - 37
    End If
    
    
    If Rank = "Bombinha de aguá" Then
    TextX = ((GetPlayerX(Index) * PIC_X + Player(Index).XOffset) - 30) - ((Len(GetPlayerName(Index)) * FONT_WIDTH) / 2)
    TextY = GetPlayerY(Index) * PIC_Y + Player(Index).YOffset - (PIC_Y \ 2) - 37
    End If
    If Rank = "Bombinha de terra" Then
    TextX = ((GetPlayerX(Index) * PIC_X + Player(Index).XOffset) - 30) - ((Len(GetPlayerName(Index)) * FONT_WIDTH) / 2)
    TextY = GetPlayerY(Index) * PIC_Y + Player(Index).YOffset - (PIC_Y \ 2) - 37
    End If
    If Rank = "Bombinha de vento" Then
    TextX = ((GetPlayerX(Index) * PIC_X + Player(Index).XOffset) - 30) - ((Len(GetPlayerName(Index)) * FONT_WIDTH) / 2)
    TextY = GetPlayerY(Index) * PIC_Y + Player(Index).YOffset - (PIC_Y \ 2) - 37
    End If
    If Rank = "Bombinha de fogo" Then
    TextX = ((GetPlayerX(Index) * PIC_X + Player(Index).XOffset) - 30) - ((Len(GetPlayerName(Index)) * FONT_WIDTH) / 2)
    TextY = GetPlayerY(Index) * PIC_Y + Player(Index).YOffset - (PIC_Y \ 2) - 37
    End If
    
    
    If Rank = "Super bomba de aguá" Then
    TextX = ((GetPlayerX(Index) * PIC_X + Player(Index).XOffset) - 35) - ((Len(GetPlayerName(Index)) * FONT_WIDTH) / 2)
    TextY = GetPlayerY(Index) * PIC_Y + Player(Index).YOffset - (PIC_Y \ 2) - 37
    End If
    If Rank = "Super bomba de terra" Then
    TextX = ((GetPlayerX(Index) * PIC_X + Player(Index).XOffset) - 35) - ((Len(GetPlayerName(Index)) * FONT_WIDTH) / 2)
    TextY = GetPlayerY(Index) * PIC_Y + Player(Index).YOffset - (PIC_Y \ 2) - 37
    End If
    If Rank = "Super bomba de vento" Then
    TextX = ((GetPlayerX(Index) * PIC_X + Player(Index).XOffset) - 35) - ((Len(GetPlayerName(Index)) * FONT_WIDTH) / 2)
    TextY = GetPlayerY(Index) * PIC_Y + Player(Index).YOffset - (PIC_Y \ 2) - 37
    End If
    If Rank = "Super bomba de fogo" Then
    TextX = ((GetPlayerX(Index) * PIC_X + Player(Index).XOffset) - 35) - ((Len(GetPlayerName(Index)) * FONT_WIDTH) / 2)
    TextY = GetPlayerY(Index) * PIC_Y + Player(Index).YOffset - (PIC_Y \ 2) - 37
    End If
   '' End If
    
    
    ' Draw name
    Call DrawText(TexthDC, TextX, TextY, Rank, QBColor(Black))
End Sub
