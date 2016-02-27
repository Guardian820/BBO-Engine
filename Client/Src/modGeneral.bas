Attribute VB_Name = "modGeneral"
Option Explicit

Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpfilename As String) As Long

' ******************************************
' **              BMO Source              **
' ******************************************

Public Sub Main()
Dim i As Long

    frmMain.height = Main_OrigHeight
    frmMain.picLoading.Visible = False
    frmMain.picMainMenu.Visible = False
    frmMain.picRegister.Visible = False
    frmMain.picLogin.Visible = False
    
    Current_SpriteNum = 1
    
    frmMain.Show
    frmMain.picLoading.Visible = True
    
    RegisterWhite = True
    RegisterSprite = 1
    RegisterAnim = 1
    
    Call SetStatus("Loading...")
    
    Load frmMainGame
    
    'Call SetStatus("Initializing TCP settings...")
    Call TcpInit
    
    'Call SetStatus("Initializing DirectX...")
    
    ' check for early binding
    If DX7 Is Nothing Then
        Set DX7 = New DirectX7  ' late binding
    End If
    
    vbQuote = ChrW$(34) ' "
    
    GettingMap = True
    
    ' randomize rnd's seed
    Randomize
    
    Call CheckTiles
    ' initialize DirectX in the background after the form appears
    Call InitDirectDraw
    Call InitSurfaces ' Initialize all needed in-game surfaces
    Call InitDirectSound
    Call InitDirectMusic
    
    DirectMusic_PlayMidi "title.mid"
    
    frmMain.Visible = True
    frmMain.picMainMenu.Visible = True
    frmMain.picLoading.Visible = False
    
    Call SetFont(FONT_NAME, FONT_SIZE)
    Walk_Speed = 4
    Run_Speed = 4
    CheaterLoop
    
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

Public Sub CheckTiles()
Dim i As Long

    i = 1

    While FileExist(GFX_PATH & "tiles" & i & GFX_EXT)
        MAX_TILESETS = MAX_TILESETS + 1
        i = i + 1
    Wend

    frmMainGame.scrlTileSet.Max = MAX_TILESETS

End Sub

Public Sub RegisterBlt()
Dim sRECT As DxVBLib.RECT
Dim dRECT As DxVBLib.RECT

    sRECT.Top = RegisterSprite * 64
    sRECT.Bottom = sRECT.Top + 64
    sRECT.Left = (8 + RegisterAnim) * 64
    sRECT.Right = sRECT.Left + 64
    
    dRECT.Top = 0
    dRECT.Bottom = 64
    dRECT.Left = 0
    dRECT.Right = 64
    
    Engine_BltToDC DDS_Sprite, sRECT, dRECT, frmMain.picDisplay
    
End Sub
Public Sub LobbyBlt()
Dim sRECT As DxVBLib.RECT
Dim dRECT As DxVBLib.RECT
If frmLobby.Visible = True Then
    sRECT.Top = (GetPlayerSprite(MyIndex)) * 64
    sRECT.Bottom = sRECT.Top + 64
    sRECT.Left = (8 + LobbyAnim) * 64
    sRECT.Right = sRECT.Left + 64
    
    dRECT.Top = 0
    dRECT.Bottom = 64
    dRECT.Left = 0
    dRECT.Right = 64
    
    Engine_BltToDC DDS_Sprite, sRECT, dRECT, frmLobby.picMyChar
    End If
End Sub
Public Sub WaitingBlt1()
Dim sRECT As DxVBLib.RECT
Dim dRECT As DxVBLib.RECT
If frmWaiting.Visible = True Then
If frmWaiting.lstSprites.List(0) = vbNullString Then Exit Sub
    sRECT.Top = frmWaiting.lstSprites.List(0) * 64
    sRECT.Bottom = sRECT.Top + 64
    sRECT.Left = (8 + WaitAnim1) * 64
    sRECT.Right = sRECT.Left + 64
    
    dRECT.Top = 0
    dRECT.Bottom = 64
    dRECT.Left = 0
    dRECT.Right = 64
    
    Engine_BltToDC DDS_Sprite, sRECT, dRECT, frmWaiting.picRoomChar1
    End If
End Sub
Public Sub WaitingBlt2()
Dim sRECT As DxVBLib.RECT
Dim dRECT As DxVBLib.RECT
If frmWaiting.Visible = True Then
If frmWaiting.lstSprites.List(1) = vbNullString Then Exit Sub
    sRECT.Top = frmWaiting.lstSprites.List(1) * 64
    sRECT.Bottom = sRECT.Top + 64
    sRECT.Left = (8 + WaitAnim2) * 64
    sRECT.Right = sRECT.Left + 64
    
    dRECT.Top = 0
    dRECT.Bottom = 64
    dRECT.Left = 0
    dRECT.Right = 64
    
    Engine_BltToDC DDS_Sprite, sRECT, dRECT, frmWaiting.picRoomChar2
    End If
End Sub
Public Sub WaitingBlt3()
Dim sRECT As DxVBLib.RECT
Dim dRECT As DxVBLib.RECT
If frmWaiting.Visible = True Then
If frmWaiting.lstSprites.List(2) = vbNullString Then Exit Sub
    sRECT.Top = frmWaiting.lstSprites.List(2) * 64
    sRECT.Bottom = sRECT.Top + 64
    sRECT.Left = (8 + WaitAnim3) * 64
    sRECT.Right = sRECT.Left + 64
    
    dRECT.Top = 0
    dRECT.Bottom = 64
    dRECT.Left = 0
    dRECT.Right = 64
    
    Engine_BltToDC DDS_Sprite, sRECT, dRECT, frmWaiting.picRoomChar3
    End If
End Sub
Public Sub WaitingBlt4()
Dim sRECT As DxVBLib.RECT
Dim dRECT As DxVBLib.RECT
If frmWaiting.Visible = True Then
If frmWaiting.lstSprites.List(3) = vbNullString Then Exit Sub
    sRECT.Top = frmWaiting.lstSprites.List(3) * 64
    sRECT.Bottom = sRECT.Top + 64
    sRECT.Left = (8 + Waitanim4) * 64
    sRECT.Right = sRECT.Left + 64
    
    dRECT.Top = 0
    dRECT.Bottom = 64
    dRECT.Left = 0
    dRECT.Right = 64
    
    Engine_BltToDC DDS_Sprite, sRECT, dRECT, frmWaiting.picRoomChar4
    End If
End Sub

Public Sub MenuState(ByVal State As Long)

    frmMain.picLoading.Visible = True
    
    Select Case State
    
        Case MENU_STATE_NEWACCOUNT
            frmMain.picRegister.Visible = False
            frmMain.height = Main_OrigHeight
            If ConnectToServer Then
                If Not FEMALE Then
                    Call SendNewAccount(frmMain.txtRegisterName.Text, frmMain.txtRegisterPass.Text, SEX_MALE, RegisterSprite)
                Else
                    Call SendNewAccount(frmMain.txtRegisterName.Text, frmMain.txtRegisterPass.Text, SEX_FEMALE, RegisterSprite)
                End If
            End If
       
        Case MENU_STATE_LOGIN
            frmMain.picLogin.Visible = False
            If ConnectToServer Then
                Call SendLogin(frmMain.txtName.Text, frmMain.txtPassword.Text)
                Exit Sub
            End If

    End Select

    If Not IsConnected Then
        frmMain.picMainMenu.Visible = True
        frmMain.picLoading.Visible = False
        frmMsg.Visible = True
        frmMsg.lblAtenção.Caption = "Atenção"
        frmMsg.lblMsg.Caption = "Não foi possível se conectar ao servidor."
    End If
    
End Sub

Public Function WindowEnumerator(ByVal app_hwnd As Long, ByVal lparam As Long) As Long
Dim buf As String * 256
Dim Title As String
Dim length As Long

    ' Get the window's title.
    length = GetWindowText(app_hwnd, buf, Len(buf))
    Title = Left$(buf, length)

    ' See if the title contains the target.
    If InStr(Title, TargetName) > 0 Then
        ' Save the hwnd and end the enumeration.
        TargetHwnd = app_hwnd
        WindowEnumerator = False
    Else
        ' Continue the enumeration.
        WindowEnumerator = True
    End If
End Function

Sub GameInit()

    frmMain.picLoading.Visible = False
    
    frmMain.Hide
    'frmMainGame.Show
    DirectMusic_StopMidi
    frmLobby.Show
    
    ' Set the focus
    Call SetFocusOnChat

    frmMainGame.picScreen.Visible = True
End Sub

Public Sub DestroyGame()
    ' break out of GameLoop
    InGame = False
    
    Call DestroyTCP
    
    'destroy objects in reverse order
    Call DestroyDirectMusic
    Call DestroyDirectSound
    Call DestroyDirectDraw
    
    ' destory DirectX7 master object
    If Not DX7 Is Nothing Then
        Set DX7 = Nothing
    End If
    
    Call UnloadAllForms
    End
End Sub

Public Sub UnloadAllForms()
    Dim frm As Form

    For Each frm In VB.Forms
        Unload frm
    Next
End Sub

Public Sub SetStatus(ByVal Caption As String)

End Sub

Public Sub AddText(ByVal Msg As String, ByVal Color As Integer)
Dim s As String
  
    s = vbNewLine & Msg
    'If frmMainGame.Visible Then
        With frmMainGame
            .txtChat.SelStart = Len(.txtChat.Text)
            .txtChat.SelColor = QBColor(Color)
            .txtChat.SelText = s
            .txtChat.SelStart = Len(.txtChat.Text) - 1
            ' Prevent players from name spoofing
            .txtChat.SelHangingIndent = 15
        End With
    'ElseIf frmLobby.Visible Then
        With frmLobby
            .txtChat.SelStart = Len(.txtChat.Text)
            .txtChat.SelColor = QBColor(Color)
            .txtChat.SelText = s
            .txtChat.SelStart = Len(.txtChat.Text) - 1
            ' Prevent players from name spoofing
            .txtChat.SelHangingIndent = 15
        End With
    'End If
            With frmWaiting
            .txtChat.SelStart = Len(.txtChat.Text)
            .txtChat.SelColor = QBColor(Color)
            .txtChat.SelText = s
            .txtChat.SelStart = Len(.txtChat.Text) - 1
            ' Prevent players from name spoofing
            .txtChat.SelHangingIndent = 15
    End With
End Sub

' Used for adding text to packet debugger
Public Sub TextAdd(ByVal Txt As TextBox, Msg As String, NewLine As Boolean)
    If NewLine Then
        Txt.Text = Txt.Text + Msg + vbCrLf
    Else
        Txt.Text = Txt.Text + Msg
    End If
    
    Txt.SelStart = Len(Txt.Text) - 1
End Sub

Public Sub SetFocusOnChat()
On Error Resume Next 'prevent RTE5, no way to handle error
    If frmMainGame.Visible Then
        frmMainGame.txtMyChat.SetFocus
    ElseIf frmLobby.Visible Then
        frmLobby.txtEnterChat.SetFocus
    End If
End Sub

Public Function Rand(ByVal Low As Long, ByVal High As Long) As Long
  Rand = Int((High - Low + 1) * Rnd) + Low
End Function

Public Sub GetPlayersOnMap()
Dim i As Long

    PlayersOnMapHighIndex = 1

    ReDim PlayersOnMap(1 To MAX_PLAYERS)
        
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            If GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
                PlayersOnMap(PlayersOnMapHighIndex) = i
                PlayersOnMapHighIndex = PlayersOnMapHighIndex + 1
            End If
        End If
    Next
    
    ' Subtract 1 to prevent subscript out of range, why? because we have one array that starts
    '' at 0, and another that starts at 1
    PlayersOnMapHighIndex = PlayersOnMapHighIndex - 1
    
End Sub

Public Function isLoginLegal(ByVal Username As String, ByVal Password As String) As Boolean
    If LenB(Trim$(Username)) >= 3 Then
        If LenB(Trim$(Password)) >= 3 Then
            isLoginLegal = True
        End If
    End If
End Function

Public Function isStringLegal(ByVal sInput As String) As Boolean
Dim i As Long

    ' Prevent high ascii chars
    For i = 1 To Len(sInput)
        If Asc(Mid$(sInput, i, 1)) < vbKeySpace Or Asc(Mid$(sInput, i, 1)) > vbKeyF15 Then
        frmMsg.Visible = True
        frmMsg.lblAtenção.Caption = "Atenção"
        frmMsg.lblMsg.Caption = "Não são permitidos caracteres especiais."
            Exit Function
        End If
    Next
    
    isStringLegal = True
        
End Function

Function SortColumn(ByVal ListViewControl As MSComctlLib.ListView, ColumnIndex As Integer, SortType As Integer, SortOrder As Integer) As Boolean

    Dim x As Integer, y As Integer
    'On Error GoTo ErrHandler
    


    Select Case SortType
        
        '*** Alphanumeric sort
        Case sortAlphanumeric


        DoSort ListViewControl, SortOrder, ColumnIndex - 1
            
            '*** Numeric Sort
            Case sortNumeric
            Dim strMax As String, strNew As String
            
            'Find the longest (whole) number string
            '     length in the column

            If ColumnIndex > 1 Then

                For x = 1 To ListViewControl.ListItems.Count


                    If Len(ListViewControl.ListItems(x).ListSubItems(ColumnIndex - 1)) <> 0 Then 'ignores 0 length strings


                        If Len(CStr(ListViewControl.ListItems(x).ListSubItems(ColumnIndex - 1))) > Len(strMax) Then
                            strMax = CStr(ListViewControl.ListItems(x).SubItems(ColumnIndex - 1))
                        End If

                    End If

                Next

            Else


                For x = 1 To ListViewControl.ListItems.Count


                    If Len(ListViewControl.ListItems(x)) <> 0 Then

                        If Len(CStr(Int(ListViewControl.ListItems(x)))) > Len(strMax) Then
                            strMax = CStr(Int(ListViewControl.ListItems(x)))
                        End If

                    End If

                Next

            End If

            
            'hide the control - speeds up the sort
            ListViewControl.Visible = False
            

            If ColumnIndex > 1 Then

                For x = 1 To ListViewControl.ListItems.Count


                    If Len(ListViewControl.ListItems(x).ListSubItems(ColumnIndex - 1)) = 0 Then
                        ListViewControl.ListItems(x).ListSubItems(ColumnIndex - 1) = "0" 'make 0 length strings = To "0"
                    ElseIf Len(CStr(ListViewControl.ListItems(x).ListSubItems(ColumnIndex - 1))) < Len(strMax) Then
                        'prefix all numbers with 0's as required
                        '
                        strNew = ListViewControl.ListItems(x).ListSubItems(ColumnIndex - 1)


                        For y = 1 To Len(strMax) - Len(CStr(ListViewControl.ListItems(x).ListSubItems(ColumnIndex - 1)))
                            strNew = "0" & strNew
                        Next

                        ListViewControl.ListItems(x).ListSubItems(ColumnIndex - 1) = strNew
                    End If

                Next

            Else


                For x = 1 To ListViewControl.ListItems.Count


                    If Len(ListViewControl.ListItems(x).Text) = 0 Then
                        ListViewControl.ListItems(x).Text = "0" 'make 0 length strings = To "0"
                    ElseIf Len(CStr(Int(ListViewControl.ListItems(x)))) < Len(strMax) Then
                        'prefix all numbers with 0's as required
                        '
                        strNew = ListViewControl.ListItems(x).Text


                        For y = 1 To Len(strMax) - Len(CStr(Int(ListViewControl.ListItems(x))))
                            strNew = "0" & strNew
                        Next

                        ListViewControl.ListItems(x).Text = strNew
                    End If

                Next

            End If

            

            DoSort ListViewControl, SortOrder, ColumnIndex - 1
                


                If ColumnIndex > 1 Then
                    'Remove preceding 0's

                    For x = 1 To ListViewControl.ListItems.Count
                        ListViewControl.ListItems(x).ListSubItems(ColumnIndex - 1) = CDbl(ListViewControl.ListItems(x).ListSubItems(ColumnIndex - 1))
                        If ListViewControl.ListItems(x).ListSubItems(ColumnIndex - 1) = 0 Then ListViewControl.ListItems(x).ListSubItems(ColumnIndex - 1) = ""
                    Next

                Else
                    'Remove preceding 0's

                    For x = 1 To ListViewControl.ListItems.Count
                        ListViewControl.ListItems(x).Text = CDbl(ListViewControl.ListItems(x).Text)
                        If ListViewControl.ListItems(x).Text = 0 Then ListViewControl.ListItems(x).Text = ""
                    Next

                End If

                ListViewControl.Visible = True
                
                '*** Date Sort
                Case sortDate
                ListViewControl.Visible = False

                If ColumnIndex > 1 Then
                    'Convert dates to format that can be sor
                    '     ted alphanumerically

                    For x = 1 To ListViewControl.ListItems.Count
                        ListViewControl.ListItems(x).ListSubItems(ColumnIndex - 1) = Format(ListViewControl.ListItems(x).ListSubItems(ColumnIndex - 1), "YYYY MM DD hh:mm:ss")
                    Next


                    DoSort ListViewControl, SortOrder, ColumnIndex - 1
                        'Convert dates back to General Date form
                        '     at

                        For x = 1 To ListViewControl.ListItems.Count
                            ListViewControl.ListItems(x).ListSubItems(ColumnIndex - 1) = Format(ListViewControl.ListItems(x).ListSubItems(ColumnIndex - 1), "General Date")
                        Next

                    Else
                        'Convert dates to format that can be sor
                        '     ted alphanumerically

                        For x = 1 To ListViewControl.ListItems.Count
                            ListViewControl.ListItems(x).Text = Format(ListViewControl.ListItems(x).Text, "YYYY MM DD hh:mm:ss")
                        Next


                        DoSort ListViewControl, SortOrder, ColumnIndex - 1
                            'Convert dates back to General Date form
                            '     at

                            For x = 1 To ListViewControl.ListItems.Count
                                ListViewControl.ListItems(x).Text = Format(ListViewControl.ListItems(x).Text, "General Date")
                            Next

                            
                        End If

                        
                        ListViewControl.Visible = True
                    End Select

                SortColumn = True
                
Exit_Function:
                Exit Function
                
ErrHandler:
                MsgBox Err.Description & " (" & Err.Number & ")", vbOKOnly + vbCritical, "ListView Sort module Error"
                SortColumn = False
                Resume Exit_Function
            End Function
            
Private Sub DoSort(ByVal ListViewControl As MSComctlLib.ListView, SortOrder As Integer, SortKey As Integer)


    If SortOrder = sortAscending Then
        ListViewControl.SortOrder = lvwAscending
    ElseIf SortOrder = sortDescending Then
        ListViewControl.SortOrder = lvwDescending
    End If

    ListViewControl.SortKey = SortKey
    ListViewControl.Sorted = True
End Sub
