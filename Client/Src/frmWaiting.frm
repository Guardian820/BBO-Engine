VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Begin VB.Form frmWaiting 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bomber ! Bomber ! Online.: Esperando..."
   ClientHeight    =   7500
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6960
   Icon            =   "frmWaiting.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmWaiting.frx":0CCA
   ScaleHeight     =   7500
   ScaleWidth      =   6960
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Timer Timer2 
      Interval        =   500
      Left            =   3480
      Top             =   240
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   3000
      Top             =   240
   End
   Begin VB.Timer tmrAnim2 
      Interval        =   200
      Left            =   1560
      Top             =   240
   End
   Begin VB.Timer tmrAnim3 
      Interval        =   200
      Left            =   2040
      Top             =   240
   End
   Begin VB.Timer tmrAnim4 
      Interval        =   200
      Left            =   2520
      Top             =   240
   End
   Begin VB.Timer tmrAnim1 
      Interval        =   200
      Left            =   1080
      Top             =   240
   End
   Begin VB.ListBox lstSprites 
      Appearance      =   0  'Flat
      Height          =   420
      Left            =   840
      TabIndex        =   18
      Top             =   8640
      Width           =   1815
   End
   Begin VB.ListBox lstInfos 
      Appearance      =   0  'Flat
      Height          =   615
      Left            =   3480
      TabIndex        =   17
      Top             =   7680
      Width           =   2295
   End
   Begin VB.PictureBox picRoomChar4 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   960
      Left            =   4800
      ScaleHeight     =   960
      ScaleWidth      =   960
      TabIndex        =   8
      Top             =   2480
      Width           =   960
   End
   Begin VB.PictureBox picRoomChar3 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   960
      Left            =   3540
      ScaleHeight     =   960
      ScaleWidth      =   960
      TabIndex        =   7
      Top             =   2480
      Width           =   960
   End
   Begin VB.PictureBox picRoomChar2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   960
      Left            =   2280
      ScaleHeight     =   960
      ScaleWidth      =   960
      TabIndex        =   6
      Top             =   2480
      Width           =   960
   End
   Begin VB.PictureBox picRoomChar1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   960
      Left            =   1020
      ScaleHeight     =   960
      ScaleWidth      =   960
      TabIndex        =   5
      Top             =   2480
      Width           =   960
   End
   Begin VB.Timer tmrAtualizador 
      Interval        =   5000
      Left            =   600
      Top             =   240
   End
   Begin VB.TextBox txtMyChat 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   210
      Left            =   600
      TabIndex        =   4
      Top             =   6440
      Width           =   6165
   End
   Begin VB.Timer tmrGameStart 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   120
      Top             =   240
   End
   Begin VB.ListBox lstPlayers 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   720
      TabIndex        =   1
      Top             =   7680
      Width           =   2175
   End
   Begin RichTextLib.RichTextBox txtChat 
      Height          =   1770
      Left            =   120
      TabIndex        =   3
      Top             =   4560
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   3122
      _Version        =   393217
      BackColor       =   4210752
      BorderStyle     =   0
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmWaiting.frx":AABCC
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblKD3 
      BackStyle       =   0  'Transparent
      Caption         =   "K/D:"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3480
      TabIndex        =   26
      Top             =   1980
      Width           =   615
   End
   Begin VB.Label lblKD4 
      BackStyle       =   0  'Transparent
      Caption         =   "K/D:"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4740
      TabIndex        =   25
      Top             =   1980
      Width           =   615
   End
   Begin VB.Label lblKD2 
      BackStyle       =   0  'Transparent
      Caption         =   "K/D:"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2240
      TabIndex        =   24
      Top             =   1980
      Width           =   615
   End
   Begin VB.Label lblKD1 
      BackStyle       =   0  'Transparent
      Caption         =   "K/D:"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   23
      Top             =   1980
      Width           =   615
   End
   Begin VB.Label lblAdd4 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   4680
      TabIndex        =   22
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label lblAdd3 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   3480
      TabIndex        =   21
      Top             =   3480
      Width           =   1095
   End
   Begin VB.Label lblAdd2 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   2160
      TabIndex        =   20
      Top             =   3480
      Width           =   1095
   End
   Begin VB.Label lblAdd1 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   840
      TabIndex        =   19
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label lblInfoPlayer2 
      BackStyle       =   0  'Transparent
      Caption         =   "999/999"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2240
      TabIndex        =   16
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label lblInfoPlayer3 
      BackStyle       =   0  'Transparent
      Caption         =   "999/999"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3500
      TabIndex        =   15
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label lblInfoPlayer4 
      BackStyle       =   0  'Transparent
      Caption         =   "999/999"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4755
      TabIndex        =   14
      Top             =   2160
      Width           =   1050
   End
   Begin VB.Label lblInfoPlayer1 
      BackStyle       =   0  'Transparent
      Caption         =   "999/999"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   13
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label lblRoomPName3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Nomeblablabla"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3480
      TabIndex        =   12
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label lblRoomPName4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Nomeblablabla"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4720
      TabIndex        =   11
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label lblRoomPName2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Nomeblablabla"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2235
      TabIndex        =   10
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label lblRoomPName1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Nomeblablabla"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   960
      TabIndex        =   9
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      Top             =   3960
      Width           =   1575
   End
   Begin VB.Label lblWait 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Aguardando..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   6960
      Width           =   6975
   End
   Begin VB.Image imgRoomRank2 
      Height          =   480
      Left            =   2780
      Top             =   1920
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgRoomRank1 
      Height          =   480
      Left            =   1560
      Top             =   1920
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgRoomRank3 
      Height          =   480
      Left            =   4080
      Top             =   1920
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgRoomRank4 
      Height          =   480
      Left            =   5280
      Top             =   1920
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "frmWaiting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private timeleft As Long

Private Sub Form_Load()
    Dim result As Long
    timeleft = 3

    result = SetWindowLong(txtChat.hWnd, GWL_EXSTYLE, WS_EX_TRANSPARENT)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'frmLobby.txtChat.Text = vbNullString
    'frmLobby.Show
    'Me.Hide
    SendData CLeaveRoom & END_CHAR
End Sub

Private Sub Label1_Click()
    SendData CLeaveRoom & END_CHAR
    frmWaiting.imgRoomRank1.Picture = Nothing
    frmWaiting.imgRoomRank2.Picture = Nothing
    frmWaiting.imgRoomRank3.Picture = Nothing
    frmWaiting.imgRoomRank4.Picture = Nothing
End Sub

Private Sub lblAdd1_Click()
If Not lblRoomPName1.Caption = vbNullString Then
frmMsg.Visible = True
frmMsg.lblMsg.Caption = "Você adicionou" & lblRoomPName1.Caption & " !"
frmMsg.lblAtenção = "Parabéns !"
Call addfriend(lblRoomPName1.Caption)
End If
End Sub

Private Sub lblAdd2_Click()
If Not lblRoomPName2.Caption = vbNullString Then
frmMsg.Visible = True
frmMsg.lblMsg.Caption = "Você adicionou" & lblRoomPName2.Caption & " !"
frmMsg.lblAtenção = "Parabéns !"
Call addfriend(lblRoomPName2.Caption)
End If
End Sub

Private Sub lblAdd3_Click()
If Not lblRoomPName3.Caption = vbNullString Then
frmMsg.Visible = True
frmMsg.lblMsg.Caption = "Você adicionou" & lblRoomPName3.Caption & " !"
frmMsg.lblAtenção = "Parabéns !"
Call addfriend(lblRoomPName3.Caption)
End If
End Sub
Private Sub lblAdd4_Click()
If Not lblRoomPName4.Caption = vbNullString Then
frmMsg.Visible = True
frmMsg.lblMsg.Caption = "Você adicionou" & lblRoomPName4.Caption & " !"
frmMsg.lblAtenção = "Parabéns !"
Call addfriend(lblRoomPName4.Caption)
End If
End Sub

Private Sub Timer1_Timer()
    If Not frmWaiting.lstInfos.List(0) = vbNullString Then
    imgRoomRank1.Visible = True
    Else
    imgRoomRank1.Visible = False
    End If
    If Not frmWaiting.lstInfos.List(1) = vbNullString Then
    imgRoomRank2.Visible = True
    Else
    frmWaiting.imgRoomRank2.Visible = False
    End If
    If Not frmWaiting.lstInfos.List(2) = vbNullString Then
    imgRoomRank3.Visible = True
    Else
    imgRoomRank3.Visible = False
    End If
    If Not frmWaiting.lstInfos.List(3) = vbNullString Then
    imgRoomRank4.Visible = True
    Else
    imgRoomRank4.Visible = False
    End If
End Sub

Private Sub Timer2_Timer()
Dim p As Long
Dim q As Long
Dim l As Long


If frmWaiting.Visible = True Then
For l = 1 To PlayersOnMapHighIndex
If Player(PlayersOnMap(l)).Kills * 5 >= 0 Then
Player(PlayersOnMap(l)).Rank = "Bombinha de aguá"
End If

If Player(PlayersOnMap(l)).Kills * 5 > 20 Then
Player(PlayersOnMap(l)).Rank = "Bomba de aguá"
End If

If Player(PlayersOnMap(l)).Kills * 5 > 50 Then
Player(PlayersOnMap(l)).Rank = "Super bomba de aguá"
End If

If Player(PlayersOnMap(l)).Kills * 5 > 120 Then
Player(PlayersOnMap(l)).Rank = "Bombinha de terra"
End If

If Player(PlayersOnMap(l)).Kills * 5 > 240 Then
Player(PlayersOnMap(l)).Rank = "Bomba de terra"
End If

If Player(PlayersOnMap(l)).Kills * 5 > 500 Then
Player(PlayersOnMap(l)).Rank = "Super bomba de terra"
End If

If Player(PlayersOnMap(l)).Kills * 5 > 1200 Then
Player(PlayersOnMap(l)).Rank = "Bombinha de vento"
End If

If Player(PlayersOnMap(l)).Kills * 5 > 2600 Then
Player(PlayersOnMap(l)).Rank = "Bomba de vento"
End If

If Player(PlayersOnMap(l)).Kills * 5 > 5400 Then
Player(PlayersOnMap(l)).Rank = "Super bomba de vento"
End If

If Player(PlayersOnMap(l)).Kills * 5 > 11000 Then
Player(PlayersOnMap(l)).Rank = "Bombinha de fogo"
End If

If Player(PlayersOnMap(l)).Kills * 5 > 23000 Then
Player(PlayersOnMap(l)).Rank = "Bomba de fogo"
End If

If Player(PlayersOnMap(l)).Kills * 5 > 50000 Then
Player(PlayersOnMap(l)).Rank = "Super bomba de fogo"
End If
Next
    
    lstSprites.Clear
    lstInfos.Clear
    
    For p = 1 To PlayersOnMapHighIndex
    frmWaiting.lstSprites.AddItem Player(PlayersOnMap(p)).sprite
    frmWaiting.lstInfos.AddItem Player(PlayersOnMap(p)).Kills & "/" & Player(PlayersOnMap(p)).Deaths
    
    Next
    
    frmWaiting.lblInfoPlayer1.Caption = frmWaiting.lstInfos.List(0)
    frmWaiting.lblInfoPlayer2.Caption = frmWaiting.lstInfos.List(1)
    frmWaiting.lblInfoPlayer3.Caption = frmWaiting.lstInfos.List(2)
    frmWaiting.lblInfoPlayer4.Caption = frmWaiting.lstInfos.List(3)
    frmWaiting.lblRoomPName1.Caption = frmWaiting.lstPlayers.List(0)
    frmWaiting.lblRoomPName2.Caption = frmWaiting.lstPlayers.List(1)
    frmWaiting.lblRoomPName3.Caption = frmWaiting.lstPlayers.List(2)
    frmWaiting.lblRoomPName4.Caption = frmWaiting.lstPlayers.List(3)
    
        For q = 1 To PlayersOnMapHighIndex
    
    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(0) And Player(PlayersOnMap(q)).Rank = "Bombinha de aguá" Then
    frmWaiting.imgRoomRank1.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank1" & ".gif")
    End If
    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(0) And Player(PlayersOnMap(q)).Rank = "Bomba de aguá" Then
    frmWaiting.imgRoomRank1.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank2" & ".gif")
    End If
    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(0) And Player(PlayersOnMap(q)).Rank = "Super bomba de aguá" Then
    frmWaiting.imgRoomRank1.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank3" & ".gif")
    End If
    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(0) And Player(PlayersOnMap(q)).Rank = "Bombinha de terra" Then
    frmWaiting.imgRoomRank1.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank4" & ".gif")
    End If
    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(0) And Player(PlayersOnMap(q)).Rank = "Bomba de terra" Then
    frmWaiting.imgRoomRank1.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank5" & ".gif")
    End If
    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(0) And Player(PlayersOnMap(q)).Rank = "Super bomba de terra" Then
    frmWaiting.imgRoomRank1.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank6" & ".gif")
    End If
    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(0) And Player(PlayersOnMap(q)).Rank = "Bombinha de vento" Then
    frmWaiting.imgRoomRank1.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank7" & ".gif")
    End If
    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(0) And Player(PlayersOnMap(q)).Rank = "Bomba de vento" Then
    frmWaiting.imgRoomRank1.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank8" & ".gif")
    End If
    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(0) And Player(PlayersOnMap(q)).Rank = "Super bomba de vento" Then
    frmWaiting.imgRoomRank1.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank9" & ".gif")
    End If
    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(0) And Player(PlayersOnMap(q)).Rank = "Bombinha de fogo" Then
    frmWaiting.imgRoomRank1.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank10" & ".gif")
    End If
    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(0) And Player(PlayersOnMap(q)).Rank = "Bomba de fogo" Then
    frmWaiting.imgRoomRank1.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank11" & ".gif")
    End If
    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(0) And Player(PlayersOnMap(q)).Rank = "Super bomba de fogo" Then
    frmWaiting.imgRoomRank1.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank12" & ".gif")
    End If


    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(1) And Player(PlayersOnMap(q)).Rank = "Bombinha de aguá" Then
    frmWaiting.imgRoomRank2.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank1" & ".gif")
    End If
    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(1) And Player(PlayersOnMap(q)).Rank = "Bomba de aguá" Then
    frmWaiting.imgRoomRank2.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank2" & ".gif")
    End If
    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(1) And Player(PlayersOnMap(q)).Rank = "Super bomba de aguá" Then
    frmWaiting.imgRoomRank2.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank3" & ".gif")
    End If
    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(1) And Player(PlayersOnMap(q)).Rank = "Bombinha de terra" Then
    frmWaiting.imgRoomRank2.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank4" & ".gif")
    End If
    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(1) And Player(PlayersOnMap(q)).Rank = "Bomba de terra" Then
    frmWaiting.imgRoomRank2.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank5" & ".gif")
    End If
    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(1) And Player(PlayersOnMap(q)).Rank = "Super bomba de terra" Then
    frmWaiting.imgRoomRank2.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank6" & ".gif")
    End If
    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(1) And Player(PlayersOnMap(q)).Rank = "Bombinha de vento" Then
    frmWaiting.imgRoomRank2.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank7" & ".gif")
    End If
    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(1) And Player(PlayersOnMap(q)).Rank = "Bomba de vento" Then
    frmWaiting.imgRoomRank2.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank8" & ".gif")
    End If
    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(1) And Player(PlayersOnMap(q)).Rank = "Super bomba de vento" Then
    frmWaiting.imgRoomRank2.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank9" & ".gif")
    End If
    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(1) And Player(PlayersOnMap(q)).Rank = "Bombinha de fogo" Then
    frmWaiting.imgRoomRank2.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank10" & ".gif")
    End If
    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(1) And Player(PlayersOnMap(q)).Rank = "Bomba de fogo" Then
    frmWaiting.imgRoomRank2.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank11" & ".gif")
    End If
    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(1) And Player(PlayersOnMap(q)).Rank = "Super bomba de fogo" Then
    frmWaiting.imgRoomRank2.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank12" & ".gif")
    End If


    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(2) And Player(PlayersOnMap(q)).Rank = "Bombinha de aguá" Then
    frmWaiting.imgRoomRank3.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank1" & ".gif")
    End If
    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(2) And Player(PlayersOnMap(q)).Rank = "Bomba de aguá" Then
    frmWaiting.imgRoomRank3.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank2" & ".gif")
    End If
    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(2) And Player(PlayersOnMap(q)).Rank = "Super bomba de aguá" Then
    frmWaiting.imgRoomRank3.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank3" & ".gif")
    End If
    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(2) And Player(PlayersOnMap(q)).Rank = "Bombinha de terra" Then
    frmWaiting.imgRoomRank3.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank4" & ".gif")
    End If
    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(2) And Player(PlayersOnMap(q)).Rank = "Bomba de terra" Then
    frmWaiting.imgRoomRank3.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank5" & ".gif")
    End If
    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(2) And Player(PlayersOnMap(q)).Rank = "Super bomba de terra" Then
    frmWaiting.imgRoomRank3.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank6" & ".gif")
    End If
    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(2) And Player(PlayersOnMap(q)).Rank = "Bombinha de vento" Then
    frmWaiting.imgRoomRank3.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank7" & ".gif")
    End If
    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(2) And Player(PlayersOnMap(q)).Rank = "Bomba de vento" Then
    frmWaiting.imgRoomRank3.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank8" & ".gif")
    End If
    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(2) And Player(PlayersOnMap(q)).Rank = "Super bomba de vento" Then
    frmWaiting.imgRoomRank3.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank9" & ".gif")
    End If
    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(2) And Player(PlayersOnMap(q)).Rank = "Bombinha de fogo" Then
    frmWaiting.imgRoomRank3.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank10" & ".gif")
    End If
    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(2) And Player(PlayersOnMap(q)).Rank = "Bomba de fogo" Then
    frmWaiting.imgRoomRank3.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank11" & ".gif")
    End If
    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(2) And Player(PlayersOnMap(q)).Rank = "Super bomba de fogo" Then
    frmWaiting.imgRoomRank3.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank12" & ".gif")
    End If


    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(3) And Player(PlayersOnMap(q)).Rank = "Bombinha de aguá" Then
    frmWaiting.imgRoomRank4.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank1" & ".gif")
    End If
    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(3) And Player(PlayersOnMap(q)).Rank = "Bomba de aguá" Then
    frmWaiting.imgRoomRank4.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank2" & ".gif")
    End If
    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(3) And Player(PlayersOnMap(q)).Rank = "Super bomba de aguá" Then
    frmWaiting.imgRoomRank4.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank3" & ".gif")
    End If
    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(3) And Player(PlayersOnMap(q)).Rank = "Bombinha de terra" Then
    frmWaiting.imgRoomRank4.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank4" & ".gif")
    End If
    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(3) And Player(PlayersOnMap(q)).Rank = "Bomba de terra" Then
    frmWaiting.imgRoomRank4.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank5" & ".gif")
    End If
    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(3) And Player(PlayersOnMap(q)).Rank = "Super bomba de terra" Then
    frmWaiting.imgRoomRank4.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank6" & ".gif")
    End If
    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(3) And Player(PlayersOnMap(q)).Rank = "Bombinha de vento" Then
    frmWaiting.imgRoomRank4.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank7" & ".gif")
    End If
    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(3) And Player(PlayersOnMap(q)).Rank = "Bomba de vento" Then
    frmWaiting.imgRoomRank4.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank8" & ".gif")
    End If
    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(3) And Player(PlayersOnMap(q)).Rank = "Super bomba de vento" Then
    frmWaiting.imgRoomRank4.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank9" & ".gif")
    End If
    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(3) And Player(PlayersOnMap(q)).Rank = "Bombinha de fogo" Then
    frmWaiting.imgRoomRank4.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank10" & ".gif")
    End If
    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(3) And Player(PlayersOnMap(q)).Rank = "Bomba de fogo" Then
    frmWaiting.imgRoomRank4.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank11" & ".gif")
    End If
    If GetPlayerName(PlayersOnMap(q)) = frmWaiting.lstPlayers.List(3) And Player(PlayersOnMap(q)).Rank = "Super bomba de fogo" Then
    frmWaiting.imgRoomRank4.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank12" & ".gif")
    End If
    Next
    End If
End Sub



Private Sub tmrAnim1_Timer()
    If frmWaiting.Visible = True Then
    picRoomChar1.Visible = True
        WaitAnim1 = WaitAnim1 + 1
        If WaitAnim1 = 4 Then WaitAnim1 = 5
        If WaitAnim1 > 7 Then WaitAnim1 = 1
     WaitingBlt1
    End If
    If lstSprites.List(0) = vbNullString Then
    picRoomChar1.Visible = False
    End If
End Sub

Private Sub tmrAnim2_Timer()
    If frmWaiting.Visible = True Then
    picRoomChar2.Visible = True
        WaitAnim2 = WaitAnim2 + 1
        If WaitAnim2 = 4 Then WaitAnim2 = 5
        If WaitAnim2 > 7 Then WaitAnim2 = 1
     WaitingBlt2
    End If
    If lstInfos.List(1) = vbNullString Then
    picRoomChar2.Visible = False
    End If
End Sub

Private Sub tmrAnim3_Timer()
    If frmWaiting.Visible = True Then
    picRoomChar3.Visible = True
        WaitAnim3 = WaitAnim3 + 1
        If WaitAnim3 = 4 Then WaitAnim3 = 5
        If WaitAnim3 > 7 Then WaitAnim3 = 1
     WaitingBlt3
    End If
    If lstSprites.List(2) = vbNullString Then
    picRoomChar3.Visible = False
    End If
End Sub

Private Sub tmrAnim4_Timer()
    If frmWaiting.Visible = True Then
    picRoomChar4.Visible = True
        Waitanim4 = Waitanim4 + 1
        If Waitanim4 = 4 Then Waitanim4 = 5
        If Waitanim4 > 7 Then Waitanim4 = 1
     WaitingBlt4
    End If
    If lstSprites.List(3) = vbNullString Then
    picRoomChar4.Visible = False
    End If
End Sub

Private Sub tmrAtualizador_Timer()
    If frmWaiting.Visible = False Then
    lblWait.Caption = "Aguardando..."
    timeleft = 4
    End If
End Sub

Private Sub tmrGameStart_Timer()
    If timeleft < 1 Then timeleft = 0: tmrGameStart.Enabled = False: frmMainGame.txtChat.Text = vbNullString
    lblWait.Caption = "Começando em ... " & timeleft
    timeleft = timeleft - 1
End Sub

Private Sub txtChat_GotFocus()
SetFocusOnChat
End Sub

Private Sub txtMyChat_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
SayMsg (txtMyChat.Text)
txtMyChat.Text = vbNullString
End If
End Sub
