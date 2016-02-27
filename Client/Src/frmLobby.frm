VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Begin VB.Form frmLobby 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bomber ! Bomber ! Online"
   ClientHeight    =   7500
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   8325
   Icon            =   "frmLobby.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmLobby.frx":0CCA
   ScaleHeight     =   500
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   555
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Interval        =   10000
      Left            =   1920
      Top             =   0
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   1440
      Top             =   0
   End
   Begin VB.PictureBox picLoja 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   4500
      Left            =   600
      Picture         =   "frmLobby.frx":CC6DC
      ScaleHeight     =   4500
      ScaleWidth      =   6960
      TabIndex        =   17
      Top             =   1200
      Visible         =   0   'False
      Width           =   6960
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   2520
         TabIndex        =   51
         Top             =   3960
         Width           =   1575
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Clique no visual desejado."
         ForeColor       =   &H000000FF&
         Height          =   855
         Left            =   5400
         TabIndex        =   50
         Top             =   2880
         Width           =   1095
      End
      Begin VB.Image Image15 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   990
         Left            =   3000
         Picture         =   "frmLobby.frx":132BCE
         Top             =   600
         Width           =   990
      End
      Begin VB.Image Image14 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   990
         Left            =   600
         Picture         =   "frmLobby.frx":135C10
         Top             =   600
         Width           =   990
      End
      Begin VB.Image Image13 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   990
         Left            =   600
         Picture         =   "frmLobby.frx":138C52
         Top             =   1680
         Width           =   990
      End
      Begin VB.Image Image12 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   990
         Left            =   600
         Picture         =   "frmLobby.frx":13BC94
         Top             =   2760
         Width           =   990
      End
      Begin VB.Image Image11 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   990
         Left            =   4200
         Picture         =   "frmLobby.frx":13ECD6
         Top             =   2760
         Width           =   990
      End
      Begin VB.Image Image10 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   990
         Left            =   3000
         Picture         =   "frmLobby.frx":141D18
         Top             =   1680
         Width           =   990
      End
      Begin VB.Image Image9 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   990
         Left            =   1800
         Picture         =   "frmLobby.frx":144D5A
         Top             =   600
         Width           =   990
      End
      Begin VB.Image Image8 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   990
         Left            =   1800
         Picture         =   "frmLobby.frx":147D9C
         Top             =   2760
         Width           =   990
      End
      Begin VB.Image Image7 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   990
         Left            =   3000
         Picture         =   "frmLobby.frx":14ADDE
         Top             =   2760
         Width           =   990
      End
      Begin VB.Image Image6 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   990
         Left            =   5400
         Picture         =   "frmLobby.frx":14DE20
         Top             =   1680
         Width           =   990
      End
      Begin VB.Image Image5 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   990
         Left            =   4200
         Picture         =   "frmLobby.frx":150E62
         Top             =   600
         Width           =   990
      End
      Begin VB.Image Image4 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   990
         Left            =   1800
         Picture         =   "frmLobby.frx":153EA4
         Top             =   1680
         Width           =   990
      End
      Begin VB.Image Image2 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   990
         Left            =   4200
         Picture         =   "frmLobby.frx":156EE6
         Top             =   1680
         Width           =   990
      End
      Begin VB.Image Image3 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   990
         Left            =   5400
         Picture         =   "frmLobby.frx":159F28
         Top             =   600
         Width           =   990
      End
   End
   Begin VB.Timer tmrProteção 
      Interval        =   5000
      Left            =   960
      Top             =   0
   End
   Begin VB.PictureBox picLojaPrincipal 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4515
      Left            =   600
      Picture         =   "frmLobby.frx":15CF6A
      ScaleHeight     =   4515
      ScaleWidth      =   6960
      TabIndex        =   34
      Top             =   1200
      Visible         =   0   'False
      Width           =   6960
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Cuidado,se comprar um item de bonus menor,você perderá o bonûs que possúi agora."
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   2880
         TabIndex        =   52
         Top             =   3360
         Width           =   3735
      End
      Begin VB.Image label14 
         Height          =   300
         Left            =   3960
         Picture         =   "frmLobby.frx":1C345C
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Image label13 
         Height          =   300
         Left            =   3960
         Picture         =   "frmLobby.frx":1C4D4E
         Top             =   2040
         Width           =   1575
      End
      Begin VB.Image label12 
         Height          =   300
         Left            =   3960
         Picture         =   "frmLobby.frx":1C6640
         Top             =   2880
         Width           =   1575
      End
      Begin VB.Image label11 
         Height          =   300
         Left            =   840
         Picture         =   "frmLobby.frx":1C7F32
         Top             =   3720
         Width           =   1575
      End
      Begin VB.Image label7 
         Height          =   300
         Left            =   840
         Picture         =   "frmLobby.frx":1C9824
         Top             =   2880
         Width           =   1575
      End
      Begin VB.Image label9 
         Height          =   300
         Left            =   840
         Picture         =   "frmLobby.frx":1CB116
         Top             =   2040
         Width           =   1575
      End
      Begin VB.Image label10 
         Height          =   300
         Left            =   840
         Picture         =   "frmLobby.frx":1CCA08
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   2520
         TabIndex        =   49
         Top             =   3960
         Width           =   1575
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "100 BC"
         Height          =   255
         Index           =   4
         Left            =   4080
         TabIndex        =   48
         Top             =   2640
         Width           =   855
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "100 BC"
         Height          =   255
         Index           =   3
         Left            =   4080
         TabIndex        =   47
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "10000 BP"
         Height          =   255
         Index           =   2
         Left            =   4080
         TabIndex        =   46
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "10000 BP"
         Height          =   255
         Index           =   1
         Left            =   960
         TabIndex        =   45
         Top             =   2640
         Width           =   855
      End
      Begin VB.Label lblLojaI1 
         BackStyle       =   0  'Transparent
         Caption         =   "Alcance de bombas +3"
         Height          =   255
         Index           =   4
         Left            =   4080
         TabIndex        =   44
         Top             =   1560
         Width           =   1815
      End
      Begin VB.Label lblItem4 
         BackStyle       =   0  'Transparent
         Caption         =   "Alcance das bombas +2"
         Height          =   255
         Index           =   3
         Left            =   960
         TabIndex        =   43
         Top             =   2400
         Width           =   1815
      End
      Begin VB.Label lblLojaI1 
         BackStyle       =   0  'Transparent
         Caption         =   "Limite de bombas +3"
         Height          =   255
         Index           =   2
         Left            =   4080
         TabIndex        =   42
         Top             =   2400
         Width           =   1815
      End
      Begin VB.Label lblLojaI1 
         BackStyle       =   0  'Transparent
         Caption         =   "Limite de bombas +2"
         Height          =   255
         Index           =   1
         Left            =   4080
         TabIndex        =   41
         Top             =   720
         Width           =   1815
      End
      Begin VB.Image imgItem6 
         Height          =   480
         Left            =   3480
         Picture         =   "frmLobby.frx":1CE2FA
         Top             =   1560
         Width           =   480
      End
      Begin VB.Image imgItem3 
         Height          =   480
         Left            =   360
         Picture         =   "frmLobby.frx":1CEF3C
         Top             =   2400
         Width           =   480
      End
      Begin VB.Image imgItem7 
         Height          =   480
         Left            =   3480
         Picture         =   "frmLobby.frx":1CFB7E
         Top             =   2400
         Width           =   480
      End
      Begin VB.Image imgItem5 
         Height          =   480
         Left            =   3480
         Picture         =   "frmLobby.frx":1D07C0
         Top             =   720
         Width           =   480
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "100 BC"
         Height          =   255
         Index           =   0
         Left            =   960
         TabIndex        =   40
         Top             =   3480
         Width           =   855
      End
      Begin VB.Label lblItem3 
         BackStyle       =   0  'Transparent
         Caption         =   "Mudar aparência"
         Height          =   255
         Left            =   960
         TabIndex        =   39
         Top             =   3240
         Width           =   1575
      End
      Begin VB.Image imgItem4 
         Height          =   480
         Left            =   360
         Picture         =   "frmLobby.frx":1D1402
         Top             =   3240
         Width           =   480
      End
      Begin VB.Label lblLojaP2 
         BackStyle       =   0  'Transparent
         Caption         =   "1000 BP"
         Height          =   255
         Left            =   960
         TabIndex        =   38
         Top             =   960
         Width           =   855
      End
      Begin VB.Label lblItem2 
         BackStyle       =   0  'Transparent
         Caption         =   "Alcance das bombas + 1"
         Height          =   255
         Left            =   960
         TabIndex        =   37
         Top             =   1560
         Width           =   1935
      End
      Begin VB.Label lblLojaP1 
         BackStyle       =   0  'Transparent
         Caption         =   "1000 BP"
         Height          =   255
         Left            =   960
         TabIndex        =   36
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label lblItem1 
         BackStyle       =   0  'Transparent
         Caption         =   "Limite de bombas +1"
         Height          =   255
         Index           =   0
         Left            =   960
         TabIndex        =   35
         Top             =   720
         Width           =   1815
      End
      Begin VB.Image imgItem1 
         Height          =   480
         Left            =   360
         Picture         =   "frmLobby.frx":1D2044
         Top             =   720
         Width           =   480
      End
      Begin VB.Image imgItem2 
         Height          =   480
         Left            =   360
         Picture         =   "frmLobby.frx":1D2C86
         Top             =   1560
         Width           =   480
      End
   End
   Begin VB.PictureBox picRemoFriend 
      BorderStyle     =   0  'None
      Height          =   2550
      Left            =   1920
      Picture         =   "frmLobby.frx":1D38C8
      ScaleHeight     =   2550
      ScaleWidth      =   4500
      TabIndex        =   26
      Top             =   2160
      Visible         =   0   'False
      Width           =   4500
      Begin VB.TextBox Text1 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   220
         Left            =   1200
         TabIndex        =   31
         Top             =   1200
         Width           =   2055
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   1320
         TabIndex        =   30
         Top             =   2040
         Width           =   1695
      End
   End
   Begin VB.PictureBox picAddFriend 
      BorderStyle     =   0  'None
      Height          =   2550
      Left            =   1920
      Picture         =   "frmLobby.frx":1F8EB2
      ScaleHeight     =   2550
      ScaleWidth      =   4500
      TabIndex        =   27
      Top             =   2160
      Visible         =   0   'False
      Width           =   4500
      Begin VB.TextBox txtFriend 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   1200
         TabIndex        =   29
         Top             =   1200
         Width           =   2055
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   1440
         TabIndex        =   28
         Top             =   2040
         Width           =   1455
      End
   End
   Begin VB.Timer tmrLoad 
      Interval        =   200
      Left            =   480
      Top             =   0
   End
   Begin VB.PictureBox picFriendList 
      BorderStyle     =   0  'None
      Height          =   2550
      Left            =   1920
      Picture         =   "frmLobby.frx":21E49C
      ScaleHeight     =   2550
      ScaleWidth      =   4500
      TabIndex        =   23
      Top             =   2160
      Visible         =   0   'False
      Width           =   4500
      Begin VB.ListBox lstFriend 
         Appearance      =   0  'Flat
         Height          =   1590
         Left            =   240
         TabIndex        =   24
         Top             =   240
         Width           =   3975
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Feche e abra a lista para atualizar."
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   240
         TabIndex        =   33
         Top             =   1830
         Width           =   3975
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   2640
         TabIndex        =   32
         Top             =   2040
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Height          =   375
         Left            =   240
         TabIndex        =   25
         Top             =   2040
         Width           =   1575
      End
   End
   Begin VB.PictureBox picMsg 
      BorderStyle     =   0  'None
      Height          =   2550
      Left            =   1920
      Picture         =   "frmLobby.frx":243A86
      ScaleHeight     =   2550
      ScaleWidth      =   4500
      TabIndex        =   19
      Top             =   2160
      Width           =   4500
      Begin VB.Label lblClose 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   1440
         TabIndex        =   21
         Top             =   2040
         Width           =   1575
      End
      Begin VB.Label lblMsg 
         BackStyle       =   0  'Transparent
         Caption         =   "Você ganhou 100 BCash de presente do ""servidor"" !"
         Height          =   255
         Left            =   360
         TabIndex        =   20
         Top             =   1080
         Width           =   3855
      End
   End
   Begin VB.Timer tmrAnim 
      Interval        =   200
      Left            =   0
      Top             =   0
   End
   Begin VB.PictureBox picMyChar 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   960
      Left            =   6750
      ScaleHeight     =   960
      ScaleWidth      =   960
      TabIndex        =   6
      Top             =   3630
      Width           =   960
   End
   Begin VB.TextBox txtEnterChat 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   240
      Left            =   600
      TabIndex        =   2
      Top             =   7170
      Width           =   7575
   End
   Begin VB.ListBox lstRooms 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   1920
      ItemData        =   "frmLobby.frx":269070
      Left            =   90
      List            =   "frmLobby.frx":269077
      TabIndex        =   0
      Top             =   480
      Width           =   6105
   End
   Begin RichTextLib.RichTextBox txtChat 
      Height          =   2055
      Left            =   120
      TabIndex        =   1
      Top             =   5025
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   3625
      _Version        =   393217
      BackColor       =   16777215
      BorderStyle     =   0
      Enabled         =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmLobby.frx":269087
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Image imgRank 
      Height          =   480
      Left            =   6330
      Picture         =   "frmLobby.frx":269102
      Top             =   2640
      Width           =   480
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   6480
      TabIndex        =   22
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label lblLoja 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   6480
      TabIndex        =   18
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label lblKillDeath 
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
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
      Left            =   7200
      TabIndex        =   16
      Top             =   2640
      Width           =   855
   End
   Begin VB.Label lblSKillDeath 
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
      Left            =   6840
      TabIndex        =   15
      Top             =   2640
      Width           =   375
   End
   Begin VB.Label lblExp 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Left            =   7200
      TabIndex        =   14
      Top             =   2880
      Width           =   855
   End
   Begin VB.Label lblSEXP 
      BackStyle       =   0  'Transparent
      Caption         =   "EXP:"
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
      Left            =   6840
      TabIndex        =   13
      Top             =   2880
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   2265
      Left            =   120
      Picture         =   "frmLobby.frx":2695C3
      Top             =   2520
      Width           =   6045
   End
   Begin VB.Label lblBC 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Left            =   7200
      TabIndex        =   12
      Top             =   3360
      Width           =   855
   End
   Begin VB.Label lblBP 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Left            =   7200
      TabIndex        =   11
      Top             =   3120
      Width           =   855
   End
   Begin VB.Label lblSBC 
      BackStyle       =   0  'Transparent
      Caption         =   "BC:"
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
      Left            =   6840
      TabIndex        =   10
      Top             =   3360
      Width           =   255
   End
   Begin VB.Label lblSBP 
      BackStyle       =   0  'Transparent
      Caption         =   "BP:"
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
      Left            =   6840
      TabIndex        =   9
      Top             =   3120
      Width           =   255
   End
   Begin VB.Label lblPlayerRank 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Super bomba de aguá"
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
      Left            =   6240
      TabIndex        =   8
      Top             =   1800
      Width           =   1935
   End
   Begin VB.Label lblPlayerName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Nome"
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
      Left            =   6315
      TabIndex        =   7
      Top             =   1620
      Width           =   1815
   End
   Begin VB.Label lblHighScores 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6360
      TabIndex        =   5
      Top             =   960
      Width           =   1875
   End
   Begin VB.Label lblExit 
      BackStyle       =   0  'Transparent
      Height          =   285
      Left            =   6480
      TabIndex        =   4
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Label lblJoinRoom 
      BackStyle       =   0  'Transparent
      Height          =   285
      Left            =   6480
      TabIndex        =   3
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "frmLobby"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)

    Call HandleKeypresses(KeyAscii)
    
    ' prevents textbox on error ding sound
    If KeyAscii = vbKeyReturn Then KeyAscii = 0
    
End Sub

Private Sub Form_Load()
Dim result As Long

    frmMain.tmrRegister.Enabled = False
    
    result = SetWindowLong(txtChat.hWnd, GWL_EXSTYLE, WS_EX_TRANSPARENT)
    
    SendData CRequestRoomList & END_CHAR
    
End Sub

Private Sub Form_Terminate()
    DestroyGame
End Sub

Private Sub Form_Unload(Cancel As Integer)
    DestroyGame
End Sub

Private Sub imgLojaI1_Click()

End Sub

Private Sub Image10_Click()
If Player(MyIndex).BCash > 99 Then
Call DealBC(100)
Call SetPlayerSprite(MyIndex, 18)
Call SendSprite(18)
frmMsg.Visible = True
frmMsg.lblMsg.Caption = "Você comprou um novo visual !"
frmMsg.lblAtenção = "Transação OK !"
Else
frmMsg.Visible = True
frmMsg.lblMsg.Caption = "Você não possui Bomber Cash suficiente."
frmMsg.lblAtenção = "Transação falhou !"
End If
End Sub

Private Sub Image11_Click()
If Player(MyIndex).BCash > 99 Then
Call DealBC(100)
Call SetPlayerSprite(MyIndex, 9)
Call SendSprite(9)
frmMsg.Visible = True
frmMsg.lblMsg.Caption = "Você comprou um novo visual !"
frmMsg.lblAtenção = "Transação OK !"
Else
frmMsg.Visible = True
frmMsg.lblMsg.Caption = "Você não possui Bomber Cash suficiente."
frmMsg.lblAtenção = "Transação falhou !"
End If
End Sub

Private Sub Image12_Click()
If Player(MyIndex).BCash > 99 Then
Call DealBC(100)
Call SetPlayerSprite(MyIndex, 15)
Call SendSprite(15)
frmMsg.Visible = True
frmMsg.lblMsg.Caption = "Você comprou um novo visual !"
frmMsg.lblAtenção = "Transação OK !"
Else
frmMsg.Visible = True
frmMsg.lblMsg.Caption = "Você não possui Bomber Cash suficiente."
frmMsg.lblAtenção = "Transação falhou !"
End If
End Sub

Private Sub Image13_Click()
If Player(MyIndex).BCash > 99 Then
Call DealBC(100)
Call SetPlayerSprite(MyIndex, 8)
Call SendSprite(8)
frmMsg.Visible = True
frmMsg.lblMsg.Caption = "Você comprou um novo visual !"
frmMsg.lblAtenção = "Transação OK !"
Else
frmMsg.Visible = True
frmMsg.lblMsg.Caption = "Você não possui Bomber Cash suficiente."
frmMsg.lblAtenção = "Transação falhou !"
End If
End Sub

Private Sub Image14_Click()
If Player(MyIndex).BCash > 99 Then
Call DealBC(100)
Call SetPlayerSprite(MyIndex, 14)
Call SendSprite(14)
frmMsg.Visible = True
frmMsg.lblMsg.Caption = "Você comprou um novo visual !"
frmMsg.lblAtenção = "Transação OK !"
Else
frmMsg.Visible = True
frmMsg.lblMsg.Caption = "Você não possui Bomber Cash suficiente."
frmMsg.lblAtenção = "Transação falhou !"
End If
End Sub

Private Sub Image15_Click()
If Player(MyIndex).BCash > 99 Then
Call DealBC(100)
Call SetPlayerSprite(MyIndex, 7)
Call SendSprite(7)
frmMsg.Visible = True
frmMsg.lblMsg.Caption = "Você comprou um novo visual !"
frmMsg.lblAtenção = "Transação OK !"
Else
frmMsg.Visible = True
frmMsg.lblMsg.Caption = "Você não possui Bomber Cash suficiente."
frmMsg.lblAtenção = "Transação falhou !"
End If
End Sub



Private Sub Image17_Click()

End Sub

Private Sub Image2_Click()
If Player(MyIndex).BCash > 99 Then
Call DealBC(100)
Call SetPlayerSprite(MyIndex, 19)
Call SendSprite(19)
frmMsg.Visible = True
frmMsg.lblMsg.Caption = "Você comprou um novo visual !"
frmMsg.lblAtenção = "Transação OK !"
Else
frmMsg.Visible = True
frmMsg.lblMsg.Caption = "Você não possui Bomber Cash suficiente."
frmMsg.lblAtenção = "Transação falhou !"
End If
End Sub

Private Sub Image3_Click()
If Player(MyIndex).BCash > 99 Then
Call DealBC(100)
Call SetPlayerSprite(MyIndex, 20)
Call SendSprite(20)
frmMsg.Visible = True
frmMsg.lblMsg.Caption = "Você comprou um novo visual !"
frmMsg.lblAtenção = "Transação OK !"
Else
frmMsg.Visible = True
frmMsg.lblMsg.Caption = "Você não possui Bomber Cash suficiente."
frmMsg.lblAtenção = "Transação falhou !"
End If
End Sub

Private Sub Image4_Click()
If Player(MyIndex).BCash > 99 Then
Call DealBC(100)
Call SetPlayerSprite(MyIndex, 17)
Call SendSprite(14)
frmMsg.Visible = True
frmMsg.lblMsg.Caption = "Você comprou um novo visual !"
frmMsg.lblAtenção = "Transação OK !"
Else
frmMsg.Visible = True
frmMsg.lblMsg.Caption = "Você não possui Bomber Cash suficiente."
frmMsg.lblAtenção = "Transação falhou !"
End If
End Sub

Private Sub Image5_Click()
If Player(MyIndex).BCash > 99 Then
Call DealBC(100)
Call SetPlayerSprite(MyIndex, 12)
Call SendSprite(12)
frmMsg.Visible = True
frmMsg.lblMsg.Caption = "Você comprou um novo visual !"
frmMsg.lblAtenção = "Transação OK !"
Else
frmMsg.Visible = True
frmMsg.lblMsg.Caption = "Você não possui Bomber Cash suficiente."
frmMsg.lblAtenção = "Transação falhou !"
End If
End Sub

Private Sub Image6_Click()
If Player(MyIndex).BCash > 99 Then
Call DealBC(100)
Call SetPlayerSprite(MyIndex, 16)
Call SendSprite(16)
frmMsg.Visible = True
frmMsg.lblMsg.Caption = "Você comprou um novo visual !"
frmMsg.lblAtenção = "Transação OK !"
Else
frmMsg.Visible = True
frmMsg.lblMsg.Caption = "Você não possui Bomber Cash suficiente."
frmMsg.lblAtenção = "Transação falhou !"
End If
End Sub

Private Sub Image7_Click()
If Player(MyIndex).BCash > 99 Then
Call DealBC(100)
Call SetPlayerSprite(MyIndex, 11)
Call SendSprite(11)
frmMsg.Visible = True
frmMsg.lblMsg.Caption = "Você comprou um novo visual !"
frmMsg.lblAtenção = "Transação OK !"
Else
frmMsg.Visible = True
frmMsg.lblMsg.Caption = "Você não possui Bomber Cash suficiente."
frmMsg.lblAtenção = "Transação falhou !"
End If
End Sub

Private Sub Image8_Click()
If Player(MyIndex).BCash > 99 Then
Call DealBC(100)
Call SetPlayerSprite(MyIndex, 10)
Call SendSprite(10)
frmMsg.Visible = True
frmMsg.lblMsg.Caption = "Você comprou um novo visual !"
frmMsg.lblAtenção = "Transação OK !"
Else
frmMsg.Visible = True
frmMsg.lblMsg.Caption = "Você não possui Bomber Cash suficiente."
frmMsg.lblAtenção = "Transação falhou !"
End If
End Sub

Private Sub Image9_Click()
If Player(MyIndex).BCash > 99 Then
Call DealBC(100)
Call SetPlayerSprite(MyIndex, 13)
Call SendSprite(13)
frmMsg.Visible = True
frmMsg.lblMsg.Caption = "Você comprou um novo visual !"
frmMsg.lblAtenção = "Transação OK !"
Else
frmMsg.Visible = True
frmMsg.lblMsg.Caption = "Você não possui Bomber Cash suficiente."
frmMsg.lblAtenção = "Transação falhou !"
End If
End Sub

Private Sub Label1_Click()
If picFriendList.Visible = False Then
picFriendList.Visible = True
lstFriend.Clear
Call UpdateFriendList
Else
picFriendList.Visible = False
End If
End Sub

Private Sub Label10_Click()
If Player(MyIndex).BPoints > 999 Then
Call DealBP(1000)
Call AddStatus(0, 1)
frmMsg.Visible = True
frmMsg.lblMsg.Caption = "Você comprou 'Limite de bombas +1'."
frmMsg.lblAtenção = "Transação OK !"
Else
frmMsg.Visible = True
frmMsg.lblMsg.Caption = "Você não possui Bomber Point suficiente."
frmMsg.lblAtenção = "Transação falhou !"
End If
End Sub

Private Sub Label11_Click()
If Player(MyIndex).BCash > 99 Then
picLoja.Visible = True
Else
frmMsg.Visible = True
frmMsg.lblMsg.Caption = "Você não possui Bomber Cash suficiente."
frmMsg.lblAtenção = "Transação falhou !"
End If
End Sub

Private Sub Label12_Click()
If Player(MyIndex).BCash > 99 Then
Call DealBC(100)
Call AddStatus(3, 0)
frmMsg.Visible = True
frmMsg.lblMsg.Caption = "Você comprou 'Limite de bombas +3'."
frmMsg.lblAtenção = "Transação OK !"
Else
frmMsg.Visible = True
frmMsg.lblMsg.Caption = "Você não possui Bomber Cash suficiente."
frmMsg.lblAtenção = "Transação falhou !"
End If
End Sub

Private Sub Label13_Click()
If Player(MyIndex).BCash > 99 Then
Call DealBC(100)
Call AddStatus(3, 0)
frmMsg.Visible = True
frmMsg.lblMsg.Caption = "Você comprou 'Alcance das bombas +3'."
frmMsg.lblAtenção = "Transação OK !"
Else
frmMsg.Visible = True
frmMsg.lblMsg.Caption = "Você não possui Bomber Cash suficiente."
frmMsg.lblAtenção = "Transação falhou !"
End If
End Sub

Private Sub Label14_Click()
If Player(MyIndex).BPoints > 9999 Then
Call DealBP(10000)
Call AddStatus(0, 2)
frmMsg.Visible = True
frmMsg.lblMsg.Caption = "Você comprou 'Limite de bombas +2'."
frmMsg.lblAtenção = "Transação OK !"
Else
frmMsg.Visible = True
frmMsg.lblMsg.Caption = "Você não possui Bomber Point suficiente."
frmMsg.lblAtenção = "Transação falhou !"
End If
End Sub

Private Sub Label15_Click()
picLojaPrincipal.Visible = False
End Sub

Private Sub Label17_Click()
picLoja.Visible = False
End Sub

Private Sub Label2_Click()
picAddFriend.Visible = True
End Sub

Private Sub Label3_Click()
If Not txtFriend.Text = vbNullString Then
Call addfriend(txtFriend.Text)
picAddFriend.Visible = False
Else
picAddFriend.Visible = False
End If
End Sub

Private Sub Label4_Click()
If Not Text1.Text = vbNullString Then
Call RemoveFriend(Text1.Text)
picRemoFriend.Visible = False
Else
picRemoFriend.Visible = False
End If
End Sub

Private Sub Label5_Click()
picRemoFriend.Visible = True
End Sub

Private Sub Label7_Click()
If Player(MyIndex).BPoints > 9999 Then
Call DealBP(10000)
Call AddStatus(2, 0)
frmMsg.Visible = True
frmMsg.lblMsg.Caption = "Você comprou 'Alcance das bombas +2'."
frmMsg.lblAtenção = "Transação OK !"
Else
frmMsg.Visible = True
frmMsg.lblMsg.Caption = "Você não possui Bomber Point suficiente."
frmMsg.lblAtenção = "Transação falhou !"
End If
End Sub

Private Sub Label9_Click()
If Player(MyIndex).BPoints > 999 Then
Call DealBP(1000)
Call AddStatus(1, 0)
frmMsg.Visible = True
frmMsg.lblMsg.Caption = "Você comprou 'Alcance das bombas +2'."
frmMsg.lblAtenção = "Transação OK !"
Else
frmMsg.Visible = True
frmMsg.lblMsg.Caption = "Você não possui Bomber Point suficiente."
frmMsg.lblAtenção = "Transação falhou !"
End If
End Sub
Private Sub lblClose_Click()
picMsg.Visible = False
End Sub

Private Sub lblExit_Click()
    DestroyGame
End Sub

Private Sub lblHighScores_Click()
    SendData CRequestHighScores & END_CHAR
End Sub

Private Sub lblJoinRoom_Click()
    frmMainGame.txtChat.Text = vbNullString
    If lstRooms.ListIndex > -1 Then
        If lstRooms.List(lstRooms.ListIndex) <> "Não há salas disponíveis." Then
            SendData CRequestJoinRoom & SEP_CHAR & lstRooms.ListIndex + 1 & END_CHAR 'Room(0) & " " & Room(1) & END_CHAR
        End If
    End If
End Sub

Private Sub lblLoja_Click()
'picMsg.Visible = True
picLojaPrincipal.Visible = True
End Sub

Private Sub lblLojaI2_Click()

End Sub

Private Sub lstRooms_DblClick()
    frmMainGame.txtChat.Text = vbNullString
    If lstRooms.ListIndex > -1 Then
        If lstRooms.List(lstRooms.ListIndex) <> "Não há salas disponíveis." Then
            SendData CRequestJoinRoom & SEP_CHAR & lstRooms.ListIndex + 1 & END_CHAR 'Room(0) & " " & Room(1) & END_CHAR
        End If
    End If
End Sub

Private Sub Timer1_Timer()
If frmLobby.Visible = True Then
If lblBC.Caption = "101" Then
picMsg.Visible = True
Timer1.Enabled = False
End If
End If
End Sub

Private Sub Timer2_Timer()
If frmLobby.Visible = True Then
Timer1.Enabled = False
Timer2.Enabled = False
End If
End Sub

Private Sub tmrProteção_Timer()
If frmLobby.Visible = True Then
If Player(MyIndex).BPoints < 0 Then
frmLobby.Visible = False
frmMsg.Visible = True
frmMsg.lblMsg.Caption = "Está conta foi banida."
frmMsg.lblAtenção = "Atenção !"
End If
End If

If frmMainGame.Visible = True Then
If Player(MyIndex).BPoints < 0 Then
frmMainGame.Visible = False
frmMsg.Visible = True
frmMsg.lblMsg.Caption = "Está conta foi banida."
frmMsg.lblAtenção = "Atenção !"
End If
End If

If frmLobby.Visible = True Then
If Player(MyIndex).BCash < 0 Then
frmLobby.Visible = False
frmMsg.Visible = True
frmMsg.lblMsg.Caption = "Está conta foi banida."
frmMsg.lblAtenção = "Atenção !"
End If
End If

If frmMainGame.Visible = True Then
If Player(MyIndex).BCash < 0 Then
frmMainGame.Visible = False
frmMsg.Visible = True
frmMsg.lblMsg.Caption = "Está conta foi banida."
frmMsg.lblAtenção = "Atenção !"
End If
End If

End Sub

Private Sub tmrAnim_Timer()
    If frmLobby.Visible = True Then
        LobbyAnim = LobbyAnim + 1
        If LobbyAnim = 4 Then LobbyAnim = 5
        If LobbyAnim > 7 Then LobbyAnim = 1
     LobbyBlt
    End If
End Sub

Private Sub tmrUpFriend_Timer()
lstFriend.Clear
Call UpdateFriendList
End Sub

Private Sub tmrLoad_Timer()
Dim Rank As String

If Player(MyIndex).Kills * 5 >= 0 Then
Rank = "Bombinha de aguá"
End If

If Player(MyIndex).Kills * 5 > 20 Then
Rank = "Bomba de aguá"
End If

If Player(MyIndex).Kills * 5 > 50 Then
Rank = "Super bomba de aguá"
End If

If Player(MyIndex).Kills * 5 > 120 Then
Rank = "Bombinha de terra"
End If

If Player(MyIndex).Kills * 5 > 240 Then
Rank = "Bomba de terra"
End If

If Player(MyIndex).Kills * 5 > 500 Then
Rank = "Super bomba de terra"
End If

If Player(MyIndex).Kills * 5 > 1200 Then
Rank = "Bombinha de vento"
End If

If Player(MyIndex).Kills * 5 > 2600 Then
Rank = "Bomba de vento"
End If

If Player(MyIndex).Kills * 5 > 5400 Then
Rank = "Super bomba de vento"
End If

If Player(MyIndex).Kills * 5 > 11000 Then
Rank = "Bombinha de fogo"
End If

If Player(MyIndex).Kills * 5 > 23000 Then
Rank = "Bomba de fogo"
End If

If Player(MyIndex).Kills * 5 > 50000 Then
Rank = "Super bomba de fogo"
End If

If frmLobby.Visible = True Then
LobbyBlt
lblPlayerName.Caption = GetPlayerName(MyIndex)
lblPlayerRank.Caption = Rank

If Rank = "Bombinha de aguá" Then
frmLobby.imgRank.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank1" & ".gif")
End If
If Rank = "Bomba de aguá" Then
frmLobby.imgRank.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank2" & ".gif")
End If
If Rank = "Super bomba de aguá" Then
frmLobby.imgRank.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank3" & ".gif")
End If
If Rank = "Bombinha de terra" Then
frmLobby.imgRank.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank4" & ".gif")
End If
If Rank = "Bomba de terra" Then
frmLobby.imgRank.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank5" & ".gif")
End If
If Rank = "Super bomba de terra" Then
frmLobby.imgRank.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank6" & ".gif")
End If
If Rank = "Bombinha de vento" Then
frmLobby.imgRank.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank7" & ".gif")
End If
If Rank = "Bomba de vento" Then
frmLobby.imgRank.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank8" & ".gif")
End If
If Rank = "Super bomba de vento" Then
frmLobby.imgRank.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank9" & ".gif")
End If
If Rank = "Bombinha de fogo" Then
frmLobby.imgRank.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank10" & ".gif")
End If
If Rank = "Bomba de fogo" Then
frmLobby.imgRank.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank11" & ".gif")
End If
If Rank = "Super bomba de fogo" Then
frmLobby.imgRank.Picture = LoadPicture(App.Path & "\graphics\ranks\" & "rank12" & ".gif")
End If
End If
End Sub

Private Sub txtChat_GotFocus()
    txtEnterChat.SetFocus
End Sub

Private Sub txtEnterChat_Change()
    MyText = txtEnterChat.Text
End Sub

Private Sub txtEnterChat_KeyPress(KeyAscii As Integer)
    Form_KeyPress KeyAscii
End Sub

