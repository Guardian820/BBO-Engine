VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bomber ! Bomber ! Online"
   ClientHeight    =   4845
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4485
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   323
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   299
   StartUpPosition =   2  'CenterScreen
   Begin InetCtlsObjects.Inet Inet 
      Left            =   4680
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.PictureBox picMainMenu 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3735
      Left            =   0
      Picture         =   "frmMain.frx":0CCA
      ScaleHeight     =   249
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   300
      TabIndex        =   0
      Top             =   0
      Width           =   4500
      Begin VB.Label lblExit 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1470
         TabIndex        =   15
         Top             =   3120
         Width           =   1635
      End
      Begin VB.Label lblNewAccount 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2280
         TabIndex        =   7
         Top             =   1920
         Width           =   1635
      End
      Begin VB.Label lblLogin 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   360
         TabIndex        =   6
         Top             =   1920
         Width           =   1635
      End
   End
   Begin VB.PictureBox picRegister 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4050
      Left            =   0
      Picture         =   "frmMain.frx":35FD4
      ScaleHeight     =   4050
      ScaleWidth      =   4590
      TabIndex        =   9
      Top             =   0
      Width           =   4590
      Begin VB.PictureBox picFemale 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   2430
         Picture         =   "frmMain.frx":6F47A
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   24
         Top             =   2865
         Width           =   315
      End
      Begin VB.PictureBox picMale 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1620
         Picture         =   "frmMain.frx":6F9FC
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   23
         Top             =   2865
         Width           =   315
      End
      Begin VB.PictureBox picYellow 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H0000FFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1620
         Picture         =   "frmMain.frx":6FF7E
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   22
         Top             =   1380
         Width           =   315
      End
      Begin VB.PictureBox picBlue 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00C00000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1620
         Picture         =   "frmMain.frx":70500
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   21
         Top             =   1020
         Width           =   315
      End
      Begin VB.PictureBox picRed 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1620
         Picture         =   "frmMain.frx":70A82
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   20
         Top             =   660
         Width           =   315
      End
      Begin VB.PictureBox picGreen 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00008000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   2010
         Picture         =   "frmMain.frx":71004
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   19
         Top             =   660
         Width           =   315
      End
      Begin VB.PictureBox picBlack 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000001&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   2010
         Picture         =   "frmMain.frx":71586
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   18
         Top             =   1380
         Width           =   315
      End
      Begin VB.PictureBox picWhite 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   2010
         Picture         =   "frmMain.frx":71B08
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   17
         Top             =   1020
         Width           =   315
      End
      Begin VB.Timer tmrRegister 
         Interval        =   200
         Left            =   3960
         Top             =   960
      End
      Begin VB.PictureBox picDisplay 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   960
         Left            =   2790
         ScaleHeight     =   960
         ScaleWidth      =   960
         TabIndex        =   16
         Top             =   705
         Width           =   960
      End
      Begin VB.TextBox txtRegisterName 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   1680
         MaxLength       =   20
         TabIndex        =   3
         Top             =   1780
         Width           =   2055
      End
      Begin VB.TextBox txtRegisterPass 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
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
         ForeColor       =   &H00000000&
         Height          =   225
         IMEMode         =   3  'DISABLE
         Left            =   1680
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   2150
         Width           =   2055
      End
      Begin VB.TextBox txtRetype 
         BackColor       =   &H00C0C0C0&
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
         ForeColor       =   &H00000000&
         Height          =   225
         IMEMode         =   3  'DISABLE
         Left            =   1680
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   5
         Top             =   2500
         Width           =   2055
      End
      Begin VB.Timer tmrBltName 
         Interval        =   1
         Left            =   3960
         Top             =   600
      End
      Begin VB.Label picLess 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2400
         TabIndex        =   26
         Top             =   1320
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.Label picMore 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2400
         TabIndex        =   25
         Top             =   840
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.Label lblRegCancel 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2400
         TabIndex        =   14
         Top             =   3480
         Width           =   1635
      End
      Begin VB.Label lblACCEPT 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   720
         TabIndex        =   13
         Top             =   3480
         Width           =   1575
      End
   End
   Begin VB.PictureBox picLogin 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3735
      Left            =   0
      Picture         =   "frmMain.frx":7208A
      ScaleHeight     =   3735
      ScaleWidth      =   4500
      TabIndex        =   10
      Top             =   0
      Width           =   4500
      Begin VB.TextBox txtName 
         BackColor       =   &H00C0C0C0&
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
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   1685
         MaxLength       =   20
         TabIndex        =   1
         Top             =   1445
         Width           =   2070
      End
      Begin VB.TextBox txtPassword 
         BackColor       =   &H00C0C0C0&
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
         ForeColor       =   &H00000000&
         Height          =   225
         IMEMode         =   3  'DISABLE
         Left            =   1685
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   1800
         Width           =   2070
      End
      Begin VB.Label lblCancel 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   2400
         TabIndex        =   12
         Top             =   2880
         Width           =   1635
      End
      Begin VB.Label lblConnect 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   600
         TabIndex        =   11
         Top             =   3000
         Width           =   1635
      End
   End
   Begin VB.PictureBox picLoading 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3735
      Left            =   0
      Picture         =   "frmMain.frx":A7394
      ScaleHeight     =   3735
      ScaleWidth      =   4590
      TabIndex        =   8
      Top             =   0
      Width           =   4590
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private LoadingAnim As Long

Private Sub Form_Terminate()
    DestroyGame
End Sub

Private Sub Form_Unload(Cancel As Integer)
    DestroyGame
End Sub

Private Sub lblACCEPT_Click()
Dim i As Long
Dim Name As String
Dim Password As String
Dim PasswordAgain As String

    Name = Trim$(txtRegisterName.Text)
    Password = Trim$(txtRegisterPass.Text)
    PasswordAgain = Trim$(txtRetype.Text)

    If isLoginLegal(Name, Password) Then
    
        If Password <> PasswordAgain Then
        frmMsg.Visible = True
        frmMsg.lblAtenção.Caption = "Atenção"
        frmMsg.lblMsg.Caption = "Sua senha não foi repetida corretamente."
            Exit Sub
        End If
        
        If Not isStringLegal(Name) Then
            Exit Sub
        End If
    
        Call MenuState(MENU_STATE_NEWACCOUNT)
        
    End If

End Sub

Private Sub lblCancel_Click()
    frmMain.picMainMenu.Visible = True
    frmMain.picLogin.Visible = False
End Sub

Private Sub lblConnect_Click()
    If ConnectToServer Then
        If isLoginLegal(txtName.Text, txtPassword.Text) Then
            'DirectMusic_StopMidi
            Call MenuState(MENU_STATE_LOGIN)
            'Call MenuState(MENU_STATE_USECHAR)
        End If
    Else
        frmMsg.Visible = True
        frmMsg.lblAtenção.Caption = "Atenção"
        frmMsg.lblMsg.Caption = "Não foi possível se conectar ao servidor."
    End If
End Sub

Private Sub lblExit_Click()
    DestroyGame
End Sub

Private Sub lblRegCancel_Click()
    frmMain.picMainMenu.Visible = True
    frmMain.picRegister.Visible = False
    frmMain.height = Main_OrigHeight
End Sub

Private Sub lblLogin_Click()

    picLoading.Visible = True
    picMainMenu.Visible = False
    
    If ConnectToServer Then
        picLogin.Visible = True
        picMainMenu.Visible = False
        txtName.SetFocus
    Else
        frmMsg.Visible = True
        frmMsg.lblAtenção.Caption = "Atenção"
        frmMsg.lblMsg.Caption = "Não foi possível se conectar ao servidor."
        picMainMenu.Visible = True
        picLoading.Visible = False
    End If
    
End Sub

Private Sub lblNewAccount_Click()

    picLoading.Visible = True
    picMainMenu.Visible = False

    If ConnectToServer Then
        Call SendGetClasses
        picLoading.Visible = True
        frmMain.picMainMenu.Visible = False
    Else
        frmMsg.Visible = True
        frmMsg.lblAtenção.Caption = "Atenção"
        frmMsg.lblMsg.Caption = "Não foi possível se conectar ao servidor."
        picMainMenu.Visible = True
        picLoading.Visible = False
    End If
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        DestroyGame
    End If
End Sub

Private Sub picBlack_Click()
    If Not RegisterBlack Then
        RegisterBlack = True: picBlack.Picture = LoadPicture(App.Path & "/Graphics/blackok.bmp")
        If RegisterWhite Then RegisterWhite = False: picWhite.Picture = LoadPicture(App.Path & "/Graphics/white.bmp")
        If RegisterGreen Then RegisterGreen = False: picGreen.Picture = LoadPicture(App.Path & "/Graphics/green.bmp")
        If RegisterRed Then RegisterRed = False: picRed.Picture = LoadPicture(App.Path & "/Graphics/red.bmp")
        If RegisterBlue Then RegisterBlue = False: picBlue.Picture = LoadPicture(App.Path & "/Graphics/blue.bmp")
        If RegisterYellow Then RegisterYellow = False: picYellow.Picture = LoadPicture(App.Path & "/Graphics/yellow.bmp")
        HandleSprite
    End If
End Sub

Private Sub picBlue_Click()
    If Not RegisterBlue Then
        RegisterBlue = True: picBlue.Picture = LoadPicture(App.Path & "/Graphics/blueok.bmp")
        If RegisterBlack Then RegisterBlack = False: picBlack.Picture = LoadPicture(App.Path & "/Graphics/black.bmp")
        If RegisterGreen Then RegisterGreen = False: picGreen.Picture = LoadPicture(App.Path & "/Graphics/green.bmp")
        If RegisterRed Then RegisterRed = False: picRed.Picture = LoadPicture(App.Path & "/Graphics/red.bmp")
        If RegisterWhite Then RegisterWhite = False: picWhite.Picture = LoadPicture(App.Path & "/Graphics/white.bmp")
        If RegisterYellow Then RegisterYellow = False: picYellow.Picture = LoadPicture(App.Path & "/Graphics/yellow.bmp")
        HandleSprite
    End If
End Sub

Private Sub picFemale_Click()
    If Not FEMALE Then
        FEMALE = True
        picFemale.Picture = LoadPicture(App.Path & "/Graphics/checked.bmp")
        picMale.Picture = LoadPicture(App.Path & "/Graphics/unchecked.bmp")
    End If
End Sub

Private Sub picGreen_Click()
    If Not RegisterGreen Then
        RegisterGreen = True: picGreen.Picture = LoadPicture(App.Path & "/Graphics/greenok.bmp")
        If RegisterBlack Then RegisterBlack = False: picBlack.Picture = LoadPicture(App.Path & "/Graphics/black.bmp")
        If RegisterWhite Then RegisterWhite = False: picWhite.Picture = LoadPicture(App.Path & "/Graphics/white.bmp")
        If RegisterRed Then RegisterRed = False: picRed.Picture = LoadPicture(App.Path & "/Graphics/red.bmp")
        If RegisterBlue Then RegisterBlue = False: picBlue.Picture = LoadPicture(App.Path & "/Graphics/blue.bmp")
        If RegisterYellow Then RegisterYellow = False: picYellow.Picture = LoadPicture(App.Path & "/Graphics/yellow.bmp")
        HandleSprite
    End If
End Sub

Private Sub picLess_Click()
    If Current_SpriteNum - 1 > 0 Then Current_SpriteNum = Current_SpriteNum - 1
    HandleSprite
End Sub

Private Sub picMale_Click()
    If FEMALE Then
        FEMALE = False
        picMale.Picture = LoadPicture(App.Path & "/Graphics/checked.bmp")
        picFemale.Picture = LoadPicture(App.Path & "/Graphics/unchecked.bmp")
    End If
End Sub

Private Sub picMore_Click()
    If Current_SpriteNum + 1 < 4 Then Current_SpriteNum = Current_SpriteNum + 1
    HandleSprite
End Sub

Private Sub picRed_Click()
    If Not RegisterRed Then
        If RegisterWhite Then RegisterWhite = False: picWhite.Picture = LoadPicture(App.Path & "/Graphics/white.bmp")
        If RegisterBlack Then RegisterBlack = False: picBlack.Picture = LoadPicture(App.Path & "/Graphics/black.bmp")
        If RegisterGreen Then RegisterGreen = False: picGreen.Picture = LoadPicture(App.Path & "/Graphics/green.bmp")
        RegisterRed = True: picRed.Picture = LoadPicture(App.Path & "/Graphics/redok.bmp")
        If RegisterBlue Then RegisterBlue = False: picBlue.Picture = LoadPicture(App.Path & "/Graphics/blue.bmp")
        If RegisterYellow Then RegisterYellow = False: picYellow.Picture = LoadPicture(App.Path & "/Graphics/yellow.bmp")
        HandleSprite
    End If
End Sub

Private Sub picWhite_Click()
    If Not RegisterWhite Then
        RegisterWhite = True: picWhite.Picture = LoadPicture(App.Path & "/Graphics/whiteok.bmp")
        If RegisterBlack Then RegisterBlack = False: picBlack.Picture = LoadPicture(App.Path & "/Graphics/black.bmp")
        If RegisterGreen Then RegisterGreen = False: picGreen.Picture = LoadPicture(App.Path & "/Graphics/green.bmp")
        If RegisterRed Then RegisterRed = False: picRed.Picture = LoadPicture(App.Path & "/Graphics/red.bmp")
        If RegisterBlue Then RegisterBlue = False: picBlue.Picture = LoadPicture(App.Path & "/Graphics/blue.bmp")
        If RegisterYellow Then RegisterYellow = False: picYellow.Picture = LoadPicture(App.Path & "/Graphics/yellow.bmp")
        HandleSprite
    End If
End Sub

Private Sub picYellow_Click()
    If Not RegisterYellow Then
        RegisterYellow = True: picYellow.Picture = LoadPicture(App.Path & "/Graphics/yellowok.bmp")
        If RegisterBlack Then RegisterBlack = False: picBlack.Picture = LoadPicture(App.Path & "/Graphics/black.bmp")
        If RegisterGreen Then RegisterGreen = False: picGreen.Picture = LoadPicture(App.Path & "/Graphics/green.bmp")
        If RegisterRed Then RegisterRed = False: picRed.Picture = LoadPicture(App.Path & "/Graphics/red.bmp")
        If RegisterBlue Then RegisterBlue = False: picBlue.Picture = LoadPicture(App.Path & "/Graphics/blue.bmp")
        If RegisterWhite Then RegisterWhite = False: picWhite.Picture = LoadPicture(App.Path & "/Graphics/white.bmp")
        HandleSprite
    End If
End Sub

Private Sub scrlSprite_Change()
    HandleSprite
End Sub

Private Sub tmrBltName_Timer()
    If picRegister.Visible Then
        If LenB(Trim$(txtRegisterName.Text)) > 0 Then DrawText frmMain.picDisplay.hdc, 0, 0, Trim$(txtRegisterName.Text), QBColor(Black), True
    End If
End Sub

Private Sub tmrRegister_Timer()
    If picRegister.Visible Then
        RegisterAnim = RegisterAnim + 1
        If RegisterAnim = 4 Then RegisterAnim = 5
        If RegisterAnim > 7 Then RegisterAnim = 1
        RegisterBlt
    End If
End Sub

Private Sub HandleSprite()

    Select Case Current_SpriteNum
    
        Case 1
            If RegisterWhite Then
                RegisterSprite = 1
            End If
            If RegisterBlack Then
                RegisterSprite = 2
            End If
            If RegisterGreen Then
                RegisterSprite = 3
            End If
            If RegisterRed Then
                RegisterSprite = 4
            End If
            If RegisterBlue Then
                RegisterSprite = 5
            End If
            If RegisterYellow Then
                RegisterSprite = 6
            End If
        Case 2
            If RegisterWhite Then
                RegisterSprite = 8
            End If
            If RegisterBlack Then
                RegisterSprite = 9
            End If
            If RegisterGreen Then
                RegisterSprite = 10
            End If
            If RegisterRed Then
                RegisterSprite = 11
            End If
            If RegisterBlue Then
                RegisterSprite = 12
            End If
            If RegisterYellow Then
                RegisterSprite = 13
            End If
        Case 3
            If RegisterWhite Then
                RegisterSprite = 15
            End If
            If RegisterBlack Then
                RegisterSprite = 16
            End If
            If RegisterGreen Then
                RegisterSprite = 17
            End If
            If RegisterRed Then
                RegisterSprite = 18
            End If
            If RegisterBlue Then
                RegisterSprite = 19
            End If
            If RegisterYellow Then
                RegisterSprite = 20
            End If
    End Select

End Sub
