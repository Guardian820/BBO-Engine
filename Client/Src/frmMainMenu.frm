VERSION 5.00
Begin VB.Form frmMainMenu 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bomberman Online"
   ClientHeight    =   3765
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4500
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMainMenu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMainMenu.frx":08CA
   ScaleHeight     =   3765
   ScaleWidth      =   4500
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblLogin 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   1320
      TabIndex        =   1
      Top             =   1920
      Width           =   1875
   End
   Begin VB.Label lblNewAccount 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   1320
      TabIndex        =   0
      Top             =   2400
      Width           =   1875
   End
End
Attribute VB_Name = "frmMainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ******************************************
' **            Mirage Source 4           **
' ******************************************

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        DestroyGame
    End If
End Sub

Private Sub Form_Load()
    
    ' initialize DirectX in the background after the form appears
    Call InitDirectDraw
    Call InitSurfaces ' Initialize all needed in-game surfaces
    Call InitDirectSound
    Call InitDirectMusic
    
    DirectMusic_PlayMidi "title.mid"
    
End Sub

Private Sub lblLogin_Click()
    If ConnectToServer Then
        frmLogin.Visible = True
        Me.Visible = False
    Else
        MsgBox "The server is offline! Check back later.", , GAME_NAME
    End If
End Sub

Private Sub lblNewAccount_Click()

    If ConnectToServer Then
        Call SendGetClasses
        frmSendGetData.Visible = True
        Me.Visible = False
    Else
        MsgBox "The server is offline! Check back later.", , GAME_NAME
    End If
End Sub
