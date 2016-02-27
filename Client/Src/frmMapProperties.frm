VERSION 5.00
Begin VB.Form frmMapProperties 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Map Properties"
   ClientHeight    =   3900
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7470
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3900
   ScaleWidth      =   7470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.HScrollBar scrlPlayerLimit 
      Height          =   255
      Left            =   6000
      Max             =   16
      Min             =   2
      TabIndex        =   22
      Top             =   120
      Value           =   4
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Caption         =   "Ligações"
      Height          =   1455
      Left            =   120
      TabIndex        =   17
      Top             =   600
      Width           =   2895
      Begin VB.TextBox txtUp 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   1080
         TabIndex        =   21
         Text            =   "0"
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox txtDown 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   1080
         TabIndex        =   20
         Text            =   "0"
         Top             =   960
         Width           =   735
      End
      Begin VB.TextBox txtRight 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   1920
         TabIndex        =   19
         Text            =   "0"
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox txtLeft 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   240
         TabIndex        =   18
         Text            =   "0"
         Top             =   600
         Width           =   735
      End
   End
   Begin VB.Frame fraMapSettings 
      Caption         =   "Configuração do mapa"
      Height          =   2055
      Left            =   3120
      TabIndex        =   11
      Top             =   600
      Width           =   4215
      Begin VB.ComboBox cmbMoral 
         Height          =   360
         ItemData        =   "frmMapProperties.frx":0000
         Left            =   960
         List            =   "frmMapProperties.frx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   360
         Width           =   2415
      End
      Begin VB.HScrollBar scrlMusic 
         Height          =   255
         Left            =   960
         Max             =   255
         TabIndex        =   12
         Top             =   840
         Value           =   1
         Width           =   2415
      End
      Begin VB.Label Label6 
         Caption         =   "Tipo"
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label10 
         Caption         =   "Musica"
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   840
         Width           =   735
      End
      Begin VB.Label lblMusic 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   375
         Left            =   3480
         TabIndex        =   14
         Top             =   840
         Width           =   495
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Renascimento"
      Height          =   1575
      Left            =   120
      TabIndex        =   4
      Top             =   2160
      Width           =   2895
      Begin VB.TextBox txtBootMap 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   1440
         TabIndex        =   7
         Text            =   "0"
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox txtBootX 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   1440
         TabIndex        =   6
         Text            =   "0"
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox txtBootY 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   1440
         TabIndex        =   5
         Text            =   "0"
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "Mapa"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label8 
         Caption         =   "X"
         Height          =   255
         Left            =   360
         TabIndex        =   9
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label9 
         Caption         =   "Y"
         Height          =   255
         Left            =   360
         TabIndex        =   8
         Top             =   1080
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   3120
      TabIndex        =   2
      Top             =   3360
      Width           =   4215
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   375
      Left            =   3120
      TabIndex        =   1
      Top             =   3000
      Width           =   4215
   End
   Begin VB.TextBox txtName 
      Height          =   300
      Left            =   840
      TabIndex        =   0
      Top             =   120
      Width           =   3255
   End
   Begin VB.Label Label2 
      Caption         =   "Limite de jogadores:"
      Height          =   255
      Left            =   4200
      TabIndex        =   24
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label lblPlayerLimit 
      Alignment       =   1  'Right Justify
      Caption         =   "4"
      Height          =   375
      Left            =   6840
      TabIndex        =   23
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Nome"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   735
   End
   Begin VB.Label lblMap 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Mapa atual:"
      Height          =   255
      Left            =   3120
      TabIndex        =   25
      Top             =   2685
      Width           =   4215
   End
End
Attribute VB_Name = "frmMapProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



' ******************************************
' **              BMO Source              **
' ******************************************

Private Sub Form_Load()
Dim x As Long
Dim y As Long
Dim i As Long

    txtName.Text = Trim$(Map.Name)
    If Map.PlayerLimit < 2 Then Map.PlayerLimit = 2
    scrlPlayerLimit.Value = Map.PlayerLimit
    txtUp.Text = CStr(Map.Up)
    txtDown.Text = CStr(Map.Down)
    txtLeft.Text = CStr(Map.Left)
    txtRight.Text = CStr(Map.Right)
    cmbMoral.ListIndex = Map.Moral
    scrlMusic.Value = Map.Music
    txtBootMap.Text = CStr(Map.BootMap)
    txtBootX.Text = CStr(Map.BootX)
    txtBootY.Text = CStr(Map.BootY)
    
    lblMap.Caption = "Mapa atual: " & GetPlayerMap(MyIndex)
    
End Sub

Private Sub scrlMusic_Change()
    lblMusic.Caption = CStr(scrlMusic.Value)
End Sub

Private Sub cmdOk_Click()
Dim i As Long
Dim sTemp As Long

    With Map
    
        .Name = Trim$(txtName.Text)
        .PlayerLimit = Val(scrlPlayerLimit.Value)
        .Up = Val(txtUp.Text)
        .Down = Val(txtDown.Text)
        .Left = Val(txtLeft.Text)
        .Right = Val(txtRight.Text)
        .Moral = cmbMoral.ListIndex
        .Music = scrlMusic.Value
        .BootMap = Val(txtBootMap.Text)
        .BootX = Val(txtBootX.Text)
        .BootY = Val(txtBootY.Text)
        .Shop = 0
        
        For i = 1 To MAX_MAP_NPCS
            .Npc(i) = 0
        Next
        
    End With
    
    Unload Me
    
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub scrlPlayerLimit_Change()
    lblPlayerLimit.Caption = scrlPlayerLimit.Value
End Sub
