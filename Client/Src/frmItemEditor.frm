VERSION 5.00
Begin VB.Form frmItemEditor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Item Editor"
   ClientHeight    =   1845
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5055
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
   ScaleHeight     =   123
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   337
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picItems 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   120
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   8
      Top             =   4440
      Width           =   480
   End
   Begin VB.PictureBox picPic 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   4440
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   7
      Top             =   480
      Width           =   480
   End
   Begin VB.HScrollBar scrlPic 
      Height          =   255
      Left            =   960
      Max             =   255
      TabIndex        =   5
      Top             =   600
      Width           =   2895
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   2640
      TabIndex        =   3
      Top             =   1320
      Width           =   2295
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   2295
   End
   Begin VB.TextBox txtName 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   960
      TabIndex        =   0
      Top             =   120
      Width           =   3975
   End
   Begin VB.Label lblPic 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   255
      Left            =   3840
      TabIndex        =   6
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label5 
      Caption         =   "Imagem"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Nome"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "frmItemEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ******************************************
' **              BMO Source              **
' ******************************************

Private Sub cmdOk_Click()
    If LenB(Trim$(txtName)) = 0 Then
        Call MsgBox("Nome requerido.")
    Else
        Call ItemEditorOk
    End If
End Sub

Private Sub cmdCancel_Click()
    Call ItemEditorCancel
End Sub

Private Sub cmbType_Click()

End Sub

Private Sub scrlPic_Change()
    lblPic.Caption = CStr(scrlPic.Value)
    Call ItemEditorBltItem
End Sub

' Equipment Data
' *********************************
Private Sub scrlDurability_Change()
    'lblDurability.Caption = CStr(scrlDurability.Value)
End Sub

Private Sub scrlStrength_Change()
    'lblStrength.Caption = CStr(scrlStrength.Value)
End Sub

' Vitals Data
' *********************************
Private Sub scrlVitalMod_Change()
    'lblVitalMod.Caption = CStr(scrlVitalMod.Value)
End Sub

' Spell Data
' *********************************
Private Sub scrlSpell_Change()
    'lblSpellName.Caption = Trim$(Spell(scrlSpell.Value).Name)
    'lblSpell.Caption = CStr(scrlSpell.Value)
End Sub
