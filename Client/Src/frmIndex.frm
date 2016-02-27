VERSION 5.00
Begin VB.Form frmIndex 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Index"
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5295
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
   ScaleHeight     =   4215
   ScaleWidth      =   5295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmddelete 
      Caption         =   "Deletar"
      Height          =   375
      Left            =   1920
      TabIndex        =   3
      Top             =   3720
      Width           =   1575
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   3720
      TabIndex        =   2
      Top             =   3720
      Width           =   1455
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   3720
      Width           =   1575
   End
   Begin VB.ListBox lstIndex 
      Height          =   3420
      ItemData        =   "frmIndex.frx":0000
      Left            =   120
      List            =   "frmIndex.frx":0002
      TabIndex        =   0
      Top             =   120
      Width           =   5055
   End
End
Attribute VB_Name = "frmIndex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ******************************************
' **              BMO Source              **
' ******************************************

Private Sub cmddelete_Click()
    EditorIndex = lstIndex.ListIndex + 1
    Call SendData(CDelete & SEP_CHAR & Editor & SEP_CHAR & EditorIndex & END_CHAR)
End Sub

Private Sub cmdOk_Click()
    EditorIndex = lstIndex.ListIndex + 1
    
    Select Case Editor
    
        Case EDITOR_ITEM
            Call SendData(CEditItem & SEP_CHAR & EditorIndex & END_CHAR)
        Case EDITOR_NPC
            Call SendData(CEditNpc & SEP_CHAR & EditorIndex & END_CHAR)
        Case EDITOR_SHOP
            Call SendData(CEditShop & SEP_CHAR & EditorIndex & END_CHAR)
        Case EDITOR_SPELL
            Call SendData(CEditSpell & SEP_CHAR & EditorIndex & END_CHAR)
            
    End Select

    Unload frmIndex
End Sub

Private Sub cmdCancel_Click()
    Editor = 0
    Unload frmIndex
End Sub
