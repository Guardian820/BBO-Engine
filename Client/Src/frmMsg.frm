VERSION 5.00
Begin VB.Form frmMsg 
   Caption         =   "Bomber ! Bomber ! Online"
   ClientHeight    =   2550
   ClientLeft      =   7320
   ClientTop       =   6180
   ClientWidth     =   4500
   Icon            =   "frmMsg.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmMsg.frx":0CCA
   ScaleHeight     =   2550
   ScaleWidth      =   4500
   Begin VB.Label lblCloseMsg 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   1440
      TabIndex        =   2
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Label lblAtenção 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   1320
      TabIndex        =   1
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label lblMsg 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Width           =   3855
   End
End
Attribute VB_Name = "frmMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub lblCloseMsg_Click()
frmMsg.Visible = False
End Sub
