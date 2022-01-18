VERSION 5.00
Begin VB.Form frmStartUp 
   BackColor       =   &H00000000&
   Caption         =   "Deal Or No Deal"
   ClientHeight    =   5445
   ClientLeft      =   10905
   ClientTop       =   7485
   ClientWidth     =   9195
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Century Schoolbook"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   5445
   ScaleWidth      =   9195
   Begin VB.Image imgExit 
      Height          =   945
      Left            =   5880
      Picture         =   "Deal or No Deal By Xie Amy.frx":0000
      Stretch         =   -1  'True
      Top             =   3720
      Width           =   2340
   End
   Begin VB.Image imgStart 
      Height          =   945
      Left            =   1080
      Picture         =   "Deal or No Deal By Xie Amy.frx":7014
      Stretch         =   -1  'True
      Top             =   3720
      Width           =   2340
   End
   Begin VB.Image Image1 
      Height          =   915
      Left            =   1320
      Picture         =   "Deal or No Deal By Xie Amy.frx":F045
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   6660
   End
End
Attribute VB_Name = "frmStartUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub imgExit_Click()
    End
End Sub

Private Sub imgStart_Click()
    Unload frmStartUp
    frmMain.Show vbModal
End Sub
