VERSION 5.00
Begin VB.Form frmHighScore 
   BackColor       =   &H00000000&
   Caption         =   "HighScore"
   ClientHeight    =   5100
   ClientLeft      =   7275
   ClientTop       =   3750
   ClientWidth     =   8565
   BeginProperty Font 
      Name            =   "Copperplate Gothic Light"
      Size            =   14.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   5100
   ScaleWidth      =   8565
   Begin VB.PictureBox picScore 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3735
      Left            =   4200
      ScaleHeight     =   3735
      ScaleWidth      =   3975
      TabIndex        =   4
      Top             =   720
      Width           =   3975
   End
   Begin VB.PictureBox picRank 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   3495
      Left            =   240
      ScaleHeight     =   3495
      ScaleWidth      =   3495
      TabIndex        =   3
      Top             =   840
      Width           =   3495
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "OK"
      Height          =   495
      Left            =   6960
      TabIndex        =   2
      Top             =   4560
      Width           =   1455
   End
   Begin VB.Label lblLabel 
      BackStyle       =   0  'Transparent
      Caption         =   "Score"
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   0
      Left            =   4320
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
   Begin VB.Line Seperater 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   4320
      X2              =   6960
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Line Seperater 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   240
      X2              =   2160
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label lblLabel 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   1
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "frmHighScore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExit_Click()
    Unload frmHighScore
End Sub

Private Sub Form_Load()
    Const MAX = 7
    Dim FileName As String
    Dim PName(1 To MAX) As String
    Dim Score(1 To MAX) As Long
    Dim X As Integer
    Dim Y As Integer
    
    X = 0
    FileName = App.Path & "\Highscore.txt"
    
    Open FileName For Input As #1
    Do While Not EOF(1)
        X = X + 1
        Input #1, PName(X), Score(X)
    Loop
    Close #1
    For Y = 1 To X
        picRank.Print Y; ". "; PName(Y)
        picScore.Print Format$(Score(Y), "$##,###,###")
    Next Y
End Sub
