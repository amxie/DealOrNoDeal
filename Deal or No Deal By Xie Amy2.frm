VERSION 5.00
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   Caption         =   "Deal OR No Deal"
   ClientHeight    =   9030
   ClientLeft      =   4320
   ClientTop       =   2610
   ClientWidth     =   11790
   FillColor       =   &H00FFFFFF&
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Copperplate Gothic Light"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   9030
   ScaleWidth      =   11790
   Begin VB.Frame fraBank 
      BackColor       =   &H00000000&
      ForeColor       =   &H00000000&
      Height          =   5775
      Left            =   2640
      TabIndex        =   63
      Top             =   1200
      Visible         =   0   'False
      Width           =   6615
      Begin VB.Label lblNextRound 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Copperplate Gothic Light"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1215
         Left            =   240
         TabIndex        =   65
         Top             =   3360
         Width           =   6135
      End
      Begin VB.Image imgNo 
         Height          =   720
         Left            =   3840
         Picture         =   "Deal or No Deal By Xie Amy2.frx":0000
         Stretch         =   -1  'True
         Top             =   4680
         Visible         =   0   'False
         Width           =   1905
      End
      Begin VB.Image imgYes 
         Height          =   735
         Left            =   720
         Picture         =   "Deal or No Deal By Xie Amy2.frx":6637
         Stretch         =   -1  'True
         Top             =   4680
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Image imgNoDeal 
         Height          =   795
         Left            =   3240
         Picture         =   "Deal or No Deal By Xie Amy2.frx":D117
         Stretch         =   -1  'True
         Top             =   4680
         Width           =   3015
      End
      Begin VB.Image imgDeal 
         Height          =   795
         Left            =   360
         Picture         =   "Deal or No Deal By Xie Amy2.frx":18930
         Stretch         =   -1  'True
         Top             =   4680
         Width           =   2415
      End
      Begin VB.Label lblOffer 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Copperplate Gothic Light"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   3015
         Left            =   240
         TabIndex        =   64
         Top             =   360
         Width           =   6135
      End
   End
   Begin VB.CommandButton cmdBriefcase 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   0
      Left            =   1440
      TabIndex        =   62
      Top             =   240
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Frame fraHold 
      BackColor       =   &H00000000&
      Caption         =   "Cases"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Light"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Index           =   3
      Left            =   9480
      TabIndex        =   34
      Top             =   7800
      Width           =   2175
      Begin VB.Label lblRemaining 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Copperplate Gothic Light"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1560
         TabIndex        =   38
         Top             =   600
         Width           =   495
      End
      Begin VB.Label lblMessage 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Remaining:"
         BeginProperty Font 
            Name            =   "Copperplate Gothic Bold"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   37
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label lblOpened 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Copperplate Gothic Light"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1560
         TabIndex        =   36
         Top             =   240
         Width           =   495
      End
      Begin VB.Label lblMessage 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Opened:"
         BeginProperty Font 
            Name            =   "Copperplate Gothic Bold"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   35
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame fraHold 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6375
      Index           =   2
      Left            =   9720
      TabIndex        =   2
      Top             =   1200
      Width           =   1935
      Begin VB.Label lblMoney 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Height          =   255
         Index           =   26
         Left            =   120
         TabIndex        =   60
         Top             =   6000
         Width           =   1695
      End
      Begin VB.Label lblMoney 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Height          =   255
         Index           =   25
         Left            =   120
         TabIndex        =   59
         Top             =   5520
         Width           =   1695
      End
      Begin VB.Label lblMoney 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Height          =   255
         Index           =   24
         Left            =   120
         TabIndex        =   58
         Top             =   5040
         Width           =   1695
      End
      Begin VB.Label lblMoney 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Height          =   255
         Index           =   23
         Left            =   120
         TabIndex        =   57
         Top             =   4560
         Width           =   1695
      End
      Begin VB.Label lblMoney 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Height          =   255
         Index           =   22
         Left            =   120
         TabIndex        =   56
         Top             =   4080
         Width           =   1695
      End
      Begin VB.Label lblMoney 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Height          =   255
         Index           =   21
         Left            =   120
         TabIndex        =   55
         Top             =   3600
         Width           =   1695
      End
      Begin VB.Label lblMoney 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Height          =   255
         Index           =   20
         Left            =   120
         TabIndex        =   54
         Top             =   3120
         Width           =   1695
      End
      Begin VB.Label lblMoney 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Height          =   255
         Index           =   19
         Left            =   120
         TabIndex        =   53
         Top             =   2670
         Width           =   1695
      End
      Begin VB.Label lblMoney 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Height          =   255
         Index           =   18
         Left            =   120
         TabIndex        =   52
         Top             =   2160
         Width           =   1695
      End
      Begin VB.Label lblMoney 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Height          =   255
         Index           =   17
         Left            =   120
         TabIndex        =   51
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label lblMoney 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Height          =   255
         Index           =   16
         Left            =   120
         TabIndex        =   50
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label lblMoney 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Height          =   255
         Index           =   15
         Left            =   120
         TabIndex        =   49
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label lblMoney 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         ForeColor       =   &H80000011&
         Height          =   255
         Index           =   14
         Left            =   120
         TabIndex        =   48
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame fraMain 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5775
      Left            =   2640
      TabIndex        =   1
      Top             =   1200
      Width           =   6615
      Begin VB.CommandButton cmdBriefcase 
         Caption         =   "26"
         BeginProperty Font 
            Name            =   "Copperplate Gothic Bold"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   26
         Left            =   5400
         TabIndex        =   61
         Top             =   4680
         Width           =   855
      End
      Begin VB.CommandButton cmdBriefcase 
         BackColor       =   &H80000007&
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "Copperplate Gothic Bold"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   9
         Left            =   4140
         TabIndex        =   46
         Top             =   1440
         Width           =   855
      End
      Begin VB.CommandButton cmdBriefcase 
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "Copperplate Gothic Bold"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   8
         Left            =   2880
         TabIndex        =   45
         Top             =   1440
         Width           =   855
      End
      Begin VB.CommandButton cmdBriefcase 
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "Copperplate Gothic Bold"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   7
         Left            =   1620
         TabIndex        =   44
         Top             =   1440
         Width           =   855
      End
      Begin VB.CommandButton cmdBriefcase 
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "Copperplate Gothic Bold"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   6
         Left            =   360
         TabIndex        =   43
         Top             =   1440
         Width           =   855
      End
      Begin VB.CommandButton cmdBriefcase 
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Copperplate Gothic Bold"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   4
         Left            =   4140
         TabIndex        =   42
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton cmdBriefcase 
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "Copperplate Gothic Bold"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   3
         Left            =   2880
         TabIndex        =   41
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton cmdBriefcase 
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Copperplate Gothic Bold"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   2
         Left            =   1620
         TabIndex        =   40
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton cmdBriefcase 
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Copperplate Gothic Bold"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   1
         Left            =   360
         TabIndex        =   39
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton cmdBriefcase 
         Caption         =   "25"
         BeginProperty Font 
            Name            =   "Copperplate Gothic Bold"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   25
         Left            =   4140
         TabIndex        =   32
         Top             =   4680
         Width           =   855
      End
      Begin VB.CommandButton cmdBriefcase 
         Caption         =   "24"
         BeginProperty Font 
            Name            =   "Copperplate Gothic Bold"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   24
         Left            =   2880
         TabIndex        =   31
         Top             =   4680
         Width           =   855
      End
      Begin VB.CommandButton cmdBriefcase 
         Caption         =   "23"
         BeginProperty Font 
            Name            =   "Copperplate Gothic Bold"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   23
         Left            =   1620
         TabIndex        =   30
         Top             =   4680
         Width           =   855
      End
      Begin VB.CommandButton cmdBriefcase 
         Caption         =   "22"
         BeginProperty Font 
            Name            =   "Copperplate Gothic Bold"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   22
         Left            =   360
         TabIndex        =   29
         Top             =   4680
         Width           =   855
      End
      Begin VB.CommandButton cmdBriefcase 
         Caption         =   "21"
         BeginProperty Font 
            Name            =   "Copperplate Gothic Bold"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   21
         Left            =   5400
         TabIndex        =   28
         Top             =   3600
         Width           =   855
      End
      Begin VB.CommandButton cmdBriefcase 
         BackColor       =   &H80000008&
         Caption         =   "20"
         BeginProperty Font 
            Name            =   "Copperplate Gothic Bold"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   20
         Left            =   4140
         MaskColor       =   &H00000000&
         TabIndex        =   27
         Top             =   3600
         Width           =   855
      End
      Begin VB.CommandButton cmdBriefcase 
         Caption         =   "19"
         BeginProperty Font 
            Name            =   "Copperplate Gothic Bold"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   19
         Left            =   2880
         TabIndex        =   26
         Top             =   3600
         Width           =   855
      End
      Begin VB.CommandButton cmdBriefcase 
         Caption         =   "18"
         BeginProperty Font 
            Name            =   "Copperplate Gothic Bold"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   18
         Left            =   1620
         TabIndex        =   25
         Top             =   3600
         Width           =   855
      End
      Begin VB.CommandButton cmdBriefcase 
         Caption         =   "17"
         BeginProperty Font 
            Name            =   "Copperplate Gothic Bold"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   17
         Left            =   360
         TabIndex        =   24
         Top             =   3600
         Width           =   855
      End
      Begin VB.CommandButton cmdBriefcase 
         Caption         =   "16"
         BeginProperty Font 
            Name            =   "Copperplate Gothic Bold"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   16
         Left            =   5520
         TabIndex        =   23
         Top             =   2520
         Width           =   855
      End
      Begin VB.CommandButton cmdBriefcase 
         Caption         =   "15"
         BeginProperty Font 
            Name            =   "Copperplate Gothic Bold"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   15
         Left            =   4464
         TabIndex        =   22
         Top             =   2520
         Width           =   855
      End
      Begin VB.CommandButton cmdBriefcase 
         Caption         =   "14"
         BeginProperty Font 
            Name            =   "Copperplate Gothic Bold"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   14
         Left            =   3408
         TabIndex        =   21
         Top             =   2520
         Width           =   855
      End
      Begin VB.CommandButton cmdBriefcase 
         Caption         =   "13"
         BeginProperty Font 
            Name            =   "Copperplate Gothic Bold"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   13
         Left            =   2352
         TabIndex        =   20
         Top             =   2520
         Width           =   855
      End
      Begin VB.CommandButton cmdBriefcase 
         Caption         =   "12"
         BeginProperty Font 
            Name            =   "Copperplate Gothic Bold"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   12
         Left            =   1296
         TabIndex        =   19
         Top             =   2520
         Width           =   855
      End
      Begin VB.CommandButton cmdBriefcase 
         Caption         =   "11"
         BeginProperty Font 
            Name            =   "Copperplate Gothic Bold"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   11
         Left            =   240
         TabIndex        =   18
         Top             =   2520
         Width           =   855
      End
      Begin VB.CommandButton cmdBriefcase 
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "Copperplate Gothic Bold"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   10
         Left            =   5400
         TabIndex        =   17
         Top             =   1440
         Width           =   855
      End
      Begin VB.CommandButton cmdBriefcase 
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "Copperplate Gothic Bold"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   5
         Left            =   5400
         TabIndex        =   16
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame fraHold 
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6375
      Index           =   1
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   1935
      Begin VB.Label lblMoney 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Height          =   255
         Index           =   13
         Left            =   120
         TabIndex        =   47
         Top             =   6000
         Width           =   1695
      End
      Begin VB.Label lblMoney 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Height          =   255
         Index           =   12
         Left            =   120
         TabIndex        =   15
         Top             =   5520
         Width           =   1695
      End
      Begin VB.Label lblMoney 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Height          =   255
         Index           =   11
         Left            =   120
         TabIndex        =   14
         Top             =   5040
         Width           =   1695
      End
      Begin VB.Label lblMoney 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Height          =   255
         Index           =   10
         Left            =   120
         TabIndex        =   13
         Top             =   4560
         Width           =   1695
      End
      Begin VB.Label lblMoney 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   12
         Top             =   4080
         Width           =   1695
      End
      Begin VB.Label lblMoney 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   11
         Top             =   3600
         Width           =   1695
      End
      Begin VB.Label lblMoney 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   10
         Top             =   3120
         Width           =   1695
      End
      Begin VB.Label lblMoney 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   9
         Top             =   2640
         Width           =   1695
      End
      Begin VB.Label lblMoney 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   8
         Top             =   2160
         Width           =   1695
      End
      Begin VB.Label lblMoney 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   7
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label lblMoney 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   6
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label lblMoney 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label lblMoney 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Label lblSpeaker 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Copperplate Gothic Light"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   1695
      Left            =   2760
      TabIndex        =   33
      Top             =   7200
      Width           =   6375
   End
   Begin VB.Label lblMessage 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      BackStyle       =   0  'Transparent
      Caption         =   "Your Briefcase"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Light"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   795
      Left            =   2737
      Picture         =   "Deal or No Deal By Xie Amy2.frx":21B2F
      Stretch         =   -1  'True
      Top             =   240
      Width           =   6420
   End
   Begin VB.Menu mnuSettings 
      Caption         =   "S&ettings"
      Begin VB.Menu mnuTheme 
         Caption         =   "Th&emes"
         Begin VB.Menu mnuNight 
            Caption         =   "N&ight Mode"
            Checked         =   -1  'True
            Shortcut        =   ^N
         End
         Begin VB.Menu mnuDay 
            Caption         =   "D&ay Mode"
            Shortcut        =   ^D
         End
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "Abo&ut..."
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnuSeperater 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
         Shortcut        =   {F12}
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Programmer: Xie, Amy
' Date: May 16, 2019
' Purpose: To allow players to enjoy a game of "Deal or No Deal".

Const MAX = 26
Dim InOrderMoney(1 To MAX) As Long
Dim RandomizedMoney(1 To MAX) As Long
Dim First As Boolean
Dim Offering As Boolean
Dim FileName As String
Dim YourMoney As Long
Dim Offer As Single
Dim Round As Integer
Dim CounterOpened As Integer
Dim CounterRem As Integer
Dim StartingCases As Integer
Dim CasesLeft As Integer
Dim X As Integer

Private Sub cmdBriefcase_Click(Index As Integer)
    Dim PName As String
    Dim NextRoundMsg As String
    Dim OfferMsg As String
    Dim Response As Integer
    
    Dim X As Integer
    
'   Resets the response for the remaining as "no" to count down the remaining briefcases.
    Response = vbNo
'   So if the player is being given a offer from the bank, there is no round displayed at the bottom.
    Offering = False
    
'   Allows player to choose their briefcase.
    If First = True Then
        cmdBriefcase(Index).Visible = False
        cmdBriefcase(0).Visible = True
        cmdBriefcase(0).Caption = cmdBriefcase(Index).Caption
        YourMoney = RandomizedMoney(Index)
        First = False
        Round = Round + 1
        CasesLeft = 6
        StartingCases = 6
    Else
'       Removes the money in the scoreboard that is assigned to briefcase.
        Search InOrderMoney(), RandomizedMoney(Index), MAX
        cmdBriefcase(Index).Visible = False
        CounterOpened = CounterOpened + 1
        CasesLeft = CasesLeft - 1
        
'       Sets up the different rounds.
        If CasesLeft = 0 Then
'           Calculates bank offer and returns the offer value to 'Offer'.
            Offering = True
            Offer = BankOffer(InOrderMoney(), MAX, Round)
            Round = Round + 1
'           To have the decreasing amount of cases to select.
            If StartingCases > 1 Then
                StartingCases = StartingCases - 1
            End If
'           Allows user to select their briefcase.
            If CounterOpened = 24 Then
                OfferMsg = "You have opened " & Str$(CounterOpened) & " briefcases!" & vbCrLf & vbCrLf & vbCrLf & "your bank offer is: " & vbCrLf & Format$(Offer, "$##,###,###")
                NextRoundMsg = "This is the last bank offer!" & vbCrLf & "You will have to open your briefcase, if declined."
            ElseIf CounterOpened = 25 Then
                OfferMsg = vbCrLf & "Your briefcase contained:" & vbCrLf & vbCrLf & vbCrLf & Format$(YourMoney, "$##,###,###")
                NextRoundMsg = "Would you like to play again?"
                imgDeal.Visible = False
                imgNoDeal.Visible = False
                imgYes.Visible = True
                imgNo.Visible = True
            Else
                OfferMsg = "You have opened " & Str$(CounterOpened) & " briefcases!" & vbCrLf & "There are " & Str$(CounterRem - 1) & " briefcases left." & vbCrLf & vbCrLf & vbCrLf & "The bank offer is:" & vbCrLf & Format$(Offer, "$##,###,###")
                NextRoundMsg = "In the next round you will have to open " & Str$(StartingCases) & " briefcases."
            End If
            CasesLeft = StartingCases
        End If
    End If
    
'   Does not decrease an extra remaining caption when it is reseted.
    If Response = vbNo Then
        CounterRem = CounterRem - 1
    End If

'   Displays the bank offer.
    If Offering = True Then
        lblSpeaker.Caption = ""
        lblOffer.Caption = OfferMsg
        lblNextRound.Caption = NextRoundMsg
    End If
'   Only is able to display the change in the counters and the differences in rounds.
    Display Offering, Round, CounterOpened, CounterRem, CasesLeft
    
End Sub

Private Sub Form_Load()
    Randomize
    FileName = App.Path & "/Money.txt"

'   Sets Up Form
    ResetForm First, CounterRem, CounterOpened, Round, MAX
    LoadMoney InOrderMoney(), RandomizedMoney(), FileName
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim Response As Integer
    
    Response = MsgBox("Are you sure you would like to quit?" & vbCrLf & "You will lose all your progress!", vbExclamation + vbYesNo + vbDefaultButton2, "Exit")
    
    If Response = vbNo Then
        Cancel = True
    End If
End Sub

Private Sub imgDeal_Click()
    imgDeal.Visible = False
    imgNoDeal.Visible = False
    imgYes.Visible = True
    imgNo.Visible = True
    
    lblOffer.Caption = "You have accepted the banks offer of:" & vbCrLf & vbCrLf & Format$(Offer, "$##,###,###") & vbCrLf & vbCrLf & "Your briefcase contained: " & Format$(YourMoney, "$##,###,###")
    lblNextRound.Caption = "Would you like to play again?"
End Sub

Private Sub imgNo_Click()
    End
End Sub

Private Sub imgNoDeal_Click()
    fraBank.Visible = False
    fraMain.Visible = True
    Offering = False
    If CounterOpened <> 24 Then
        Display Offering, Round, CounterOpened, CounterRem, CasesLeft
    Else
        LastCase MAX
    End If
End Sub

Private Sub imgYes_Click()
    ResetForm First, CounterRem, CounterOpened, Round, MAX
    RandomizeArray InOrderMoney(), RandomizedMoney(), MAX
End Sub

Private Sub mnuAbout_Click()
    MsgBox "Programmer: Xie, Amy" & vbCrLf & "Date Created: May 2019", vbOKOnly + vbInformation, "About"
End Sub

Private Sub mnuDay_Click()
    Dim X As Integer
    Dim Y As Integer
    
    mnuDay.Checked = True
    mnuNight.Checked = False
    
    For X = 1 To 3
        lblMessage(X).BackColor = vbWhite
        lblMessage(X).ForeColor = vbBlack
        fraHold(X).BackColor = vbWhite
        fraHold(X).ForeColor = vbBlack
    Next X
'   For the number to stay visible to the user when changing between themes in the middle of the game.
    For Y = 1 To 26
        If lblMoney(Y).BackStyle = 0 Then
            lblMoney(Y).ForeColor = vbBlack
        End If
    Next Y
    
    frmMain.BackColor = vbWhite
    fraMain.BackColor = vbWhite
    fraMain.ForeColor = vbBlack
    fraBank.BackColor = vbWhite
    fraBank.ForeColor = vbBlack
    lblOffer.ForeColor = vbBlack
    lblOpened.ForeColor = vbBlack
    lblSpeaker.ForeColor = vbBlack
    lblRemaining.ForeColor = vbBlack
    lblNextRound.ForeColor = vbBlack
End Sub

Private Sub mnuExit_Click()
    Dim Response As Integer
    
    Response = MsgBox("Are you sure you would like to quit?" & vbCrLf & "You will lose all your progress!", vbExclamation + vbYesNo + vbDefaultButton2, "Exit")
    
    If Response = vbYes Then
        End
    End If
End Sub

Private Sub mnuNight_Click()
    Dim X As Integer
    Dim Y As Integer
    Dim Z As Integer

'   To indicate to the user which theme it is in.
    mnuDay.Checked = False
    mnuNight.Checked = True
    
    For X = 1 To 3
        lblMessage(X).BackColor = vbBlack
        lblMessage(X).ForeColor = vbWhite
        fraHold(X).BackColor = vbBlack
        fraHold(X).ForeColor = vbWhite
    Next X
'   For the number to stay visible to the user when changing between themes in the middle of the game.
    For Y = 1 To 26
        If lblMoney(Y).BackStyle = 0 Then
            lblMoney(Y).ForeColor = vbWhite
        End If
    Next Y
    
    frmMain.BackColor = vbBlack
    fraMain.BackColor = vbBlack
    fraMain.ForeColor = vbWhite
    fraBank.BackColor = vbBlack
    fraBank.ForeColor = vbWhite
    lblOffer.ForeColor = vbWhite
    lblOpened.ForeColor = vbWhite
    lblSpeaker.ForeColor = vbWhite
    lblRemaining.ForeColor = vbWhite
    lblNextRound.ForeColor = vbWhite
End Sub
