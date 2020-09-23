VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6DE6E6DD-C656-11D2-B052-444553540000}#3.0#0"; "VBCARDS.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00008000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Video Draw Poker"
   ClientHeight    =   6255
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   9000
   FillColor       =   &H00008000&
   FillStyle       =   0  'Solid
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6255
   ScaleWidth      =   9000
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdBet1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Bet $1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2490
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   5640
      Width           =   1095
   End
   Begin VB.CommandButton cmdBetMax 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Bet Max"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3810
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5640
      Width           =   1095
   End
   Begin VB.Frame fraWinningHand 
      BackColor       =   &H00FFFF00&
      Height          =   2415
      Left            =   15
      TabIndex        =   11
      Top             =   0
      Width           =   2370
      Begin VB.Label lblPoints 
         BackColor       =   &H00FFFF00&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "   Royal Flush"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   0
         TabIndex        =   35
         Top             =   240
         Width           =   2355
      End
      Begin VB.Label lblPoints 
         BackColor       =   &H00FFFF00&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "   Straight Flush"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   0
         TabIndex        =   34
         Top             =   480
         Width           =   2355
      End
      Begin VB.Label lblPoints 
         BackColor       =   &H00FFFF00&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "   Four of a Kind"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   0
         TabIndex        =   33
         Top             =   720
         Width           =   2355
      End
      Begin VB.Label lblPoints 
         BackColor       =   &H00FFFF00&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "   Full House"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   0
         TabIndex        =   32
         Top             =   960
         Width           =   2355
      End
      Begin VB.Label lblPoints 
         BackColor       =   &H00FFFF00&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "   Flush"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   0
         TabIndex        =   31
         Top             =   1200
         Width           =   2355
      End
      Begin VB.Label lblPoints 
         BackColor       =   &H00FFFF00&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "   Straight"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   0
         TabIndex        =   30
         Top             =   1440
         Width           =   2355
      End
      Begin VB.Label lblPoints 
         BackColor       =   &H00FFFF00&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "   Three of a Kind"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   0
         TabIndex        =   29
         Top             =   1680
         Width           =   2355
      End
      Begin VB.Label lblPoints 
         BackColor       =   &H00FFFF00&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "   Two Pair"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   0
         TabIndex        =   28
         Top             =   1920
         Width           =   2355
      End
      Begin VB.Label lblPoints 
         BackColor       =   &H00FFFF00&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "   One Pair"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   27
         Top             =   2160
         Width           =   2355
      End
      Begin VB.Label lblHandDrawn 
         BackColor       =   &H00FF8080&
         Caption         =   "   Hand Drawn"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   0
         TabIndex        =   20
         Top             =   0
         Width           =   2415
      End
   End
   Begin VB.Frame fraPayout 
      BackColor       =   &H00C0FFFF&
      Height          =   2415
      Index           =   4
      Left            =   7650
      TabIndex        =   10
      Top             =   0
      Width           =   1335
      Begin VB.Label lblRoyalFlush 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1250"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   0
         TabIndex        =   80
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblStrFlush 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "250"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   0
         TabIndex        =   75
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label lbl4Kind 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "125"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   0
         TabIndex        =   70
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label lblFullHouse 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "40"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   0
         TabIndex        =   65
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label lblFlush 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "25"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   0
         TabIndex        =   60
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label lblStraight 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "20"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   0
         TabIndex        =   55
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label lbl3Kind 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "15"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   0
         TabIndex        =   50
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label lbl2Pair 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   0
         TabIndex        =   45
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label lbl1Pair 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   0
         TabIndex        =   40
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label lblWagerAmount 
         Alignment       =   2  'Center
         BackColor       =   &H00FF8080&
         Caption         =   "$5.00 Bet"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   4
         Left            =   0
         TabIndex        =   19
         Top             =   0
         Width           =   1335
      End
   End
   Begin VB.Frame fraPayout 
      BackColor       =   &H00C0FFFF&
      Height          =   2415
      Index           =   3
      Left            =   6330
      TabIndex        =   9
      Top             =   0
      Width           =   1335
      Begin VB.Label lblRoyalFlush 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   0
         TabIndex        =   79
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblStrFlush 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "200"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   0
         TabIndex        =   74
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label lbl4Kind 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "100"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   0
         TabIndex        =   69
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label lblFullHouse 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "32"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   0
         TabIndex        =   64
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label lblFlush 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "20"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   0
         TabIndex        =   59
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label lblStraight 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "16"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   0
         TabIndex        =   54
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label lbl3Kind 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "12"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   0
         TabIndex        =   49
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label lbl2Pair 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   0
         TabIndex        =   44
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label lbl1Pair 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   0
         TabIndex        =   39
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label lblWagerAmount 
         Alignment       =   2  'Center
         BackColor       =   &H00FF8080&
         Caption         =   "$4.00 Bet"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   3
         Left            =   0
         TabIndex        =   18
         Top             =   0
         Width           =   1335
      End
   End
   Begin VB.Frame fraPayout 
      BackColor       =   &H00C0FFFF&
      Height          =   2415
      Index           =   2
      Left            =   5010
      TabIndex        =   8
      Top             =   0
      Width           =   1335
      Begin VB.Label lblRoyalFlush 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "750"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   0
         TabIndex        =   78
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblStrFlush 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "150"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   0
         TabIndex        =   73
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label lbl4Kind 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "75"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   0
         TabIndex        =   68
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label lblFullHouse 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "24"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   0
         TabIndex        =   63
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label lblFlush 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "15"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   0
         TabIndex        =   58
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label lblStraight 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "12"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   0
         TabIndex        =   53
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label lbl3Kind 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   0
         TabIndex        =   48
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label lbl2Pair 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   0
         TabIndex        =   43
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label lbl1Pair 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   0
         TabIndex        =   38
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label lblWagerAmount 
         Alignment       =   2  'Center
         BackColor       =   &H00FF8080&
         Caption         =   "$3.00 Bet"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   2
         Left            =   0
         TabIndex        =   17
         Top             =   0
         Width           =   1335
      End
   End
   Begin VB.Frame fraPayout 
      BackColor       =   &H00C0FFFF&
      Height          =   2415
      Index           =   1
      Left            =   3690
      TabIndex        =   7
      Top             =   0
      Width           =   1335
      Begin VB.Label lblRoyalFlush 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "500"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   0
         TabIndex        =   77
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblStrFlush 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "100"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   0
         TabIndex        =   72
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label lbl4Kind 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "50"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   0
         TabIndex        =   67
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label lblFullHouse 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "16"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   0
         TabIndex        =   62
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label lblFlush 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   0
         TabIndex        =   57
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label lblStraight 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   0
         TabIndex        =   52
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label lbl3Kind 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   0
         TabIndex        =   47
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label lbl2Pair 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   0
         TabIndex        =   42
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label lbl1Pair 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   0
         TabIndex        =   37
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label lblWagerAmount 
         Alignment       =   2  'Center
         BackColor       =   &H00FF8080&
         Caption         =   "$2.00 Bet"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   1
         Left            =   0
         TabIndex        =   16
         Top             =   0
         Width           =   1335
      End
   End
   Begin VB.Frame fraPayout 
      BackColor       =   &H00C0FFFF&
      Height          =   2415
      Index           =   0
      Left            =   2370
      TabIndex        =   6
      Top             =   0
      Width           =   1335
      Begin VB.Label lblRoyalFlush 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "250"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   76
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblStrFlush 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "50"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   71
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label lbl4Kind 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "25"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   66
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label lblFullHouse 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   61
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label lblFlush 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   56
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label lblStraight 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   51
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label lbl3Kind 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   46
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label lbl2Pair 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   41
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label lbl1Pair 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   36
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label lblWagerAmount 
         Alignment       =   2  'Center
         BackColor       =   &H00FF8080&
         Caption         =   "$1.00 Bet"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   15
         Top             =   0
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmdHold 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Hold"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   4
      Left            =   7770
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4800
      Width           =   1095
   End
   Begin VB.CommandButton cmdHold 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Hold"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   6450
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4800
      Width           =   1095
   End
   Begin VB.CommandButton cmdHold 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Hold"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   5130
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4800
      Width           =   1095
   End
   Begin VB.CommandButton cmdHold 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Hold"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   3810
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4800
      Width           =   1095
   End
   Begin VB.CommandButton cmdHold 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Hold"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   2490
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4800
      Width           =   1095
   End
   Begin VB.CommandButton cmdDeal 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7770
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5640
      Width           =   1095
   End
   Begin VBCards.Deck Deck1 
      Left            =   60
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   1032
      Picture         =   "frmMain.frx":0A8A
   End
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   60
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picCasino 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   3840
      Left            =   0
      Picture         =   "frmMain.frx":29DC
      ScaleHeight     =   3840
      ScaleWidth      =   2370
      TabIndex        =   21
      Top             =   2415
      Width           =   2370
   End
   Begin VB.Label lblMsg 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0FF&
      Height          =   375
      Left            =   2400
      TabIndex        =   81
      Top             =   2400
      Width           =   6615
   End
   Begin VB.Label lblHold 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      Caption         =   "HOLD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   360
      Index           =   4
      Left            =   7770
      TabIndex        =   26
      Top             =   2760
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label lblHold 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      Caption         =   "HOLD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   360
      Index           =   3
      Left            =   6450
      TabIndex        =   25
      Top             =   2760
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label lblHold 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      Caption         =   "HOLD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   360
      Index           =   2
      Left            =   5130
      TabIndex        =   24
      Top             =   2760
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label lblHold 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      Caption         =   "HOLD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   360
      Index           =   1
      Left            =   3810
      TabIndex        =   23
      Top             =   2760
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label lblHold 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      Caption         =   "HOLD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   360
      Index           =   0
      Left            =   2490
      TabIndex        =   22
      Top             =   2760
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label lblPurse 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   5250
      TabIndex        =   14
      Top             =   5640
      Width           =   2175
   End
   Begin VB.Image imgHand 
      Height          =   1455
      Index           =   4
      Left            =   7770
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Image imgHand 
      Height          =   1455
      Index           =   3
      Left            =   6450
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Image imgHand 
      Height          =   1455
      Index           =   2
      Left            =   5130
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Image imgHand 
      Height          =   1455
      Index           =   1
      Left            =   3810
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Image imgHand 
      Height          =   1455
      Index           =   0
      Left            =   2490
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Shape shpPurseBox 
      BorderColor     =   &H00E0E0E0&
      BorderWidth     =   2
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   5130
      Top             =   5520
      Width           =   2415
   End
   Begin VB.Shape shpHold 
      BorderColor     =   &H000000C0&
      BorderWidth     =   3
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   1560
      Index           =   0
      Left            =   2430
      Top             =   3180
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Shape shpHold 
      BorderColor     =   &H000000C0&
      BorderWidth     =   3
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   1560
      Index           =   3
      Left            =   6390
      Top             =   3180
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Shape shpHold 
      BorderColor     =   &H000000C0&
      BorderWidth     =   3
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   1560
      Index           =   2
      Left            =   5070
      Top             =   3180
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Shape shpHold 
      BorderColor     =   &H000000C0&
      BorderWidth     =   3
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   1560
      Index           =   1
      Left            =   3750
      Top             =   3180
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Shape shpHold 
      BorderColor     =   &H000000C0&
      BorderWidth     =   3
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   1560
      Index           =   4
      Left            =   7710
      Top             =   3180
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Menu mnuGame 
      Caption         =   "&Game"
      Begin VB.Menu mnuGameNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuGameOptions 
         Caption         =   "&Options..."
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGameExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpContents 
         Caption         =   "&Contents..."
      End
      Begin VB.Menu mnuHelpBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About Video Poker..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function OSWinHelp% Lib "user32" Alias "WinHelpA" (ByVal hwnd&, ByVal HelpFile$, ByVal wCommand%, dwData As Any)
Public x, y, CountBet, HoldBackValue As Integer
Dim FrameToChange, TotalBet As Integer
Public FrameColor1, FrameColor2, LabelColor1, LabelColor2, TextColor1, TextColor2

Private Sub cmdBet1_Click()
    lblMsg.Caption = ""
    CountBet = CountBet + 1
    TotalBet = CountBet
    lblPurse.Caption = Format(lblPurse.Caption - 1, "$###,##0.00")
    If lblPurse.Caption + 0 = 0 Then
        cmdBet1.Enabled = False
        cmdBetMax.Enabled = False
    End If
    If cmdDeal.Caption = "Draw" Then
        cmdDeal.Caption = "Deal"
        ResetHandColor
        Deck1.ChangeCard = 55
        For x = 0 To 4
            fraPayout(x).BackColor = FrameColor1
            lblWagerAmount(x).BackColor = LabelColor1
            lblWagerAmount(x).ForeColor = TextColor1
            lblHold(x).Visible = False
            shpHold(x).Visible = False
            imgHand(x).Picture = Deck1.Picture
        Next x
        x = 0
    End If
    If x = 0 Then
        ShuffleDeck 'Shuffle the cards
        'Highlight the $1 payout frame
        fraPayout(x).BackColor = FrameColor2
        lblWagerAmount(x).BackColor = LabelColor2
        lblWagerAmount(x).ForeColor = TextColor2
        cmdDeal.Enabled = True
        x = 1
    Else
        'Move the highlight to the next payout frame
        fraPayout(x).BackColor = FrameColor2
        lblWagerAmount(x).BackColor = LabelColor2
        lblWagerAmount(x).ForeColor = TextColor2
        fraPayout(x - 1).BackColor = FrameColor1
        lblWagerAmount(x - 1).BackColor = LabelColor1
        lblWagerAmount(x - 1).ForeColor = TextColor1
        x = x + 1
        'Player can't wager more than $5 per game
        If x = 5 Then
            cmdBet1.Enabled = False
            cmdBetMax.Enabled = False
        End If
    End If
End Sub

Private Sub cmdBetMax_Click()
    FrameToChange = 0
    ShuffleDeck
    ResetHandColor
    If cmdDeal.Caption = "Draw" Then
        cmdDeal.Caption = "Deal"
    End If
    Deck1.ChangeCard = 55
    'Resets the color for the $1-$5 payout frames
    For x = 0 To 4
        fraPayout(x).BackColor = FrameColor1
        lblWagerAmount(x).BackColor = LabelColor1
        lblWagerAmount(x).ForeColor = TextColor1
        lblHold(x).Visible = False
        shpHold(x).Visible = False
        imgHand(x).Picture = Deck1.Picture
    Next x
        x = 0

    'Disables the cmdBet1 button
    cmdBet1.Enabled = False
    'Disables the cmdBetMax Button
    cmdBetMax.Enabled = False
    'Enables the Deal Button
    cmdDeal.Enabled = True
    'Place the complete wager
    For x = 1 To (5 - CountBet)
        If lblPurse.Caption + 0 <= 0 Then
            FrameToChange = FrameToChange + CountBet
            'Changes the color for the maximum payout frame
            fraPayout(FrameToChange - 1).BackColor = FrameColor2
            lblWagerAmount(FrameToChange - 1).BackColor = LabelColor2
            lblWagerAmount(FrameToChange - 1).ForeColor = TextColor2
            Exit Sub
        End If
        TotalBet = TotalBet + 1
        FrameToChange = FrameToChange + 1
        lblPurse.Caption = Format(lblPurse.Caption - 1, "$###,##0.00")
    Next x
    'Changes the color for the maximum payout frame
    fraPayout(4).BackColor = FrameColor2
    lblWagerAmount(4).BackColor = LabelColor2
    lblWagerAmount(4).ForeColor = TextColor2
    CountBet = 0
    cmdDeal_Click
End Sub

Private Sub cmdDeal_Click()
Dim xx As Integer
    If cmdDeal.Caption = "Deal" Then
        'Disable the wagering buttons
        cmdBet1.Enabled = False
        cmdBetMax.Enabled = False
        'Draw the initial hand to play
        For x = 0 To 4
            Deck1.ChangeCard = Card(x)
            imgHand(x).Picture = Deck1.Picture
            cmdHold(x).Enabled = True
            SortVariable(x) = Card(x)
        Next x
        cmdDeal.Caption = "Draw"
        BubbleSort
        ReadHand
    Else
        'Draw the replacement set of cards
        xx = 5
        For x = 0 To 4
            If lblHold(x).Visible = True Then
                SortVariable(x) = Card(x)
            Else
                If lblHold(x).Visible = False Then
                    Deck1.ChangeCard = Card(xx)
                    imgHand(x).Picture = Deck1.Picture
                    SortVariable(x) = Card(xx)
                    xx = xx + 1
                End If
            End If
        Next x
        BubbleSort
        'Disable the cmdHold buttons
        x = 0
        For x = 0 To 4
            cmdHold(x).Enabled = False
        Next x
    'If a winning hand, this code calls the procedure
    'that pays the winner the prize.
        'Enable the wagering buttons if the purse > 0
        cmdDeal.Enabled = False
        If lblPurse.Caption + 0 <= 0 Then
            cmdBet1.Enabled = False
            cmdBetMax.Enabled = False
        Else
            CountBet = 0
            FrameToChange = 0
            cmdBet1.Enabled = True
            cmdBetMax.Enabled = True
        End If
        ReadHand
        PayThePlayer
        TotalBet = 0
    End If
'    ReadHand
    ResetHandColor
    ColorHandDrawn
End Sub

Private Sub cmdHold_Click(Index As Integer)
    If lblHold(Index).Visible = False Then
        lblHold(Index).Visible = True
        shpHold(Index).Visible = True
    Else
        lblHold(Index).Visible = False
        shpHold(Index).Visible = False
    End If
End Sub

Private Sub Form_Load()
    Me.Left = 105
    Me.Top = -180
    Me.Width = 9090
    Me.Height = 6915
    
    FrameColor1 = (&HC0FFFF)
    FrameColor2 = (&HFFC0FF)
    LabelColor1 = (&HFF8080)
    LabelColor2 = (&HFF0000)
    TextColor1 = (&HFFFF&)
    TextColor2 = (&HC0FFC0)
    cmdDeal.Caption = "Deal"
    lblPurse.Caption = 25
    Deck1.ChangeCard = 55
    lblPurse.Caption = Format(lblPurse.Caption, "$###,##0.00")
    For x = 0 To 4
        imgHand(x).Picture = Deck1.Picture
        lbl1Pair(x).Caption = Format(lbl1Pair(x).Caption, "$###,##0.00")
        lbl2Pair(x).Caption = Format(lbl2Pair(x).Caption, "$###,##0.00")
        lbl3Kind(x).Caption = Format(lbl3Kind(x).Caption, "$###,##0.00")
        lblStraight(x).Caption = Format(lblStraight(x).Caption, "$###,##0.00")
        lblFlush(x).Caption = Format(lblFlush(x).Caption, "$###,##0.00")
        lblFullHouse(x).Caption = Format(lblFullHouse(x), "$###,##0.00")
        lbl4Kind(x).Caption = Format(lbl4Kind(x).Caption, "$###,##0.00")
        lblStrFlush(x).Caption = Format(lblStrFlush(x).Caption, "$###,##0.00")
        lblRoyalFlush(x).Caption = Format(lblRoyalFlush(x).Caption, "$###,##0.00")
    Next x
    x = 0
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Dim i As Integer
    'close all sub forms
    For i = Forms.Count - 1 To 1 Step -1
        Unload Forms(i)
    Next
    End
End Sub

Private Sub mnuHelpAbout_Click()
    frmAbout.Show vbModal, Me
End Sub

Private Sub mnuHelpContents_Click()
    Dim nRet As Integer
    'if there is no helpfile for this project display a message to the user
    'you can set the HelpFile for your application in the
    'Project Properties dialog
    If Len(App.HelpFile) = 0 Then
        MsgBox "Unable to display Help Contents. There is no Help associated with this project.", vbInformation, Me.Caption
    Else
        On Error Resume Next
        nRet = OSWinHelp(Me.hwnd, App.HelpFile, 3, 0)
        If Err Then
            MsgBox Err.Description
        End If
    End If

End Sub

Private Sub mnuGameExit_Click()
    'unload the form
    Unload Me
    End
End Sub

Private Sub mnuGameNew_Click()
ResetHandColor
lblMsg.Caption = ""
cmdBet1.Enabled = True
cmdBetMax.Enabled = True
cmdDeal.Enabled = False
cmdDeal.Caption = "Deal"
Deck1.ChangeCard = 55
lblPurse.Caption = Format(25, "$###,##0.00")
For x = 0 To 4
    cmdHold(x).Enabled = False 'True
    lblHold(x).Visible = False
    shpHold(x).Visible = False
    imgHand(x).Picture = Deck1.Picture
    fraPayout(x).BackColor = FrameColor1
    lblWagerAmount(x).BackColor = LabelColor1
    lblWagerAmount(x).ForeColor = TextColor1
Next x
x = 0
y = 0
CountBet = 0
FrameToChange = 0
TotalBet = 0
End Sub

Public Sub ReadHand()
Dim HandValue As Boolean
    HandValue = DoIsFlush()
    If HandValue = True Then
        ReduceCards
        HandValue = DoIsRoyal()
        If HandValue = True Then
            lblMsg.Caption = "Royal Flush"
            Exit Sub
        End If
        HandValue = DoIsStraight()
        If HandValue = True Then
            lblMsg.Caption = "Straight Flush"
            Exit Sub
        End If
        lblMsg.Caption = "Flush"
        Exit Sub
    End If
    ReduceCards
    BubbleSort
    HandValue = DoIsRoyal()
    If HandValue = True Then
        lblMsg.Caption = "Straight"
        Exit Sub
    End If
    HandValue = DoIsStraight()
    If HandValue = True Then
        lblMsg.Caption = "Straight"
        Exit Sub
    End If
    HandValue = DoIs4Kind()
    If HandValue = True Then
        lblMsg.Caption = "Four of a Kind"
        Exit Sub
    End If
    HandValue = DoIs3Kind()
    If HandValue = True Then
        HandValue = DoIsFullHouse
        If HandValue = True Then
            lblMsg.Caption = "Full House"
            Exit Sub
        End If
        lblMsg.Caption = "Three of a Kind"
        Exit Sub
    End If
    HandValue = DoIsPair()
    If HandValue = True Then
        HandValue = DoIsTwoPair()
        If HandValue = True Then
            lblMsg.Caption = "Two Pair"
            Exit Sub
        End If
        lblMsg.Caption = "One Pair"
        Exit Sub
    End If
    lblMsg.Caption = ""
End Sub

Public Sub ResetHandColor()
    For y = 0 To 8
        lblPoints(y).BackColor = (&HFFFF00)
        lblPoints(y).ForeColor = (&H80000012)
    Next y
    For y = 0 To 4
        lbl1Pair(y).BackColor = (&HC0FFFF)
        lbl2Pair(y).BackColor = (&HC0FFFF)
        lbl3Kind(y).BackColor = (&HC0FFFF)
        lblStraight(y).BackColor = (&HC0FFFF)
        lblFlush(y).BackColor = (&HC0FFFF)
        lblFullHouse(y).BackColor = (&HC0FFFF)
        lbl4Kind(y).BackColor = (&HC0FFFF)
        lblStrFlush(y).BackColor = (&HC0FFFF)
        lblRoyalFlush(y).BackColor = (&HC0FFFF)
        lbl1Pair(y).BackStyle = 0
        lbl2Pair(y).BackStyle = 0
        lbl3Kind(y).BackStyle = 0
        lblStraight(y).BackStyle = 0
        lblFlush(y).BackStyle = 0
        lblFullHouse(y).BackStyle = 0
        lbl4Kind(y).BackStyle = 0
        lblStrFlush(y).BackStyle = 0
        lblRoyalFlush(y).BackStyle = 0
    Next y
End Sub

Public Sub ColorHandDrawn()
Dim Dummy As String
Dummy = lblMsg.Caption
    Select Case Dummy
    Case "One Pair":
        lblPoints(0).BackStyle = 1
        lblPoints(0).BackColor = (&HC0FFFF)
        For y = 0 To 4
            lbl1Pair(y).BackStyle = 1
            lbl1Pair(y).BackColor = (&HFFFF00)
        Next y
    Case "Two Pair":
        lblPoints(1).BackStyle = 1
        lblPoints(1).BackColor = (&HC0FFFF)
        For y = 0 To 4
            lbl2Pair(y).BackStyle = 1
            lbl2Pair(y).BackColor = (&HFFFF00)
        Next y
    Case "Three of a Kind":
        lblPoints(2).BackStyle = 1
        lblPoints(2).BackColor = (&HC0FFFF)
        For y = 0 To 4
            lbl3Kind(y).BackStyle = 1
            lbl3Kind(y).BackColor = (&HFFFF00)
        Next y
    Case "Straight":
        lblPoints(3).BackStyle = 1
        lblPoints(3).BackColor = (&HC0FFFF)
        For y = 0 To 4
            lblStraight(y).BackStyle = 1
            lblStraight(y).BackColor = (&HFFFF00)
        Next y
    Case "Flush":
        lblPoints(4).BackStyle = 1
        lblPoints(4).BackColor = (&HC0FFFF)
        For y = 0 To 4
            lblFlush(y).BackStyle = 1
            lblFlush(y).BackColor = (&HFFFF00)
        Next y
    Case "Full House":
        lblPoints(5).BackStyle = 1
        lblPoints(5).BackColor = (&HC0FFFF)
        For y = 0 To 4
            lblFullHouse(y).BackStyle = 1
            lblFullHouse(y).BackColor = (&HFFFF00)
        Next y
    Case "Four of a Kind":
        lblPoints(6).BackStyle = 1
        lblPoints(6).BackColor = (&HC0FFFF)
        For y = 0 To 4
            lbl4Kind(y).BackStyle = 1
            lbl4Kind(y).BackColor = (&HFFFF00)
        Next y
    Case "Straight Flush":
        lblPoints(7).BackStyle = 1
        lblPoints(7).BackColor = (&HC0FFFF)
        For y = 0 To 4
            lblStrFlush(y).BackStyle = 1
            lblStrFlush(y).BackColor = (&HFFFF00)
        Next y
    Case "Royal Flush":
        lblPoints(8).BackStyle = 1
        lblPoints(8).BackColor = (&HC0FFFF)
        For y = 0 To 4
            lblRoyalFlush(y).BackStyle = 1
            lblRoyalFlush(y).BackColor = (&HFFFF00)
        Next y
    End Select
End Sub

Public Sub PayThePlayer()
Dim MyValue As Integer

MyValue = Int(lblPurse.Caption)
Select Case lblMsg.Caption
    Case "One Pair":
            MyValue = MyValue + Int(lbl1Pair(TotalBet - 1).Caption)
    Case "Two Pair":
            MyValue = MyValue + Int(lbl2Pair(TotalBet - 1).Caption)
    Case "Three of a Kind":
            MyValue = MyValue + Int(lbl3Kind(TotalBet - 1).Caption)
    Case "Straight":
            MyValue = MyValue + Int(lblStraight(TotalBet - 1).Caption)
    Case "Flush":
            MyValue = MyValue + Int(lblFlush(TotalBet - 1).Caption)
    Case "Full House":
            MyValue = MyValue + Int(lblFullHouse(TotalBet - 1).Caption)
    Case "Four of a Kind":
            MyValue = MyValue + Int(lbl4Kind(TotalBet - 1).Caption)
    Case "Straight Flush":
            MyValue = MyValue + Int(lblStrFlush(TotalBet - 1).Caption)
    Case "Royal Flush":
            MyValue = MyValue + Int(lblRoyalFlush(TotalBet - 1).Caption)
End Select

lblPurse.Caption = Format(MyValue, "$###,##0.00")
    If lblPurse.Caption + 0 > 0 Then
        cmdBet1.Enabled = True
        cmdBetMax.Enabled = True
    End If

End Sub
