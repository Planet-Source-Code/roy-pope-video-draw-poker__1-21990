VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00008000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About"
   ClientHeight    =   4440
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6480
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4440
   ScaleWidth      =   6480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pctIcon 
      AutoSize        =   -1  'True
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   240
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   2
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Height          =   375
      Left            =   5160
      TabIndex        =   1
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Label lblLegal 
      BackColor       =   &H00008000&
      ForeColor       =   &H00FFFFC0&
      Height          =   1215
      Left            =   120
      TabIndex        =   7
      Top             =   3120
      Width           =   4815
   End
   Begin VB.Line Line4 
      X1              =   0
      X2              =   6480
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Line Line3 
      BorderColor     =   &H0080FF80&
      BorderWidth     =   3
      X1              =   0
      X2              =   6480
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Label lblRights 
      BackColor       =   &H00008000&
      Caption         =   "ALL RIGHTS RESERVED"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1320
      TabIndex        =   6
      Top             =   1440
      Width           =   3135
   End
   Begin VB.Label lblCopyright 
      BackColor       =   &H00008000&
      Caption         =   "Copyright"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1320
      TabIndex        =   5
      Top             =   1200
      Width           =   3135
   End
   Begin VB.Label lblVersion 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      Caption         =   "Version"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   1320
      TabIndex        =   4
      Top             =   720
      Width           =   5055
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      Caption         =   "App Test"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   1320
      TabIndex        =   3
      Top             =   240
      Width           =   5055
   End
   Begin VB.Line Line2 
      X1              =   1200
      X2              =   1200
      Y1              =   0
      Y2              =   2880
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0080FF80&
      BorderWidth     =   3
      X1              =   1200
      X2              =   1200
      Y1              =   0
      Y2              =   2880
   End
   Begin VB.Label lblWarning 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      Height          =   810
      Left            =   1320
      TabIndex        =   0
      Top             =   1920
      Width           =   4605
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'--------------------------------------------------
'Unload the form
'--------------------------------------------------
Private Sub cmdOK_Click()
    Unload Me
End Sub

'----------------------------------------------------------
'Load the form and set the caption
'----------------------------------------------------------
Private Sub Form_Load()
'BEGIN: AESTHETIC DESIGN AND CONTROL AREA.
    pctIcon.Picture = frmMain.Icon
    Line1.X1 = pctIcon.Left + pctIcon.Width + 150
    Line1.X2 = pctIcon.Left + pctIcon.Width + 150
    Line2.X1 = pctIcon.Left + pctIcon.Width + 150
    Line2.X2 = pctIcon.Left + pctIcon.Width + 150
    lblTitle.Caption = App.Title
    lblTitle.Left = Line2.X2 + 150
    lblTitle.Width = 6480 - (lblTitle.Left + 150)
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblVersion.Left = Line2.X2 + 150
    lblVersion.Width = 6480 - (lblVersion.Left + 150)
    lblCopyright.Caption = App.LegalCopyright
    lblWarning.Left = Line2.X2 + 150
    lblWarning.Width = 6480 - (lblWarning.Left + 150)
    lblLegal.Caption = "Warning:  This computer program is protected by copyright law and international treaties. Unauthorized reproduction or distribution of this program, or any portion of it, may result in severe civil and criminal penalties and will be prosecuted to the maximum extent under the law."
    lblCopyright.Left = Line2.X2 + 150
    lblCopyright.Width = 6480 - (lblCopyright.Left + 150)
    lblRights.Left = Line2.X2 + 150
    lblRights.Width = 6480 - (lblRights.Left + 150)
    lblWarning.Caption = "Code distributed by rd-soft.com for educational and demonstrative purposes only. You can find me at http://www.rd-soft.com"
    Me.Caption = App.Title & " v." & App.Major & "." & App.Minor & "." & App.Revision
End Sub

