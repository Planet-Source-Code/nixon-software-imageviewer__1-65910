VERSION 5.00
Begin VB.Form frmAbout 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "About NIXON ImageViewer"
   ClientHeight    =   4440
   ClientLeft      =   1545
   ClientTop       =   2220
   ClientWidth     =   4755
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmAbout.frx":0000
   ScaleHeight     =   4440
   ScaleWidth      =   4755
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   105
      TabIndex        =   2
      Top             =   3915
      Width           =   1215
   End
   Begin VB.Label lblWebsite 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Visit Nixon on the web at members.shaw.ca/nixon.com"
      Height          =   465
      Left            =   225
      TabIndex        =   4
      Top             =   3270
      Width           =   4005
   End
   Begin VB.Label lblContact 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Questions? Comments? Complaints? Suggestions? E-mail lcaa9@netscape.net"
      Height          =   465
      Left            =   225
      TabIndex        =   3
      Top             =   2760
      Width           =   4005
   End
   Begin VB.Label lblCopyright 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Copyright © 2005-2006 NIXON Software Corporation. Some rights reserved."
      Height          =   465
      Left            =   225
      TabIndex        =   1
      Top             =   2250
      Width           =   4005
   End
   Begin VB.Label lblVersion 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Version"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   1815
      Width           =   3975
   End
   Begin VB.Image imgIcon 
      Height          =   720
      Left            =   3960
      Picture         =   "frmAbout.frx":0037
      Top             =   690
      Width           =   720
   End
   Begin VB.Image imgTitle 
      Height          =   1500
      Left            =   90
      Picture         =   "frmAbout.frx":02EB
      Top             =   150
      Width           =   3675
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub Form_Load()
lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
End Sub

