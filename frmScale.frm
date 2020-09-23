VERSION 5.00
Begin VB.Form frmScale 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Scale"
   ClientHeight    =   2940
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   3855
   Icon            =   "frmScale.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2940
   ScaleWidth      =   3855
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkMaintain 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Preserve Aspect Ratio"
      Height          =   270
      Left            =   285
      TabIndex        =   10
      Top             =   1905
      Width           =   3315
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   360
      Left            =   2190
      TabIndex        =   9
      Top             =   2280
      Width           =   915
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   360
      Left            =   1230
      TabIndex        =   8
      Top             =   2280
      Width           =   915
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   360
      Left            =   270
      TabIndex        =   7
      Top             =   2280
      Width           =   915
   End
   Begin VB.ComboBox cboType 
      Height          =   315
      ItemData        =   "frmScale.frx":038A
      Left            =   2640
      List            =   "frmScale.frx":039D
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   1050
      Width           =   870
   End
   Begin VB.TextBox txtHeight 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Height          =   315
      Left            =   1215
      TabIndex        =   1
      Top             =   1470
      Width           =   1350
   End
   Begin VB.TextBox txtWidth 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Height          =   315
      Left            =   1215
      TabIndex        =   0
      Top             =   1050
      Width           =   1350
   End
   Begin VB.Label lblRes 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Resolution: 0×0 pixels per inch"
      Height          =   300
      Left            =   285
      TabIndex        =   5
      Top             =   600
      Width           =   3300
   End
   Begin VB.Label lblSize 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Original Size: 0×0 pixels"
      Height          =   300
      Left            =   285
      TabIndex        =   4
      Top             =   270
      Width           =   3315
   End
   Begin VB.Label lblHeight 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Height:"
      Height          =   255
      Left            =   300
      TabIndex        =   3
      Top             =   1470
      Width           =   855
   End
   Begin VB.Label lblWidth 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Width:"
      Height          =   255
      Left            =   300
      TabIndex        =   2
      Top             =   1050
      Width           =   855
   End
End
Attribute VB_Name = "frmScale"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ImageRatio As Double
Dim OtherChanging As Boolean

Private Sub cboType_Click()
ChangeMode
RefreshSize
frmMain.ScaleMode = vbTwips
End Sub

Private Sub chkMaintain_Click()
txtWidth_Change
End Sub

Private Sub cmdApply_Click()
ChangeMode
Pan1 = False
frmMain.mnuViewPan.Checked = False
frmMain.Image1.Stretch = True
frmMain.Image1.Width = txtWidth.Text
frmMain.Image1.Height = txtHeight.Text
frmMain.ScaleMode = vbTwips
frmMain.FitWindowtoImage
frmMain.CenterImage
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
cmdApply_Click
Unload Me
End Sub

Private Sub Form_Load()
cboType.ListIndex = 0
frmMain.ScaleMode = vbPixels
lblSize.Caption = "Original Size: " & frmMain.Image1.Width & " " & Chr$(215) & " " & frmMain.Image1.Height & " pixels"
RefreshSize
frmMain.ScaleMode = vbTwips
lblRes.Caption = "Resolution: " & 1440 / Screen.TwipsPerPixelX & " " & Chr$(215) & " " & 1440 / Screen.TwipsPerPixelY & " pixels per inch"
End Sub

Private Function RefreshSize()
txtWidth.Text = frmMain.Image1.Width
txtHeight.Text = frmMain.Image1.Height
End Function
Private Function ChangeMode()
Select Case cboType.ListIndex
Case 0
frmMain.ScaleMode = vbPixels
Case 1
frmMain.ScaleMode = vbInches
Case 2
frmMain.ScaleMode = vbCentimeters
Case 3
frmMain.ScaleMode = vbMillimeters
Case 4
frmMain.ScaleMode = vbPoints
End Select
End Function
Private Function SetImage()
ChangeMode
frmMain.Image1.Width = txtWidth.Text
frmMain.Image1.Height = txtHeight.Text
frmMain.FitWindowtoImage
End Function

Private Sub txtHeight_Change()
On Error GoTo 10
If txtWidth.Text = "" Or OtherChanging = True Then Exit Sub
If chkMaintain.Value = Checked Then
   ImageRatio = frmMain.Image1.Width / frmMain.Image1.Height
   OtherChanging = True
   txtWidth.Text = txtHeight.Text * ImageRatio
   OtherChanging = False
End If
10:
End Sub

Private Sub txtHeight_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
Case Asc("0") To Asc("9")
Case Asc(".")
Case 8
Case Else
KeyAscii = 0
End Select
End Sub

Private Sub txtWidth_Change()
On Error GoTo 10
If txtWidth.Text = "" Or OtherChanging = True Then Exit Sub
If chkMaintain.Value = Checked Then
   ImageRatio = frmMain.Image1.Height / frmMain.Image1.Width
   OtherChanging = True
   txtHeight.Text = txtWidth.Text * ImageRatio
   OtherChanging = False
End If
10:
End Sub

Private Sub txtWidth_KeyPress(KeyAscii As Integer)
 Select Case KeyAscii
Case Asc("0") To Asc("9")
Case Asc(".")
Case 8
Case Else
KeyAscii = 0
End Select
End Sub
