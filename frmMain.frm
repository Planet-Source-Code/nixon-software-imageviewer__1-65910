VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   Caption         =   "NIXON ImageViewer"
   ClientHeight    =   4140
   ClientLeft      =   2505
   ClientTop       =   2625
   ClientWidth     =   6060
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MouseIcon       =   "frmMain.frx":1CFA
   ScaleHeight     =   4140
   ScaleWidth      =   6060
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   6060
      TabIndex        =   1
      Top             =   3840
      Visible         =   0   'False
      Width           =   6060
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmMain.frx":25C4
         Left            =   0
         List            =   "frmMain.frx":2604
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   0
         Width           =   1095
      End
      Begin VB.Label lblFileSize 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "File Size: 0 KB"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3045
         TabIndex        =   5
         Top             =   30
         Width           =   990
      End
      Begin VB.Label lblImageSize 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Image Size: 0 × 0 px"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1410
         TabIndex        =   4
         Top             =   30
         Width           =   1500
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1110
         TabIndex        =   3
         Top             =   30
         Width           =   255
      End
   End
   Begin MSComDlg.CommonDialog dlgCD 
      Left            =   2445
      Top             =   2925
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1410
      Top             =   2400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2673
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2A0D
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2DA7
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3099
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3633
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":39CD
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3D67
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4089
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":43AB
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":460D
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":495F
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4BC1
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbToolbar 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   6060
      _ExtentX        =   10689
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Wrappable       =   0   'False
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   12
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Open"
            Object.ToolTipText     =   "Open"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Save"
            Description     =   "Save"
            Object.ToolTipText     =   "Save"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Properties"
            Object.ToolTipText     =   "Properties"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Print"
            Object.ToolTipText     =   "Print"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Copy"
            Object.ToolTipText     =   "Copy"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Paste"
            Object.ToolTipText     =   "Paste"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Zoom In"
            Object.ToolTipText     =   "Zoom In"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Zoom Out"
            Object.ToolTipText     =   "Zoom Out"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Normal"
            Object.ToolTipText     =   "Normal"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ScaletoForm"
            Object.ToolTipText     =   "ScaletoForm"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Scale"
            Object.ToolTipText     =   "Scale"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "SizeInPixels"
            Object.ToolTipText     =   "SizeInPixels"
            ImageIndex      =   10
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   3855
      Left            =   60
      Picture         =   "frmMain.frx":4F13
      Top             =   225
      Width           =   5985
   End
   Begin VB.Menu FILE 
      Caption         =   "&File"
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open…"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "&Save As…"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileProperties 
         Caption         =   "Get Info…"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuFileLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "&Print…"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileLine3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileQuit 
         Caption         =   "&Quit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditCopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewToolbar 
         Caption         =   "&Toolbar"
      End
      Begin VB.Menu mnuViewSizeAndInfo 
         Caption         =   "&Size Toolbar"
      End
      Begin VB.Menu mnuViewPan 
         Caption         =   "&Pan"
      End
      Begin VB.Menu mnuViewSize 
         Caption         =   "&Size"
         Begin VB.Menu mnuViewSizeInPixels 
            Caption         =   "In Pixels…"
         End
         Begin VB.Menu mnuViewSizeInInches 
            Caption         =   "In Inches…"
         End
         Begin VB.Menu mnuViewSizeInMM 
            Caption         =   "In Millimetres…"
         End
         Begin VB.Menu mnuViewSizeInTwips 
            Caption         =   "In Twips…"
         End
      End
   End
   Begin VB.Menu mnuSize 
      Caption         =   "&Size"
      Begin VB.Menu mnuSizeZoomIn 
         Caption         =   "Zoom &In"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuSizeZoomOut 
         Caption         =   "Zoom &Out"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuSizeActualPixels 
         Caption         =   "&Actual Pixels"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuSizeFittoScreen 
         Caption         =   "Fit to Screen"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuSizeFittoWindow 
         Caption         =   "Fit to Window"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuSizeFittoPage 
         Caption         =   "Fit to Page"
         Begin VB.Menu mnuSizeFitPageL 
            Caption         =   "Landscape"
         End
         Begin VB.Menu mnuSizeFitPageP 
            Caption         =   "Portrait"
         End
      End
      Begin VB.Menu mnuSizeSplit 
         Caption         =   "&Split"
         Begin VB.Menu mnuSizeSplitHeight 
            Caption         =   "&Height"
            Shortcut        =   ^T
         End
         Begin VB.Menu mnuSizeSplitWidth 
            Caption         =   "&Width"
            Shortcut        =   ^Y
         End
      End
      Begin VB.Menu mnuSizeDouble 
         Caption         =   "&Double"
         Begin VB.Menu mnuSizeDoubleHeight 
            Caption         =   "&Height"
            Shortcut        =   ^H
         End
         Begin VB.Menu mnuSizeDoubleWidth 
            Caption         =   "&Width"
            Shortcut        =   ^W
         End
      End
      Begin VB.Menu mnuSizeRatio 
         Caption         =   "&Ratio"
         Begin VB.Menu mnuSizeRatio43 
            Caption         =   "4:3"
         End
         Begin VB.Menu mnuSizeRatio169 
            Caption         =   "16:9"
         End
         Begin VB.Menu mnuSizeRatio1610 
            Caption         =   "16:10"
         End
      End
      Begin VB.Menu mnuSizeMakeSquare 
         Caption         =   "Make Square from"
         Begin VB.Menu mnuSizeMakeSquareWidth 
            Caption         =   "Width"
         End
         Begin VB.Menu mnuSizeMakeSquareHeight 
            Caption         =   "Height"
         End
         Begin VB.Menu mnuSizeMakeSquareAverage 
            Caption         =   "Median"
         End
      End
      Begin VB.Menu mnuSizeScale 
         Caption         =   "Scale…"
         Shortcut        =   ^L
      End
   End
   Begin VB.Menu About 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    'Functions
    'Strings
    Dim sFile As String 'Filename
    'Booleans
    Dim Clicked As Boolean
    Dim FitWindow As Boolean
    Dim DoNotMove As Boolean
    Dim DoNotChange As Boolean
    'Integers
    Dim TBHeight As Long
    Dim PanFromX As Long
    Dim PanFromY As Long
    Dim PanToX As Long
    Dim PanToY As Long
    Dim OriginalWidth As Long
    Dim OriginalHeight As Long
    Dim DoRefresh As Integer
    
    
Private Sub About_Click()
On Error GoTo 10
frmAbout.Show
10:
End Sub



Private Sub Combo1_Click()
If DoNotChange = True Then
    DoNotChange = False
    Exit Sub
End If
Image1.Stretch = False
If Combo1.ListIndex = 1 Then
 Image1.Height = Image1.Height / 2
 Image1.Width = Image1.Width / 2
 CenterImage
 Image1.Stretch = True
 Exit Sub
End If
If Combo1.ListIndex = 0 Then
 Image1.Height = Image1.Height / 4
 Image1.Width = Image1.Width / 4
 CenterImage
 Image1.Stretch = True
 Exit Sub
End If
Image1.Height = Image1.Height * (Combo1.ListIndex - 1)
Image1.Width = Image1.Width * (Combo1.ListIndex - 1)
CenterImage
Image1.Stretch = True
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Combo1_Click
End Sub

Private Sub mnuEditCopy_Click()
On Error GoTo 10
Clipboard.Clear
Clipboard.SetData Image1.Picture
Exit Sub
10:
MsgBox Err.Number & " : " & Err.Description & ".", vbExclamation + vbOKOnly, "Error"
End Sub

Private Sub mnuFileQuit_Click()
End
End Sub

Private Sub Form_Load()
On Error GoTo 10
Form_Resize
If Command <> "" Then
    dlgCD.FileName = Mid(Command, 2, Len(Command) - 2)
    Image1.Picture = LoadPicture(dlgCD.FileName)
    Image1.Stretch = False
    Clicked = True
End If
    DoNotMove = False
    Image1.Left = 0
    Image1.Top = 0
10:
If Err.Number = 481 Then MsgBox "Invalid Picture", vbExclamation + vbOKOnly, "Error"
End Sub

Private Sub Form_Resize()
On Error GoTo 10
Dim MeWidth As Long
Dim MeHeight As Long
Dim ImageWidth As Long
Dim ImageHeight As Long
    CenterImage
    'Me.ScaleMode = vbPixels
    MeWidth = Me.Width / Screen.TwipsPerPixelX
    MeHeight = Me.Height / Screen.TwipsPerPixelY
    ImageHeight = Image1.Height / Screen.TwipsPerPixelY
    ImageWidth = Image1.Width / Screen.TwipsPerPixelX
    If Image1.Stretch = False Then
    If Pan1 = False Then
    If MeWidth < ImageWidth Then Me.Width = Image1.Width '* Screen.TwipsPerPixelX
    If MeHeight < ImageHeight Then Me.Height = Image1.Height '* Screen.TwipsPerPixelY
    End If
    End If
    If FitWindow = True Then
    Image1.Height = Me.Height - 100
    Image1.Width = Me.Width
    End If
    If Me.Width < 190 * Screen.TwipsPerPixelX Then
    Me.Width = 190 * Screen.TwipsPerPixelY
    End If
    If Me.Height < 70 * Screen.TwipsPerPixelY Then
    Me.Height = 70 * Screen.TwipsPerPixelY
    End If
Me.ScaleMode = vbTwips
10:
End Sub
Public Sub CenterImage()
If Pan1 = True Or FitWindow = True Then Exit Sub
Dim X As Single
Dim Y As Single
Dim realHeight As Single
ScaleMode = vbTwips
realHeight = Me.Height - TBHeight
If Image1.Height < Screen.Height Then
Y = (realHeight - Image1.Height - 740) / 2
Else
Y = 0 - (Image1.Height - realHeight + 740) / 2
End If
If Image1.Width < Screen.Width Then
X = (Me.Width - Image1.Width - 112) / 2
Else
X = 0 - (Image1.Width - Me.Width + 112) / 2
End If
Image1.Left = X
Image1.Top = Y + TBHeight
End Sub



Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub mnuSizeDoubleHeight_Click()
Image1.Height = Image1.Height * 2
Image1.Stretch = True
CenterImage
End Sub

Private Sub mnuSizeSplitHeight_Click()
Image1.Height = Image1.Height / 2
Image1.Stretch = True
CenterImage
End Sub

Private Sub mnuSizeViewSizeInMM_Click()

End Sub

Private Sub mnuViewSizeAndInfo_Click()
Picture1.Visible = Not (Picture1.Visible)
mnuViewSizeAndInfo.Checked = Picture1.Visible
DoNotChange = True
Combo1.ListIndex = 2
If Picture1.Visible Then
    RefreshImgSize
    If tbToolbar.Visible = True Then
    TBHeight = 300
    Else
    TBHeight = -300
    End If
Else
    If Me.WindowState = 0 Then Me.Height = Me.Height - 300
    If tbToolbar.Visible = True Then
    TBHeight = 300
    Else
    TBHeight = 0
    End If
End If
    If Me.WindowState = 2 Then
        If Picture1.Visible Then
            If Image1.Top = 0 Then Image1.Top = -300
        Else
            Image1.Top = 0
            CenterImage
        End If
    End If
CenterImage
End Sub

Private Sub mnuViewSizeInPixels_Click()
On Error GoTo 10
Dim IWidth As String
Dim IHeight As String
ScaleMode = 3
IWidth = Image1.Width
IHeight = Image1.Height
MsgBox "Image Width: " + IWidth + Chr$(10) + "Image Height: " + IHeight, vbInformation + vbOKOnly, "Image Attributes"
ScaleMode = 1
10:
End Sub

Private Sub Image1_Click()
If Pan1 = True Then Exit Sub
If Clicked = False Then mnuFileOpen_Click
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
PanFromX = X
PanFromY = Y
If Pan1 = True Then
DoNotMove = True
Image1.MousePointer = 99
Me.MousePointer = 99
End If
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long
DoRefresh = DoRefresh + 1
If DoRefresh >= 5 Then
    If DoNotMove <> False Then
        PanToX = X
        PanToY = Y
        If Button = vbLeftButton Then
            If Image1.Height > Me.Height Then
                If PanToY < PanFromY Then
                    Image1.Top = Image1.Top + (PanToY - PanFromY) / 1.44
                End If
                If PanToY > PanFromY Then
                    Image1.Top = Image1.Top - (PanFromY - PanToY) / 1.44
                End If
            End If
            If Image1.Width > Me.Width Then
                If PanToX < PanFromX Then
                    Image1.Left = Image1.Left + (PanToX - PanFromX) / 1.44
                End If
                If PanToX > PanFromX Then
                    Image1.Left = Image1.Left - (PanFromX - PanToX) / 1.44
                End If
            End If
        End If
                If Image1.Top > 0 Then
                Image1.Top = 0
                'Image1.Refresh
                End If
                If Image1.Left > 0 Then
                Image1.Left = 0
                'Image1.Refresh
                End If
                    If Image1.Height > Me.Height Then
                        If Image1.Top < 0 - (Image1.Height - Me.Height) Then Image1.Top = 0 - (Image1.Height - Me.Height)
                        'Image1.Refresh
                    End If
                    If Image1.Width > Me.Width Then
                    i = Image1.Left - 2 * Screen.TwipsPerPixelX
                        If i < 128 * Screen.TwipsPerPixelX - Image1.Width Then Image1.Left = 128 * Screen.TwipsPerPixelX - Image1.Width
                        'Image1.Refresh
                    End If
    End If
DoRefresh = 0
End If
End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
DoNotMove = False
Me.ScaleMode = vbTwips
End Sub

Private Sub mnuViewSizeInMM_Click()
On Error GoTo 10
Dim IWidth As String
Dim IHeight As String
ScaleMode = 6
IWidth = Image1.Width
IHeight = Image1.Height
MsgBox "Image Width: " + IWidth + Chr$(10) + "Image Height: " + IHeight, vbInformation + vbOKOnly, "Image Attributes"
ScaleMode = 1
10:
End Sub

Private Sub mnuViewSizeInInches_Click()
On Error GoTo 10
Dim IWidth As String
Dim IHeight As String
ScaleMode = 5
IWidth = Image1.Width
IHeight = Image1.Height
MsgBox "Image Width: " + IWidth + Chr$(10) + "Image Height: " + IHeight, vbInformation + vbOKOnly, "Image Attributes"
ScaleMode = 1
10:
End Sub

Private Sub mnuViewSizeInTwips_Click()
On Error GoTo 10
Dim IWidth As String
Dim IHeight As String
ScaleMode = 1
IWidth = Image1.Width
IHeight = Image1.Height
MsgBox "Image Width: " + IWidth + Chr$(10) + "Image Height: " + IHeight, vbInformation + vbOKOnly, "Image Attributes"
10:
End Sub


Private Sub mnuFileOpen_Click()
'Dim YesNo%
On Error GoTo 10
ScaleMode = vbPixels
Me.ScaleMode = vbTwips
    Clicked = True
    With dlgCD
        .DialogTitle = "Open"
        .CancelError = False
        .Flags = cdlOFNOverwritePrompt
        .Filter = "Supported Picture Files (*.gif, *.jpg, *.tga, *.bmp, *.dib, *.wmf, *.ico, *.cur)|*.gif;*.jpg;*.tga;*.bmp;*.dib;*.wmf;*.ico;*.cur|All files (*.*)|*.*"
        .ShowOpen
        If Len(.FileName) = 0 Then
            Clicked = False
            Exit Sub
        End If
        sFile = .FileName
    End With
    Image1.Picture = LoadPicture(sFile)
    Image1.Stretch = False
    Pan1 = False
    mnuViewPan.Checked = False
    RefreshImgSize
    FitWindowtoImage
    Image1.Left = 0
    Image1.Top = 0
    CenterImage
    Me.Caption = dlgCD.FileTitle + " - NIXON ImageViewer"
    If Image1.Width > Screen.Width Or Image1.Height > Screen.Height Then
        WindowState = 2
    End If
    If Picture1.Visible = True Then Combo1_Click
Exit Sub
10:
ErrorTrap
End Sub

Private Sub mnuFileSaveAs_Click()
Dim sFile As String
    With dlgCD
        .DialogTitle = "Save"
        .CancelError = False
        .Filter = "Bitmaps (*.bmp, *.dib)|*.bmp;*.dib"
        .ShowSave
        If Len(.FileName) = 0 Then
            Exit Sub
        End If
        sFile = .FileName
    End With
SavePicture Image1.Picture, sFile
End Sub





Public Sub MakeSquare()
On Error GoTo 10
Dim Square1 As Integer
'ScaleMode = vbPixels
'Me.ScaleMode = vbPixels
Pan1 = False
mnuViewPan.Checked = False
If Image1.Width = Image1.Height Then
 MsgBox "The image is already a square.", vbInformation + vbOKOnly, "Information"
 Exit Sub
End If
If Image1.Width > Image1.Height Then
   Square1 = Image1.Width - Image1.Height
   Square1 = Square1 / 2
   Image1.Width = Image1.Width - Square1
   Image1.Height = Image1.Height + Square1
Else
   Square1 = Image1.Height - Image1.Width
   Square1 = Square1 / 2
   Image1.Width = Square1
   Image1.Height = Square1
   Image1.Width = Image1.Width + Square1
   Image1.Height = Image1.Height - Square1
End If
Image1.Stretch = True
CenterImage
Exit Sub
10:
ErrorTrap
End Sub

Private Sub mnuSizeFitPageL_Click()
Image1.Stretch = False
Image1.Width = 15840
Image1.Height = 12240
Image1.Stretch = True
Pan1 = False
mnuViewPan.Checked = False
CenterImage
End Sub

Private Sub mnuSizeFitPageP_Click()
Image1.Stretch = False
Image1.Width = 12240
Image1.Height = 15840
Image1.Stretch = True
Pan1 = False
mnuViewPan.Checked = False
CenterImage
End Sub

Private Sub mnuSizeMakeSquareAverage_Click()
MakeSquare
End Sub

Private Sub mnuSizeMakeSquareHeight_Click()
If Image1.Width = Image1.Height Then
 MsgBox "The image is already a square.", vbInformation + vbOKOnly, "Information"
 Exit Sub
End If
Pan1 = False
mnuViewPan.Checked = False
Image1.Width = Image1.Height
Image1.Stretch = True
CenterImage
End Sub

Private Sub mnuSizeMakeSquareWidth_Click()
If Image1.Width = Image1.Height Then
 MsgBox "The image is already a square.", vbInformation + vbOKOnly, "Information"
 Exit Sub
End If
Pan1 = False
mnuViewPan.Checked = False
Image1.Height = Image1.Width
Image1.Stretch = True
CenterImage
End Sub

Private Sub mnuSizeRatio1610_Click()
MakeRatio 16, 10
End Sub

Private Sub mnuSizeRatio169_Click()
MakeRatio 16, 9
End Sub

Private Sub mnuSizeRatio43_Click()
MakeRatio 4, 3
End Sub
Public Sub MakeRatio(Width1 As Integer, Height1 As Integer)
On Error GoTo 10
Dim RatioWidth As Integer
Dim RatioHeight As Integer
Pan1 = False
mnuViewPan.Checked = False
Image1.Stretch = False
RatioWidth = Image1.Height / Width1
RatioHeight = RatioWidth * Height1
Image1.Height = RatioHeight
Image1.Stretch = True
CenterImage
Exit Sub
10:
ErrorTrap
End Sub



Private Sub mnuSizeZoom_Click()

End Sub

Private Sub mnuViewPan_Click()
mnuViewPan.Checked = Not (mnuViewPan.Checked)
Pan1 = mnuViewPan.Checked
If Pan1 = False Then
Image1.MousePointer = 0
Me.MousePointer = 0
End If
End Sub

Private Sub mnuViewReset_Click()

End Sub

Private Sub mnuViewToolbar_Click()
    mnuViewToolbar.Checked = Not mnuViewToolbar.Checked
    tbToolbar.Visible = mnuViewToolbar.Checked
    Ischecked = mnuViewToolbar.Checked = True
    If tbToolbar.Visible = True Then
        If Picture1.Visible = True Then
            TBHeight = 600
            If WindowState = 0 Then Me.Height = Me.Height + 300
        Else
            TBHeight = 300
            If WindowState = 0 Then Me.Height = Me.Height + 0
        End If
    Else
        If Picture1.Visible = True Then
            TBHeight = -300
            Me.Height = Me.Height - 300
        Else
            TBHeight = 0
        End If
    End If
    CenterImage
End Sub

Private Sub mnuSizeActualPixels_Click()
FitWindow = False
Pan1 = False
mnuViewPan.Checked = False
Image1.Stretch = False
Image1.Top = 0
Image1.Left = 0
Form_Resize
CenterImage
mnuSizeFittoWindow.Checked = False
End Sub





Private Sub mnuFileProperties_Click()
frmFileOperations.Show
End Sub





Private Sub mnuEditPaste_Click()
On Error GoTo 10
Pan1 = False
mnuViewPan.Checked = False
Image1.Picture = Clipboard.GetData
ScaleMode = 1
If Image1.Height >= Screen.Height Or Image1.Width >= Screen.Width Then
    WindowState = 2
Else
    WindowState = 0
    FitWindowtoImage
    CenterImage
    RefreshImgSize
End If
If Clicked = False Then Clicked = True
10:
ErrorTrap
End Sub

Private Sub mnuFilePrint_Click()
On Error GoTo 10
    If Image1.Picture = LoadPicture() Then
        MsgBox "No picture loaded", vbExclamation + vbOKOnly, "Error"
        Exit Sub
    End If
    With dlgCD
        .CancelError = True
        .ShowPrinter
    End With
    MousePointer = 11 'Display Hourglass
    Printer.PaintPicture Image1.Picture, 0, 0 'Print the Picture
    Printer.EndDoc
    MousePointer = 1
Exit Sub
10:
If Err.Number = 32755 Then Exit Sub
ErrorTrap
End Sub



Private Sub mnuSizeScale_Click()
On Error GoTo 10
'Dim HStr As Long, VStr As Long
'Me.ScaleMode = vbPixels
'VStr = InputBox("Image Width:", "Scale")
'HStr = InputBox("Image Height:", "Scale")
'Image1.Height = HStr
'Image1.Width = WStr
'Image1.Stretch = True
frmScale.Show vbModal, frmMain
10:
If Err.Number = 13 Then Exit Sub
ErrorTrap
End Sub

Private Sub mnuSizeFittoWindow_Click()
    Pan1 = False
    mnuViewPan.Checked = False
     Image1.Height = Me.Height
     Image1.Width = Me.Width
     Image1.Top = 0
     Image1.Left = 0
     Image1.Stretch = True
     mnuSizeFittoWindow.Checked = Not (mnuSizeFittoWindow.Checked)
     FitWindow = mnuSizeFittoWindow.Checked
End Sub

Private Sub mnuSizeFittoScreen_Click()
ScaleMode = 1
Pan1 = False
mnuViewPan.Checked = False
     Image1.Height = Screen.Height
     Image1.Width = Screen.Width
     Image1.Top = 0
     Image1.Left = 0
     Image1.Stretch = True
     Me.WindowState = 2
End Sub

Private Sub mnuSizeZoomOut_Click()
mnuSizeSplitHeight_Click
mnuSizeSplitWidth_Click
CenterImage
End Sub




Private Function RefreshImgSize()
If Picture1.Visible Then
    If Me.WindowState = 0 Then Me.Height = Me.Height + 300
    If dlgCD.FileName <> "" Then lblFileSize.Caption = frmFileOperations.GetFileSize(dlgCD.FileName)
    Me.ScaleMode = vbPixels
    lblImageSize.Caption = "Image Size: " & Image1.Width & " " & Chr$(215) & " " & Image1.Height & " px"
    lblFileSize.Left = lblImageSize.Left + lblImageSize.Width + 240
End If
End Function
Private Sub tbToolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
 Select Case Button.Key
    Case "Open"
    mnuFileOpen_Click
    Case "Save"
    mnuFileSaveAs_Click
    Case "Properties"
    mnuFileProperties_Click
    Case "Print"
    mnuFilePrint_Click
    Case "Copy"
    mnuEditCopy_Click
    Case "Paste"
    mnuEditPaste_Click
    Case "Zoom In"
    mnuSizeZoomIn_Click
    Case "Zoom Out"
    mnuSizeZoomOut_Click
    Case "Normal"
    mnuSizeActualPixels_Click
    Case "ScaletoForm"
    mnuSizeFittoWindow_Click
    Case "Scale"
    mnuSizeScale_Click
    Case "SizeInPixels"
    mnuViewSizeInPixels_Click
  End Select
End Sub

Private Sub mnuSizeDoubleWidth_Click()
Image1.Width = Image1.Width * 2
Image1.Stretch = True
CenterImage
End Sub

Private Sub mnuSizeSplitWidth_Click()
Image1.Width = Image1.Width / 2
Image1.Stretch = True
CenterImage
End Sub

Private Sub mnuSizeZoomIn_Click()
mnuSizeDoubleHeight_Click
mnuSizeDoubleWidth_Click
CenterImage
If Image1.Width > Me.Width Or Image1.Height > Me.Height Then
If Me.WindowState = 2 Then Exit Sub
FitWindowtoImage
End If
End Sub

Public Sub FitWindowtoImage()
    If Image1.Height >= Screen.Height Or Image1.Width >= Screen.Width Then
        WindowState = 2
    Else
        WindowState = 0
        If Image1.Width < 2551 Then
            Me.Height = Image1.Height + 1000
        Me.Width = Image1.Width + 116
        Else
            Me.Height = Image1.Height + 736
            Me.Width = Image1.Width + 116
        End If
    End If
End Sub

