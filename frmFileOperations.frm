VERSION 5.00
Begin VB.Form frmFileOperations 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "File Properties/Operations"
   ClientHeight    =   4470
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6825
   Icon            =   "frmFileOperations.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4470
   ScaleWidth      =   6825
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkNormal 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Normal"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3210
      TabIndex        =   19
      Top             =   3150
      Width           =   1230
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply Attributes"
      Height          =   375
      Left            =   4545
      TabIndex        =   12
      Top             =   3225
      Width           =   1350
   End
   Begin VB.CheckBox chkTemporary 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Temporary"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3210
      TabIndex        =   11
      Top             =   3435
      Width           =   1230
   End
   Begin VB.CheckBox chkSystem 
      BackColor       =   &H00FFFFFF&
      Caption         =   "System"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   2085
      TabIndex        =   10
      Top             =   3435
      Width           =   1065
   End
   Begin VB.CheckBox chkArchive 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Archive"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   945
      TabIndex        =   9
      Top             =   3435
      Width           =   1065
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Done"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   195
      TabIndex        =   8
      Top             =   3930
      Width           =   945
   End
   Begin VB.CheckBox chkHidden 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Hidden"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   2085
      TabIndex        =   7
      Top             =   3150
      Width           =   1065
   End
   Begin VB.CheckBox chkReadOnly 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Read-Only"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   945
      TabIndex        =   6
      Top             =   3150
      Width           =   1065
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "Move/Rename File to…"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3105
      TabIndex        =   5
      Top             =   1920
      Width           =   1905
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "Copy File to…"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1905
      TabIndex        =   4
      Top             =   1920
      Width           =   1185
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete File"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   945
      TabIndex        =   3
      Top             =   1920
      Width           =   945
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load…"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5595
      TabIndex        =   2
      Top             =   1920
      Width           =   945
   End
   Begin VB.Label lblAttr 
      BackColor       =   &H00FFFFFF&
      Caption         =   "No File Loaded"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Left            =   2325
      TabIndex        =   18
      Top             =   2400
      Width           =   4215
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Last Modified/Created:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   945
      TabIndex        =   17
      Top             =   1365
      Width           =   1305
   End
   Begin VB.Label lblModified 
      BackColor       =   &H00FFFFFF&
      Caption         =   "--/--/---- --:--:-- --"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   2325
      TabIndex        =   16
      Top             =   1365
      Width           =   4200
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Attributes:"
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
      Left            =   945
      TabIndex        =   15
      Top             =   2400
      Width           =   1305
   End
   Begin VB.Label lblSize 
      BackColor       =   &H00FFFFFF&
      Caption         =   "0 Bytes"
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
      Left            =   2325
      TabIndex        =   14
      Top             =   960
      Width           =   4200
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "File Size:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   945
      TabIndex        =   13
      Top             =   960
      Width           =   1305
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   225
      Picture         =   "frmFileOperations.frx":030A
      Top             =   135
      Width           =   570
   End
   Begin VB.Label lblName 
      BackColor       =   &H00FFFFFF&
      Caption         =   "No File"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   2325
      TabIndex        =   1
      Top             =   135
      Width           =   4215
   End
   Begin VB.Label lblFileName 
      BackColor       =   &H00FFFFFF&
      Caption         =   "File Name:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   945
      TabIndex        =   0
      Top             =   150
      Width           =   1290
   End
End
Attribute VB_Name = "frmFileOperations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetFileAttributes Lib "kernel32" Alias "SetFileAttributesA" _
    (ByVal lpFileSpec As String, ByVal dwFileAttributes As Long) As Long

Public FileNameStr As String
Public PicFile As String




Private Sub chkHidden_Validate(Cancel As Boolean)
chkHidden_Click
End Sub

Private Sub chkNormal_Click()
If FileNameStr = "" Then
MsgBox "No file loaded.", vbExclamation + vbOKOnly, "Error"
Exit Sub
End If
If chkNormal.Value = Checked Then
Uncheck
SetAttr FileNameStr, vbNormal
End If
End Sub

Public Sub chkReadOnly_Click()
End Sub

Public Sub chkHidden_Click()
End Sub

Private Sub cmdApply_Click()
Dim attr As Long
    If chkReadOnly.Value = Checked Then attr = vbReadOnly
    If chkArchive.Value = Checked Then attr = attr + vbArchive
    If chkSystem.Value = Checked Then attr = attr + vbSystem
    If chkHidden.Value = Checked Then attr = attr + vbHidden
    If chkTemporary.Value = Checked Then attr = attr + vbTemporary
    SetFileAttributes FileNameStr, attr
    GetTextAttributes
End Sub

Private Sub cmdLoad_Click()
On Error GoTo 10
    With frmMain.dlgCD
        .DialogTitle = "Open"
        .CancelError = False
        .Filter = "All Files (*.*)|*.*"
        .ShowOpen
        If Len(.FileName) = 0 Then
            Exit Sub
        End If
        FileNameStr = .FileName
    End With

    GetFileInfo
    ToggleEnabled (True)
10:
    ErrorTrap
End Sub

Public Sub cmdDelete_Click()
    On Error GoTo 10
    Dim MsgBoxReturn As Integer
    If FileNameStr = "" Then
        MsgBox "No file loaded.", vbExclamation + vbOKOnly, "Error"
        Exit Sub
    End If
    MsgBoxReturn = MsgBox("Are you sure you want to delete this file?", vbQuestion + vbYesNo, "Question")
    If MsgBoxReturn = vbYes Then
    Kill FileNameStr
    If FileNameStr = PicFile Then PicFile = ""
        If PicFile <> "" Then
            FileNameStr = PicFile
            lblName.Caption = FileNameStr
            GetFileAttributes
            GetTextAttributes
            lblModified.Caption = FileDateTime(FileNameStr)
            lblSize.Caption = GetFileSize(FileNameStr)
            ToggleEnabled (True)
        Else
            FileNameStr = ""
            lblName.Caption = "No File"
            Uncheck
            chkNormal.Value = 0
            lblModified.Caption = "--/--/---- --:--:-- --"
            lblSize.Caption = "0 Bytes"
            lblAttr.Caption = "No File Loaded"
            ToggleEnabled (False)
        End If
    End If
    Exit Sub
10:
    ErrorTrap
End Sub
Public Sub Uncheck()
    chkReadOnly.Value = Unchecked
    chkArchive.Value = Unchecked
    chkSystem.Value = Unchecked
    chkHidden.Value = Unchecked
    chkTemporary.Value = Unchecked
End Sub
Public Sub cmdCopy_Click()
    On Error GoTo 10
    Dim ToFile As String
    If FileNameStr = "" Then
        MsgBox "No file loaded.", vbExclamation + vbOKOnly, "Error"
        Exit Sub
    End If
        With frmMain.dlgCD
        .Flags = cdlOFNOverwritePrompt
        .DialogTitle = "Copy to" & Chr$(133)
        .CancelError = True
        .ShowSave
        If Len(.FileName) = 0 Then Exit Sub
        ToFile = .FileName
    End With
    If Len(frmMain.dlgCD.FileName) = 0 Then Exit Sub
    ToFile = frmMain.dlgCD.FileName
    FileCopy FileNameStr, ToFile
    Exit Sub
10:
If Err.Number = 32755 Then Exit Sub
    ErrorTrap
End Sub
Public Sub cmdMove_Click()
   On Error GoTo 10
   Dim ToFile As String
    If FileNameStr = "" Then
        MsgBox "No file loaded.", vbExclamation + vbOKOnly, "Error"
        Exit Sub
    End If
    With frmMain.dlgCD
        .Flags = cdlOFNOverwritePrompt
        .DialogTitle = "Move to..."
        .CancelError = True
        .ShowSave
        If Len(.FileName) = 0 Then Exit Sub
        ToFile = .FileName
    End With
    FileCopy FileNameStr, ToFile
    Kill FileNameStr
   FileNameStr = ToFile
   GetFileInfo
   Exit Sub
10:
If Err.Number = 32755 Then Exit Sub
    ErrorTrap
End Sub

Private Sub cmdProperties_Click()
frmProperties.FileNameString = FileNameStr
frmProperties.Show
End Sub

Private Sub cmdOK_Click()
Unload Me
End Sub

Private Sub Command1_Click()
    If chkReadOnly.Value = Checked Then chkReadOnly.Value = Unchecked
    If chkArchive.Value = Checked Then chkArchive.Value = Unchecked
    If chkSystem.Value = Checked Then chkSystem.Value = Unchecked
    If chkHidden.Value = Checked Then chkHidden.Value = Unchecked
    If chkTemporary.Value = Checked Then chkTemporary.Value = Unchecked
    If chkCompressed.Value = Checked Then chkCompressed.Value = Unchecked
    SetAttr FileNameStr, vbNormal
End Sub

Private Sub Form_Load()
On Error GoTo 10
Dim nofile As Boolean
PicFile = frmMain.dlgCD.FileName
FileNameStr = PicFile
lblName.Caption = frmMain.dlgCD.FileName
If frmMain.dlgCD.FileName = "" Then
nofile = True
ToggleEnabled (False)
Exit Sub
End If
GetFileInfo
10:
    ErrorTrap
End Sub
Public Sub GetFileInfo()
lblName.Caption = FileNameStr
GetFileAttributes
GetTextAttributes
lblModified.Caption = FileDateTime(FileNameStr)
lblSize.Caption = GetFileSize(FileNameStr)
End Sub

Public Function GetFileAttributes()
Dim FileAttr As Long
If FileNameStr = "" Then Exit Function
FileAttr = GetAttr(FileNameStr)
    If FileAttr And vbReadOnly Then
        chkReadOnly.Value = Checked
    End If
    If FileAttr And vbArchive Then
        chkArchive.Value = Checked
    End If
    If FileAttr And vbSystem Then
        chkSystem.Value = Checked
    End If
    If FileAttr And vbHidden Then
        chkHidden.Value = Checked
    End If
    If FileAttr And vbNormal Then
        chkNormal.Value = Checked
    End If
    If FileAttr And vbTemporary Then
        chkTemporary.Value = Checked
    End If
End Function

Public Function GetTextAttributes()
If FileNameStr = "" Then Exit Function
lblAttr.Caption = ""
If GetAttr(FileNameStr) And vbAlias Then
    If lblAttr.Caption = "" Then
        lblAttr.Caption = lblAttr.Caption & "Alias"
    Else
        lblAttr.Caption = lblAttr.Caption & ", Alias"
    End If
End If
If GetAttr(FileNameStr) And vbArchive Then
    If lblAttr.Caption = "" Then
        lblAttr.Caption = lblAttr.Caption & "Archive"
    Else
        lblAttr.Caption = lblAttr.Caption & ", Archive"
    End If
End If
If GetAttr(FileNameStr) And vbDirectory Then
    If lblAttr.Caption = "" Then
        lblAttr.Caption = lblAttr.Caption & "Directory"
    Else
        lblAttr.Caption = lblAttr.Caption & ", Directory"
    End If
End If
If GetAttr(FileNameStr) And vbHidden Then
    If lblAttr.Caption = "" Then
        lblAttr.Caption = lblAttr.Caption & "Hidden"
    Else
        lblAttr.Caption = lblAttr.Caption & ", Hidden"
    End If
End If
If GetAttr(FileNameStr) And vbNormal Then
    If lblAttr.Caption = "" Then
        lblAttr.Caption = lblAttr.Caption & "Normal"
    Else
        lblAttr.Caption = lblAttr.Caption & ", Normal"
    End If
End If
If GetAttr(FileNameStr) And vbReadOnly Then
    If lblAttr.Caption = "" Then
        lblAttr.Caption = lblAttr.Caption & "Read-Only"
    Else
        lblAttr.Caption = lblAttr.Caption & ", Read-Only"
    End If
End If
If GetAttr(FileNameStr) And vbSystem Then
    If lblAttr.Caption = "" Then
        lblAttr.Caption = lblAttr.Caption & "System"
    Else
        lblAttr.Caption = lblAttr.Caption & ", System"
    End If
End If
If GetAttr(FileNameStr) And vbVolume Then
    If lblAttr.Caption = "" Then
        lblAttr.Caption = lblAttr.Caption & "Volume"
    Else
        lblAttr.Caption = lblAttr.Caption & ", Volume"
    End If
End If
If GetAttr(FileNameStr) And vbCompressed Then
    If lblAttr.Caption = "" Then
        lblAttr.Caption = lblAttr.Caption & "Compressed"
    Else
        lblAttr.Caption = lblAttr.Caption & ", Compressed"
    End If
End If
If GetAttr(FileNameStr) And vbTemporary Then
    If lblAttr.Caption = "" Then
        lblAttr.Caption = lblAttr.Caption & "Temporary"
    Else
        lblAttr.Caption = lblAttr.Caption & ", Temporary"
    End If
End If
End Function
Public Function GetFileSize(FileName) As String
    On Error GoTo 10
    Dim sTemp As String
    sTemp = FileLen(FileName)
    If sTemp >= "1024" Then
        sTemp = CCur(sTemp / 1024) & " KB"
    Else
        If sTemp >= "1048576" Then
            sTemp = CCur(sTemp / (1024 * 1024)) & " KB"
        Else
            sTemp = CCur(sTemp) & "Bytes"
        End If
    End If
    GetFileSize = sTemp
10:
    If Err.Number = 0 Then Exit Function
    GetFileSize = "0 Bytes"
    ErrorTrap
End Function
Public Function ToggleEnabled(TrueFalse As Boolean)
cmdMove.Enabled = TrueFalse
cmdCopy.Enabled = TrueFalse
cmdDelete.Enabled = TrueFalse
cmdApply.Enabled = TrueFalse
chkNormal.Enabled = TrueFalse
chkReadOnly.Enabled = TrueFalse
chkHidden.Enabled = TrueFalse
chkArchive.Enabled = TrueFalse
chkSystem.Enabled = TrueFalse
chkTemporary.Enabled = TrueFalse
End Function
