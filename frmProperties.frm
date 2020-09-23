VERSION 5.00
Begin VB.Form frmProperties 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Properties"
   ClientHeight    =   3030
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5535
   Icon            =   "frmProperties.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   5535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      Height          =   375
      Left            =   2940
      TabIndex        =   6
      Top             =   2520
      Width           =   1185
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   4170
      TabIndex        =   5
      Top             =   2520
      Width           =   1185
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
      Height          =   315
      Left            =   435
      TabIndex        =   8
      Top             =   2085
      Width           =   1695
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
      Left            =   2145
      TabIndex        =   7
      Top             =   2085
      Width           =   3225
   End
   Begin VB.Label lblAttr 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   2130
      TabIndex        =   4
      Top             =   1500
      Width           =   3225
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
      Left            =   420
      TabIndex        =   3
      Top             =   1500
      Width           =   1695
   End
   Begin VB.Label lblModified 
      BackColor       =   &H00FFFFFF&
      Caption         =   "No file loaded"
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
      Left            =   2130
      TabIndex        =   2
      Top             =   1110
      Width           =   3225
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
      Height          =   315
      Left            =   420
      TabIndex        =   1
      Top             =   1110
      Width           =   1695
   End
   Begin VB.Label lblFileName 
      BackColor       =   &H00FFFFFF&
      Caption         =   "-Error reading filename-"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   870
      TabIndex        =   0
      Top             =   210
      Width           =   4485
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   210
      Picture         =   "frmProperties.frx":030A
      Top             =   210
      Width           =   570
   End
End
Attribute VB_Name = "frmProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const vbCompressed = &H800
Public FileNameString As String


Private Sub cmdOK_Click()
Unload Me
End Sub

Private Sub cmdRefresh_Click()
Form_Load
End Sub

Private Sub Form_Load()
On Error GoTo 10
'Dim nofile As Boolean
'nofile = FileNameString = ""
If FileNameString = "" Then Exit Sub
lblFileName.Caption = FileNameString
lblModified.Caption = FileDateTime(FileNameString)
lblSize.Caption = GetFileSize(FileNameString)
10:
If Err.Number = "53" Then
lblAttr.Caption = "No File Attributes"
lblModified.Caption = "File hasn't been saved yet"
lblSize.Caption = "0 Bytes"
End If
Exit Sub
End Sub


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
    frmFileOperations.ErrorTrap
End Function


