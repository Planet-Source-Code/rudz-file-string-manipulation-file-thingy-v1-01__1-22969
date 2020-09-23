VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "FileString Manipulation v1.01 - By Rudy Alex Kohn"
   ClientHeight    =   1365
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9060
   ClipControls    =   0   'False
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1365
   ScaleWidth      =   9060
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command8 
      Caption         =   "GetFileNoExt"
      Height          =   375
      Left            =   5400
      TabIndex        =   9
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Quit"
      Height          =   375
      Left            =   8040
      TabIndex        =   8
      Top             =   840
      Width           =   975
   End
   Begin VB.CommandButton Command6 
      Caption         =   "About"
      Height          =   375
      Left            =   6960
      TabIndex        =   7
      Top             =   840
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Get Drive"
      Height          =   375
      Left            =   4080
      TabIndex        =   5
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Get Path"
      Height          =   375
      Left            =   2760
      TabIndex        =   4
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Get Extension"
      Height          =   375
      Left            =   1440
      TabIndex        =   3
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Get Filename"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   8640
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   "Select File"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8895
      Begin VB.CommandButton Command5 
         Caption         =   ".."
         Height          =   285
         Left            =   8520
         TabIndex        =   6
         Top             =   240
         Width           =   255
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   8295
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
  If LenB(Text1.Text) <> 0 Then MsgBox GetFileName(Text1), 64, "GetFileName"
End Sub

Private Sub Command2_Click()
  If LenB(Text1) <> 0 Then
    MsgBox GetFileExtension(Text1, False), 64, "GetFileExtension"
    MsgBox GetFileExtension(Text1), 64, "GetFileExtension - Lowercase"
  End If
End Sub

Private Sub Command3_Click()
  If LenB(Text1) <> 0 Then
    MsgBox GetFilePath(Text1), 64, "GetFilePath - w/ Drive"
    MsgBox GetFilePath(Text1, False), 64, "GetFilePath - wo/ Drive"
  End If
End Sub

Private Sub Command4_Click()
    If LenB(Text1.Text) <> 0 Then
        MsgBox GetDrive(Text1.Text), 64, "GetDrive"
        MsgBox GetDrive(Text1.Text, True), 64, "GetDrive - w/ bslash"
    End If
End Sub

Private Sub Command5_Click()
With cd
    .CancelError = False
    .FileName = vbNullString
    .ShowOpen
    If LenB(.FileName) = 0 Then Exit Sub
    Text1.Text = .FileName
End With
End Sub

Private Sub Command6_Click()
    MsgBox "Example by Rudy Alex Kohn." & vbCr & _
           "Made in a hurry for quick use, developement time < 5 min. =)" & vbCr & _
           "Contact me at rudyalexkohn@hotmail.com", 64, Me.Caption
End Sub

Private Sub Command7_Click()
    Unload Me
    End
End Sub

Private Sub Command8_Click()
  If LenB(Text1) <> 0 Then MsgBox GetFileNoExtension(Text1), 64, "GetFileNameNoExtension"
End Sub

Private Sub Form_Load()
    Text1 = App.Path & "\" & App.EXEName & ".Exe"
End Sub
