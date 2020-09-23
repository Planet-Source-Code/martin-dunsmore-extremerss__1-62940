VERSION 5.00
Begin VB.Form frmSave 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Save new feed"
   ClientHeight    =   1200
   ClientLeft      =   3945
   ClientTop       =   3330
   ClientWidth     =   5535
   Icon            =   "frmSave.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1200
   ScaleWidth      =   5535
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4200
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   375
      Left            =   2880
      TabIndex        =   1
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox txtSave 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   5295
   End
   Begin VB.Label lblSave 
      Caption         =   "Enter name for new feed"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "frmSave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim FSys As New FileSystemObject
Dim FSysFile As Object
Dim FSysFolder As Object

Private Sub cmdSave_Click()

    On Error GoTo Error_Handler
    
    txtSave.Text = Trim(txtSave.Text)
    
    If txtSave.Text = "" Then
        MsgBox "Warning: File Name must not be blank!", vbInformation, "Save error"
        Exit Sub
    End If
    
    If FSys.FileExists(frmMain.cboCategory & "\" & txtSave.Text) Then
        MsgBox "File already exists, Please choose another name!", vbInformation, "File already exists"
        Exit Sub
    Else
        GoTo CreateFile
    End If
    
CreateFile:

    Open frmMain.cboCategory & "\" & txtSave.Text For Output As #1
    Print #1, Trim(frmMain.cboAddress.Text)
    Close #1
    frmMain.fileFeeds.Refresh
    Unload frmSave
    
Error_Handler:
    
    If Err = 52 Then
        MsgBox "Bad file name, please choose another name!", vbInformation, "Bad file name"
        Exit Sub
    End If
    
    If Err = 76 Then
        MsgBox "Bad file name, please choose another name!", vbInformation, "Bad file name"
        Exit Sub
    End If
    
End Sub

Private Sub cmdCancel_Click()

    Unload frmSave
    
End Sub

Private Sub Form_Load()

    txtSave = ""
    
End Sub

Private Sub txtSave_KeyPress(KeyAscii As Integer)
    
    On Error Resume Next
    
    If KeyAscii = vbKeyReturn Then
        cmdSave_Click
    End If

End Sub
