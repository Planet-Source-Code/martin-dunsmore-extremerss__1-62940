VERSION 5.00
Begin VB.Form frmRename 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Rename feed"
   ClientHeight    =   1200
   ClientLeft      =   4395
   ClientTop       =   4245
   ClientWidth     =   5535
   Icon            =   "frmRename.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1200
   ScaleWidth      =   5535
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4200
      TabIndex        =   3
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton cmdRename 
      Caption         =   "&Rename"
      Height          =   375
      Left            =   2880
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox txtRename 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   5295
   End
   Begin VB.Label lblRename 
      Caption         =   "Enter new name"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "frmRename"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim FileName As String
Dim NewFileName As String

Private Sub cmdRename_Click()

    On Error GoTo Error_Handler

    txtRename.Text = Trim(txtRename.Text)
    
    FileName = frmMain.cboCategory & "\" & frmMain.fileFeeds.FileName
    NewFileName = frmMain.cboCategory & "\" & txtRename.Text
    Name FileName As NewFileName
    frmMain.fileFeeds.Refresh
    Unload frmRename
    
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

    txtRename.Text = ""
    Unload frmRename
    
End Sub

Private Sub txtRename_KeyPress(KeyAscii As Integer)

    On Error Resume Next
    
    If KeyAscii = vbKeyReturn Then
        cmdRename_Click
    End If

End Sub
