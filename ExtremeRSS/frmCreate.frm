VERSION 5.00
Begin VB.Form frmCreate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Create new folder"
   ClientHeight    =   1200
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5535
   Icon            =   "frmCreate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1200
   ScaleWidth      =   5535
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4200
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton cmdCreate 
      Caption         =   "&Save"
      Height          =   375
      Left            =   2880
      TabIndex        =   1
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox txtFolderName 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   5295
   End
   Begin VB.Label lblFolderName 
      Caption         =   "Enter new folder name"
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1905
   End
End
Attribute VB_Name = "frmCreate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim FSys As New FileSystemObject
Dim FSysFile As Object
Dim FSysFolder As Object

Private Sub cmdCancel_Click()

    Unload frmCreate

End Sub

Private Sub cmdCreate_Click()

    On Error GoTo Error_Handler
    
    txtFolderName.Text = Trim(txtFolderName.Text)

    FSys.CreateFolder "C:\RSS\" & txtFolderName.Text
    
    frmMain.cboCategory.Clear
    
    For Each FSysFolder In FSys.GetFolder("C:\RSS\").SubFolders
        frmMain.cboCategory.AddItem FSysFolder
    Next
    
    frmMain.cboCategory.AddItem "C:\RSS"
    frmMain.cboCategory.SelText = "C:\RSS\" & txtFolderName.Text
    frmMain.fileFeeds.Path = frmMain.cboCategory
    txtFolderName.Text = ""
    Unload frmCreate

Error_Handler:

    If Err = 58 Then
        MsgBox "Folder '" & txtFolderName.Text & "' Already exists. Please choose another name.", vbInformation, "Folder create error"
        txtFolderName.Text = ""
        Exit Sub
    End If
    
    If Err = 52 Then
        MsgBox "Bad file name, please choose another name!", vbInformation, "Bad file name"
        Exit Sub
    End If
    
    If Err = 76 Then
        MsgBox "Bad file name, please choose another name!", vbInformation, "Bad file name"
        Exit Sub
    End If
    
End Sub

Private Sub txtFolderName_KeyPress(KeyAscii As Integer)
    
    On Error Resume Next
    
    If KeyAscii = vbKeyReturn Then
        cmdCreate_Click
    End If
    
End Sub
