VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Extreme RSS"
   ClientHeight    =   9000
   ClientLeft      =   2415
   ClientTop       =   735
   ClientWidth     =   12000
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboCategory 
      Height          =   315
      Left            =   80
      TabIndex        =   1
      Text            =   "C:\RSS"
      Top             =   1065
      Width           =   3975
   End
   Begin MSComctlLib.ImageList ImgList 
      Left            =   4560
      Top             =   7800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":09EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0F84
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":131E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":16B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1A52
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1DEC
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2186
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2520
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2ABA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin SHDocVwCtl.WebBrowser webFeeds 
      Height          =   7320
      Left            =   4100
      TabIndex        =   4
      Top             =   1065
      Width           =   7835
      ExtentX         =   13820
      ExtentY         =   12912
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.ListBox lstHeadlines 
      Height          =   4155
      Left            =   80
      TabIndex        =   3
      Top             =   4530
      Width           =   3975
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   7
      Top             =   8700
      Width           =   12000
      _ExtentX        =   21167
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1958
            MinWidth        =   1411
            TextSave        =   "22/08/2005"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4410
            MinWidth        =   4410
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   14684
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.FileListBox fileFeeds 
      Height          =   2625
      Left            =   80
      TabIndex        =   2
      Top             =   1620
      Width           =   3975
   End
   Begin ComCtl3.CoolBar CoolBar2 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   6
      Top             =   390
      Width           =   12000
      _ExtentX        =   21167
      _ExtentY        =   741
      BandCount       =   2
      FixedOrder      =   -1  'True
      _CBWidth        =   12000
      _CBHeight       =   420
      _Version        =   "6.7.9782"
      Caption1        =   "Address"
      Child1          =   "cboAddress"
      MinHeight1      =   360
      Width1          =   1095
      NewRow1         =   0   'False
      Child2          =   "ToolBar2"
      MinHeight2      =   330
      NewRow2         =   0   'False
      BandStyle2      =   1
      Begin VB.ComboBox cboAddress 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   735
         TabIndex        =   0
         Top             =   30
         Width           =   11145
      End
      Begin MSComctlLib.Toolbar ToolBar2 
         Height          =   330
         Left            =   11970
         TabIndex        =   9
         Top             =   45
         Width           =   30
         _ExtentX        =   53
         _ExtentY        =   582
         ButtonWidth     =   1138
         ButtonHeight    =   582
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "ImgList"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Go"
               Key             =   "Go"
               Object.ToolTipText     =   "Go to RSS feed"
               ImageIndex      =   1
            EndProperty
         EndProperty
      End
   End
   Begin ComCtl3.CoolBar CoolBar1 
      Align           =   1  'Align Top
      Height          =   390
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   12000
      _ExtentX        =   21167
      _ExtentY        =   688
      BandCount       =   1
      _CBWidth        =   12000
      _CBHeight       =   390
      _Version        =   "6.7.9782"
      Child1          =   "ToolBar1"
      MinHeight1      =   330
      Width1          =   3975
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar ToolBar1 
         Height          =   330
         Left            =   30
         TabIndex        =   8
         Top             =   30
         Width           =   11880
         _ExtentX        =   20955
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
         ImageList       =   "ImgList"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   9
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Save"
               Object.ToolTipText     =   "Save feed"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Rename"
               Object.ToolTipText     =   "Rename feed"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Delete"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Open"
               Object.ToolTipText     =   "Open feed"
               ImageIndex      =   9
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Create"
               Object.ToolTipText     =   "Create new folder"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "DeleteFolder"
               Object.ToolTipText     =   "Delete folder"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Exit"
               Object.ToolTipText     =   "Exit Extreme RSS"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "About"
               Object.ToolTipText     =   "About Extreme RSS"
               ImageIndex      =   5
            EndProperty
         EndProperty
      End
   End
   Begin VB.Label lblDefault 
      AutoSize        =   -1  'True
      Caption         =   "To open this feed in your default browser, click here."
      Height          =   195
      Left            =   7020
      TabIndex        =   14
      Top             =   8460
      Width           =   4545
   End
   Begin VB.Image ImgURL 
      Height          =   240
      Left            =   11640
      Picture         =   "frmMain.frx":2E54
      Top             =   8445
      Width           =   240
   End
   Begin VB.Label lblSubscribed 
      Caption         =   "Subscribed Feeds"
      Height          =   255
      Left            =   80
      TabIndex        =   13
      Top             =   1410
      Width           =   2055
   End
   Begin VB.Label lblCategory 
      Caption         =   "Category(s)"
      Height          =   195
      Left            =   80
      TabIndex        =   12
      Top             =   855
      Width           =   1305
   End
   Begin VB.Label lblDescription 
      Caption         =   "Description"
      Height          =   255
      Left            =   4110
      TabIndex        =   11
      Top             =   855
      Width           =   945
   End
   Begin VB.Label lblHeadlines 
      Caption         =   "Feed headlines"
      Height          =   165
      Left            =   80
      TabIndex        =   10
      Top             =   4320
      Width           =   1935
   End
   Begin VB.Menu File 
      Caption         =   "&File"
      Begin VB.Menu Close 
         Caption         =   "&Close"
      End
   End
   Begin VB.Menu Help 
      Caption         =   "&Help"
      Begin VB.Menu About 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oRSS As MSXML2.DOMDocument
Dim oItemList() As MSXML2.IXMLDOMNode

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Dim strURL As String
Dim strFeed As String
Dim strPubDate As String
Dim strHeadlines As String
Dim FeedURL As String

' Global FileSystemObject settings
Dim FSys As New FileSystemObject
Dim FSysFile As Object
Dim FSysFolder As Object

Private WithEvents m_frmSysTray As frmSysTray
Attribute m_frmSysTray.VB_VarHelpID = -1

Private Sub Form_Load()
    
    ' Set the caption for the form.
    Me.Caption = "Extreme RSS v" & App.Major & "." & App.Minor & "." & _
                 App.Revision
                 
    ' Set a few bits up.
    CoolBar2.Bands(2).MinWidth = 700
    StatusBar.Panels(1).Width = 100
    fileFeeds.Pattern = "*"
    cboCategory.AddItem "C:\RSS"
    
    ' Check the HTML directory is there.
    Call CheckHTML
    ' Call SystemTray to enable system tray goodness.
    Call SystemTray
    ' Write the HTML file incase it has been deleted.
    Call WriteHTML
    ' Write the HTML file incase it has been delete.
    Call WriteFeeds
    ' Get the directory where all the saved feeds will be stored.
    Call GetDirectory
    ' Now populate the Category Combo with the sub directories in C:\RSS
    Call FillCategory
    
    ' Now navigate to the default HTML page.
    webFeeds.Navigate App.Path & "\HTML\ExtremeRSS.html"
    
End Sub

Private Function GetRSS()
    
    ' This just makes sure everything is nice and clean.
    lstHeadlines.Clear
    webFeeds.Navigate "about:blank"
    strHeadlines = ""
    strURL = ""
    strFeed = ""
    strPubDate = ""
    DoEvents
    
    ' Disbale fileFeeds and let the user know we are getting the feeds.
    fileFeeds.Enabled = False
    StatusBar.Panels(2).Text = "Please wait getting feeds..."
    DoEvents
    
    Dim oItems As MSXML2.IXMLDOMNodeList
    Dim i As Integer
    Dim oNode As IXMLDOMNode
    
    Set oRSS = New MSXML2.DOMDocument
    oRSS.async = False
    oRSS.Load (cboAddress.Text)
    
    Set oItems = oRSS.selectNodes("rss/channel/item")

    i = -1
    
    ReDim oItemList(oItems.length)
    
    For Each oNode In oItems
        i = i + 1
        lstHeadlines.AddItem oNode.selectSingleNode("title").Text
        Set oItemList(i) = oNode
    Next oNode
    
    ' Let the user know we are done.
    fileFeeds.Enabled = True
    StatusBar.Panels(2).Text = "Retrieved " & lstHeadlines.ListCount & " feeds."
    DoEvents
    
    ' Display a ballon tip! :)
    m_frmSysTray.ShowBalloonTip _
    "Retrieved " & lstHeadlines.ListCount & " feeds.", _
    "Extreme RSS", _
    NIIF_INFO

End Function

Private Function GetHeadlines()

    Dim oNode As MSXML2.IXMLDOMNode
    Set oNode = oItemList(lstHeadlines.ListIndex)

        strHeadlines = oNode.selectSingleNode("title").Text
        strURL = oNode.selectSingleNode("link").Text
        strFeed = oNode.selectSingleNode("description").Text
    
    
    Dim oTestNode As MSXML2.IXMLDOMNode
    Set oTestNode = oNode.selectSingleNode("pubDate|dc:date")
    If oTestNode Is Nothing Then
    
        ' If the Node does not exist we simply display No Information.
        strPubDate = "No Information"
        
    Else
    
        ' Next statement finds the first node matching any of the tags in the list
        strPubDate = oNode.selectSingleNode("pubDate|dc:date").Text
        
    End If
       
    ' Now we write the info to the HTML page.
    Call WriteFeeds

    ' Now we display the web page to the user.
    webFeeds.Navigate App.Path & "\HTML\ExtremeFeeds.html"
   
End Function

Private Function OpenFeed()

    ' This just put the address of you favorite feeds in the address bar
    ' when you open them.
    Dim InStream As TextStream
    Set InStream = FSys.OpenTextFile(frmMain.cboCategory & "\" & fileFeeds.FileName)
    Dim noth As String
        FeedURL = InStream.ReadLine
        InStream.Close '< -EOF
        
    cboAddress.Text = FeedURL
    cboAddress.Text = Replace(cboAddress.Text, Chr(10), "")
    cboAddress.Text = Replace(cboAddress.Text, Chr(13), "")
    cboAddress.Text = Trim(cboAddress.Text)
    fileFeeds.Refresh
    
    ' Now we check that the address has gone in ok.
    Call CheckAddress
    
End Function

Private Function CheckHTML()

    ' Make sure the HTML directory is there.
    If FSys.FolderExists(App.Path & "\HTML") Then
        Exit Function
    Else
        FSys.CreateFolder (App.Path & "\HTML")
    End If

End Function


Private Sub lstHeadlines_Click()
    
    ' This gets the headlines and displayed them in webFeeds.
    Call GetHeadlines
    
End Sub

Private Function GetDirectory()

    ' Make sure that C:\RSS is there, if not create it.
    If FSys.FolderExists(cboCategory) Then
        fileFeeds.Path = cboCategory
    Else
        FSys.CreateFolder ("C:\RSS")
        fileFeeds.Path = cboCategory
    End If
    
End Function

Private Sub FillCategory()

    ' Clear the Combo just incase.
    cboCategory.Clear
    
    ' Now loop through C:\RSS and add the folder to the Combo.
    For Each FSysFolder In FSys.GetFolder("C:\RSS\").SubFolders
        cboCategory.AddItem FSysFolder
    Next
    
    ' Now add the Default directory and select it.
    cboCategory.AddItem "C:\RSS"
    cboCategory.SelText = "C:\RSS"

End Sub

Private Function DeleteFeed()

    ' This deletes the selected feed.
    Dim iDeleteFeed As Integer
    
    If fileFeeds.FileName = "" Then
        MsgBox "No feed to delete!", vbInformation, "Cannot delete"
    Else
        iDeleteFeed = MsgBox("Are you sure you want to delete " & fileFeeds.FileName & "?", vbYesNo + vbQuestion, "Confirm delete")
        If iDeleteFeed = vbYes Then
            FSys.DeleteFile (frmMain.cboCategory & "\" & fileFeeds.FileName)
            Call FillCategory
            fileFeeds.Refresh
        End If
    End If
    
End Function

Private Function DeleteFolder()
    
    ' This deletes the selected folder.
    Dim iDeleteFolder As Integer
    
    If cboCategory.Text = "" Then
        MsgBox "No folder to delete!", vbInformation, "Cannot delete"
    Else
        iDeleteFolder = MsgBox("Are you sure you want to delete " & cboCategory & "?", vbYesNo + vbQuestion, "Confirm delete")
        If iDeleteFolder = vbYes Then
            FSys.DeleteFolder cboCategory
            Call FillCategory
            fileFeeds.Refresh
        End If
    End If

End Function

Private Sub cboAddress_Click()

    ' This calls GetRSS which gets the feed from the internet.
    Call GetRSS

End Sub

Private Sub cboAddress_KeyPress(KeyAscii As Integer)
    
    ' Call GetRSS when return is pressed.
    On Error Resume Next
    If KeyAscii = vbKeyReturn Then
        Call GetRSS
    End If
    
End Sub

Private Sub cboCategory_Click()
    
    ' Tidy up a bit.
    lstHeadlines.Clear
    cboAddress.Text = ""
    webFeeds.Navigate "about:blank"
    strHeadlines = ""
    strURL = ""
    strFeed = ""
    strPubDate = ""
    DoEvents
    
    ' Set the fileFeeds path to that listed in cboCategory.
    If FSys.FolderExists(cboCategory) Then
        fileFeeds.Path = cboCategory
    End If
    
End Sub

Private Sub ImgURL_Click()
    
    ' This is for users who do not want the full article to be opened in IE.
    ' this just opens the URL in the users default browser when the button is clicked.
    Dim retValue As Long
    
    If strURL <> "" Then
        retValue = ShellExecute(frmMain.hwnd, "Open", strURL, 0&, 0&, 0&)
    Else
        MsgBox "No feed to open!", vbInformation, "No feed available"
    End If
    
End Sub

Private Function CheckAddress()
    
    ' This just checks that the address is not blank when called.
    If cboAddress.Text = "" Then
        MsgBox "No URL to open!", vbInformation, "No URL"
        Exit Function
    End If

End Function

Private Sub About_Click()
    
    ' Display about information when About is clicked.
    MsgBox "Extreme RSS v" & App.Major & "." & App.Minor & "." & _
            App.Revision & vbNewLine & vbNewLine & _
           "Written by Martin Dunsmore", vbInformation, "Extreme RSS"

End Sub

Private Sub Close_Click()
    
    ' Do I have to really?
    Unload frmCreate
    Unload frmRename
    Unload frmSave
    Unload frmMain
    End
    
End Sub

Private Sub ToolBar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    
    ' This the Toolbar buttons.
    On Error Resume Next
     
    Select Case Button.Key
        Case "Save"
            ' Call the save feeds form to save a feed.
            Call SaveFeed
        Case "Rename"
            ' Call the rename form to rename a feed.
            If fileFeeds.FileName = "" Then
                MsgBox "You must select a file before you can rename one!", vbInformation, "No file selected"
            Else
                frmRename.Show
            End If
        Case "Delete"
            ' This will delete a feed.
            Call DeleteFeed
        Case "Open"
            ' Open feed again.
            Call OpenFeed
            Call GetRSS
        Case "Create"
            ' Create a new folder.
            frmCreate.Show
        Case "DeleteFolder"
            ' Delete a folder.
            Call DeleteFolder
        Case "Exit"
            ' Really?
            Unload frmCreate
            Unload frmRename
            Unload frmSave
            Unload frmMain
            End
        Case "About"
            ' Same as other About message.
            MsgBox "Extreme RSS v" & App.Major & "." & App.Minor & "." & _
                   App.Revision & vbNewLine & vbNewLine & _
                   "Written by Martin Dunsmore", vbInformation, "Extreme RSS"
    End Select
    
End Sub

Private Function SaveFeed()
    
    ' Call the save feed form. First we check that the address is not blank.
    If cboAddress <> "" Then
        frmSave.Show
    Else
        MsgBox "No RSS feed to save!", vbInformation, "No feed to save"
    End If

End Function

Private Sub ToolBar2_ButtonClick(ByVal Button As MSComctlLib.Button)
    
    ' This the go button, this will open a feed that has been typed in.
    On Error Resume Next
     
    Select Case Button.Key
        Case "Go"
            Call GetRSS
    End Select
    
End Sub

Private Sub fileFeeds_Click()
    
    ' Just display what feed is currently active in the title abr.
    Me.Caption = "Extreme RSS - " & cboCategory & "\" & fileFeeds.FileName
    
End Sub
Private Sub fileFeeds_DblClick()
    
    ' Just display what feed is currently active in the title abr.
    Me.Caption = "Extreme RSS - " & cboCategory & "\" & fileFeeds.FileName
    
    ' As it is a double click we will open the feed.
    Call OpenFeed
    Call GetRSS
    
End Sub

Private Sub SystemTray()

    ' Minimize to System Tray stuff
    Set m_frmSysTray = New frmSysTray
    With m_frmSysTray
        .AddMenuItem "&Open Extreme RSS", "open", True
        .AddMenuItem "&Minimize to Tray", "minimize"
        .AddMenuItem "-"
        .AddMenuItem "&Close", "close"
        .ToolTip = "Extreme RSS"
        .IconHandle = Me.Icon.Handle
    End With

End Sub

Private Sub webFeeds_StatusTextChange(ByVal Text As String)
    
    ' Show status of webFeeds in the StatusBar.
    StatusBar.Panels(3).Text = Text

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    ' Minimize to System Tray stuff.
    Unload m_frmSysTray
    Set m_frmSysTray = Nothing
    
End Sub

Private Sub m_frmSysTray_MenuClick(ByVal lIndex As Long, ByVal sKey As String)
    
    ' Minimize to System Tray stuff.
    Select Case sKey
        Case "open"
            Me.Show
            Me.ZOrder
        Case "minimize"
            Me.Hide
        Case "close"
            Unload Me
    End Select

End Sub

Private Sub m_frmSysTray_SysTrayDoubleClick(ByVal eButton As MouseButtonConstants)
    
    ' Minimize to System Tray stuff.
    If frmMain.Visible = False Then
        Me.Show
        Me.ZOrder
    Else
        Me.Hide
    End If
    
End Sub

Private Sub m_frmSysTray_SysTrayMouseDown(ByVal eButton As MouseButtonConstants)
    
    ' Minimize to System Tray stuff.
    If (eButton = vbRightButton) Then
        m_frmSysTray.ShowMenu
    End If
    
End Sub

Private Function WriteFeeds()
    
    ' This is the HTML that will display the feed.
    Open App.Path & "\HTML\ExtremeFeeds.html" For Output As #1
    Print #1, "<html>"
    Print #1, "<head>"
    Print #1, "<title>" & strHeadlines & "</title>"
    Print #1, "<style type=""text/css"">"
    Print #1, "<!--"
    Print #1, "body,td,th {color: #383C45;font-family: Verdana, Arial, Helvetica, sans-serif;}"
    Print #1, "body {background-color: #F4F5F9;margin-left: 5px;margin-top: 5px;margin-right: 5px;margin-bottom: 5px;}"
    Print #1, "a:link {color: #2F4D8B;text-decoration: none;}"
    Print #1, "a:visited {text-decoration: none;color: #2F4D8B;}"
    Print #1, "a:hover {text-decoration: underline;color: #2F4D8B;}"
    Print #1, "a:active {text-decoration: none;color: #2F4D8B;}"
    Print #1, ".style2 {font-size: xx-small;color: #797C83;}"
    Print #1, "-->"
    Print #1, "</style></head>"
    Print #1, "<body><table width=""100%"">"
    Print #1, "<tr>"
    Print #1, "<td width=""2%""><img src=""Bullet.jpg"" border=""0""></td>"
    Print #1, "<td width=""98%""><a href=" & strURL & " target=""_blank""><strong>" & strHeadlines & "</strong></a></td>"
    Print #1, "</tr>"
    Print #1, "<tr>"
    Print #1, "<td>&nbsp;</td>"
    Print #1, "<td><span class=""style2""><strong>Published Date:</strong> " & strPubDate & "</span></td>"
    Print #1, "</tr>"
    Print #1, "<tr>"
    Print #1, "<td>&nbsp;</td>"
    Print #1, "<td>" & strFeed & "</td>"
    Print #1, "</tr>"
    Print #1, "</table>"
    Print #1, "</body>"
    Print #1, "</html>"
    Close #1

End Function

Private Function WriteHTML()

    ' This is the HTML you see when you first open the application.
    Open App.Path & "\HTML\ExtremeRSS.html" For Output As #1
    Print #1, "<html>"
    Print #1, "<head>"
    Print #1, "<title>Extreme RSS</title>"
    Print #1, "<style type=""text/css"">"
    Print #1, "<!--"
    Print #1, "body {background-color: #F4F5F9;margin-left: 5px;margin-top: 5px;margin-right: 5px;margin-bottom: 5px;font-family: Verdana, Arial, Helvetica, sans-serif;}"
    Print #1, ".style1 {font-size: xx-large;font-weight: bold;}"
    Print #1, "-->"
    Print #1, "</style></head>"
    Print #1, "<body>"
    Print #1, "<div align=""center"">"
    Print #1, "<p class=""style1""><u>Extreme RSS</u></p>"
    Print #1, "<p>Written by Martin Dunsmore</p>"
    Print #1, "<p>v" & App.Major & "." & App.Minor & "." & App.Revision & "</p>"
    Print #1, "</div>"
    Print #1, "</body>"
    Print #1, "</html>"
    Close #1

End Function
