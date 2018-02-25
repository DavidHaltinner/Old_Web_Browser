VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmBrowser 
   Caption         =   "WebExpress"
   ClientHeight    =   7140
   ClientLeft      =   -150
   ClientTop       =   690
   ClientWidth     =   11055
   Icon            =   "browserform.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7140
   ScaleWidth      =   11055
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.ComboBox cboAddress 
      Height          =   315
      Left            =   3360
      TabIndex        =   3
      Text            =   "Type URL here"
      Top             =   120
      Width           =   7035
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6240
      Top             =   1920
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   393216
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   225
      Left            =   0
      TabIndex        =   2
      Top             =   6915
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   397
      Style           =   1
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
      EndProperty
   End
   Begin VB.Timer timTimer 
      Enabled         =   0   'False
      Interval        =   120
      Left            =   6180
      Top             =   2880
   End
   Begin SHDocVwCtl.WebBrowser brwWebBrowser 
      Height          =   6480
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   9480
      ExtentX         =   16722
      ExtentY         =   11430
      ViewMode        =   1
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   0
      AutoArrange     =   -1  'True
      NoClientEdge    =   -1  'True
      AlignLeft       =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.PictureBox picAddress 
      BorderStyle     =   0  'None
      Height          =   1030
      Left            =   0
      ScaleHeight     =   1035
      ScaleWidth      =   8775
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1560
      Width           =   8772
   End
   Begin VB.Image Image12 
      Height          =   440
      Left            =   1200
      Picture         =   "browserform.frx":030A
      Stretch         =   -1  'True
      Top             =   80
      Width           =   440
   End
   Begin VB.Image Image11 
      Height          =   440
      Left            =   1920
      Picture         =   "browserform.frx":074C
      Stretch         =   -1  'True
      Top             =   80
      Width           =   440
   End
   Begin VB.Image Image10 
      Height          =   440
      Left            =   600
      Picture         =   "browserform.frx":0B8E
      Stretch         =   -1  'True
      Top             =   80
      Width           =   440
   End
   Begin VB.Image Image9 
      Height          =   440
      Left            =   120
      Picture         =   "browserform.frx":0FD0
      Stretch         =   -1  'True
      Top             =   80
      Width           =   440
   End
   Begin VB.Image Favb 
      Height          =   440
      Left            =   2400
      Picture         =   "browserform.frx":1412
      Stretch         =   -1  'True
      Top             =   80
      Width           =   440
   End
   Begin VB.Image Normal 
      Height          =   440
      Left            =   2880
      Picture         =   "browserform.frx":1854
      Stretch         =   -1  'True
      Top             =   80
      Width           =   440
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   10560
      Picture         =   "browserform.frx":1C96
      ToolTipText     =   "Ahh!!!!!! Get that off of me!!!!! I hate mouse pointers!!!!!"
      Top             =   0
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   9720
      Picture         =   "browserform.frx":1FA0
      Top             =   1560
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   9720
      Picture         =   "browserform.frx":22AA
      Top             =   1920
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image4 
      Height          =   480
      Left            =   9600
      Picture         =   "browserform.frx":25B4
      Top             =   2280
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image5 
      Height          =   480
      Left            =   9600
      Picture         =   "browserform.frx":28BE
      Top             =   2640
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image6 
      Height          =   480
      Left            =   9600
      Picture         =   "browserform.frx":2BC8
      Top             =   3000
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image7 
      Height          =   480
      Left            =   9720
      Picture         =   "browserform.frx":2ED2
      Top             =   3360
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image8 
      Height          =   480
      Left            =   9720
      Picture         =   "browserform.frx":31DC
      Top             =   3840
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Menu file 
      Caption         =   "&File"
      Begin VB.Menu open 
         Caption         =   "Open File"
         Shortcut        =   ^O
      End
      Begin VB.Menu save 
         Caption         =   "Save Page"
         Shortcut        =   ^S
      End
      Begin VB.Menu space 
         Caption         =   "-"
      End
      Begin VB.Menu print 
         Caption         =   "Print Page"
         Shortcut        =   ^P
      End
      Begin VB.Menu space1 
         Caption         =   "-"
      End
      Begin VB.Menu exit 
         Caption         =   "Exit"
         Shortcut        =   ^Z
      End
   End
   Begin VB.Menu command 
      Caption         =   "&Commands"
      Begin VB.Menu back 
         Caption         =   "Back"
         Shortcut        =   ^B
      End
      Begin VB.Menu favorites 
         Caption         =   "Favorites"
         Shortcut        =   ^F
      End
      Begin VB.Menu forward 
         Caption         =   "Forward"
         Shortcut        =   ^W
      End
      Begin VB.Menu stop 
         Caption         =   "Stop"
         Shortcut        =   ^T
      End
      Begin VB.Menu refresh 
         Caption         =   "Refresh"
         Shortcut        =   ^R
      End
      Begin VB.Menu home 
         Caption         =   "Home"
         Shortcut        =   ^H
      End
      Begin VB.Menu search 
         Caption         =   "Search"
         Shortcut        =   ^Y
      End
   End
   Begin VB.Menu fav 
      Caption         =   "&Favorites"
   End
   Begin VB.Menu help 
      Caption         =   "&Help"
      Begin VB.Menu how 
         Caption         =   "How To Use"
      End
      Begin VB.Menu onlinehelp 
         Caption         =   "&Online Help"
      End
      Begin VB.Menu d 
         Caption         =   "-"
      End
      Begin VB.Menu about 
         Caption         =   "About Program"
      End
   End
End
Attribute VB_Name = "frmBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public StartingAddress As String
Dim mbDontNavigateNow As Boolean

Private Sub about_Click()
frmabout.Show
End Sub

Private Sub Back_Click()
brwWebBrowser.GoBack
End Sub

Private Sub brwWebBrowser_DownloadComplete()
Me.cboAddress = brwWebBrowser.LocationURL
End Sub

Private Sub exit_Click()
End
End Sub

Private Sub fav_Click()
Favs.Show
End Sub

Private Sub Favb_Click()
Favs.Show
End Sub

Private Sub favorites_Click()
Favs.Show
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Me.Show
    frmBrowser.Visible = True
    Image1.Left = Me.Width - 700
    brwWebBrowser.Width = Me.ScaleWidth
    brwWebBrowser.Height = Me.ScaleHeight - StatusBar1.Height
    StartingAddress = "http://www.snoopy.com/#thestrip"
    timTimer.Interval = 120

    If Len(StartingAddress) > 0 Then
        cboAddress.Text = StartingAddress
        cboAddress.AddItem cboAddress.Text
        timTimer.Enabled = True
        brwWebBrowser.Navigate StartingAddress
    End If


End Sub



Private Sub brwWebBrowser_NavigateComplete(ByVal URL As String)
    Dim i As Integer
    Dim bFound As Boolean
    Me.Caption = brwWebBrowser.LocationName & " - WebExpress"
    For i = 0 To cboAddress.ListCount - 1
        If cboAddress.List(i) = brwWebBrowser.LocationURL Then
            bFound = True
            Exit For
        End If
    Next i
    mbDontNavigateNow = True
    If bFound Then
        cboAddress.RemoveItem i
    End If
    cboAddress.AddItem brwWebBrowser.LocationURL, 0
    cboAddress.ListIndex = 0
    mbDontNavigateNow = False
End Sub


Private Sub cboAddress_Click()
    If mbDontNavigateNow Then Exit Sub
    timTimer.Enabled = True
    brwWebBrowser.Navigate cboAddress.Text
End Sub


Private Sub cboAddress_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = vbKeyReturn Then
        cboAddress_Click
    End If
End Sub

Private Sub Forward_Click()
brwWebBrowser.GoForward
End Sub

Private Sub Home_Click()
brwWebBrowser.Navigate StartingAddress
End Sub

Private Sub how_Click()
MsgBox ("WebExpress is very easy to use.  Use the tool bar to go to your homepage, go to a search page, go to the page you were at last, go forward to a page, stop, and look in your favorites.  Or use the menu bar where you can also save, open, and print pages.  The favorites may be a bit strange.  Just Copy-Paste address back and forth to get to them.  *hint* Use the right mouse button a lot!!!!")
End Sub

Private Sub Image1_DblClick()
MsgBox ("Stop touching me!!!!")
End Sub


Private Sub Image10_Click()
brwWebBrowser.GoForward
End Sub

Private Sub Image11_Click()
brwWebBrowser.Navigate StartingAddress
End Sub

Private Sub Image12_Click()
brwWebBrowser.Stop
Me.Caption = brwWebBrowser.LocationName
End Sub

Private Sub Image9_Click()
 brwWebBrowser.GoBack
End Sub

Private Sub Normal_Click()
fullBrowser.Show
Me.Visible = False
End Sub

Private Sub onlinehelp_Click()
MsgBox ("none yet")
End Sub

Private Sub open_Click()
CommonDialog1.ShowOpen
cboAddress.Text = "file://" & CommonDialog1.FileName
brwWebBrowser.Navigate cboAddress.Text
End Sub

Private Sub print_Click()
MsgBox ("Im sorry for the trouble, but i have not figured out how to print the web page from here, so right click on the page and choose print.  Its easy to do.")
End Sub

Private Sub refresh_Click()
brwWebBrowser.refresh
End Sub

Private Sub save_Click()
CommonDialog1.ShowSave
End Sub

Private Sub setupp_Click()
On Error Resume Next
    CommonDialog1.Flags = &H40
    CommonDialog1.ShowPrinter
End Sub

Private Sub search_Click()
brwWebBrowser.GoSearch
End Sub

Private Sub timTimer_Timer()
    If brwWebBrowser.Busy = False Then

        Me.Caption = brwWebBrowser.LocationName & " - WebExpress"
        StatusBar1.SimpleText = brwWebBrowser.LocationName
    Else
        Me.Caption = "Finding Page: " & cboAddress.Text & "......"
        StatusBar1.SimpleText = cboAddress.Text & "......"
        Image1.Picture = Image2.Picture
        Image2.Picture = Image3.Picture
        Image3.Picture = Image4.Picture
        Image4.Picture = Image5.Picture
        Image5.Picture = Image6.Picture
        Image6.Picture = Image7.Picture
        Image7.Picture = Image8.Picture
        Image8.Picture = Image1.Picture
    End If
End Sub


Private Sub tbToolBar_ButtonClick(ByVal Button As Button)
    On Error Resume Next
      

    timTimer.Enabled = True
      

    Select Case Button.Key
        Case "Back"
            brwWebBrowser.GoBack
        Case "Forward"
            brwWebBrowser.GoForward
        Case "Refresh"
            brwWebBrowser.refresh
        Case "Home"
            brwWebBrowser.Navigate StartingAddress
        Case "Search"
            brwWebBrowser.GoSearch
        Case "Favorites"
            Favs.Show
        Case "Stop"
            brwWebBrowser.Stop
            Me.Caption = brwWebBrowser.LocationName
        Case "Screen"
            fullBrowser.Show
            Me.Visible = False
    End Select


End Sub
