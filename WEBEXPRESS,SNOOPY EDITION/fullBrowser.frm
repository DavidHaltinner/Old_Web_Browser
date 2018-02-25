VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form fullBrowser 
   BorderStyle     =   0  'None
   ClientHeight    =   7365
   ClientLeft      =   345
   ClientTop       =   -345
   ClientWidth     =   15360
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7365
   ScaleWidth      =   15360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.ComboBox cboAddress 
      Height          =   315
      Left            =   3120
      TabIndex        =   2
      Text            =   "Type URL Here"
      Top             =   30
      Width           =   7050
   End
   Begin VB.Timer timTimer 
      Enabled         =   0   'False
      Interval        =   120
      Left            =   6180
      Top             =   1500
   End
   Begin VB.PictureBox picAddress 
      BorderStyle     =   0  'None
      Height          =   370
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   4995
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   960
      Visible         =   0   'False
      Width           =   5000
   End
   Begin SHDocVwCtl.WebBrowser brwWebBrowser 
      Height          =   8895
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   9930
      ExtentX         =   17515
      ExtentY         =   15690
      ViewMode        =   1
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   -1  'True
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.Image Favb 
      Height          =   435
      Left            =   1920
      Picture         =   "fullBrowser.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   435
   End
   Begin VB.Image Back 
      Height          =   435
      Left            =   120
      Picture         =   "fullBrowser.frx":0442
      Stretch         =   -1  'True
      Top             =   0
      Width           =   435
   End
   Begin VB.Image Forward 
      Height          =   435
      Left            =   480
      Picture         =   "fullBrowser.frx":0884
      Stretch         =   -1  'True
      Top             =   0
      Width           =   435
   End
   Begin VB.Image Home 
      Height          =   435
      Left            =   1440
      Picture         =   "fullBrowser.frx":0CC6
      Stretch         =   -1  'True
      Top             =   0
      Width           =   435
   End
   Begin VB.Image Stop 
      Height          =   435
      Left            =   960
      Picture         =   "fullBrowser.frx":1108
      Stretch         =   -1  'True
      Top             =   0
      Width           =   435
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   14880
      TabIndex        =   3
      Top             =   0
      Width           =   255
   End
   Begin VB.Image Image9 
      Height          =   480
      Left            =   120
      Picture         =   "fullBrowser.frx":154A
      Top             =   3720
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image8 
      Height          =   480
      Left            =   120
      Picture         =   "fullBrowser.frx":1854
      Top             =   3240
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image7 
      Height          =   480
      Left            =   0
      Picture         =   "fullBrowser.frx":1B5E
      Top             =   2880
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image6 
      Height          =   480
      Left            =   0
      Picture         =   "fullBrowser.frx":1E68
      Top             =   2520
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image5 
      Height          =   480
      Left            =   0
      Picture         =   "fullBrowser.frx":2172
      Top             =   2160
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image4 
      Height          =   480
      Left            =   120
      Picture         =   "fullBrowser.frx":247C
      Top             =   1800
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   120
      Picture         =   "fullBrowser.frx":2786
      Top             =   1440
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   14160
      Picture         =   "fullBrowser.frx":2A90
      ToolTipText     =   "Ahh!!!!!! Get that off of me!!!!! I hate mouse pointers!!!!!"
      Top             =   0
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   132
      Left            =   5280
      Top             =   1200
      Width           =   372
   End
   Begin VB.Image Normal 
      Height          =   480
      Left            =   2520
      Picture         =   "fullBrowser.frx":2D9A
      Top             =   0
      Width           =   480
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   525
      Left            =   0
      Top             =   0
      Width           =   15300
   End
   Begin ComctlLib.ImageList imlIcons 
      Left            =   5880
      Top             =   960
      _ExtentX        =   688
      _ExtentY        =   688
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   10
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "fullBrowser.frx":31DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "fullBrowser.frx":38EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "fullBrowser.frx":4000
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "fullBrowser.frx":4712
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "fullBrowser.frx":4E24
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "fullBrowser.frx":5536
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "fullBrowser.frx":5C48
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "fullBrowser.frx":5F62
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "fullBrowser.frx":627C
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "fullBrowser.frx":6596
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "fullBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public StartingAddress As String
Dim mbDontNavigateNow As Boolean

Private Sub Back_Click()
brwWebBrowser.GoBack
End Sub

Private Sub Favb_Click()
Favs.Show
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Me.Show
    fullBrowser.Visible = True
    cboAddress.Visible = True
    brwWebBrowser.Height = Me.Height
    brwWebBrowser.Width = Me.Width
    Shape1.Width = Me.Width
    cboAddress.Text = frmBrowser.cboAddress.Text
    brwWebBrowser.Navigate (cboAddress.Text)
    brwWebBrowser.FullScreen = True
    brwWebBrowser.TheaterMode = True
    'Image1.Left = Me.Width - 700
    Label1.Left = Me.Width - 400

    If Len(StartingAddress) > 0 Then
        cboAddress.Text = StartingAddress
        cboAddress.AddItem cboAddress.Text
        timTimer.Enabled = True
        brwWebBrowser.Navigate StartingAddress
    End If


End Sub





Private Sub brwWebBrowser_DownloadComplete()
    On Error Resume Next
    Me.Caption = brwWebBrowser.LocationName
    cboAddress.Text = brwWebBrowser.LocationURL
End Sub




Private Sub brwWebBrowser_NavigateComplete(ByVal URL As String)
    Dim i As Integer
    Dim bFound As Boolean
    Me.Caption = brwWebBrowser.LocationName
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


Private Sub Form_Resize()
    brwWebBrowser.Width = Me.ScaleWidth - 100
    brwWebBrowser.Height = Me.ScaleHeight - (picAddress.Top + picAddress.Height) - 100
End Sub


Private Sub Forward_Click()
brwWebBrowser.GoForward
End Sub

Private Sub Home_Click()
brwWebBrowser.GoHome
End Sub

Private Sub Label1_Click()
End
End Sub

Private Sub Normal_Click()
frmBrowser.Visible = True
            frmBrowser.cboAddress.Text = fullBrowser.cboAddress.Text
            frmBrowser.brwWebBrowser.Navigate (fullBrowser.cboAddress.Text)
            Unload Me
End Sub

Private Sub Stop_Click()
timTimer.Enabled = False
            brwWebBrowser.Stop
            Me.Caption = brwWebBrowser.LocationName
End Sub

Private Sub timTimer_Timer()
    If brwWebBrowser.Busy = False Then
        Me.Caption = brwWebBrowser.LocationName & " - WebExpress"
    Else
        Me.Caption = "Finding Page: " & cboAddress.Text & "......"
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

