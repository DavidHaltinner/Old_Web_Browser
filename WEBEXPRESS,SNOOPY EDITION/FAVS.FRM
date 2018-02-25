VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form Favs 
   Caption         =   "Favorites"
   ClientHeight    =   5770
   ClientLeft      =   40
   ClientTop       =   270
   ClientWidth     =   6270
   Icon            =   "favs.frx":0000
   LinkTopic       =   "Form2"
   MinButton       =   0   'False
   ScaleHeight     =   5770
   ScaleWidth      =   6270
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Add Current Page"
      Height          =   370
      Left            =   2880
      TabIndex        =   4
      Top             =   840
      Width           =   2050
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Go To Selection"
      Height          =   370
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   1690
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Program Files\WebExpress\WebExpress.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   492
      Left            =   4560
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Favorites"
      Top             =   4080
      Visible         =   0   'False
      Width           =   1030
   End
   Begin MSDBGrid.DBGrid DBGrid2 
      Bindings        =   "favs.frx":0442
      Height          =   4330
      Left            =   0
      OleObjectBlob   =   "favs.frx":0452
      TabIndex        =   0
      Top             =   1440
      Width           =   6240
   End
   Begin VB.Line Line1 
      X1              =   2520
      X2              =   2520
      Y1              =   0
      Y2              =   1440
   End
   Begin VB.Label Label2 
      Caption         =   "To add the current page, just hit the button next to the next open slot, and then hit the ""Add Current Page"" button"
      Height          =   730
      Left            =   2880
      TabIndex        =   3
      Top             =   120
      Width           =   3010
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Click on the button next to your choice, and then click the button to go there."
      Height          =   610
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2050
   End
End
Attribute VB_Name = "Favs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If frmBrowser.Visible = True Then
    frmBrowser.cboAddress.Text = DBGrid2.SelText
    frmBrowser.brwWebBrowser.Navigate (frmBrowser.cboAddress.Text)
Else
    fullBrowser.cboAddress.Text = DBGrid2.SelText
    fullBrowser.brwWebBrowser.Navigate (frmBrowser.cboAddress.Text)
End If
Unload Me
End Sub

Private Sub Command2_Click()
If frmBrowser.Visible = True Then
    DBGrid2.SelText = frmBrowser.cboAddress.Text
Else
    DBGrid2.SelText = fullBrowser.cboAddress.Text
End If
End Sub

Private Sub Form_Load()
DBGrid2.Height = Me.Height - 900
DBGrid2.Width = Me.Width
End Sub

Private Sub Form_LostFocus()
Unload Me
End Sub

Private Sub Form_Resize()
DBGrid2.Height = Me.Height - 900
DBGrid2.Width = Me.Width
End Sub
