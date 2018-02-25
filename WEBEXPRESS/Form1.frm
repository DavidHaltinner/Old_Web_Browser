VERSION 5.00
Begin VB.Form frmabout 
   BackColor       =   &H00FFFFFF&
   Caption         =   "About WebExpress"
   ClientHeight    =   2544
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   6432
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2544
   ScaleWidth      =   6432
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   370
      Left            =   4800
      TabIndex        =   3
      Top             =   1560
      Width           =   1330
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Build 1.0.0"
      Height          =   252
      Left            =   2400
      TabIndex        =   4
      Top             =   1560
      Width           =   1812
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Made by David Urban Haltinner"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   370
      Left            =   2280
      TabIndex        =   2
      Top             =   1080
      Width           =   4570
   End
   Begin VB.Label lblProductName 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "WebExpress"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   32.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   770
      Left            =   2280
      TabIndex        =   0
      Top             =   240
      Width           =   3890
   End
   Begin VB.Label lblCompanyProduct 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Version 1.0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   430
      Left            =   2280
      TabIndex        =   1
      Top             =   480
      Visible         =   0   'False
      Width           =   1890
   End
   Begin VB.Image imgLogo 
      Height          =   2110
      Left            =   0
      Picture         =   "Form1.frx":030A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2180
   End
End
Attribute VB_Name = "frmabout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
imgLogo.Height = lblCompanyProduct.Width
imgLogo.Width = lblCompanyProduct.Width
End Sub
