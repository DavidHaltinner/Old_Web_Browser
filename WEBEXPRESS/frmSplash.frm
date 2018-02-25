VERSION 5.00
Begin VB.Form frmSplash 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   2076
   ClientLeft      =   228
   ClientTop       =   1380
   ClientWidth     =   2076
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2076
   ScaleWidth      =   2076
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   5880
      Top             =   1200
   End
   Begin VB.Image imgLogo 
      Height          =   1890
      Left            =   120
      Picture         =   "frmSplash.frx":000C
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1890
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit



Private Sub Timer1_Timer()
frmBrowser.Show
Unload Me
End Sub
