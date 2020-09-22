VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4965
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6300
   LinkTopic       =   "Form1"
   ScaleHeight     =   4965
   ScaleWidth      =   6300
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H8000000D&
      Caption         =   "Created By El Jefe -- Close"
      Height          =   285
      Left            =   45
      Picture         =   "frmAbout.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4575
      Width           =   6195
   End
   Begin VB.Image Image8 
      Height          =   4440
      Left            =   0
      Picture         =   "frmAbout.frx":13D2
      Top             =   0
      Width           =   6300
   End
   Begin VB.Image Image1 
      Height          =   2250
      Left            =   -6735
      Picture         =   "frmAbout.frx":18197
      Top             =   3255
      Width           =   19200
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub
