VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMessage 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Encrypted Message"
   ClientHeight    =   1875
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8085
   Icon            =   "frmMessage.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1875
   ScaleWidth      =   8085
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cdSave 
      Left            =   6105
      Top             =   1305
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Save Encrypted File"
      Filter          =   "Encrypted File (*.enc)|*.enc"
   End
   Begin VB.TextBox txtPadName 
      Height          =   330
      Left            =   30
      TabIndex        =   2
      Top             =   1275
      Width           =   3825
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save to Disk"
      Height          =   600
      Left            =   6015
      TabIndex        =   1
      Top             =   1245
      Width           =   2055
   End
   Begin VB.TextBox txtMessage 
      Appearance      =   0  'Flat
      BackColor       =   &H00000080&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0FF&
      Height          =   1200
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   8085
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   600
      Left            =   3975
      TabIndex        =   3
      Top             =   1245
      Width           =   2055
   End
   Begin VB.Label lblPadName 
      Caption         =   "Encryption Pad Name"
      Height          =   255
      Left            =   165
      TabIndex        =   4
      Top             =   1635
      Width           =   1635
   End
End
Attribute VB_Name = "frmMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdSave_Click()
Dim i As Integer
Dim strName As String

LoadedMessage(1) = txtPadName.Text
LoadedMessage(2) = txtMessage.Text

cdSave.ShowSave
If cdSave.FileName = "" Then
    Exit Sub
End If
Open cdSave.FileName For Output As #1
For i = 1 To 2
    Write #1, LoadedMessage(i)
Next i
Close #1
End Sub
