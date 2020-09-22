VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmPad 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "One-Time Pad"
   ClientHeight    =   8115
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8085
   Icon            =   "frmPad.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8115
   ScaleWidth      =   8085
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cdLoad2 
      Left            =   7170
      Top             =   2955
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Load File to Decrypt"
      Filter          =   "Encrypted File (*.enc)|*.enc|All Files (*.*)|*.*"
   End
   Begin MSComDlg.CommonDialog cdSave 
      Left            =   4335
      Top             =   4815
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Save Encryption Pad"
      Filter          =   "Encryption Pad (*.pad)|*.pad"
   End
   Begin MSComDlg.CommonDialog cdLoad 
      Left            =   7185
      Top             =   2115
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Load Encryption Pad"
      Filter          =   "Encryption Pad (*.pad)|*.pad|All Files (*.*)|*.*"
   End
   Begin VB.CommandButton cmdDecrypt 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Decrypt Message"
      Height          =   300
      Left            =   4140
      MaskColor       =   &H00C0FFC0&
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   0
      Width           =   2085
   End
   Begin VB.CommandButton cmdEncrypt 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Encrypt Message"
      Height          =   300
      Left            =   2070
      MaskColor       =   &H00C0C0FF&
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   0
      Width           =   2085
   End
   Begin VB.Frame Frame3 
      Caption         =   "Options"
      Height          =   3075
      Left            =   5475
      TabIndex        =   4
      Top             =   1635
      Width           =   2550
      Begin VB.CommandButton cmdAbout 
         Caption         =   "About"
         Height          =   915
         Left            =   90
         Picture         =   "frmPad.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   2055
         Width           =   2370
      End
      Begin VB.CommandButton cmdLoadDecrypt 
         Caption         =   "Load File For Decryption"
         Height          =   915
         Left            =   90
         Picture         =   "frmPad.frx":1194
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1155
         Width           =   2370
      End
      Begin VB.CommandButton cmdLoad 
         Caption         =   "Load Pad For Decryption"
         Height          =   915
         Left            =   90
         Picture         =   "frmPad.frx":1A5E
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   255
         Width           =   2370
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Generated Pad"
      Height          =   3705
      Left            =   0
      TabIndex        =   2
      Top             =   1620
      Width           =   5310
      Begin VB.CommandButton cmdSavePad 
         Caption         =   "Save Pad"
         Height          =   255
         Left            =   150
         TabIndex        =   11
         Top             =   3360
         Width           =   4995
      End
      Begin VB.TextBox txtPad 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   5.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   2535
         Left            =   630
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   3
         Top             =   615
         Width           =   3960
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   2805
         Left            =   525
         Top             =   495
         Width           =   4170
      End
   End
   Begin VB.TextBox txtMessage 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   1230
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   345
      Width           =   8085
   End
   Begin VB.CommandButton cmdGenerate 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Generate Pad"
      Height          =   300
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Width           =   2085
   End
   Begin VB.Label Label1 
      Caption         =   "theguys@tyler.net"
      ForeColor       =   &H00C00000&
      Height          =   225
      Left            =   6780
      TabIndex        =   10
      Top             =   15
      Width           =   1320
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000F&
      BackStyle       =   1  'Opaque
      Height          =   4260
      Left            =   -120
      Top             =   1200
      Width           =   8280
   End
   Begin VB.Image Image3 
      Appearance      =   0  'Flat
      Height          =   6750
      Left            =   -3690
      Picture         =   "frmPad.frx":2328
      Top             =   1545
      Width           =   12000
   End
End
Attribute VB_Name = "frmPad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAbout_Click()
    frmAbout.Show
End Sub

Private Sub cmdGenerate_Click()
    txtPad.Text = ""
    txtPad.Text = GeneratePad()
End Sub

Private Sub cmdEncrypt_Click()
        ' Limit text to characters that I have set to be translated
    For X = 1 To Len(txtMessage.Text)
        If InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789 ,." & vbCrLf, Mid(txtMessage.Text, X, 1)) < 1 Then
            MsgBox "Please limit your text to Letters A-Z, Numbers 0-9, Spaces, Commas, And Periods.", vbOKOnly, "Encryption Error"
            Exit Sub
        End If
    Next
    
    If txtPad.Text = "" Then
        Call cmdGenerate_Click
    End If
    
    frmMessage.txtMessage.Text = Encrypt(frmPad.txtMessage.Text)
    frmMessage.Show
End Sub

Private Sub cmdDecrypt_Click()
    frmMessage.txtMessage.Text = Decrypt(frmPad.txtMessage.Text)
    frmMessage.Show
End Sub

Private Sub cmdLoadDecrypt_Click()
Dim fnum As Integer
Dim num_lines As Long
Dim i As Long
    cdLoad2.ShowOpen
    If cdLoad2.FileName = "" Then
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    DoEvents

    fnum = FreeFile
    On Error GoTo errorhandler
    Open cdLoad2.FileName For Input As fnum
    For i = 1 To 2
        Input #fnum, LoadedMessage(i)
    Next i
    Close #fnum
    
    Screen.MousePointer = vbDefault
    
txtMessage.Text = LoadedMessage(2)
MsgBox "The Pad File is Titled: " & LoadedMessage(1), vbOKOnly, "Pad File"
frmMessage.txtPadName.Text = LoadedMessage(1)
Exit Sub
errorhandler:
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdSavePad_Click()
Dim i As Integer
Dim strName As String
cdSave.ShowSave
If cdSave.FileName = "" Then
    Exit Sub
End If
Open cdSave.FileName For Output As #1
For i = 1 To 20
    Write #1, SplicedPad(1, i), SplicedPad(2, i), SplicedPad(3, i), SplicedPad(4, 1), SplicedPad(5, i), SplicedPad(6, i), SplicedPad(7, i), SplicedPad(8, i), SplicedPad(9, i)
Next i
Close #1
frmMessage.txtPadName.Text = cdSave.FileTitle
End Sub

Private Sub cmdLoad_Click()
    cdLoad.ShowOpen
    If cdLoad.FileName = "" Then
        Exit Sub
    End If
    LoadPad cdLoad.FileName
    txtPad.Text = DisplaySplicedPad
End Sub

Private Sub Form_Load()
    Randomize
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

