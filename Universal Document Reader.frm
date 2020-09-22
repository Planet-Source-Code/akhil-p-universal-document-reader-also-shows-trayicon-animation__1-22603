VERSION 5.00
Object = "{EEE78583-FE22-11D0-8BEF-0060081841DE}#1.0#0"; "XVoice.dll"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.PictureBox Picture3 
      Height          =   540
      Left            =   2625
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   4
      Top             =   1500
      Width           =   540
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   3435
      Top             =   2130
   End
   Begin VB.PictureBox Picture1 
      Height          =   540
      Left            =   1890
      Picture         =   "Universal Document Reader.frx":0000
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   3
      Top             =   1485
      Width           =   540
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   2145
      OLEDropMode     =   2  'Automatic
      Picture         =   "Universal Document Reader.frx":030A
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   2
      Top             =   2460
      Width           =   540
   End
   Begin VB.TextBox Text1 
      Height          =   1095
      Left            =   90
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   90
      Width           =   3930
   End
   Begin ACTIVEVOICEPROJECTLibCtl.DirectSS DirectSS1 
      Height          =   405
      Left            =   465
      OleObjectBlob   =   "Universal Document Reader.frx":0614
      TabIndex        =   0
      Top             =   2220
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuFileRead 
         Caption         =   "Start &Reading"
      End
      Begin VB.Menu mnuFileStop 
         Caption         =   "&Stop Reading"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Akhil As CWinAPI

Private Sub Form_Load()
Set Akhil = New CWinAPI
Akhil.TrayIcon shelliconAdd, Picture2.hwnd, "Akhil's Agent", Picture2.Picture

End Sub

Private Sub Form_Unload(Cancel As Integer)
Akhil.TrayIcon shelliconDelete, Picture2.hwnd, "Akhil's Agent", Picture2.Picture
End Sub

Private Sub mnuFileExit_Click()
Akhil.TrayIcon shelliconDelete, Picture2.hwnd, "Akhil's Agent", Picture2.Picture
End
End Sub

Private Sub mnuFileRead_Click()
Text1.Text = Clipboard.GetText
    DirectSS1.Speak Text1.Text
End Sub

Private Sub mnuFileStop_Click()
DirectSS1.Speak ""
End Sub

Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim lng As Long
    lng = Akhil.TryIcnRtnMsg(X)
If lng = xyzLeftButtonDown Then
    'DirectSS1.Speak "I am going to read E-Mail"
    Text1.Text = Clipboard.GetText
    DirectSS1.Speak Text1.Text
ElseIf lng = xyzRightButtonDown Then
    PopupMenu mnuFile
End If

End Sub

Private Sub Timer1_Timer()
If Picture2.Picture <> Picture1.Picture Then
    Picture3.Picture = Picture2.Picture
    Picture2.Picture = Picture1.Picture
Else
    Picture2.Picture = Picture3.Picture
End If

Akhil.TrayIcon shelliconModify, Picture2.hwnd, "Akhil's Agent", Picture2.Picture
End Sub
