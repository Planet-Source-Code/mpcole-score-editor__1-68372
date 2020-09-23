VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Score Editor"
   ClientHeight    =   3510
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   10065
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   10065
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.TextBox txtComposer 
      Height          =   285
      Left            =   6480
      TabIndex        =   13
      Top             =   720
      Width           =   3495
   End
   Begin VB.TextBox txtTitle 
      Height          =   285
      Left            =   960
      TabIndex        =   11
      Top             =   720
      Width           =   3495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Neu"
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   615
   End
   Begin MSComDlg.CommonDialog Cmd 
      Left            =   9000
      Top             =   2280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Verbinden"
      Height          =   255
      Left            =   5520
      TabIndex        =   8
      Top             =   360
      Width           =   1095
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Trenn en"
      Height          =   255
      Left            =   5520
      TabIndex        =   9
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Bild speichern"
      Height          =   375
      Left            =   2760
      TabIndex        =   7
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Akkord"
      Height          =   375
      Left            =   4200
      TabIndex        =   6
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Stop"
      Height          =   375
      Left            =   8400
      TabIndex        =   5
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Abspielen"
      Height          =   375
      Left            =   6840
      TabIndex        =   4
      Top             =   120
      Width           =   1455
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      LargeChange     =   1000
      Left            =   120
      Max             =   20000
      SmallChange     =   75
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   3120
      Width           =   9855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Laden"
      Height          =   375
      Left            =   840
      TabIndex        =   2
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Speichern"
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin Projekt1.ctlScoreEdit ScoreEditor1 
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   9855
      _ExtentX        =   10610
      _ExtentY        =   10398
   End
   Begin VB.Label Label2 
      Caption         =   "Komponist:"
      Height          =   255
      Left            =   5400
      TabIndex        =   14
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Titel:"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   720
      Width           =   735
   End
   Begin VB.Menu mnuPitch 
      Caption         =   "Lage"
      Begin VB.Menu o 
         Caption         =   "Oktave tiefer"
         Shortcut        =   {F2}
      End
      Begin VB.Menu m 
         Caption         =   "Mittel"
         Shortcut        =   {F3}
      End
      Begin VB.Menu h 
         Caption         =   "Oktave höher"
         Shortcut        =   {F4}
      End
      Begin VB.Menu n 
         Caption         =   "Noch höher"
         Shortcut        =   {F5}
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Call ScoreEditor1.NewSheet
End Sub

Private Sub Command10_Click()
Call ScoreEditor1.JoinSelection
End Sub

Private Sub Command11_Click()
Call ScoreEditor1.SplitSelection
End Sub

Private Sub Command2_Click()
On Error GoTo Error
Cmd.Filter = "Binary File(*.bin)|*.bin"
Cmd.ShowSave

If Cmd.Filename <> "" Then Call ScoreEditor1.Save(Cmd.Filename)
Error:
End Sub

Private Sub Command3_Click()
On Error GoTo Error
Cmd.Filter = "Binary File(*.bin)|*.bin"
Cmd.ShowOpen

If Cmd.Filename <> "" Then Call ScoreEditor1.Load(Cmd.Filename)
txtTitle.Text = ScoreEditor1.Title
txtComposer.Text = ScoreEditor1.Composer
Error:
End Sub



Private Sub Command4_Click()
ScoreEditor1.ChangeSignature 2
End Sub

Private Sub Command6_Click()
Call ScoreEditor1.Play(120)
End Sub

Private Sub Command7_Click()
Call ScoreEditor1.StopPlay
End Sub

Private Sub Command8_Click()
Dim Chord As String
Chord = InputBox("Akkord eingeben:")
Call ScoreEditor1.AddChord(Chord)
End Sub

Private Sub Command9_Click()
'Call ScoreEditor1.SavePictureFile("D:\Eigene Dateien\Test.bmp", 2)
On Error GoTo Error
Cmd.Filter = "Bitmap(*.bmp)|*.bmp"
Cmd.ShowSave

If Cmd.Filename <> "" Then Call ScoreEditor1.SavePictureFile(Cmd.Filename, 3)
Error:
End Sub

Private Sub h_Click()
Call ScoreEditor1.ChangePitch(1)
End Sub

Private Sub HScroll1_Change()
Call ScoreEditor1.ScrollTo(HScroll1.Value)
End Sub

Private Sub HScroll1_Scroll()
Call ScoreEditor1.ScrollTo(HScroll1.Value)
End Sub

Private Sub m_Click()
Call ScoreEditor1.ChangePitch(0)
End Sub

Private Sub n_Click()
Call ScoreEditor1.ChangePitch(2)
End Sub

Private Sub o_Click()
Call ScoreEditor1.ChangePitch(-1)
End Sub

Private Sub ScoreEditor1_ScrollChanged(Value As Long)
HScroll1.Value = HScroll1.Value + Value
End Sub

Private Sub txtComposer_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then ScoreEditor1.Composer = txtComposer.Text
End Sub

Private Sub txtTitle_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then ScoreEditor1.Title = txtTitle.Text
End Sub
