VERSION 5.00
Begin VB.UserControl ctlScoreEdit 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fest Einfach
   ClientHeight    =   7710
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11880
   BeginProperty Font 
      Name            =   "capella3-invertiert"
      Size            =   20.25
      Charset         =   2
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   514
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   792
   Begin VB.PictureBox picBuffer 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'Kein
      BeginProperty Font 
         Name            =   "capella3"
         Size            =   20.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   13500
      Left            =   1560
      ScaleHeight     =   900
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   620
      TabIndex        =   0
      Top             =   6240
      Visible         =   0   'False
      Width           =   9300
   End
   Begin VB.Timer tmrScroll 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   5280
      Top             =   4320
   End
   Begin VB.Timer tmrMetro 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   4440
      Top             =   4200
   End
   Begin VB.Timer tmrPlay 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3960
      Top             =   4200
   End
   Begin VB.Image imgBow2 
      Height          =   60
      Left            =   5880
      Picture         =   "ScoreEditor.ctx":0000
      Top             =   2880
      Width           =   600
   End
   Begin VB.Image imgBow 
      Height          =   60
      Left            =   5880
      Picture         =   "ScoreEditor.ctx":0222
      Top             =   2400
      Width           =   600
   End
End
Attribute VB_Name = "ctlScoreEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim Space As Single

Dim Chords(21) As String
Dim Tones(36) As String

Private Type Note
    NotePos As Integer
    Note As Integer
    Length As Integer
    Up As Boolean
    Down As Boolean
    No As Boolean
    Point As Boolean
    Join As Boolean
    Text As String
    Bar As Integer
    Triple As Boolean
    Bow As Boolean
    Chord As String
    Splitter As Boolean
End Type

Dim MaxLength As Long

Dim Notes As New Collection
Dim NoteCollection() As Note

Dim NoteLength As Integer
Dim Up As Boolean
Dim Down As Boolean
Dim No As Boolean
Dim Point As Boolean
Dim Join As Boolean
Dim TripleNote As Boolean
Dim Bow As Boolean
Dim Pitch As Integer

Dim CursorX As Long
Dim SelBoundLeft As Long
Dim SelBoundRight As Long

Dim TimeSignature As Single
Dim Signature As Integer
Dim Scroll As Long
Dim Additional As Long

Dim Tone As Integer
Dim PlayTempo As Integer
Dim NoteCounter As Long
Dim PlayCounter1 As Single
Dim PlayCounter2 As Long
Dim AktivFis As String
Dim AktivB As String
Dim CurrentChord As String

Dim mTitle As String * 60
Dim mComposer As String * 60

Private Declare Function CreateCaret Lib "user32" (ByVal hwnd As Long, ByVal hBitmap As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function ShowCaret Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetCaretPos Lib "user32" (ByVal x As Long, ByVal Y As Long) As Long
Private Declare Function HideCaret Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function LockWindowUpdate Lib "user32.dll" (ByVal hwndLock As Long) As Long

Public Event ScrollChanged(Value As Long)

Public Property Let Title(ByVal Data As String)
mTitle = Data
End Property

Public Property Get Title() As String
Title = TrimString(mTitle)
End Property

Public Property Let Composer(ByVal Data As String)
mComposer = Data
End Property

Public Property Get Composer() As String
Composer = TrimString(mComposer)
End Property

Public Sub Save(Filename As String)
Dim i As Long
ReDim NoteCollection(1 To Notes.Count)
For i = 1 To UBound(NoteCollection)
    NoteCollection(i).NotePos = Notes(i).NotePos
    NoteCollection(i).Note = Notes(i).Note
    NoteCollection(i).Length = Notes(i).Length
    NoteCollection(i).Up = Notes(i).Up
    NoteCollection(i).Down = Notes(i).Down
    NoteCollection(i).No = Notes(i).No
    NoteCollection(i).Point = Notes(i).Point
    NoteCollection(i).Join = Notes(i).Join
    NoteCollection(i).Text = Notes(i).Text
    NoteCollection(i).Bar = Notes(i).Bar
    NoteCollection(i).Triple = Notes(i).Triple
    NoteCollection(i).Bow = Notes(i).Bow
    NoteCollection(i).Chord = Notes(i).Chord
    NoteCollection(i).Splitter = Notes(i).Splitter
Next i
Open Filename For Binary Access Write As #1
    Put #1, , TimeSignature
    Put #1, , Signature
    Put #1, , UBound(NoteCollection)
    Put #1, , NoteCollection
    Put #1, , mTitle
    Put #1, , mComposer
Close #1
End Sub

Public Sub Load(Filename As String)
Dim i As Long
Dim Count As Long

Set Notes = New Collection
Open Filename For Binary Access Read As #1
        'Get #1, , mTitle
        Get #1, , TimeSignature
        Get #1, , Signature
        Get #1, , Count
        ReDim NoteCollection(1 To Count)
        Get #1, , NoteCollection
        Get #1, , mTitle
        Get #1, , mComposer
Close #1
mTitle = TrimString(mTitle)
mComposer = TrimString(mComposer)

For i = 1 To UBound(NoteCollection)
    Dim Note As New clsNotes
    Note.NotePos = NoteCollection(i).NotePos
    Note.Note = NoteCollection(i).Note
    Note.Length = NoteCollection(i).Length
    Note.Up = NoteCollection(i).Up
    Note.Down = NoteCollection(i).Down
    Note.No = NoteCollection(i).No
    Note.Point = NoteCollection(i).Point
    Note.Join = NoteCollection(i).Join
    Note.Text = NoteCollection(i).Text
    Note.Bar = NoteCollection(i).Bar
    Note.Triple = NoteCollection(i).Triple
    Note.Bow = NoteCollection(i).Bow
    Note.Chord = NoteCollection(i).Chord
    Note.Splitter = NoteCollection(i).Splitter
    Notes.Add Note
    Set Note = Nothing
Next i
CursorX = 1
SelBoundLeft = 0
SelBoundRight = 0
Call DrawLines
End Sub

Public Sub NewSheet()
Set Notes = New Collection
TimeSignature = 1
Signature = 0
CursorX = 1
Call DrawLines
Call UserControl.SetFocus
End Sub

Public Sub SavePictureFile(Filename As String, BarsPerRow As Integer)
Dim i As Long
Dim x As Long
Dim Value As Long
Dim Counter1 As Single
Dim Counter2 As Long
Dim Last As Long
Dim Top As Long

Call LockWindowUpdate(UserControl.hwnd)
picBuffer.FontBold = True
picBuffer.Font.Size = 16
picBuffer.Font.Name = "Arial"
picBuffer.CurrentY = 5
picBuffer.CurrentX = (595 / 2) - (picBuffer.TextWidth(Title) / 2)
picBuffer.Print Title
picBuffer.FontBold = False
picBuffer.Font.Size = 10
picBuffer.CurrentY = 30
picBuffer.CurrentX = (595 / 2) - (picBuffer.TextWidth(Composer) / 2)
picBuffer.Print Composer

Last = 0
Counter1 = 0
Counter2 = 0
Top = 50
For i = 1 To Notes.Count
    Value = Notes(i).Length
    If Notes(i).Triple = True Then
        Counter1 = Counter1 + (2 * (1 / Value)) / 3
    Else
        Counter1 = Counter1 + (1 / Value)
    End If

    If Notes(i).Point = True Then Counter1 = Counter1 + ((1 / Value) / 2)

    If Counter1 > 0 And Counter1 = TimeSignature Then
        Counter1 = 0
        Counter2 = Counter2 + 1

        If Counter2 = BarsPerRow Then
            Space = (595 - 45 - Additional) / (i - Last)
            Exit For
        End If
    End If
Next i
Call DrawLines(Space * (Last))

Last = i

Call picBuffer.PaintPicture(UserControl.Image, 10, Top, , , , , 590)
picBuffer.Line (10, 50 + Top)-(picBuffer.Width, 50 + Top)
picBuffer.Line (10, 57 + Top)-(picBuffer.Width, 57 + Top)
picBuffer.Line (10, 64 + Top)-(picBuffer.Width, 64 + Top)
picBuffer.Line (10, 71 + Top)-(picBuffer.Width, 71 + Top)
picBuffer.Line (10, 78 + Top)-(picBuffer.Width, 78 + Top)
If Counter2 = BarsPerRow Then picBuffer.Line (picBuffer.Width - 1, 50 + Top)-(picBuffer.Width - 1, 78 + Top)
Top = Top + 130

Do Until Last >= Notes.Count
    Counter1 = 0
    Counter2 = 0
    For i = Last + 1 To Notes.Count
        Value = Notes(i).Length
        If Notes(i).Triple = True Then
            Counter1 = Counter1 + (2 * (1 / Value)) / 3
        Else
            Counter1 = Counter1 + (1 / Value)
        End If

        If Notes(i).Point = True Then Counter1 = Counter1 + ((1 / Value) / 2)

        If Counter1 > 0 And Counter1 = TimeSignature Then
            Counter1 = 0
            Counter2 = Counter2 + 1

            If Counter2 = BarsPerRow Then
                Space = (595) / (i - Last)
                Exit For
            End If
        End If
    Next i
   
    Call DrawLines(((Space) * (Last + 1)) + 16 + Additional)
    Last = i
    Call picBuffer.PaintPicture(UserControl.Image, 10, Top, , , , , 590)
    picBuffer.Line (0, 50 + Top)-(picBuffer.Width, 50 + Top)
    picBuffer.Line (0, 57 + Top)-(picBuffer.Width, 57 + Top)
    picBuffer.Line (0, 64 + Top)-(picBuffer.Width, 64 + Top)
    picBuffer.Line (0, 71 + Top)-(picBuffer.Width, 71 + Top)
    picBuffer.Line (0, 78 + Top)-(picBuffer.Width, 78 + Top)
    If Counter2 = BarsPerRow Then picBuffer.Line (picBuffer.Width - 1, 50 + Top)-(picBuffer.Width - 1, 78 + Top)
    
    Top = Top + 130
Loop
Call SavePicture(picBuffer.Image, Filename)
Space = 45
Call DrawLines
Call LockWindowUpdate(False)
End Sub

Private Sub tmrPlay_Timer()
Dim Beats As Single
Dim Add As Integer
Dim Key As Integer
Dim val2 As Long

On Error Resume Next

If Notes(NoteCounter).Bow = False Then
    Call StopNote(Tone)
End If
If NoteCounter >= Notes.Count Then tmrPlay.Enabled = False: tmrMetro.Enabled = False: tmrScroll.Enabled = False
NoteCounter = NoteCounter + 1
Beats = 60 / PlayTempo '120
Beats = 1000 * Beats

If Notes(NoteCounter).Length = 4 Then
    If Notes(NoteCounter).Point = True Then tmrPlay.Interval = Beats + (Beats / 2) Else tmrPlay.Interval = Beats
ElseIf Notes(NoteCounter).Length = 2 Then
   If Notes(NoteCounter).Point = True Then tmrPlay.Interval = (Beats * 2) + (Beats) Else tmrPlay.Interval = Beats * 2
ElseIf Notes(NoteCounter).Length = 1 Then
   If Notes(NoteCounter).Point = True Then tmrPlay.Interval = (Beats * 4) + ((Beats * 4) / 2) Else tmrPlay.Interval = Beats * 4
ElseIf Notes(NoteCounter).Length = 8 Then
   If Notes(NoteCounter).Point = True Then tmrPlay.Interval = (Beats / 2) + ((Beats / 2) / 2) Else tmrPlay.Interval = Beats / 2
ElseIf Notes(NoteCounter).Length = 16 Then
   If Notes(NoteCounter).Point = True Then tmrPlay.Interval = (Beats / 4) + ((Beats / 4) / 2) Else tmrPlay.Interval = Beats / 4
End If
If Notes(NoteCounter).Triple = True Then tmrPlay.Interval = (tmrPlay.Interval / 3) * 2
    
If Notes(NoteCounter).Chord <> "" Then
    If NoteCounter <= Notes.Count Then
        CurrentChord = Notes(NoteCounter).Chord
        Call SetChord(CurrentChord)
    End If
End If
    
If InStr(1, AktivFis, Notes(NoteCounter).Note) > 0 Then
    If Notes(NoteCounter).Up = False Then Add = 1
    If Notes(NoteCounter).No = True Then
        AktivFis = Replace(AktivFis, Notes(NoteCounter).Note, "")
        Add = 0
    End If
ElseIf Notes(NoteCounter).Up = True Then
    Add = 1
End If

If InStr(1, AktivB, Notes(NoteCounter).Note) > 0 Then
    If Notes(NoteCounter).Down = False Then Add = -1
    If Notes(NoteCounter).No = True Then
        AktivB = Replace(AktivB, Notes(NoteCounter).Note, "")
        Add = 0
    End If
ElseIf Notes(NoteCounter).Down = True Then
    Add = -1
End If
If Notes(NoteCounter).Up = True Then AktivFis = AktivFis & " " & Notes(NoteCounter).Note
If Notes(NoteCounter).Down = True Then AktivB = AktivB & " " & Notes(NoteCounter).Note

Key = Notes(NoteCounter).Note
Select Case Signature
    Case 1
        If Key = -3 Then Add = 1
        If Key = 4 Then Add = 1
        If Key = 11 Then Add = 1
    Case 2
        If Key = -3 Then Add = 1
        If Key = 4 Then Add = 1
        If Key = 11 Then Add = 1
        
        If Key = -6 Then Add = 1
        If Key = 1 Then Add = 1
        If Key = 8 Then Add = 1
    Case 3
        If Key = -3 Then Add = 1
        If Key = 4 Then Add = 1
        If Key = 11 Then Add = 1
        
        If Key = -6 Then Add = 1
        If Key = 1 Then Add = 1
        If Key = 8 Then Add = 1
        
        If Key = -3 Then Add = 1
        If Key = 5 Then Add = 1
        If Key = 12 Then Add = 1
    Case 4
        If Key = -3 Then Add = 1
        If Key = 4 Then Add = 1
        If Key = 11 Then Add = 1
        
        If Key = -7 Then Add = 1
        If Key = 1 Then Add = 1
        If Key = 8 Then Add = 1
        
        If Key = -3 Then Add = 1
        If Key = 5 Then Add = 1
        If Key = 12 Then Add = 1
        
        If Key = -6 Then Add = 1
        If Key = 2 Then Add = 1
        If Key = 9 Then Add = 1
    Case 5
        If Key = -3 Then Add = 1
        If Key = 4 Then Add = 1
        If Key = 11 Then Add = 1
        
        If Key = -7 Then Add = 1
        If Key = 1 Then Add = 1
        If Key = 8 Then Add = 1
        
        If Key = -3 Then Add = 1
        If Key = 5 Then Add = 1
        If Key = 12 Then Add = 1
        
        If Key = -6 Then Add = 1
        If Key = 2 Then Add = 1
        If Key = 9 Then Add = 1
        
        If Key = -1 Then Add = 1
        If Key = 6 Then Add = 1
        If Key = 13 Then Add = 1
End Select

Select Case Notes(NoteCounter).Note
Case -6
     Tone = 12 + Add
Case -5
      Tone = 14 + Add
Case -4
 Tone = 16 + Add
Case -3
      Tone = 17 + Add
Case -2
   Tone = 19 + Add
Case -1
Tone = 21 + Add
    Case 0
Tone = 23 + Add
    Case 1
Tone = 24 + Add
    Case 2
Tone = 26 + Add
Case 3
Tone = 28 + Add
 Case 4
Tone = 29 + Add
   Case 5
Tone = 31 + Add
   Case 6
Tone = 33 + Add
  Case 7
Tone = 35 + Add
  Case 8
Tone = 36 + Add
  Case 9
Tone = 38 + Add
   Case 10
Tone = 40 + Add
   Case 11
Tone = 41 + Add
  Case 12
Tone = 43 + Add
   Case 13
Tone = 45 + Add
  Case 14
Tone = 47 + Add
    Case 15
Tone = 48 + Add
    End Select
Tone = Tone + 12

If NoteCounter > Notes.Count Then tmrPlay.Enabled = False: Exit Sub
If Notes(NoteCounter).Note <> 1000 Then
    If Notes(NoteCounter - 1).Bow = False Then
    Call StartNote(Tone, 127)
    End If
End If
CursorX = NoteCounter
Call SetCaretPos((Space * CursorX) + 15 - Scroll + Additional, 45)

val2 = Notes(NoteCounter).Length
If Notes(NoteCounter).Triple = True Then
    PlayCounter1 = PlayCounter1 + (2 * (1 / val2)) / 3
Else
    PlayCounter1 = PlayCounter1 + (1 / val2)
End If

If Notes(NoteCounter).Point = True Then PlayCounter1 = PlayCounter1 + ((1 / val2) / 2)

If PlayCounter1 > 0 And PlayCounter1 = TimeSignature Then
    PlayCounter1 = 0
    AktivFis = ""
    AktivB = ""
End If
End Sub

Private Sub tmrMetro_Timer()
Dim Beats As Single

On Error Resume Next
PlayCounter2 = PlayCounter2 + 1
If PlayCounter2 = 4 Then PlayCounter2 = 0
Beats = 60 / PlayTempo '120
Beats = 1000 * Beats
tmrMetro.Interval = Beats
Call StopMetro(9)
Call StartMetro(9)
If PlayCounter2 = 1 Then
    'Call StopDong(10)
    'Call StartDong(10)
End If
End Sub

Private Sub tmrScroll_Timer()
If (Space * CursorX) + 15 - Scroll > (UserControl.Width / Screen.TwipsPerPixelX) - 50 Then
    RaiseEvent ScrollChanged((UserControl.Width / Screen.TwipsPerPixelX) - 100)
End If
End Sub

Private Sub UserControl_DblClick()
Dim Text As String
Text = InputBox("Text eingeben (" & CursorX & "):", "Liedtext eingeben", Notes(CursorX).Text)
If Text <> "" Then Notes(CursorX).Text = Text
Call CreateCaret(UserControl.hwnd, 0, 2, 40)
Call DrawLines(Scroll)
Call ShowCaret(UserControl.hwnd)
Call UserControl.SetFocus
End Sub

Public Sub AddChord(Chord As String)
Notes(CursorX).Chord = Chord
Call CreateCaret(UserControl.hwnd, 0, 2, 40)
Call DrawLines(Scroll)
Call ShowCaret(UserControl.hwnd)
Call UserControl.SetFocus
End Sub

Public Sub StopPlay()
NoteCounter = 0
tmrPlay.Enabled = False
'tmrMetro.Enabled = False
tmrScroll.Enabled = False
Call StopNote(Tone)
End Sub

Public Sub Play(Tempo As Integer)
Call MidiInitialize
NoteCounter = CursorX - 1
AktivFis = ""
AktivB = ""
PlayCounter1 = 0
PlayCounter2 = 0
PlayTempo = Tempo
If Notes(1).Chord <> "" Then CurrentChord = Notes(1).Chord Else CurrentChord = ""
tmrPlay.Interval = 1
'tmrMetro.Interval = 1
'tmrMetro.Enabled = True
tmrPlay.Enabled = True
tmrScroll.Enabled = True
End Sub

Private Sub UserControl_EnterFocus()
Call CreateCaret(UserControl.hwnd, 0, 2, 40)
Call ShowCaret(UserControl.hwnd)
End Sub

Private Sub UserControl_Initialize()
Chords(0) = ":1;5;8"
Chords(1) = "6:1;5;8;10"
Chords(2) = "7:1;5;8;11"
Chords(3) = "maj7:1;5;8;12"
Chords(4) = "9:1;5;8;11;15"
Chords(5) = "m:1;4;8"
Chords(6) = "m6:1;4;8;10"
Chords(7) = "m7:1;4;8;11"
Chords(8) = "m maj7:1;4;8;12"
Chords(9) = "m9:1;4;8;11;15"
Chords(10) = "dim:1;4;7"
Chords(11) = "aug:1;5;9"
Chords(12) = "sus4:1;6;8"
Chords(13) = "add9:1;5;8;15"
Chords(14) = "11:1;5;8;11;15;18"
Chords(15) = "13:1;5;8;11;15;18;22"
Chords(16) = "6add9:1;5;8;10;15"
Chords(17) = "-5:1;5;7"
Chords(18) = "7-5:1;5;7;11"
Chords(19) = "7 maj5:1;5;9;11"
Chords(20) = "7 sus4:1;6;8;11"
Chords(21) = "maj9:1;5;8;12;15"

Tones(0) = "C"
Tones(1) = "C#/Db"
Tones(2) = "D"
Tones(3) = "D#/Eb"
Tones(4) = "E"
Tones(5) = "F"
Tones(6) = "F#/Gb"
Tones(7) = "G"
Tones(8) = "G#/Ab"
Tones(9) = "A"
Tones(10) = "A#/Bb"
Tones(11) = "B"
Tones(12) = "C"
Tones(13) = "C#/Db"
Tones(14) = "D"
Tones(15) = "D#/Eb"
Tones(16) = "E"
Tones(17) = "F"
Tones(18) = "F#/Gb"
Tones(19) = "G"
Tones(20) = "G#/Ab"
Tones(21) = "A"
Tones(22) = "A#/Bb"
Tones(23) = "B"
Tones(24) = "C"
Tones(25) = "C#/Db"
Tones(26) = "D"
Tones(27) = "D#/Eb"
Tones(28) = "E"
Tones(29) = "F"
Tones(30) = "F#/Gb"
Tones(31) = "G"
Tones(32) = "G#/Ab"
Tones(33) = "A"
Tones(34) = "A#/Bb"
Tones(35) = "B"
Tones(36) = "C"

Space = 45
TimeSignature = 1
Signature = 0
Call DrawLines
End Sub

Sub SetChord(Chord As String)
Dim i As Integer
Dim x As Integer
Dim Chord1() As String
Dim Chord2() As String
Dim Base As Integer
Dim NewChord As String
Dim NewChord2 As String
Dim Upper As Integer
Dim Pitch As Integer
Dim B() As String
Dim T() As String
Dim Tone As String

On Error Resume Next
If Chord = "" Then Exit Sub
Pitch = 12

If Mid$(Chord, 2, 1) = "#" Or Mid$(Chord, 2, 1) = "b" Then
    NewChord = Left$(Chord, 2)
    NewChord2 = Right(Chord, Len(Chord) - 2)
Else
    NewChord = Left$(Chord, 1)
    NewChord2 = Right(Chord, Len(Chord) - 1)
End If

For i = 0 To 21
    T = Split(Tones(i), "/")
    If UBound(T) > 0 Then
        If T(0) = NewChord Then
            Base = i
            Exit For
        ElseIf T(1) = NewChord Then
            Base = i
            Exit For
        End If
    Else
        If Tones(i) = NewChord Then
            Base = i
            Exit For
        End If
    End If
Next i

For i = 0 To 21
    B = Split(Chords(i), ":")
    If B(0) = NewChord2 Then
        
        Upper = i
        Exit For
    End If
Next i

Chord1 = Split(Chords(Upper), ":")
If UBound(Chord1) > 0 Then Chord2 = Split(Chord1(1), ";")

For i = 0 To 36 + 12 + Pitch
    Call StopNote(i)
Next i

For i = 0 To 36
    For x = 0 To UBound(Chord2)
        If i + 1 = Chord2(x) + Base Then
            Call StartNote(i + 12 + Pitch, 80)
        End If
    Next x
Next i
End Sub

Public Sub ChangeTime(Value As Single)
TimeSignature = Value
Call DrawLines(Scroll)
End Sub

Public Sub ChangeSignature(Value As Single)
Signature = Value
Call DrawLines(Scroll)
End Sub

Private Sub DrawLines(Optional Val As Long, Optional Start As Long = 1)
Dim i As Long
Dim u As Long
Dim e As Long
Dim Values As Single
Dim Value As Long
Dim NotePos As Integer

On Error Resume Next
'Bildschirm löschen und Notenschlüssel zeichnen
Call UserControl.Cls
UserControl.DrawWidth = 1
UserControl.CurrentX = 5 - Val
UserControl.CurrentY = 35
UserControl.Print Chr$(65)

'Vorzeichen
Select Case Signature
    Case 0
        Additional = 0
    Case 1
        UserControl.CurrentY = 14
        UserControl.CurrentX = 30 - Val
        UserControl.Print Chr(83)
        Additional = 9
    Case 2
        UserControl.CurrentY = 14
        UserControl.CurrentX = 30 - Val
        UserControl.Print Chr(83)
        
        UserControl.CurrentY = 24
        UserControl.CurrentX = 38 - Val
        UserControl.Print Chr(83)
        Additional = 18
    Case 3
        UserControl.CurrentY = 14
        UserControl.CurrentX = 30 - Val
        UserControl.Print Chr(83)
        
        UserControl.CurrentY = 24
        UserControl.CurrentX = 38 - Val
        UserControl.Print Chr(83)
        
        UserControl.CurrentY = 9
        UserControl.CurrentX = 46 - Val
        UserControl.Print Chr(83)
        Additional = 27
    Case 4
        UserControl.CurrentY = 14
        UserControl.CurrentX = 30 - Val
        UserControl.Print Chr(83)
        
        UserControl.CurrentY = 24
        UserControl.CurrentX = 38 - Val
        UserControl.Print Chr(83)
        
        UserControl.CurrentY = 9
        UserControl.CurrentX = 46 - Val
        UserControl.Print Chr(83)
        
        UserControl.CurrentY = 21
        UserControl.CurrentX = 54 - Val
        UserControl.Print Chr(83)
        Additional = 36
    Case 5
        UserControl.CurrentY = 14
        UserControl.CurrentX = 30 - Val
        UserControl.Print Chr(83)
        
        UserControl.CurrentY = 24
        UserControl.CurrentX = 38 - Val
        UserControl.Print Chr(83)
        
        UserControl.CurrentY = 9
        UserControl.CurrentX = 46 - Val
        UserControl.Print Chr(83)
        
        UserControl.CurrentY = 21
        UserControl.CurrentX = 54 - Val
        UserControl.Print Chr(83)
        
        UserControl.CurrentY = 31
        UserControl.CurrentX = 62 - Val
        UserControl.Print Chr(83)
        Additional = 45
End Select

UserControl.CurrentY = 21
UserControl.CurrentX = 30 - Val + Additional

'Takt zeichnen (Zähler)
If TimeSignature = 1 Then
    UserControl.Print Chr(52)
ElseIf TimeSignature = 0.75 Then
    UserControl.Print Chr(51)
ElseIf TimeSignature = 0.5 Then
    UserControl.Print Chr(50)
End If
'Takt zeichnen (Nenner)
UserControl.CurrentY = 35
UserControl.CurrentX = 30 - Val + Additional
UserControl.Print Chr(52)

'Notenlinien zeichnen
UserControl.Line (0, 50)-(UserControl.Width, 50)
UserControl.Line (0, 57)-(UserControl.Width, 57)
UserControl.Line (0, 64)-(UserControl.Width, 64)
UserControl.Line (0, 71)-(UserControl.Width, 71)
UserControl.Line (0, 78)-(UserControl.Width, 78)
MaxLength = 0

'Alle Noten zeichnen
For i = Start To Notes.Count
    NotePos = Notes(i).NotePos
    UserControl.DrawWidth = 1
    
    'Maximale Balkenlänge errechnen
    For u = i To Notes.Count
        If Notes(u).Join = True Then
            If Notes(u).Note > MaxLength Then
               MaxLength = Notes(u).Note
            Else
                MaxLength = MaxLength
            End If
        Else
            Exit For
        End If
    Next u
    'Nur Noten zeichnen, die auf den Bildschirm passen
    If (i * Space + 25 + Additional + NotePos) >= Val - 60 And (i * Space + 25 + Additional + NotePos) + UserControl.Width <= Val + UserControl.Width + 2500 Then
        'Zeichnen
        Call DrawNote(Notes(i).Note, Notes(i).Length, Notes(i).No, Notes(i).Up, Notes(i).Down, Notes(i).Point, Notes(i).Bow, Notes(i).Join, i - Start + 1, Start, Val)
    End If
    If Notes(i).Join = False Then MaxLength = 0
Next i

'Taktstriche zeichnen
For e = 1 To Notes.Count
    'Notenwerte addieren
    Value = Notes(e).Length
    If Notes(e).Triple = True Then
        Values = Values + (2 * (1 / Value)) / 3
    Else
        Values = Values + (1 / Value)
    End If

    If Notes(e).Point = True Then Values = Values + ((1 / Value) / 2)

    'Taktstrich setzen, wenn der Takt gefüllt ist
    If Values > 0 And Values = TimeSignature Then
        If e >= Start Then
            UserControl.Line (((e + 1) * Space) + 15 - Val + Additional, 50)-(((e + 1) * Space) + 15 - Val + Additional, 78)
        End If
        Values = 0
    End If
    
    'Feste Taktstiche und Wiederholungszeichen
    If Notes(e).Bar = 1 Then
        UserControl.Line (((e + 1) * Space) + 15 - c, 50)-(((e + 1) * Space) + 15 - c, 78)
        Values = 0
    ElseIf Notes(e).Bar = 2 Then
            UserControl.DrawWidth = 3
            UserControl.Line (((e + 1) * Space) + 14 - Val + Additional, 51)-(((e + 1) * Space) + 14 - Val + Additional, 77)
            UserControl.DrawWidth = 1
            UserControl.Line (((e + 1) * Space) + 10 - Val + Additional, 51)-(((e + 1) * Space) + 10 - Val + Additional, 78)
            UserControl.CurrentX = ((e + 1) * Space) + 6 + Additional
            UserControl.CurrentY = 25
            UserControl.Print Chr$(33)
            UserControl.CurrentX = ((e + 1) * Space) + 6 + Additional
            UserControl.CurrentY = 32
            UserControl.Print Chr$(33)
            Values = 0
    End If
Next e

'Triolen einzeichnen
For e = 1 To Notes.Count
    If e > 1 And e < Notes.Count Then
        If Notes(e).Triple = True And Notes(e - 1).Triple = True And Notes(e + 1).Triple = True Then
            If Notes(e).Note > 6 And Notes(e).Join = False Then
                UserControl.CurrentX = (Space * e) + 23 - Val + Additional
                UserControl.CurrentY = 40 - (Notes(e).Note * 3.5)
                UserControl.Print Chr$(249)
            Else
                UserControl.CurrentX = (Space * e) + 23 - Val + Additional
                UserControl.CurrentY = 65 - (Notes(e).Note * 3.5)
                UserControl.Print Chr$(249)
            End If
            e = e + 2
        End If
    End If
Next e

'Cursor setzen
NotePos = Notes(CursorX).NotePos
Call SetCaretPos((Space * CursorX) + 15 + Additional - Val + NotePos, 45)
'Bildschirm aktualisieren
Call UserControl.Refresh
End Sub

Public Sub ScrollTo(Value As Long)
Scroll = Value
Call DrawLines(Value)
End Sub

Private Sub DrawNote(Note As Integer, Length As Integer, No As Boolean, Up As Boolean, Down As Boolean, Point As Boolean, Bow As Boolean, Join As Boolean, Index As Long, Start As Long, c As Long)
Dim i As Long
Dim x As Long
Dim MaxLength2 As Long
Dim NotePos1 As Integer
Dim NotePos2 As Integer
Dim Text As String

On Error Resume Next
'Notenwerte in ASCII-Nummern umrechnen
Select Case Length
    Case 1
        'Ganze
        Length = 227
        If Note = 1000 Then Length = 73
    Case 2
        'Halbe
        Length = 160
        If Note = 1000 Then Length = 74
    Case 4
        'Viertel
        Length = 162
        If Note = 1000 Then Length = 75
    Case 8
        'Achtel
        Length = 164
        If Note = 1000 Then Length = 76
    Case 16
        'Sechszenhntel
        Length = 166
        If Note = 1000 Then Length = 77
'    Case 32
'        Length = 168
'        If Note = 1000 Then Length = 78
    Case Else
        Exit Sub
End Select

If Join = True Then
    Length = 162
End If
    
If Note > 6 And Note <> 1000 And Length <> 227 Then
    If Notes(Index).Join = False Then Length = Length + 1
End If

'X-Position errechnen
NotePos1 = Notes(Index).NotePos
NotePos2 = Notes(Index + 1).NotePos
x = (Space * Index) + 25 - c + Additional + NotePos1
UserControl.CurrentX = x

'Pausen
If Note = 1000 Then
    UserControl.CurrentY = 28
    UserControl.Print Chr$(CLng(Length))
'Alle anderen Noten
Else
    UserControl.CurrentY = 53 - (Note * 3.5)
    If Join = True Then UserControl.Print Chr$(CLng(Length))
    'Freistehende Noten
    If Join = False Then
        'Hohe Noten mit Hilfslinien
        If 53 - (Note * 3.5) > 53 Then
            If Length = 227 Then
                UserControl.Print Chr$(227)
                GoTo Quit
            ElseIf Length = 160 Then
                UserControl.Print Chr$(160)
            Else
                UserControl.Print Chr$(162)
            End If
            'Linien nachzeichnen
            UserControl.Line (x + 7, 60)-(x + 7, UserControl.CurrentY - 40), , BF
            UserControl.CurrentY = 24
            If Length = 164 Then
                UserControl.Print Chr$(230)
            ElseIf Length = 166 Then
                UserControl.Print Chr$(231)
            End If
        'Tiefe Noten mit Hilfslinien
        ElseIf 53 - (Note * 3.5) < 4 Then
            If Length = 227 Then
                UserControl.Print Chr$(227)
                GoTo Quit
            ElseIf Length = 161 Then
                UserControl.Print Chr$(161)
            Else
                UserControl.Print Chr$(163)
            End If
            'Linien nachzeichnen
            UserControl.Line (x, 60)-(x, UserControl.CurrentY - Space + 10), , BF
            UserControl.CurrentY = 28
            If Length = 165 Then
                UserControl.Print Chr(234)
            ElseIf Length = 167 Then
                UserControl.Print Chr(235)
            End If
        Else
            'Alle anderen Noten
            UserControl.Print Chr(CLng(Length))
            'Linien nachzeichnen
            If Length <> 227 And Note < 7 Then UserControl.Line (x + 7, UserControl.CurrentY - 60)-(x + 7, UserControl.CurrentY - 40), , BF
            If Length <> 227 And Note >= 7 Then UserControl.Line (x, UserControl.CurrentY - 14)-(x, UserControl.CurrentY - 34), , BF
        End If
    'Balkengruppen
    Else
        'Maximale Balkenhöhe errechnen
        MaxLength2 = 64 - (MaxLength * 3.5)
        'Linien nachzeichnen
        UserControl.Line (x + 7, MaxLength2)-(x + 7, UserControl.CurrentY - 40), , BF
        'X-Position errechnen
        x = (Space * Index) + 25 - c + Additional
        If Index <> 0 Then
            If Notes(Index).Splitter = False Then
                UserControl.DrawWidth = 4
                If Notes(Index).Join = True And Notes(Index + 1).Join = True Then
                    If Notes(Index).Length = 8 And Notes(Index + 1).Length = 8 Then
                        If Index + 1 <= Notes.Count Then UserControl.Line (x + 7 + NotePos1, MaxLength2 + 2)-(x + Space + 8 + NotePos2, MaxLength2 + 2), , BF
                    ElseIf Notes(Index).Length = 8 And Notes(Index + 1).Length = 16 Then
                        UserControl.Line (x + 7 + NotePos1, MaxLength2 + 2)-(x + Space + 7 + NotePos2, MaxLength2 + 2), , BF
                        If Notes(Index + 2).Length = 8 Or Notes(Index + 2).Join = False Then UserControl.Line (x + Space + NotePos2, MaxLength2 + 7)-(x + Space + 8 + NotePos2, MaxLength2 + 8), , BF
                    ElseIf Notes(Index).Length = 16 And Notes(Index + 1).Length = 8 Then
                        UserControl.Line (x + 7 + NotePos1, MaxLength2 + 2)-(x + Space + 7 + NotePos2, MaxLength2 + 2), , BF
                        If Notes(Index - 1).Join = False Or Notes(Index - 1).Length = 4 Or Notes(Index - 1).Length = 2 Or Notes(Index - 1).Length = 1 Or Notes(Index - 1).Note = 1000 Then UserControl.Line (x + 7 + NotePos1, MaxLength2 + 8)-(x + 14 + NotePos1, MaxLength2 + 8), , BF
                    ElseIf Notes(Index).Length = 16 And Notes(Index + 1).Length = 16 Then
                        UserControl.Line (x + 7 + NotePos1, MaxLength2 + 2)-(x + Space + 8 + NotePos2, MaxLength2 + 2), , BF
                        UserControl.Line (x + 7 + NotePos1, MaxLength2 + 8)-(x + Space + 8 + NotePos2, MaxLength2 + 8), , BF
                    End If
                End If
            Else
                UserControl.DrawWidth = 4
                If Notes(Index).Length = 16 And Notes(Index + 1).Length = 16 Then
                    UserControl.Line (x + NotePos1, MaxLength2 + 8)-(x + 8 + NotePos1, MaxLength2 + 8), , BF
                    UserControl.Line (x + Space + NotePos2 + 7, MaxLength2 + 8)-(x + Space + 14 + NotePos2, MaxLength2 + 8), , BF
                ElseIf Notes(Index).Length = 8 And Notes(Index + 1).Length = 16 Then
                    UserControl.Line (x + Space + NotePos2 + 7, MaxLength2 + 8)-(x + Space + 14 + NotePos2, MaxLength2 + 8), , BF
                End If
            End If
        End If
        UserControl.DrawWidth = 1
    End If
End If

'X-Position errechnen
x = (Space * Index) + 25 - c + Additional + NotePos1
Quit:
'Aufgelöste Noten
If No = True Then
    UserControl.CurrentX = x - 6 '(Space * Index) + 19 - c + Additional + NotePos1
    UserControl.CurrentY = 53 - (Note * 3.5)
    UserControl.Print Chr(82)
End If
'#-Vorzeichen
If Up = True Then
    UserControl.CurrentX = x - 6 '(Space * Index) + 19 - c + Additional + NotePos1
    UserControl.CurrentY = 53 - (Note * 3.5)
    UserControl.Print Chr(83)
End If
'b-Vorzeichen
If Down = True Then
    UserControl.CurrentX = x - 6 '(Space * Index) + 19 - c + Additional + NotePos1
    UserControl.CurrentY = 53 - (Note * 3.5)
    UserControl.Print Chr(81)
End If
'Punktierung
If Point = True Then
    UserControl.CurrentX = x + 11 '(Space * Index) + 36 - c + Additional + NotePos1
    If Note Mod 2 = 0 Then UserControl.CurrentY = 53 - (Note * 3.5) Else UserControl.CurrentY = 50 - (Note * 3.5)
    If Note = 1000 Then UserControl.CurrentY = 25
    UserControl.Print Chr(33)
End If
'Haltebögen
If Bow = True Then
    'Hohe Noten
    If Note > 6 Then
        If Join = False Then
            Call UserControl.PaintPicture(imgBow2.Picture, x + 7, 78 - (Note * 3.5))
        Else
            Call UserControl.PaintPicture(imgBow.Picture, x + 7, 95 - (Note * 3.5))
        End If
    'Tiefe Noten
    Else
        Call UserControl.PaintPicture(imgBow.Picture, x + 7, 95 - (Note * 3.5))
    End If
End If

'Hilfslinie zeichnen
'Tiefe Noten
If Note < 2 Then
    For i = 78 To 90 - (Note * 3.5) Step 7
        If Length = 227 Then
            UserControl.Line (x - 3, i)-(x + 14, i)
        Else
            UserControl.Line (x - 3, i)-(x + 11, i)
        End If
    Next i
'Hohe Noten
ElseIf Note > 12 And Note <> 1000 Then
    For i = 50 To 87 - (Note * 3.5) Step -7
        If Length = 227 Then
            UserControl.Line (x - 3, i)-(x + 14, i)
        Else
            UserControl.Line (x - 3, i)-(x + 11, i)
        End If
    Next i
End If

'Existenz von Liedtext prüfen
If Len(Notes(Index).Text) > 1 And Right$(Notes(Index).Text, 1) = "-" Then
    Text = Left$(Notes(Index).Text, Len(Notes(Index).Text) - 1)
Else
    Text = Notes(Index).Text
End If

'Liedtext und Akkorde
UserControl.Font.Name = "Arial"
UserControl.Font.Size = 8
UserControl.CurrentX = ((Space * Index) + 30) - Int(((UserControl.TextWidth(Text) / 2))) - c + Additional + NotePos1
UserControl.CurrentY = 100
UserControl.Print Text
UserControl.CurrentX = ((Space * Index) + 30) - Int(((UserControl.TextWidth(Notes(Index + Start - 1).Chord) / 2))) - c + Additional + NotePos1
UserControl.CurrentY = 10
UserControl.Print Notes(Index).Chord
If Len(Notes(Index).Text) > 1 Then
    UserControl.CurrentX = Int(((Space * (Index + 1)) - c + Additional + NotePos2) - (((Space * Index)) - c + Additional + NotePos1)) / 2
    UserControl.CurrentX = UserControl.CurrentX + ((Space * Index) - c + Additional + NotePos1) + 28
    UserControl.CurrentY = 100
    If Right$(Notes(Index).Text, 1) = "-" Then UserControl.Print "-"
    
    If Right$(Notes(Index).Text, 1) = "_" Then
        UserControl.Line (((Space * (Index)) - c + Additional + NotePos1) + 40, 98 + UserControl.TextHeight(Notes(Index).Text))-(((Space * (Index + 1)) - c + Additional + NotePos2) + 35, 98 + UserControl.TextHeight(Notes(Index).Text))
    End If
End If
UserControl.Font.Name = "capella3-invertiert"
UserControl.Font.Size = 20
End Sub

Public Sub InsertBar(Value As Integer)
On Error Resume Next
Notes(CursorX - 1).Bar = Value
Call DrawLines(Scroll)
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
Dim NotePosition As Integer

On Error GoTo Error
Select Case KeyCode
    Case 226
        If Shift = 1 Then
            If Notes(CursorX).Length > 1 Then
                Notes(CursorX).Length = Notes(CursorX).Length / 2
                If Notes(CursorX).Join = True Then
                    If Notes(CursorX).Length < 8 Then Notes(CursorX).Length = 8
                End If
            Call DrawLines(Scroll)
            End If
        Else
            If Notes(CursorX).Length < 16 Then
                Notes(CursorX).Length = Notes(CursorX).Length * 2
                Call DrawLines(Scroll)
            End If
        End If
    Case 38
        If Shift = 1 Then
            If SelBoundRight = 0 Then
                If Notes(CursorX).Note >= 21 Then Exit Sub
                Notes(CursorX).Note = Notes(CursorX).Note + 1
                Call DrawLines(Scroll)
            Else
                Call ShiftSelectionUpwards
            End If
        End If
    Case 40
        If Shift = 1 Then
            If SelBoundRight = 0 Then
                If Notes(CursorX).Note <= -6 Then Exit Sub
                Notes(CursorX).Note = Notes(CursorX).Note - 1
                Call DrawLines(Scroll)
            Else
                Call ShiftSelectionDownwards
            End If
        End If
    Case 8
        Call Notes.Remove(CursorX - 1)
        CursorX = CursorX - 1
        Call DrawLines(Scroll)
    Case 39
        If Shift = 1 Then
            Notes(CursorX).NotePos = Notes(CursorX).NotePos + 1
            Call DrawLines(Scroll)
            Exit Sub
        End If
        If CursorX < Notes.Count + 1 Then CursorX = CursorX + 1
        NotePosition = Notes(CursorX).NotePos
        Call SetCaretPos((Space * CursorX) + 15 - Scroll + Additional + NotePosition, 45)
        If (Space * CursorX) + 15 - Scroll > UserControl.Width / Screen.TwipsPerPixelX Then
             RaiseEvent ScrollChanged(100)
        End If
    Case 37
        If Shift = 1 Then
            Notes(CursorX).NotePos = Notes(CursorX).NotePos - 1
            Call DrawLines(Scroll)
            Exit Sub
        End If
        
        If CursorX > 1 Then CursorX = CursorX - 1
        NotePosition = Notes(CursorX).NotePos
        Call SetCaretPos((Space * CursorX) + 15 - Scroll + Additional + NotePosition, 45)
        If (Space * CursorX) + 15 - Scroll < 0 Then
            RaiseEvent ScrollChanged(-100)
        End If
    Case vbKeyDelete
        Call DeleteSelection
End Select
Error:
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
On Error Resume Next
Dim Tone As Integer
Dim Note As New clsNotes

Select Case Chr$(KeyAscii)
    Case "="
        Bow = True
        Exit Sub
    Case "*"
        No = True
        Exit Sub
    Case "+"
        Down = False
        Up = True
        Exit Sub
    Case "-"
        Up = False
        Down = True
        Exit Sub
    Case "."
        Point = True
        Exit Sub
    Case "&"
        Join = True
        Exit Sub
    Case Chr(32)
        Down = False
        Up = False
        Tone = 1000
    Case "c"
        Tone = 1 + Pitch
    Case "d"
        Tone = 2 + Pitch
    Case "e"
        Tone = 3 + Pitch
    Case "f"
        Tone = 4 + Pitch
    Case "g"
        Tone = 5 + Pitch
    Case "a"
        Tone = 6 + Pitch
    Case "h"
        Tone = 7 + Pitch
    Case "3"
        TripleNote = True
        Exit Sub
    Case "1", "2", "4"
        NoteLength = CInt(Chr$(KeyAscii))
        Exit Sub
    Case "6"
        NoteLength = 16
        Exit Sub
    Case "8"
        NoteLength = 8
        Exit Sub
    Case Else
        Exit Sub
End Select

'Note hinzufügen
CursorX = CursorX + 1
Note.Note = Tone
Note.Length = NoteLength
Note.Up = Up
Note.Down = Down
Note.No = No
Note.Point = Point
Note.Join = Join
Note.Text = ""
Note.Triple = TripleNote
Note.Bow = Bow
Note.NotePos = 0
If Notes.Count = 0 Then
    Notes.Add Note
Else
    If CursorX - 2 = 0 Then
        Call Notes.Add(Note, , 1)
    Else
        Call Notes.Add(Note, , , CursorX - 2)
    End If
End If
Call DrawLines(Scroll)

Up = False
Down = False
No = False
Point = False
Join = False
TripleNote = False
Bow = False
End Sub

Public Sub ChangePitch(Value As Integer)
Select Case Value
    Case -1
        Pitch = -7
    Case 0
        Pitch = 0
    Case 1
        Pitch = 7
    Case 2
        Pitch = 14
End Select
End Sub

Public Sub JoinSelection()
Dim i As Long

On Error Resume Next
For i = SelBoundLeft To SelBoundRight - 1
    If Notes(i).Length = 8 Or Notes(i).Length = 16 Then Notes(i).Join = True: Notes(i).Splitter = False
Next i
Call DrawLines(Scroll)
End Sub

Public Sub SplitSelection()
Dim i As Long

On Error Resume Next
If SelBoundRight = 0 Then Notes(CursorX - 1).Splitter = True: Call DrawLines(Scroll): Exit Sub
For i = SelBoundLeft To SelBoundRight - 1
    Notes(i).Join = False
Next i
Call DrawLines(Scroll)
End Sub

Public Sub ShiftSelectionUpwards()
Dim i As Long

On Error Resume Next
If Notes(CursorX).Note >= 21 Then Exit Sub
For i = SelBoundLeft To SelBoundRight - 1
    Notes(i).Note = Notes(i).Note + 1
Next i
Call DrawLines(Scroll)
End Sub

Public Sub ShiftSelectionDownwards()
Dim i As Long

On Error Resume Next
If Notes(CursorX).Note <= -6 Then Exit Sub
For i = SelBoundLeft To SelBoundRight - 1
    Notes(i).Note = Notes(i).Note - 1
Next i
Call DrawLines(Scroll)
End Sub

Public Sub DeleteSelection()
Dim i As Long

On Error Resume Next
For i = SelBoundLeft To SelBoundRight - 1
    Call Notes.Remove(SelBoundLeft)
Next i
Call DrawLines(Scroll)
Call CreateCaret(UserControl.hwnd, 0, 2, 40)
Call ShowCaret(UserControl.hwnd)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
Dim NotePosition As Integer

On Error Resume Next
CursorX = Round((Scroll + x - 45 - Additional) / Space) + 1
If CursorX < 1 Then CursorX = 1
If CursorX > Notes.Count + 1 Then CursorX = Notes.Count + 1
Call CreateCaret(UserControl.hwnd, 0, 2, 40)
Call ShowCaret(UserControl.hwnd)
SelBoundRight = 0
SelBoundLeft = CursorX
NotePosition = Notes(CursorX).NotePos
Call SetCaretPos((Space * CursorX) + 15 + Additional - Scroll + NotePosition, 45)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Dim NotePosition1 As Integer
Dim NotePosition2 As Integer
Dim Cursor2 As Long
Dim Width As Long

On Error Resume Next
If Button = vbLeftButton Then
    Cursor2 = Round((Scroll + x - 45 - Additional) / Space) + 1
    If Cursor2 > Notes.Count + 1 Then Cursor2 = Notes.Count + 1
    If Cursor2 < 1 Then Cursor2 = 1
    np = Notes(CursorX).NotePos
    np2 = Notes(Cursor2).NotePos
    Width = ((Space * Cursor2) + Additional + NotePosition2 - Scroll) - ((Space * CursorX) + Additional + NotePosition1 - Scroll)
    If Width < 2 Then Width = 2
    Call CreateCaret(UserControl.hwnd, 0, Width, 40)
    Call ShowCaret(UserControl.hwnd)
    Call SetCaretPos((Space * CursorX) + 15 + Additional - Scroll + NotePosition, 45)
    SelBoundRight = Cursor2
    Call UserControl.Refresh
End If
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
NoteLength = 4
Call SetCaretPos(45, 45)
CursorX = 1
End Sub

Private Sub UserControl_Resize()
Call DrawLines(Scroll)
End Sub

Private Sub UserControl_Terminate()
Call MidiReset
End Sub

Private Function TrimString(Text As String) As String
Dim LeftEnd As Long
Dim RightEnd As Long

'Start suchen
For LeftEnd = 1 To Len(Text)
    If Asc(Mid(Text, LeftEnd, 1)) > vbKeySpace Then Exit For
Next LeftEnd
'Ende suchen
For RightEnd = Len(Text) To LeftEnd + 1 Step -1
    If Asc(Mid(Text, RightEnd, 1)) > vbKeySpace Then Exit For
Next RightEnd
'Fertig
TrimString = Mid(Text, LeftEnd, RightEnd - LeftEnd + 1)
End Function
