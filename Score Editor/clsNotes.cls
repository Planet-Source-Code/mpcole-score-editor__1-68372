VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsNotes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim mNote As Integer
Dim mLength As Integer
Dim mUp As Boolean
Dim mDown As Boolean
Dim mPoint As Boolean
Dim mJoin As Boolean
Dim mText As String
Dim mBar As Integer
Dim mTriple As Boolean
Dim mNo As Boolean
Dim mBow As Boolean
Dim mChord As String
Dim mNotePos As Integer
Dim mSplitter As Boolean

Public Property Let NotePos(ByVal Data As Integer)
mNotePos = Data
End Property

Public Property Get NotePos() As Integer
NotePos = mNotePos
End Property

Public Property Let Note(ByVal Data As Integer)
mNote = Data
End Property

Public Property Get Note() As Integer
Note = mNote
End Property

Public Property Let Length(ByVal Data As Integer)
mLength = Data
End Property

Public Property Get Length() As Integer
Length = mLength
End Property

Public Property Let Up(ByVal Data As Boolean)
mUp = Data
End Property

Public Property Get Up() As Boolean
Up = mUp
End Property

Public Property Let Down(ByVal Data As Boolean)
mDown = Data
End Property

Public Property Get Down() As Boolean
Down = mDown
End Property

Public Property Let Point(ByVal Data As Boolean)
mPoint = Data
End Property

Public Property Get Point() As Boolean
Point = mPoint
End Property

Public Property Let Join(ByVal Data As Boolean)
mJoin = Data
End Property

Public Property Get Join() As Boolean
Join = mJoin
End Property

Public Property Let Splitter(ByVal Data As Boolean)
mSplitter = Data
End Property

Public Property Get Splitter() As Boolean
Splitter = mSplitter
End Property

Public Property Let Bar(ByVal Data As Integer)
mBar = Data
End Property

Public Property Get No() As Boolean
No = mNo
End Property

Public Property Let No(ByVal Data As Boolean)
mNo = Data
End Property

Public Property Get Bar() As Integer
Bar = mBar
End Property

Public Property Let Text(ByVal Data As String)
mText = Data
End Property

Public Property Get Text() As String
Text = mText
End Property

Public Property Let Triple(ByVal Data As Boolean)
mTriple = Data
End Property

Public Property Get Triple() As Boolean
Triple = mTriple
End Property

Public Property Let Bow(ByVal Data As Boolean)
mBow = Data
End Property

Public Property Get Bow() As Boolean
Bow = mBow
End Property

Public Property Let Chord(ByVal Data As String)
mChord = Data
End Property

Public Property Get Chord() As String
Chord = mChord
End Property

