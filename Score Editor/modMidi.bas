Attribute VB_Name = "modMidi"
Public Const INVALID_NOTE = -1     ' Code for keyboard keys that we don't handle
'for piano play
Public numDevices As Long      ' number of midi output devices
Public curDevice As Long       ' current midi device
Public hmidi As Long           ' midi output handle
Public rc As Long              ' return code
Public midimsg As Long         ' midi output message buffer
Public mipublicsg As Long      ' midi output message buffer
Public channel As Integer      ' midi output channel
Public Volume As Integer       ' midi volume
Public baseNote As Integer     ' the first note on our "piano"

Public bUnload As Boolean

Public Sub MidiInitialize()
Dim i As Long
Dim caps As MIDIOUTCAPS

'Get the rest of the midi devices
numDevices = midiOutGetNumDevs()
    
For i = 0 To (numDevices - 1)
    Call midiOutGetDevCaps(i, caps, Len(caps))
Next
   
'Select the MIDI Mapper as the default device
Call MidiDevice
   
'Set the default channel
channel = 0
'Set the base note
baseNote = 24
'Set volume range
Volume = 127
End Sub

Public Sub MidiDevice()
rc = midiOutClose(hmidi)
rc = midiOutOpen(hmidi, curDevice, 0, 0, 0)
End Sub

Public Sub MidiReset()
rc = midiOutClose(hmidi)
End Sub

Public Sub StartNote(Index As Integer, Optional Volume As Integer)
midimsg = &H90 + ((baseNote + Index) * &H100) + (Volume * &H10000) + channel
midiOutShortMsg hmidi, midimsg
End Sub

Public Sub StopNote(Index As Integer)
midimsg = &H80 + ((baseNote + Index) * &H100) + channel
midiOutShortMsg hmidi, midimsg
End Sub

Public Sub StartMetro(Index As Integer)
midimsg = &H90 + ((baseNote + Index) * &H100) + (Volume * &H10000) + 9
midiOutShortMsg hmidi, midimsg
End Sub

Public Sub StopMetro(Index As Integer)
midimsg = &H80 + ((baseNote + Index) * &H100) + 9 '10
midiOutShortMsg hmidi, midimsg
End Sub

Public Sub StartDong(Index As Integer)
midimsg = &H90 + ((baseNote + Index) * &H100) + (Volume * &H10000) + 9
midiOutShortMsg hmidi, midimsg
End Sub

Public Sub StopDong(Index As Integer)
midimsg = &H80 + ((baseNote + Index) * &H100) + 9
midiOutShortMsg hmidi, midimsg
End Sub
