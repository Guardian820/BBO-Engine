Attribute VB_Name = "modDirectMusic7"
Option Explicit

' ******************************************
' ** DirectMusic7, plays .MID music files **
' ******************************************

' DirectMusic variables
Public Performance As DirectMusicPerformance ' this controls the music
Public Segment As DirectMusicSegment ' this stores the music in memory
Private Loader As DirectMusicLoader 'parses the data from a file on the hard drive to the area in memory.

Public Sub InitDirectMusic()
    Set Loader = DX7.DirectMusicLoaderCreate
    Set Performance = DX7.DirectMusicPerformanceCreate
    
    Call Performance.Init(Nothing, frmMainGame.hWnd)
    Call Performance.SetPort(-1, 80)
                                    ' adjust volume 0-100
    Call Performance.SetMasterVolume(75 * 42 - 3000)
    Call Performance.SetMasterAutoDownload(True)
End Sub

Public Sub DirectMusic_PlayMidi(FileName As String)
'    Set Segment = Loader.LoadSegment(App.Path & MUSIC_PATH & FileName)
    
    If Err.Number <> DD_OK Then
        Call AddText("Error: Could not load MIDI file!", BrightRed)
    End If
    
    ' repeat midi file
'    Segment.SetLoopPoints 0, 0
 '   Segment.SetRepeats 100
    
 '   Performance.PlaySegment Segment, 0, 0
End Sub

Public Sub DirectMusic_StopMidi()
    Performance.Stop Segment, Nothing, 0, 0
End Sub

Public Sub DestroyDirectMusic()
    Set Performance = Nothing
    Set Loader = Nothing
End Sub

