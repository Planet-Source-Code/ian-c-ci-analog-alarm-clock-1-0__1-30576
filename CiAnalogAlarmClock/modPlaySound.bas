Attribute VB_Name = "modPlaySound"
Option Explicit

'/**
'* @project       TUM VB Praktikum WS 2001/2002
'* @spec          Aufgabe 4 - Analog Clock
'* @date          22.11.2001
'* @supervisor    D. Golias
'* @author        I. Carter, ian@computersciences.de (mail if you need support or want to comment)
'* @class         Module modMoveTopRight
'/**


'/**
'* @api           Long sndPlaySoundA
'* @description   (SYSTEM32\WINMM.DLL) plays a wavefile
'/**
Private Declare Function sndPlaySoundA Lib "winmm" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long


'/**
'* @enum          hex SoundPlayingConstants
'* @description   api flag constants modifying the style of playback
'/**
Public Enum SoundPlayingConstants
    enmSync = &H0       '*play synchronously (default)
    enmAsync = &H1      '*play in background
    enmNoDefault = &H2  '*no default
    enmLoop = &H8       '*loop the sound until next sndPlaySound
    enmNonStop = &H10   '*loop
End Enum


'/**
'* @method        long playSound
'* @description   calls the sndPlaySound API to play a wave file
'/**
Public Function playSound(strURL As String, Optional enmStyle As SoundPlayingConstants = enmSync, Optional sNAME = "") As Long
    On Error Resume Next
    If strURL <> "" Then playSound = sndPlaySoundA(strURL, enmStyle)
End Function





