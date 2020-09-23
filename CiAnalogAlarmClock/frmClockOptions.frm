VERSION 5.00
Begin VB.Form frmClockOptions 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1890
   ClientLeft      =   255
   ClientTop       =   45
   ClientWidth     =   3210
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H00000000&
   Icon            =   "frmClockOptions.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   126
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   214
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "Options"
      ClipControls    =   0   'False
      DragMode        =   1  'Automatic
      Height          =   1635
      Left            =   135
      TabIndex        =   0
      Top             =   135
      Width           =   2940
      Begin VB.ComboBox cmbMelody 
         Height          =   315
         ItemData        =   "frmClockOptions.frx":030A
         Left            =   945
         List            =   "frmClockOptions.frx":0317
         TabIndex        =   7
         Text            =   "sndKuckuck.wav"
         ToolTipText     =   "Eine WAV Datei im Anwendungsverzeichnis"
         Top             =   1170
         Width           =   1815
      End
      Begin VB.CheckBox chkSound 
         Alignment       =   1  'Right Justify
         Caption         =   "Tick"
         Height          =   195
         Left            =   180
         TabIndex        =   3
         Top             =   270
         Width           =   960
      End
      Begin VB.TextBox txtAppointment 
         Height          =   285
         Left            =   945
         TabIndex        =   2
         Text            =   "06:35:00"
         Top             =   540
         Width           =   1815
      End
      Begin VB.TextBox txtMemo 
         Height          =   285
         Left            =   945
         TabIndex        =   1
         Text            =   "Hi there take a break! :-) - (c) 2002 Ian Carter"
         Top             =   855
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "Alarm:"
         Height          =   240
         Left            =   225
         TabIndex        =   6
         Top             =   1230
         Width           =   600
      End
      Begin VB.Label Label2 
         Caption         =   "Memo:"
         Height          =   240
         Left            =   225
         TabIndex        =   5
         Top             =   900
         Width           =   600
      End
      Begin VB.Label Label1 
         Caption         =   "Time:"
         Height          =   240
         Left            =   225
         TabIndex        =   4
         Top             =   585
         Width           =   600
      End
   End
End
Attribute VB_Name = "frmClockOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'/**
'* @project       TUM VB Praktikum WS 2001/2002
'* @spec          Aufgabe 4 - Analog Clock
'* @date          22.11.2001
'* @supervisor    D. Golias
'* @author        I. Carter, ian@computersciences.de (mail if you need support or want to comment)
'* @class         Form frmClockOptions
'/**


'/**
'* @event         void Form_Load
'* @description   sets the alarm time to 15 seconds after loading
'/**
Private Sub Form_Load()
   Me.txtAppointment = Format(Now + CDate("00:00:15"), "HH:MM:SS")
End Sub
