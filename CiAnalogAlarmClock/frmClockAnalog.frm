VERSION 5.00
Begin VB.Form frmClockAnalog 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00800028&
   Caption         =   "Ian Carter - Aufgabe 4"
   ClientHeight    =   1950
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2040
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmClockAnalog.frx":0000
   LinkTopic       =   "frmClockAnalog"
   Picture         =   "frmClockAnalog.frx":030A
   ScaleHeight     =   1950
   ScaleWidth      =   2040
   Tag             =   "2000x2200"
   Begin VB.Timer tmrSound 
      Interval        =   999
      Left            =   2520
      Top             =   2520
   End
   Begin VB.PictureBox picClockIco 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Index           =   7
      Left            =   810
      Picture         =   "frmClockAnalog.frx":E7E4
      ScaleHeight     =   465
      ScaleWidth      =   510
      TabIndex        =   11
      Top             =   3510
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.PictureBox picClockIco 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Index           =   8
      Left            =   1350
      Picture         =   "frmClockAnalog.frx":EAEE
      ScaleHeight     =   465
      ScaleWidth      =   510
      TabIndex        =   10
      Top             =   3510
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.PictureBox picClockIco 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Index           =   9
      Left            =   270
      Picture         =   "frmClockAnalog.frx":EDF8
      ScaleHeight     =   465
      ScaleWidth      =   510
      TabIndex        =   9
      Top             =   4005
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.PictureBox picClockIco 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Index           =   10
      Left            =   810
      Picture         =   "frmClockAnalog.frx":F102
      ScaleHeight     =   465
      ScaleWidth      =   510
      TabIndex        =   8
      Top             =   4005
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.PictureBox picClockIco 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Index           =   11
      Left            =   1350
      Picture         =   "frmClockAnalog.frx":F40C
      ScaleHeight     =   465
      ScaleWidth      =   510
      TabIndex        =   7
      Top             =   4005
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.PictureBox picClockIco 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Index           =   6
      Left            =   270
      Picture         =   "frmClockAnalog.frx":F716
      ScaleHeight     =   465
      ScaleWidth      =   510
      TabIndex        =   6
      Top             =   3510
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.PictureBox picClockIco 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Index           =   0
      Left            =   270
      Picture         =   "frmClockAnalog.frx":FA20
      ScaleHeight     =   465
      ScaleWidth      =   510
      TabIndex        =   5
      Top             =   2520
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.PictureBox picClockIco 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Index           =   5
      Left            =   1350
      Picture         =   "frmClockAnalog.frx":FD2A
      ScaleHeight     =   465
      ScaleWidth      =   510
      TabIndex        =   4
      Top             =   3015
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.PictureBox picClockIco 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Index           =   4
      Left            =   810
      Picture         =   "frmClockAnalog.frx":10034
      ScaleHeight     =   465
      ScaleWidth      =   510
      TabIndex        =   3
      Top             =   3015
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.PictureBox picClockIco 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Index           =   3
      Left            =   270
      Picture         =   "frmClockAnalog.frx":1033E
      ScaleHeight     =   465
      ScaleWidth      =   510
      TabIndex        =   2
      Top             =   3015
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.PictureBox picClockIco 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Index           =   2
      Left            =   1350
      Picture         =   "frmClockAnalog.frx":10648
      ScaleHeight     =   465
      ScaleWidth      =   510
      TabIndex        =   1
      Top             =   2520
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.PictureBox picClockIco 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Index           =   1
      Left            =   810
      Picture         =   "frmClockAnalog.frx":10952
      ScaleHeight     =   465
      ScaleWidth      =   510
      TabIndex        =   0
      Top             =   2520
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Timer tmr 
      Interval        =   2
      Left            =   1980
      Top             =   2520
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1545
      Left            =   210
      TabIndex        =   12
      Tag             =   "(c) 2001 - Ian Carter"
      ToolTipText     =   "(c) 2001 Ian Carter"
      Top             =   225
      Width           =   1500
   End
End
Attribute VB_Name = "frmClockAnalog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*import (CircleObject, modEnsureMinimumHeight, modEnsureMinimumWidth,
'*        modResizeWithParent, modGetSmallerValue, modMoveTopRight,
'*        modPlaySound

Option Explicit

'/**
'* @project       TUM VB Praktikum WS 2001/2002
'* @spec          Aufgabe 4 - Analog Clock
'* @date          22.11.2001
'* @supervisor    D. Golias
'* @author        I. Carter, ian@computersciences.de (mail if you need support or want to comment)
'* @download    http://www.planet-source-code.com/xq/ASP/txtCodeId.30576/lngWId.1/qx/vb/scripts/ShowCode.htm
'* @class         Form frmClockAnalog
'* @bugs          on my pentiumII win98 the seconds-hand sometimes
'*                jumped one second irregularly if interval was set
'*                to 1000. Solution: interval = 2
'/**


Private m_objClock As CircleObject
Private m_objOptions As frmClockOptions


'/**
'* @method        void plotHand
'* @description   plots a clock-hand via the circle object
'/**
Public Sub plotHand(Winkel, Laenge, Farbe)
   With m_objClock
      .angle = Winkel
      .radius = Laenge
      .foreColor = Farbe
      .plot enmRadius, 1 + Int(0.004 * Int(getSmallerValue(.currentX, .currentY) - Int(Laenge) + 1))
   End With
End Sub


'/**
'* @event         void Form_Load
'* @description   constructor: creates a circle and an options object
'/**
Private Sub Form_Load()
   moveTopRight Me
   Set m_objClock = New CircleObject
   Set m_objOptions = New frmClockOptions
   Set m_objClock.client = Me
   Form_Resize
End Sub


'/**
'* @event         void Form_GotFocus
'* @description   hides the options dialog if it looses focus
'/**
Private Sub Form_GotFocus()
   m_objOptions.Hide
End Sub


'/**
'* @event         void Form_Resize
'* @description   ensures minimum dimensions, resizes the infolabel
'*                and recalculates the forms center
'/**
Private Sub Form_Resize()
   Me.Cls
   ensureMinimumHeight Me, 8
   ensureMinimumWidth Me, 17
   
   resizeWithParent lblInfo, 0
   
   With m_objClock
      .currentX = Me.ScaleWidth / 2
      .currentY = Me.ScaleHeight / 2
   End With
   
End Sub


'/**
'* @event         void tmr_Timer
'* @description   updates icon, caption, hands, dial, frame and date
'*                every 2 milliseconds (~constantly)
'/**
Private Sub tmr_Timer()
   On Error Resume Next
   Dim lngPartLength As Long
   Dim i As Integer

   Me.Caption = Time
   Me.Icon = Me.picClockIco(Round(Second(Now()) / 5, 0))

   If Me.WindowState = vbMinimized Then Exit Sub '*dont plot if nobody sees it ;-)
   Me.Cls
   lngPartLength = getSmallerValue(Me.ScaleHeight, Me.ScaleWidth) / 10
   
   '*alarm: like hour only with static user set time
   plotHand ((Hour(m_objOptions.txtAppointment) Mod 12) * 30) + (Int(Minute(m_objOptions.txtAppointment) / 12) * 6), 3.5 * lngPartLength, vbBlue
   
   '*hour: convert from 24h to 12h, multiply by 30° (30*12=360) and adjust to minutes (fraction of hour)
   plotHand ((Hour(Now) Mod 12) * 30) + (Int(Minute(Now) / 12) * 6), 2 * lngPartLength, vbWhite
         
   '*minute: multiply by 6° (6*60=360)
   plotHand Minute(Time()) * 6, 3 * lngPartLength, vbWhite
            
   '*second: multiply by 6° (6*60=360)
   plotHand Second(Time()) * 6, 3.8 * lngPartLength, vbRed
   
   m_objClock.radius = 4 * lngPartLength
   m_objClock.foreColor = &H993030
   m_objClock.plot enmCircle, 1

   m_objClock.sections = 60
   m_objClock.plot enmSections, 1

   m_objClock.sections = 12
   m_objClock.plot enmSections, 2

   m_objClock.radius = 4.3 * lngPartLength
   m_objClock.plot enmScale, CByte(0.01 * Int(lngPartLength))

   m_objClock.foreColor = vbRed
   m_objClock.plot enmCenter, 0.02 * lngPartLength

   '*frame
   Line (0, 0)-(Me.ScaleWidth, Me.ScaleHeight), &HFF8080, B

   '*date
   Me.currentX = 0.8 * Me.ScaleWidth
   Me.currentY = 0.02 * Me.ScaleHeight
   Print Format(Date, "dddd DD. MMMM YYYY")
End Sub


'/**
'* @event         void tmrSound_Timer
'* @description   updates sound and infolabel and evaluates alarm-time
'/**
Private Sub tmrSound_Timer()
   Dim i As Integer
   
   If m_objOptions.chkSound.Value = 1 Then playSound App.Path & "\sndTick.wav", enmAsync
   
   '*UpdateDate
   lblInfo.ToolTipText = Format(Date, "dddd DD. MMMM YYYY") & " - (rightclick to edit options)"

   
   If m_objOptions.txtAppointment = Me.Caption Then
      Me.WindowState = vbMaximized
      Me.SetFocus
      tmr_Timer
      
      If m_objOptions.cmbMelody.Text = "Beep" Then
         For i = 0 To 300
            Beep
         Next
      Else
         playSound App.Path & "\" & m_objOptions.cmbMelody.Text, enmAsync + enmLoop
      End If
      
      MsgBox m_objOptions.txtMemo, vbInformation + vbSystemModal, Format(Time, "HH:MM") & "!"
      playSound App.Path & "\-.wav", enmAsync '*ensure that the alarmsound quits

   End If
End Sub


'/**
'* @event         void lblInfo_MouseDown
'* @description   on rightclick shows optional settings
'/**
Private Sub lblInfo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = 2 Then m_objOptions.Show , Me
End Sub


'/**
'* @event         void Form_Unload
'* @description   destructor: cleans things up
'/**
Private Sub Form_Unload(Cancel As Integer)
   playSound App.Path & "\-.wav", enmAsync '*ensure no background noise
   Unload m_objOptions '*ensure unload of the options dialog
End Sub
