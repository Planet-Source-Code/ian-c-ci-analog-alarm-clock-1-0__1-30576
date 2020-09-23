VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CircleObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*import (modCoordinate, modArc)

Option Explicit

'/**
'* @project       TUM VB Praktikum WS 2001/2002
'* @spec          Aufgabe 4 - Analog Clock
'* @date          22.11.2001
'* @supervisor    D. Golias
'* @author        I. Carter, ian@computersciences.de (mail if you need support or want to comment)
'* @class         CircleObject
'* @description   geometrical circle
'/**

'*not yet implemented:
'Private m_lngFillColor As Long
'Private m_dblArea As Double
'Private m_dblCircumfence As Double

Private Const m_dblPi As Double = 3.14159265358979
Private m_lngRadius As Long
Private m_lngOrigin As Coordinate2D
Private m_lngForeColor As Long
Private m_dblAngle As Double
Private m_lngSections As Long


'/**
'* @attribute     form client
'* @description   an object (form) where the circle can be ploted
'/**
Public client As Object


'/**
'* @enum          long CirclePlottingConstants
'* @description   options what should be plotted
'/**
Public Enum CirclePlottingConstants
   enmCircle = 0
   enmRadius = 1
   enmCenter = 2
   enmSections = 3
   enmScale = 4
   enmDiameter = 5
End Enum


'/**
'* @attribute     double pi
'* @description   the circle constant pi
'/**
Public Property Get pi() As Double
   pi = m_dblPi
End Property


'/**
'* @attribute     long currentX
'* @description   x center of the circle
'/**
Public Property Let currentX(ByVal lng As Long)
   m_lngOrigin.X = lng
End Property
Public Property Get currentX() As Long
   currentX = m_lngOrigin.X
End Property


'/**
'* @attribute     long currentY
'* @description   y center of the circle
'/**
Public Property Let currentY(ByVal lng As Long)
   m_lngOrigin.Y = lng
End Property
Public Property Get currentY() As Long
   currentY = m_lngOrigin.Y
End Property


'/**
'* @attribute     long radius
'* @description   radius of the circle
'/**
Public Property Let radius(ByVal lng As Long)
   m_lngRadius = lng
End Property
Public Property Get radius() As Long
   radius = m_lngRadius
End Property


'/**
'* @attribute     long diameter
'* @description   diameter of the circle
'/**
Public Property Get diameter() As Long
   diameter = 2 * m_lngRadius
End Property


'/**
'* @attribute     long sections
'* @description   sections of the circle
'/**
Public Property Let sections(ByVal lng As Long)
   m_lngSections = lng
End Property
Public Property Get sections() As Long
   sections = m_lngSections
End Property


'/**
'* @attribute     double angle
'* @description   angle of the radius of the circle
'/**
Public Property Let angle(ByVal dbl As Double)
   m_dblAngle = dbl
End Property
Public Property Get angle() As Double
   angle = m_dblAngle
End Property


'/**
'* @attribute     long foreColor
'* @description   color of ploted graphics
'/**
Public Property Let foreColor(ByVal lng As Long)
   m_lngForeColor = lng
End Property
Public Property Get foreColor() As Long
   foreColor = m_lngForeColor
End Property


'/**
'* @method        void plot
'* @description   plots circle-graphics on the client
'/**
Public Sub plot(Optional str As CirclePlottingConstants, Optional bytWidth As Byte = 1)
   Dim i As Integer, lngX As Long, lngY As Long
   client.DrawWidth = bytWidth
   
   Select Case LCase(str)
      
      Case 1 '*"radius"
         client.Line (m_lngOrigin.X, m_lngOrigin.Y)- _
                     (m_lngOrigin.X + Cos(arc(angle + 270)) * radius, _
                      m_lngOrigin.Y + Sin(arc(angle + 270)) * radius), m_lngForeColor
      
      Case 2 '*"center"
         client.PSet (m_lngOrigin.X, m_lngOrigin.Y), m_lngForeColor
      
      Case 3 '*"sections"
         For i = 1 To m_lngSections
            lngX = m_lngRadius * Cos(1.74532925199433E-02 * ((360 / m_lngSections) * i - 90)) + m_lngOrigin.X
            lngY = m_lngRadius * Sin(1.74532925199433E-02 * ((360 / m_lngSections) * i - 90)) + m_lngOrigin.Y
            client.PSet (lngX, lngY)
            '*Print I
         Next i
      
      Case 4 '*"scale"
         '*use bytWidth as fontsize if displaying scale
         client.Font.Size = bytWidth

         For i = 1 To m_lngSections
            client.currentX = m_lngRadius * Cos(1.74532925199433E-02 * ((360 / m_lngSections) * i - 90)) + m_lngOrigin.X - 70
            client.currentY = m_lngRadius * Sin(1.74532925199433E-02 * ((360 / m_lngSections) * i - 90)) + m_lngOrigin.Y - 70
            client.Print i
         Next i

      Case 5 '*"diameter"
'*         client.Line (m_lngOrigin.X, m_lngOrigin.Y)- _
'*                  (m_lngOrigin.X + Cos(arc(angle + 270)) * radius, _
'*                   m_lngOrigin.Y + Sin(arc(angle + 270)) * radius), m_lngForeColor
'*         client.Line (m_lngOrigin.X, m_lngOrigin.Y)- _
'*                  (m_lngOrigin.X - Cos(arc(angle + 270)) * radius, _
'*                   m_lngOrigin.Y - Sin(arc(angle + 270)) * radius), m_lngForeColor
                   
      Case Else '*0 - circle
         client.Circle (m_lngOrigin.X, m_lngOrigin.Y), m_lngRadius, m_lngForeColor
   
   End Select
End Sub




