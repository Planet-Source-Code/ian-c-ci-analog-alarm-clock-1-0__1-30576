Attribute VB_Name = "modCoordinate"
Option Explicit

'/**
'* @project       TUM VB Praktikum WS 2001/2002
'* @spec          Aufgabe 4 - Analog Clock
'* @date          22.11.2001
'* @supervisor    D. Golias
'* @author        I. Carter, ian@computersciences.de (mail if you need support or want to comment)
'* @class         Module modCoordinate
'/**


'/**
'* @type          Coordinate2D
'* @description   holds two long coordinates representing x and y
'/**
Public Type Coordinate2D
   X As Long
   Y As Long
End Type


'/**
'* @type          Coordinate3D
'* @description   holds three long coordinates representing x,y and z
'/**
Public Type Coordinate3D
   X As Long
   Y As Long
   z As Long
End Type


'/**--
'* @type          Coordinate4D
'* @description   holds four long coordinates representing x,y,z and time
'/**--
Public Type Coordinate4D
   X As Long
   Y As Long
   z As Long
   t As Long
End Type
