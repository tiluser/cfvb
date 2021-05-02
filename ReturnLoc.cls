VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ReturnLoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Creole Forth for VB
' Version 0.01
' Copyright 2021 Joseph M. O'Connor
' This Source Code Form is subject to the terms of the Mozilla Public
' License, v. 2.0. If a copy of the MPL was not distributed with this
' file, You can obtain one at http://mozilla.org/MPL/2.0/.

Option Explicit
Dim lngReturnVal As Long

Dim rlCurrWord As CreoleWord
Dim lngParamFieldAddr As Long

Public Property Get CurrWord() As CreoleWord
    Set CurrWord = rlCurrWord
End Property

Public Property Set CurrWord(Value As CreoleWord)
    Set rlCurrWord = Value
End Property

Public Property Get ParamFieldAddr() As Long
    ParamFieldAddr = lngParamFieldAddr
End Property

Public Property Let ParamFieldAddr(Value As Long)
    lngParamFieldAddr = Value
End Property
