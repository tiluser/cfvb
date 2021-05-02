VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AppSpec"
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
Dim iReturnVal As Integer

' ( -- ) Do whatever you want here
Function DoTest(poGSP As GlobalSimpleProps)
    MsgBox "Testing"
    DoTest = 0
End Function
