VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CompileInfo"
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

Dim sFQNameField As String
Dim sToken As String
Dim sCompileAction As String

' Colon definitions are built into the PAD area - each new entry
' is a triplet consisting of the word's fully qualified name, its
' tokenization (either dictionary address or literal), and associated
' compilation action.

Public Property Let FQNameField(Value As String)
    sFQNameField = Value
End Property

Public Property Get FQNameField() As String
    FQNameField = sFQNameField
End Property

Public Property Let Token(Value As String)
    sToken = Value
End Property

Public Property Get Token() As String
    Token = sToken
End Property

Public Property Let CompileAction(Value As String)
    sCompileAction = Value
End Property

Public Property Get CompileAction() As String
    CompileAction = sCompileAction
End Property
