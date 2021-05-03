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
