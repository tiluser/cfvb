VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CreoleWord"
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

Dim sNameField As String
Dim sClassName As String
Dim sCodeField As String
Dim sVocabulary As String
Dim sFQNameField As String
Dim sCompileActionField As String
Dim sHelpField As String
Dim lngPrevRowLocField As Long
Dim lngRowLocField As Long
Dim lngLinkField As Long
Dim lngIndexField As Long
Dim cwParamField As New ArrayList
Dim cwDataField As New ArrayList
Dim lngParamFieldStart As Long
Dim lngParamFieldStartFrom As Long

Sub init()
    lngParamFieldStart = 0
    lngParamFieldStartFrom = 0
End Sub

Public Property Let NameField(Value As String)
    sNameField = Value
    sFQNameField = Value & "." & sVocabulary
End Property

Public Property Get NameField() As String
    NameField = sNameField
End Property

Public Property Let ClassName(Value As String)
    sClassName = Value
End Property

Public Property Get ClassName() As String
    ClassName = sClassName
End Property

Public Property Let CodeField(Value As String)
    sCodeField = Value
End Property

Public Property Get CodeField() As String
    CodeField = sCodeField
End Property

Public Property Let Vocabulary(Value As String)
    sVocabulary = Value
    sFQNameField = sNameField & "." & Value
End Property

Public Property Get Vocabulary() As String
    Vocabulary = sVocabulary
End Property

Public Property Let FQNameField(Value As String)
    sFQNameField = Value
End Property

Public Property Get FQNameField() As String
    FQNameField = sFQNameField
End Property

Public Property Let CompileActionField(Value As String)
    sCompileActionField = Value
End Property

Public Property Get CompileActionField() As String
    CompileActionField = sCompileActionField
End Property

Public Property Let HelpField(Value As String)
    sHelpField = Value
End Property

Public Property Get HelpField() As String
    HelpField = sHelpField
End Property

' Use to locate previously defined words with same name
Public Property Let PrevRowLocField(Value As Long)
    lngPrevRowLocField = Value
End Property

Public Property Get PrevRowLocField() As Long
    PrevRowLocField = lngPrevRowLocField
End Property

Public Property Let RowLocField(Value As Long)
    lngRowLocField = Value
End Property

Public Property Get RowLocField() As Long
    RowLocField = lngRowLocField
End Property

Public Property Let LinkField(Value As Long)
    lngLinkField = Value
End Property

Public Property Get LinkField() As Long
    LinkField = lngLinkField
End Property

Public Property Let IndexField(Value As Long)
    lngIndexField = Value
End Property

Public Property Get IndexField() As Long
    IndexField = lngIndexField
End Property

Public Property Set ParamField(ByRef Value As ArrayList)
    Set cwParamField = Value
End Property

Public Property Get ParamField() As ArrayList
    Set ParamField = cwParamField
End Property

Public Property Set DataField(ByRef Value As ArrayList)
    Set cwDataField = Value
End Property

Public Property Get DataField() As ArrayList
    Set DataField = cwDataField
End Property

Public Property Let ParamFieldStart(Value As Long)
    lngParamFieldStart = Value
End Property

Public Property Get ParamFieldStart() As Long
    ParamFieldStart = lngParamFieldStart
End Property

Public Property Let ParamFieldStartFrom(Value As Long)
    lngParamFieldStartFrom = Value
End Property

Public Property Get ParamFieldStartFrom() As Long
    ParamFieldStartFrom = lngParamFieldStartFrom
End Property
