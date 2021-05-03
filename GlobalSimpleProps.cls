' Creole Forth for VB
' Version 0.01
' Copyright 2021 Joseph M. O'Connor
' This Source Code Form is subject to the terms of the Mozilla Public
' License, v. 2.0. If a copy of the MPL was not distributed with this
' file, You can obtain one at http://mozilla.org/MPL/2.0/.

Dim gspDataStack As New ArrayList
Dim gspReturnStack As New ArrayList
Dim gspVocabStack As New ArrayList
Dim gspPADArea As New ArrayList
Dim gspCompiledList As ArrayList
Dim gspParsedInput As New ArrayList
Dim gspCfb As New CreoleForthBundle
Public objScratch As Object
Dim objDataStack As New ArrayList
Public Scratch As Variant
Dim sInputArea As String
Dim sOutputArea As String
Dim gspCurrWord As CreoleWord
Dim sContextVocab As String
Dim lngExecPtr As Long
Dim lngReturnVal As Long
Dim lngOuterPtr As Long
Dim lngParamFieldPtr As Long
Dim lngLoopLabelPtr As Long
Dim saLoopLabels(3) As String
Dim lngPADAreaPtr As Long
Dim laLoopCurrIndexes(3) As Long
Dim sHelpCommentField As String

Sub Push(ByVal pStack As ArrayList)
    Call pStack.Add(Scratch)
End Sub

Sub Pop(ByVal pStack As ArrayList)
    Dim lngEndIndex As Long
    If pStack.Count = 0 Then
        MsgBox "Stack underflow"
        Call CleanFields
        GoTo Done
    End If
    lngEndIndex = pStack.Count() - 1

    If IsObject(pStack.Items(lngEndIndex)) Then
        Set Scratch = pStack.Items(lngEndIndex)
    Else
        Scratch = pStack.Items(lngEndIndex)
    End If
    Call pStack.Remove(lngEndIndex)
Done:
End Sub

Sub Unshift(ByVal pStack As ArrayList)
    lngReturnVal = pStack.Insert(0, Scratch)
End Sub

Sub Shift(ByVal pStack As ArrayList)
    Scratch = pStack.Items(0)
    Call pStack.Remove(0)
End Sub

Sub CleanFields()
    Call gspDataStack.Clear
    Call gspReturnStack.Clear
    Call gspPADArea.Clear
    Call gspParsedInput.Clear
    iLoopLabelPtr = 0
    laLoopCurrIndexes(0) = 0
    laLoopCurrIndexes(1) = 0
    laLoopCurrIndexes(2) = 0
    iOuterPtr = 0
    iExecPtr = 0
    iParamFieldPtr = 0
    sInputArea = ""
    sOutputArea = ""
    sHelpCommentField = ""
End Sub

Public Property Set ParsedInput(ByRef Value As ArrayList)
    Set gspParsedInput = Value
End Property

Public Property Get ParsedInput() As ArrayList
    Set ParsedInput = gspParsedInput
End Property

Public Property Set DataStack(ByRef Value As ArrayList)
    Set gspDataStack = Value
End Property

Public Property Get DataStack() As ArrayList
    Set DataStack = gspDataStack
End Property

Public Property Set ReturnStack(ByRef Value As ArrayList)
    Set gspReturnStack = Value
End Property

Public Property Get ReturnStack() As ArrayList
    Set ReturnStack = gspReturnStack
End Property

Public Property Set VocabStack(ByRef Value As ArrayList)
    Set gspVocabStack = Value
End Property

Public Property Get VocabStack() As ArrayList
    Set VocabStack = gspVocabStack
End Property

Public Property Set PADArea(ByRef Value As ArrayList)
    Set gspPADArea = Value
End Property

Public Property Get PADArea() As ArrayList
    Set PADArea = gspPADArea
End Property

Public Property Set CompiledList(ByRef Value As ArrayList)
    Set gspCompiledList = Value
End Property

Public Property Get CompiledList() As ArrayList
    Set CompiledList = gspCompiledList
End Property

Public Property Get OuterPtr() As Long
    OuterPtr = lngOuterPtr
End Property

Public Property Let OuterPtr(Value As Long)
    lngOuterPtr = Value
End Property

Public Property Get InputArea() As String
    InputArea = sInputArea
End Property

Public Property Let InputArea(Value As String)
    sInputArea = Value
End Property

Public Property Let OutputArea(Value As String)
    sOutputArea = Value
End Property

Public Property Get OutputArea() As String
    OutputArea = sOutputArea
End Property

Public Property Get CurrWord() As CreoleWord
    Set CurrWord = gspCurrWord
End Property

Public Property Set CurrWord(ByRef Value As CreoleWord)
    Set gspCurrWord = Value
End Property

Public Property Get ExecPtr() As Long
    ExecPtr = lngExecPtr
End Property

Public Property Let ExecPtr(Value As Long)
    lngExecPtr = Value
End Property

Public Property Get ParamFieldPtr() As Long
    ParamFieldPtr = lngParamFieldPtr
End Property

Public Property Let ParamFieldPtr(Value As Long)
    lngParamFieldPtr = Value
End Property

Public Property Get LoopLabelPtr() As Long
    LoopLabelPtr = lngLoopLabelPtr
End Property

Public Property Get LoopLabels(Index As Long) As String
    saLoopLabels(0) = "I"
    saLoopLabels(1) = "J"
    saLoopLabels(2) = "K"
    LoopLabels = saLoopLabels(Index)
End Property

Public Property Get LoopCurrIndexes(Index As Long) As Long
    LoopCurrIndexes = laLoopCurrIndexes(Index)
End Property

Public Property Let LoopCurrIndexes(Index As Long, Value As Long)
    laLoopCurrIndexes(Index) = Value
End Property

Public Property Get Vec(Index As Long) As Double
    Vec = v(Index)
End Property
Public Property Let Vec(Index As Long, MyValue As Double)
    v(Index) = MyValue
End Property

Public Property Let PADAreaPtr(Value As Long)
    lngPADAreaPtr = Value
End Property

Public Property Get PADAreaPtr() As Long
    PADAreaPtr = lngPADAreaPtr
End Property

Public Property Let LoopLabelPtr(Value As Long)
    lngLoopLabelPtr = Value
End Property

Public Property Let HelpCommentField(Value As String)
    sHelpCommentField = Value
End Property

Public Property Get HelpCommentField() As String
    HelpCommentField = sHelpCommentField
End Property

Public Property Get cfb() As CreoleForthBundle
    Set cfb = gspCfb
End Property

Public Property Set cfb(ByRef Value As CreoleForthBundle)
    Set gspCfb = Value
End Property





