' Creole Forth for VB
' Version 0.01
' Copyright 2021 Joseph M. O'Connor
' This Source Code Form is subject to the terms of the Mozilla Public
' License, v. 2.0. If a copy of the MPL was not distributed with this
' file, You can obtain one at http://mozilla.org/MPL/2.0/.

Option Explicit
Dim lngReturnVal As Long

' splits the input into individual words
Function DoParseInput(ByRef poGSP As GlobalSimpleProps)
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim sToParse As String
    Dim saToParse() As String
    sToParse = ""
    Dim s As String
    
    poGSP.OuterPtr = 0
    Call poGSP.ParsedInput.Clear
    saToParse = Split(Trim(poGSP.InputArea))
    For i = 0 To UBound(saToParse)
        If Trim(saToParse(i)) <> "" Then
            Call poGSP.ParsedInput.Add(saToParse(i))
        End If
    Next i
    
    For k = 0 To poGSP.ParsedInput.Count - 1
        s = poGSP.ParsedInput.Items(k)
    Next k
    DoParseInput = 0
End Function

' Returns word in vocabulary closest to the current vocabulary (on top of the stack).
' If search fails, simply returns the word itself
Function LookupWord(psRawWord As String, pDict As Dictionary, _
                            pVocabStack As ArrayList) As String
    Dim sReturnedVal As String
    Dim sFullyQualifiedWord As String
    Dim key As Variant
    Dim i As Long
    
    sReturnedVal = psRawWord
    
    For i = 0 To pVocabStack.Count - 1
        sFullyQualifiedWord = psRawWord & "." & pVocabStack.Items(i)
        If pDict.Exists(sFullyQualifiedWord) Then
            sReturnedVal = sFullyQualifiedWord
        End If
    Next i
   
    LookupWord = sReturnedVal
End Function

Function DoOneOuter(ByRef poGSP As GlobalSimpleProps)
    Dim sVocab As String
    Dim sRawWord As String
    Dim sFoundWord As String
    
    sRawWord = UCase(poGSP.ParsedInput.Items(poGSP.OuterPtr))
    sFoundWord = LookupWord(sRawWord, poGSP.cfb.Dict, poGSP.VocabStack)
    If (sRawWord <> sFoundWord) Then
        poGSP.ExecPtr = poGSP.cfb.Dict(sFoundWord).IndexField
        lngReturnVal = DoRunWord(poGSP)
    Else
        poGSP.Scratch = sRawWord
        Call poGSP.Push(poGSP.DataStack)
    End If
    DoOneOuter = 0
End Function

Function DoOuter(ByRef poGSP As GlobalSimpleProps)
    lngReturnVal = DoParseInput(poGSP)
    While poGSP.OuterPtr < poGSP.ParsedInput.Count
        lngReturnVal = DoOneOuter(poGSP)
        poGSP.OuterPtr = poGSP.OuterPtr + 1
    Wend
    DoOuter = 0
End Function

' Looks up the word based on its list index and executes whatever is in its code field
Function DoRunWord(ByRef poGSP As GlobalSimpleProps)
    Dim objClassModule As Object

    Set poGSP.CurrWord = poGSP.cfb.address.Items(poGSP.ExecPtr)
    Set objClassModule = poGSP.cfb.ClassModules.Item(poGSP.CurrWord.ClassName)
    CallByName objClassModule, poGSP.CurrWord.CodeField, VbMethod, poGSP
    DoRunWord = 0
End Function

' Run-time code for colon definitions
Function DoColon(ByRef poGSP As GlobalSimpleProps)
    Dim lngAddrInPf As Long
    Dim cwCurrContextWord As CreoleWord
    Dim sCodeField As String
    Dim rLoc As ReturnLoc
    Dim objClassModule As Object
    Dim arParamField As ArrayList
    poGSP.ParamFieldPtr = poGSP.CurrWord.ParamFieldStart
    Set arParamField = poGSP.CurrWord.ParamField
    While (poGSP.ParamFieldPtr < arParamField.Count)
        lngAddrInPf = poGSP.CurrWord.ParamField.Items(poGSP.ParamFieldPtr)
        Set cwCurrContextWord = poGSP.CurrWord
        Set poGSP.CurrWord = poGSP.cfb.address.Items(lngAddrInPf)
        sCodeField = poGSP.cfb.address.Items(lngAddrInPf).CodeField
        poGSP.ParamFieldPtr = poGSP.ParamFieldPtr + 1
        Set rLoc = New ReturnLoc
        Set rLoc.CurrWord = cwCurrContextWord
        rLoc.ParamFieldAddr = poGSP.ParamFieldPtr
        Call poGSP.ReturnStack.Add(rLoc)
        Set objClassModule = poGSP.cfb.ClassModules.Item(poGSP.CurrWord.ClassName)
        CallByName objClassModule, poGSP.CurrWord.CodeField, VbMethod, poGSP
        Call poGSP.Pop(poGSP.ReturnStack)
        Set rLoc = poGSP.Scratch
        Set poGSP.CurrWord = rLoc.CurrWord
        poGSP.ParamFieldPtr = rLoc.ParamFieldAddr
    Wend

    DoColon = 0
End Function

' ( -- ) Empties the vocabulary stack, then puts ONLY on it
Function DoOnly(ByRef poGSP As GlobalSimpleProps)
    poGSP.VocabStack.Clear
    lngReturnVal = poGSP.VocabStack.Add("ONLY")
    DoOnly = 0
End Function
 
 ' ( -- ) Puts FORTH on the vocabulary stack
Function DoForth(ByRef poGSP As GlobalSimpleProps)
    lngReturnVal = poGSP.VocabStack.Add("FORTH")
    DoForth = 0
End Function

 ' ( -- ) Puts APPSPEC on the vocabulary stack
Function DoAppSpec(ByRef poGSP As GlobalSimpleProps)
    lngReturnVal = poGSP.VocabStack.Add("APPSPEC")
    DoAppSpec = 0
End Function

' ( -- ) Single-line comment handling
Function DoSingleLineCmts(ByRef poGSP As GlobalSimpleProps)
    While (poGSP.ParsedInput.Items(poGSP.OuterPtr) <> _
        "__#EOL#__")
        poGSP.OuterPtr = Gsp.OuterPtr + 1
    Wend
    DoSingleLineCmts = 0
End Function

' ( -- ) Multiline comment handling
Function DoParenCmts(ByRef poGSP As GlobalSimpleProps)
    Dim sCheckRightParen As String
    sCheckRightParen = _
        Right(poGSP.ParsedInput.Items(poGSP.OuterPtr), 1)
        
    While (sCheckRightParen <> ")")
        poGSP.OuterPtr = poGSP.OuterPtr + 1
        sCheckRightParen = _
            Right(poGSP.ParsedInput.Items(poGSP.OuterPtr), 1)
    Wend
    
    DoParenCmts = 0
End Function

' ( -- list ) List compiler
Function CompileList(ByRef poGSP As GlobalSimpleProps)
    Dim alCompiledList As New ArrayList
    Dim i As Long
    Dim sList As String
    
    Set poGSP.CompiledList = alCompiledList
    poGSP.OuterPtr = poGSP.OuterPtr + 1

    While (poGSP.ParsedInput.Items(poGSP.OuterPtr) <> "}")
        Call Gsp.CompiledList.Add(poGSP.ParsedInput.Items(poGSP.OuterPtr))
        poGSP.OuterPtr = poGSP.OuterPtr + 1
    Wend
    For i = 0 To poGSP.CompiledList.Count - 1
        sList = sList & " " & poGSP.CompiledList.Items(i)
    Next i
    poGSP.Scratch = Trim(sList)
    Call poGSP.Push(poGSP.DataStack)
    
    CompileList = 0
End Function

    

