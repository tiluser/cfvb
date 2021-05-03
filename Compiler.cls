' Creole Forth for VB
' Version 0.01
' Copyright 2021 Joseph M. O'Connor
' This Source Code Form is subject to the terms of the Mozilla Public
' License, v. 2.0. If a copy of the MPL was not distributed with this
' file, You can obtain one at http://mozilla.org/MPL/2.0/.

Option Explicit
Dim lngReturnVal As Long

' ( n -- ) Compiles value off the TOS into the next parameter field cell
Function DoComma(ByRef poGSP As GlobalSimpleProps)
    Dim lngNewRow As Long
    Dim lngToken As Long
    Dim cwNewCreoleWord As CreoleWord
      
    lngNewRow = poGSP.cfb.address.Count - 1
    Call poGSP.Pop(poGSP.DataStack)
    lngToken = poGSP.Scratch
    Set cwNewCreoleWord = poGSP.cfb.address.Items(lngNewRow)
    Call cwNewCreoleWord.ParamField.Add(lngToken)
    cwNewCreoleWord.ParamFieldStart = cwNewCreoleWord.ParamFieldStart + 1
    poGSP.ParamFieldPtr = cwNewCreoleWord.ParamField.Count - 1

    DoComma = 0
End Function
   
' ( address -- ) Executes the word corresponding to the address on the stack
Function DoExecute(ByRef poGSP As GlobalSimpleProps)
    Dim objClassModule As Object
    
    Call poGSP.Pop(poGSP.DataStack)
    poGSP.ExecPtr = poGSP.Scratch
    Set poGSP.CurrWord = poGSP.cfb.address.Items(poGSP.ExecPtr)
    Set objClassModule = poGSP.cfb.ClassModules.Item(poGSP.CurrWord.ClassName)
    CallByName objClassModule, poGSP.CurrWord.CodeField, VbMethod, poGSP

    DoExecute = 0
End Function

' ( -- location ) Returns address of the next available dictionary location
Function DoHere(ByRef poGSP As GlobalSimpleProps)
    Dim iHereLoc As Long
    
    iHereLoc = poGSP.cfb.address.Count
    Call poGSP.DataStack.Add(iHereLoc)
    
    DoHere = 0
End Function

' Used internally by doCreate - is not compiled into the dictionary
Function DoMyAddress(ByRef poGSP As GlobalSimpleProps)
    Dim cw As CreoleWord
    
    Set cw = poGSP.cfb.address.Items(poGSP.ExecPtr)
    Call poGSP.DataStack.Add(cw.IndexField)
    
    DoMyAddress = 0
End Function

' CREATE <name>. Adds a named entry into the dictionary
Function DoCreate(ByRef poGSP As GlobalSimpleProps)
    Dim lngHereLoc As Long
    Dim sName As String
    Dim alParams As ArrayList
    Dim alData As ArrayList
    Dim cw As CreoleWord
    
    lngHereLoc = poGSP.cfb.address.Count
    sName = poGSP.ParsedInput.Items(poGSP.OuterPtr + 1)
    Set alParams = New ArrayList
    Set alData = New ArrayList
    Set cw = New CreoleWord
    cw.NameField = sName
    cw.Vocabulary = poGSP.cfb.CurrentVocab
    cw.ClassName = "Compiler"
    cw.CodeField = "DoMyAddress"
    cw.CompileActionField = "COMPINPF"
    cw.HelpField = "TODO: "
    ' TODO: Code for redefining
    cw.PrevRowLocField = lngHereLoc
    cw.RowLocField = lngHereLoc
    cw.LinkField = lngHereLoc - 1
    cw.IndexField = lngHereLoc
    Set cw.ParamField = alParams
    Set cw.DataField = alData
    poGSP.cfb.Dict.Add key:=cw.FQNameField, Item:=cw
    Call poGSP.cfb.address.Add(cw)
    poGSP.OuterPtr = poGSP.OuterPtr + 2
    DoCreate = 0
End Function

' Looks up a word and returns its compilation info
Function ExtractCompileInfo(psRawWord As String, pDict As Dictionary, _
                            pVocabStack As ArrayList) As CompileInfo
    Dim sReturnedVal As String
    Dim sFullyQualifiedWord As String
    Dim key As Variant
    Dim i As Long
    Dim cmpCompInfo As New CompileInfo
    Dim cw As CreoleWord
    
    sReturnedVal = psRawWord
    
    For i = 0 To pVocabStack.Count - 1
        sFullyQualifiedWord = psRawWord & "." & pVocabStack.Items(i)
        If pDict.Exists(sFullyQualifiedWord) Then
            sReturnedVal = sFullyQualifiedWord
        End If
    Next i
    cmpCompInfo.FQNameField = sReturnedVal
    If (sReturnedVal = psRawWord) Then
        cmpCompInfo.CompileAction = "COMPLIT"
        cmpCompInfo.Token = sReturnedVal
    Else
        Set cw = pDict.Item(sReturnedVal)
        cmpCompInfo.CompileAction = cw.CompileActionField
        cmpCompInfo.Token = CStr(cw.IndexField)
    End If
    
    Set ExtractCompileInfo = cmpCompInfo
End Function

Function GetSemiIndex(pParsedInput As ArrayList) As Long
    Dim i As Long
    Dim lngSemiIndex As Long
    i = 0
    lngSemiIndex = -1
    While (i < pParsedInput.Count)
        If pParsedInput.Items(i) = ";" Then
            lngSemiIndex = i
        End If
        i = i + 1
    Wend
    GetSemiIndex = lngSemiIndex
End Function

' Almost like BuildPrimitive except it doesn't populate the dictionary
Function StartHighLvlDef(ByRef cfb As CreoleForthBundle, psName As String, _
    psVocab As String, psCompAction As String, psHelp As String) As CreoleWord
    
    Dim cw As New CreoleWord
    Dim params As New ArrayList
    Dim dataVals As New ArrayList
  
    cw.NameField = psName
    cw.ClassName = "Interpreter"
    cw.CodeField = "DoColon"
    cw.Vocabulary = psVocab
    cw.CompileActionField = psCompAction
    cw.HelpField = psHelp
    cw.PrevRowLocField = cfb.address.Count - 1
    cw.RowLocField = cfb.address.Count
    cw.LinkField = cfb.address.Count - 1
    cw.IndexField = cfb.address.Count
    Set cw.ParamField = params
    cw.ParamField.Clear
    cw.ParamFieldStart = 0
    cw.ParamFieldStartFrom = 0
    Set cw.DataField = dataVals
    cw.DataField.Clear
    Set StartHighLvlDef = cw
   
End Function

Function DoInitColonDefGsp(poGSPSrc As GlobalSimpleProps) As GlobalSimpleProps
    Dim gspTarget As GlobalSimpleProps
    Dim i As Long
    Dim k As Variant
    
    Set gspTarget = New GlobalSimpleProps
    For i = 0 To poGSPSrc.VocabStack.Count - 1
        Call gspTarget.VocabStack.Add(poGSPSrc.VocabStack.Items(i))
    Next i
    
    For Each k In poGSPSrc.cfb.Dict.Keys
        gspTarget.cfb.Dict.Add key:=k, Item:=poGSPSrc.cfb.Dict(k)
    Next
    For i = 0 To poGSPSrc.cfb.address.Count - 1
        Call gspTarget.cfb.address.Add(poGSPSrc.cfb.address.Items(i))
    Next i
    Call gspTarget.CleanFields
    Dim coreprimsmodule As New coreprims
    Dim interpreterModule As New interpreter
    Dim compilerModule As New Compiler
    Dim logicOpsModule As New LogicOps
    Dim appSpecModule As New AppSpec
   
    gspTarget.cfb.ClassModules.Add key:="CorePrims", Item:=coreprimsmodule
    gspTarget.cfb.ClassModules.Add key:="Interpreter", Item:=interpreterModule
    gspTarget.cfb.ClassModules.Add key:="Compiler", Item:=compilerModule
    gspTarget.cfb.ClassModules.Add key:="LogicOps", Item:=logicOpsModule
    gspTarget.cfb.ClassModules.Add key:="AppSpec", Item:=appSpecModule
   
    Set DoInitColonDefGsp = gspTarget
End Function
    

' ( -- ) Starts compilation of a colon definition
Function CompileColon(ByRef poGSP As GlobalSimpleProps)
    Dim lngHereLoc As Long
    Dim sName As String
    Dim sRawWord As String
    Dim ciCompInfo As CompileInfo
    Dim gspComp As GlobalSimpleProps
    Dim sCodeField As String
    Dim cw As CreoleWord
    Dim sFqName As String
    Dim sFqNameSmudged As String
    Dim cmpCompileInfo As CompileInfo
    Dim interpLocal As New interpreter
    Dim lngSemiIndex As Long
        
    lngHereLoc = poGSP.cfb.address.Count
    sName = poGSP.ParsedInput.Items(poGSP.OuterPtr + 1)
   
    poGSP.PADArea.Clear
    sFqName = sName & "." & poGSP.cfb.CurrentVocab
    lngSemiIndex = GetSemiIndex(poGSP.ParsedInput)
    If (lngSemiIndex = -1) Then
        MsgBox "Error: colon definition must be terminated with a semicolon"
        Call poGSP.CleanFields
        GoTo Done
    End If
    
    ' Compilation is started when the IMMEDIATE vocabulary is pushed onto the vocabulary stack. No need for the usual Forth STATE flag.
    Call poGSP.VocabStack.Add("IMMEDIATE")
    Set cw = StartHighLvlDef(poGSP.cfb, sName, poGSP.cfb.CurrentVocab, _
                             "COMPINPF", "TODO:")

    ' Smudge the lookup field temporarily to avoid accidental recursion
    cw.FQNameField = cw.NameField & "." & cw.Vocabulary & ".SMUDGED"
    poGSP.cfb.Dict.Add key:=cw.FQNameField, Item:=cw
    Call poGSP.cfb.address.Add(cw)
    poGSP.OuterPtr = poGSP.OuterPtr + 2
    ' Parameter field contents are set up in the PAD area. Each word is looked up one at a time in the dictionary, and its name, address, and
    ' compilation action are placed in the CompileInfo triplet.
    While (poGSP.OuterPtr <= lngSemiIndex)
        sRawWord = poGSP.ParsedInput.Items(poGSP.OuterPtr)
        Set cmpCompileInfo = ExtractCompileInfo(sRawWord, poGSP.cfb.Dict, poGSP.VocabStack)
        Call poGSP.PADArea.Add(cmpCompileInfo)
        poGSP.OuterPtr = poGSP.OuterPtr + 1
    Wend
    poGSP.PADAreaPtr = 0
    
    Set gspComp = DoInitColonDefGsp(poGSP)

    While (poGSP.PADAreaPtr < poGSP.PADArea.Count)
        Set cmpCompileInfo = poGSP.PADArea.Items(poGSP.PADAreaPtr)
        gspComp.InputArea = gspComp.InputArea & " " & cmpCompileInfo.Token & _
            " " & cmpCompileInfo.CompileAction
        poGSP.PADAreaPtr = poGSP.PADAreaPtr + 1
    Wend

    lngReturnVal = interpLocal.DoOuter(gspComp)
    Set cw = gspComp.cfb.address.Items(lngHereLoc)
    cw.ParamFieldStart = 0
    poGSP.cfb.Dict.Remove key:=cw.FQNameField
    cw.FQNameField = Replace(cw.FQNameField, ".SMUDGED", "")
    poGSP.cfb.Dict.Add key:=cw.FQNameField, Item:=cw
    poGSP.cfb.address.Items(poGSP.cfb.address.Count - 1).FQNameField = cw.FQNameField
    Call poGSP.PADArea.Clear
    poGSP.OuterPtr = GetSemiIndex(poGSP.ParsedInput)
    Call poGSP.Pop(poGSP.VocabStack)
Done:
    CompileColon = 0
End Function

' ( -- ) Compiles doLiteral and a literal into the parameter field
Function CompileLiteral(ByVal poGSP As GlobalSimpleProps)
    Dim lngNewAddress As Long
    Dim cw As CreoleWord
    Dim lngDoLitAddr As Long
    Dim litVal As Variant
    
    lngNewAddress = poGSP.cfb.address.Count - 1
    Set cw = poGSP.cfb.address.Items(lngNewAddress)
    lngDoLitAddr = poGSP.cfb.Dict.Item("DoLiteral.IMMEDIATE").IndexField
    Call poGSP.Pop(poGSP.DataStack)
    Call cw.ParamField.Add(lngDoLitAddr)
    Call cw.ParamField.Add(poGSP.Scratch)
    
    Set poGSP.cfb.Dict.Item(cw.FQNameField).ParamField = cw.ParamField
    Set poGSP.cfb.address.Items(lngNewAddress).ParamField = cw.ParamField
       
    poGSP.ParamFieldPtr = cw.ParamField.Count - 1
    
    CompileLiteral = 0
End Function

' ( -- ) Displays the contents of the most recently defined definition.
Function DoDisplayParamField(ByVal poGSP As GlobalSimpleProps)
    Dim lngNewAddress As Long
    Dim cw As CreoleWord
    Dim lngDoLitAddr As Long
    Dim litVal As Variant
    Dim i As Long
    Dim sVals As String
    
    lngNewAddress = poGSP.cfb.address.Count - 1
    Set cw = poGSP.cfb.address.Items(lngNewAddress)
    For i = 0 To cw.ParamField.Count - 1
        sVals = sVals & " " & cw.ParamField.Items(i)
    Next i
    
    DoDisplayParamField = 0
End Function

' ( -- lit ) Run-time code that pushes a literal onto the stack
Function DoLiteral(ByVal poGSP As GlobalSimpleProps)
    Dim rLoc As ReturnLoc
    Dim litVal As Variant
    Dim arParamField As ArrayList
    
    Call poGSP.Pop(poGSP.ReturnStack)
    Set rLoc = poGSP.Scratch
    Set arParamField = rLoc.CurrWord.ParamField
    If IsObject(arParamField.Items(rLoc.ParamFieldAddr)) Then
        Set litVal = arParamField.Items(rLoc.ParamFieldAddr)
    Else
        litVal = arParamField.Items(rLoc.ParamFieldAddr)
    End If
    rLoc.ParamFieldAddr = rLoc.ParamFieldAddr + 1
    Call poGSP.DataStack.Add(litVal)
    Call poGSP.ReturnStack.Add(rLoc)
    
    DoLiteral = 0
End Function

'  ( addr -- val ) Fetches the value in the param field  at addr)
Function DoFetch(ByRef poGSP As GlobalSimpleProps)
    Dim lngAddress As Long
    Call poGSP.Pop(poGSP.DataStack)
    lngAddress = poGSP.Scratch
    Call poGSP.DataStack.Add(poGSP.cfb.address.Items(lngAddress).ParamField.Items(0))
    DoFetch = 0
End Function

' ( val addr -- ) Stores the value in the param field  at addr
Function DoStore(ByRef poGSP As GlobalSimpleProps)
    Dim lngAddress As Long
    Call poGSP.Pop(poGSP.DataStack)
    lngAddress = poGSP.Scratch
    Call poGSP.DataStack.Pop(poGSP.DataStack)
    Call poGSP.cfb.address.Items(lngAddress).ParamField.Items.Clear
    Call poGSP.cfb.address.Items(lngAddress).ParamField.Add(poGSP.Scratch)
    DoStore = 0
End Function

' (  -- ) Sets the current (compilation) vocabulary to the context vocabulary (the one on top of the vocabulary stack)
Function DoSetCurrentToContext(ByRef poGSP As GlobalSimpleProps)
    Dim lngAddress As Long
    poGSP.cfb.CurrentVocab = poGSP.VocabStack.Items(poGSP.VocabStack.Count - 1)
    Debug.Print "Current vocab is now " + poGSP.cfb.CurrentVocab
    DoSetCurrentToContext = 0
End Function

' ( -- ) Flags a word as immediate (executes instead of compiling inside a colon definition)
Function DoImmediate(ByRef poGSP As GlobalSimpleProps)
    Dim cw As CreoleWord
    
    Set cw = poGSP.cfb.address.Items(poGSP.cfb.address.Count - 1)
    poGSP.cfb.Dict.Item(cw.FQNameField).CompileActionField = "EXECUTE"
    poGSP.cfb.Dict.Item(cw.FQNameField).Vocabulary = "IMMEDIATE"
    poGSP.cfb.address.Items(poGSP.cfb.address.Count - 1).CompileActionField = "EXECUTE"
    poGSP.cfb.address.Items(poGSP.cfb.address.Count - 1).Vocabulary = "IMMEDIATE"
    
    Debug.Print cw.FQNameField & " flagged as immediate"
    DoImmediate = 0
End Function

' ( -- location ) Compile-time code for IF
Function CompileIf(ByRef poGSP As GlobalSimpleProps)
    Dim lngNewAddress As Long
    Dim lngZeroBranchAddr As Long
    Dim alParamField As ArrayList
    Dim sFQNameField As String
        
    lngNewAddress = poGSP.cfb.address.Count - 1
    Set alParamField = poGSP.cfb.address.Items(lngNewAddress).ParamField
    sFQNameField = poGSP.cfb.address.Items(lngNewAddress).FQNameField
    lngZeroBranchAddr = poGSP.cfb.Dict.Item("0BRANCH.IMMEDIATE").IndexField
    Call alParamField.Add(lngZeroBranchAddr)
    Call alParamField.Add(-1)

    poGSP.ParamFieldPtr = alParamField.Count - 1
    Call poGSP.DataStack.Add(poGSP.ParamFieldPtr)
    Set poGSP.cfb.address.Items(lngNewAddress).ParamField = alParamField
    Set poGSP.cfb.Dict.Item(sFQNameField).ParamField = alParamField
    
    CompileIf = 0
End Function
        
' ( -- location ) Compile-time code for ELSE
Function CompileElse(ByRef poGSP As GlobalSimpleProps)
    Dim lngNewAddress As Long
    Dim lngJumpAddr As Long
    Dim lngElseAddr As Long
    Dim alParamField As ArrayList
    Dim sFQNameField As String
    Dim lngJumpAddrPFLoc As Long
    Dim lngZeroBrAddrPFLoc As Long
    
    lngNewAddress = poGSP.cfb.address.Count - 1
    lngJumpAddr = poGSP.cfb.Dict.Item("JUMP.IMMEDIATE").IndexField
    lngElseAddr = poGSP.cfb.Dict.Item("DoElse.IMMEDIATE").IndexField
    Set alParamField = poGSP.cfb.address.Items(lngNewAddress).ParamField
    sFQNameField = poGSP.cfb.address.Items(lngNewAddress).FQNameField
    Call alParamField.Add(lngJumpAddr)
    Call alParamField.Add(-1)

    lngJumpAddrPFLoc = alParamField.Count - 1
    Call alParamField.Add(lngElseAddr)
    Call poGSP.Pop(poGSP.DataStack)
    lngZeroBrAddrPFLoc = CLng(poGSP.Scratch)
    Call alParamField.SetItem(lngZeroBrAddrPFLoc, alParamField.Count - 1)
    Call poGSP.DataStack.Add(lngJumpAddrPFLoc)
    poGSP.ParamFieldPtr = poGSP.cfb.address.Items(lngNewAddress).ParamField.Count - 1
    Set poGSP.cfb.address.Items(lngNewAddress).ParamField = alParamField
    Set poGSP.cfb.Dict.Item(sFQNameField).ParamField = alParamField
    
    CompileElse = 0
End Function
           
' ( -- location ) Compile-time code for THEN
Function CompileThen(ByRef poGSP As GlobalSimpleProps)
    Dim lngNewAddress As Long
    Dim lngBranchPFLoc As Long
    Dim lngThenAddr As Long
    Dim alParamField As ArrayList
    Dim sFQNameField As String
    Dim lngJumpAddrPFLoc As Long
    Dim lngZeroBrAddrPFLoc
    
    Call poGSP.Pop(poGSP.DataStack)
    lngBranchPFLoc = poGSP.Scratch
    lngNewAddress = poGSP.cfb.address.Count - 1
    lngThenAddr = poGSP.cfb.Dict.Item("DoThen.IMMEDIATE").IndexField
    Set alParamField = poGSP.cfb.address.Items(lngNewAddress).ParamField
    sFQNameField = poGSP.cfb.address.Items(lngNewAddress).FQNameField
    Call alParamField.Add(lngThenAddr)
    Call alParamField.SetItem(lngBranchPFLoc, alParamField.Count - 1)
    Set poGSP.cfb.address.Items(lngNewAddress).ParamField = alParamField
    Set poGSP.cfb.Dict.Item(sFQNameField).ParamField = alParamField
        
    CompileThen = 0
End Function

' ( flag -- ) Run-time code for IF
Function Do0Branch(ByRef poGSP As GlobalSimpleProps)
    Dim alParamField As ArrayList
    Dim rLoc As ReturnLoc
    Dim lngJumpAddr As Long
    Dim lngBranchFlag As Long
    
    Call poGSP.Pop(poGSP.ReturnStack)
    Set rLoc = poGSP.Scratch
    Set alParamField = rLoc.CurrWord.ParamField
    lngJumpAddr = alParamField.Items(rLoc.ParamFieldAddr)
    Call poGSP.Pop(poGSP.DataStack)
    lngBranchFlag = poGSP.Scratch
    If (lngBranchFlag = 0) Then
        poGSP.ParamFieldPtr = lngJumpAddr
    Else
        poGSP.ParamFieldPtr = poGSP.ParamFieldPtr + 1
    End If
    rLoc.ParamFieldAddr = poGSP.ParamFieldPtr
    Call poGSP.ReturnStack.Add(rLoc)
    
    Do0Branch = 0
End Function

' ( -- ) Jumps unconditionally to the parameter field location next to it and is compiled by ELSE
Function DoJump(ByRef poGSP As GlobalSimpleProps)
    Dim alParamField As ArrayList
    Dim lngJumpAddr As Long
    Dim rLoc As ReturnLoc
    Dim cw As CreoleWord
    
    Call poGSP.Pop(poGSP.ReturnStack)
    Set rLoc = poGSP.Scratch
    Set alParamField = rLoc.CurrWord.ParamField
    lngJumpAddr = alParamField.Items(poGSP.ParamFieldPtr + 1)
    poGSP.ParamFieldPtr = lngJumpAddr
    rLoc.ParamFieldAddr = poGSP.ParamFieldPtr
    Call poGSP.ReturnStack.Add(rLoc)
    
    DoJump = 0
End Function

' ( -- beginLoc ) Compile-time code for BEGIN
Function CompileBegin(ByRef poGSP As GlobalSimpleProps)
    Dim alParamField As ArrayList
    Dim sFQNameField As String
    Dim lngBeginAddr As Long
    Dim lngBeginLoc As Long
    
    Set alParamField = poGSP.cfb.address.Items(poGSP.cfb.address.Count - 1).ParamField
    sFQNameField = poGSP.cfb.address.Items(poGSP.cfb.address.Count - 1).FQNameField
    lngBeginAddr = poGSP.cfb.Dict.Item("DoBegin.IMMEDIATE").IndexField
    Call alParamField.Add(lngBeginAddr)
    lngBeginLoc = alParamField.Count - 1
    Call poGSP.DataStack.Add(lngBeginLoc)
    Set poGSP.cfb.address.Items(poGSP.cfb.address.Count - 1).ParamField = alParamField
    Set poGSP.cfb.Dict.Item(sFQNameField).ParamField = alParamField

    CompileBegin = 0
End Function
   
' ( beginLoc -- ) Compile-time code for UNTIL
Function CompileUntil(ByRef poGSP As GlobalSimpleProps)
    Dim alParamField As ArrayList
    Dim sFQNameField As String
    Dim lngZeroBranchAddr As Long
    Dim lngBeginLoc As Long
    
    Set alParamField = poGSP.cfb.address.Items(poGSP.cfb.address.Count - 1).ParamField
    sFQNameField = poGSP.cfb.address.Items(poGSP.cfb.address.Count - 1).FQNameField
    Call poGSP.Pop(poGSP.DataStack)
    lngBeginLoc = poGSP.Scratch
    lngZeroBranchAddr = poGSP.cfb.Dict.Item("0BRANCH.IMMEDIATE").IndexField
    Call alParamField.Add(lngZeroBranchAddr)
    Call alParamField.Add(lngBeginLoc)
    Set poGSP.cfb.address.Items(poGSP.cfb.address.Count - 1).ParamField = alParamField
    Set poGSP.cfb.Dict.Item(sFQNameField).ParamField = alParamField

    CompileUntil = 0
End Function

' ( -- beginLoc ) Compile-time code for DO
Function CompileDo(ByRef poGSP As GlobalSimpleProps)
    Dim lngNewAddress As Long
    Dim sFQNameField As String
    Dim lngDoStartDoAddr As Long
    Dim lngDoAddr As Long
    Dim lngDoLoc As Long
    Dim alParamField As ArrayList
       
    Set alParamField = poGSP.cfb.address.Items(poGSP.cfb.address.Count - 1).ParamField
    lngNewAddress = poGSP.cfb.address.Count - 1
    sFQNameField = poGSP.cfb.address.Items(lngNewAddress).FQNameField
    lngDoStartDoAddr = poGSP.cfb.Dict.Item("DoStartDo.IMMEDIATE").IndexField
    lngDoAddr = poGSP.cfb.Dict.Item("DoDo.IMMEDIATE").IndexField
       
    Call alParamField.Add(lngDoStartDoAddr)
    Call alParamField.Add(lngDoAddr)
    Set poGSP.cfb.address.Items(poGSP.cfb.address.Count - 1).ParamField = alParamField
    Set poGSP.cfb.Dict.Item(sFQNameField).ParamField = alParamField
    
    lngDoLoc = poGSP.cfb.address.Items(lngNewAddress).ParamField.Count - 1
    Call poGSP.DataStack.Add(lngDoLoc)
    
    CompileDo = 0
End Function

' ( doLoc -- ) Compile-time code for LOOP
Function CompileLoop(ByRef poGSP As GlobalSimpleProps)
    Dim lngNewAddress As Long
    Dim sFQNameField As String
    Dim lngLoopAddr As Long
    Dim lngDoLoc As Long
    Dim alParamField As ArrayList
    
    Set alParamField = poGSP.cfb.address.Items(poGSP.cfb.address.Count - 1).ParamField
    lngNewAddress = poGSP.cfb.address.Count - 1
    sFQNameField = poGSP.cfb.address.Items(lngNewAddress).FQNameField
    lngLoopAddr = poGSP.cfb.Dict.Item("DoLoop.IMMEDIATE").IndexField
    Call poGSP.Pop(poGSP.DataStack)
    lngDoLoc = poGSP.Scratch
    Call alParamField.Add(lngLoopAddr)
    Call alParamField.Add(lngDoLoc)
    Set poGSP.cfb.address.Items(poGSP.cfb.address.Count - 1).ParamField = alParamField
    Set poGSP.cfb.Dict.Item(sFQNameField).ParamField = alParamField
    
    CompileLoop = 0
End Function

' ( doLoc -- ) Compile-time code for +LOOP
Function CompilePlusLoop(ByRef poGSP As GlobalSimpleProps)
    Dim lngNewAddress As Long
    Dim sFQNameField As String
    Dim lngLoopAddr As Long
    Dim lngDoLoc As Long
    Dim alParamField As ArrayList
    
    Set alParamField = poGSP.cfb.address.Items(poGSP.cfb.address.Count - 1).ParamField
    
    lngNewAddress = poGSP.cfb.address.Count - 1
    lngLoopAddr = poGSP.cfb.Dict.Item("DoPlusLoop.IMMEDIATE").IndexField
    Call poGSP.Pop(poGSP.DataStack)
    poGSP.Scratch = lngDoLoc
         
    Call alParamField.Add(lngLoopAddr)
    Call alParamField.Add(lngDoLoc)
    Set poGSP.cfb.address.Items(poGSP.cfb.address.Count - 1).ParamField = alParamField
    Set poGSP.cfb.Dict.Item(sFQNameField).ParamField = alParamField
    
    CompilePlusLoop = 0
End Function

' ( start end -- ) Starts off the Do by getting the start and end
Function DoStartDo(ByRef poGSP As GlobalSimpleProps)
    Dim rLoc As ReturnLoc
    Dim lngStartIndex As Long
    Dim lngLoopEnd As Long
    Dim li As LoopInfo
    
    Call poGSP.Pop(poGSP.ReturnStack)
    Set rLoc = poGSP.Scratch
    Call poGSP.Pop(poGSP.DataStack)
    lngStartIndex = poGSP.Scratch
    Call poGSP.Pop(poGSP.DataStack)
    lngLoopEnd = poGSP.Scratch
    Set li = New LoopInfo
    li.Label = poGSP.LoopLabels(poGSP.LoopLabelPtr)
    li.index = lngStartIndex
    li.Limit = lngLoopEnd
    poGSP.LoopCurrIndexes(0) = 0
    poGSP.LoopCurrIndexes(1) = 0
    poGSP.LoopCurrIndexes(2) = 0
    poGSP.LoopLabelPtr = poGSP.LoopLabelPtr + 1
    Call poGSP.ReturnStack.Add(li)
    Call poGSP.ReturnStack.Add(rLoc)
    Debug.Print "In DoStartDo"
    DoStartDo = 0
End Function

' ( inc -- ) Loops back to DoDo until the start >= the end and increments with inc
Function DoPlusLoop(ByRef poGSP As GlobalSimpleProps)
    Dim lngIncVal As Long
    Dim alParamField As ArrayList
    Dim rLoc As ReturnLoc
    Dim li As LoopInfo
    Dim lngJumpAddr As Long
    Dim lngLoopLimit As Long
    Dim sLoopLabel As String
    Dim lngCurrIndex As Long
    Dim i As Long
    
    Call poGSP.Pop(poGSP.DataStack)
    lngIncVal = poGSP.Scratch
    Call poGSP.Pop(poGSP.ReturnStack)
    Set rLoc = poGSP.Scratch
    Set alParamField = rLoc.CurrWord.ParamField
    Call poGSP.Pop(poGSP.ReturnStack)
    Set li = poGSP.Scratch
    lngJumpAddr = alParamField.Items(rLoc.ParamFieldAddr)
    lngLoopLimit = li.Limit
    sLoopLabel = li.Label
    lngCurrIndex = li.index
    If (lngIncVal < 0) Then
        lngLoopLimit = lngLoopLimit + lngIncVal
    Else
        lngLoopLimit = lngLoopLimit - lngIncVal
    End If
    
    If ((lngIncVal > 0) And (lngCurrIndex >= lngLoopLimit)) Or _
       ((lngIncVal < 0) And (lngCurrIndex <= lngLoopLimit)) Then
       poGSP.ParamFieldPtr = poGSP.ParamFieldPtr + 1
       rLoc.ParamFieldAddr = poGSP.ParamFieldPtr
       poGSP.LoopLabelPtr = poGSP.LoopLabelPtr - 1
    Else
        poGSP.ParamFieldPtr = lngJumpAddr
        lngCurrIndex = lngCurrIndex + lngIncVal
        li.index = lngCurrIndex
        rLoc.ParamFieldAddr = poGSP.ParamFieldPtr
        Call poGSP.ReturnStack.Add(li)
    End If
    
    If (sLoopLabel = "I") Then
        poGSP.LoopCurrIndexes(0) = lngCurrIndex
    ElseIf (sLoopLabel = "J") Then
        poGSP.LoopCurrIndexes(1) = lngCurrIndex
    ElseIf (sLoopLabel = "K") Then
        poGSP.LoopCurrIndexes(2) = lngCurrIndex
    Else
        Debug.Print "Error: invalid loop label"
    End If
    Call poGSP.ReturnStack.Add(rLoc)
    
    DoPlusLoop = 0
End Function

' doLoop is treated as a special case of doPlusLoop
' ( -- ) Loops back to doDo until the start equals the end
Function DoLoop(ByRef poGSP As GlobalSimpleProps)
    Call poGSP.DataStack.Add(1)
    lngReturnVal = DoPlusLoop(poGSP)
    DoLoop = 0
End Function

' ( -- index ) Returns the index of I
Function DoIndexI(ByRef poGSP As GlobalSimpleProps)
    Call poGSP.DataStack.Add(poGSP.LoopCurrIndexes(0))
    DoIndexI = 0
End Function

' ( -- index ) Returns the index of J
Function DoIndexJ(ByRef poGSP As GlobalSimpleProps)
    Call poGSP.DataStack.Add(poGSP.LoopCurrIndexes(1))
    DoIndexJ = 0
End Function

' ( -- index ) Returns the index of K
Function DoIndexK(ByRef poGSP As GlobalSimpleProps)
    Call poGSP.DataStack.Add(poGSP.LoopCurrIndexes(2))
    DoIndexK = 0
End Function

' ( address -- ) Run-time code for DOES>
Function DoDoes(ByRef poGSP As GlobalSimpleProps)
    Dim interpLocal As New interpreter
    Call poGSP.DataStack.Add(poGSP.CurrWord.IndexField)
    lngReturnVal = interpLocal.DoColon(poGSP)
    Set interpLocal = Nothing
    DoDoes = 0
End Function

' # DOES> <list of runtime actions>. When defining word is created, copies code following it into the child definition
Function CompileDoes(ByRef poGSP As GlobalSimpleProps)
    Dim rLoc As ReturnLoc
    Dim lngParentRow As Long
    Dim lngNewRow As Long
    Dim cwParent As CreoleWord
    Dim cwChild As CreoleWord
    Dim sFQNameField As String
    Dim lngDoesAddr As Long
    Dim i As Long
    Dim lngStartCopyPoint As Long
    
    Call poGSP.Pop(poGSP.ReturnStack)
    Set rLoc = poGSP.Scratch
    lngParentRow = rLoc.CurrWord.IndexField
    lngNewRow = poGSP.cfb.address.Count - 1
    Set cwParent = poGSP.cfb.address.Items(lngParentRow)
    Set cwChild = poGSP.cfb.address.Items(lngNewRow)
    sFQNameField = cwChild.FQNameField
    lngDoesAddr = poGSP.cfb.Dict("DOES>.FORTH").IndexField
    i = 0
    cwChild.ClassName = "Compiler"
    cwChild.CodeField = "DoDoes"
    ' Find the location of the does address in the parent definition
    While (i < cwParent.ParamField.Count)
        If (cwParent.ParamField.Items(i) = lngDoesAddr) Then
            lngStartCopyPoint = i + 1
            GoTo Done
        Else
            i = i + 1
        End If
    Wend
Done:
    ' Need the definition's address do doDoes can get it easily either when it's being
    ' called from the interpreter from within a compiled definition
    Call cwChild.ParamField.Add(lngNewRow)
    cwChild.ParamFieldStart = cwChild.ParamField.Count
    i = 0
    While (lngStartCopyPoint < cwParent.ParamField.Count)
        Call cwChild.ParamField.Add(cwParent.ParamField.Items(lngStartCopyPoint))
        lngStartCopyPoint = lngStartCopyPoint + 1
        i = i + 1
    Wend
    rLoc.ParamFieldAddr = rLoc.ParamFieldAddr + i
    Call poGSP.ReturnStack.Add(rLoc)
    Call poGSP.cfb.address.Remove(lngNewRow)
    Call poGSP.cfb.address.Add(cwChild)
    Call poGSP.cfb.Dict.Remove(sFQNameField)
    Call poGSP.cfb.Dict.Add(sFQNameField, cwChild)
    
    CompileDoes = 0
End Function

