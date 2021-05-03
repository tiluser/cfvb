' Creole Forth for VB
' Version 0.01
' Copyright 2021 Joseph M. O'Connor
' This Source Code Form is subject to the terms of the Mozilla Public
' License, v. 2.0. If a copy of the MPL was not distributed with this
' file, You can obtain one at http://mozilla.org/MPL/2.0/.

Option Explicit
Dim lngReturnVal As Long

' ( -- ) Do-nothing primitive which is surprisingly useful
Function DoNOP(ByRef poGSP As GlobalSimpleProps)
    DoNOP = 0
End Function

' ( n1 n2 -- sum ) Adds two numbers on the stack
Function DoPlus(ByRef poGSP As GlobalSimpleProps)
    Dim dblVal1 As Double
    Dim dblVal2 As Double
    Dim dblSum  As Double
    
    Call poGSP.Pop(poGSP.DataStack)
    dblVal1 = CDbl(poGSP.Scratch)
    Call poGSP.Pop(poGSP.DataStack)
    dblVal2 = CDbl(poGSP.Scratch)
    dblSum = dblVal1 + dblVal2
    poGSP.Scratch = dblSum
    Call poGSP.Push(poGSP.DataStack)
    
    DoPlus = 0
End Function

' ( n1 n2 -- difference ) Subtracts two numbers on the stack
Function DoMinus(ByRef poGSP As GlobalSimpleProps)
    Dim dblVal1 As Double
    Dim dblVal2 As Double
    Dim dblDifference  As Double
    
    Call poGSP.Pop(poGSP.DataStack)
    dblVal2 = CDbl(poGSP.Scratch)
    Call poGSP.Pop(poGSP.DataStack)
    dblVal1 = CDbl(poGSP.Scratch)
    dblDifference = dblVal1 - dblVal2
    poGSP.Scratch = dblDifference
    Call poGSP.Push(poGSP.DataStack)
    
    DoMinus = 0
End Function

' ( n1 n2 -- product ) Multiplies two numbers on the stack
Function DoMultiply(ByRef poGSP As GlobalSimpleProps)
    Dim dblVal1 As Double
    Dim dblVal2 As Double
    Dim dblProduct  As Double
    
    Call poGSP.Pop(poGSP.DataStack)
    dblVal1 = CDbl(poGSP.Scratch)
    Call poGSP.Pop(poGSP.DataStack)
    dblVal2 = CDbl(poGSP.Scratch)
    dblProduct = dblVal1 * dblVal2
    poGSP.Scratch = dblProduct
    Call poGSP.Push(poGSP.DataStack)
    DoMultiply = 0
End Function

' ( n1 n2 -- quotient ) Divides two numbers on the stack
Function DoDivide(ByRef poGSP As GlobalSimpleProps)
    Dim dblVal1 As Double
    Dim dblVal2 As Double
    Dim dblQuotient  As Double
    
    Call poGSP.Pop(poGSP.DataStack)
    dblVal2 = CDbl(poGSP.Scratch)
    Call poGSP.Pop(poGSP.DataStack)
    dblVal1 = CDbl(poGSP.Scratch)
    dblQuotient = dblVal1 / dblVal2
    poGSP.Scratch = dblQuotient
    Call poGSP.Push(poGSP.DataStack)
    DoDivide = 0
End Function

' ( n1 n2 -- remainder ) Returns remainder of division operation
Function DoMod(ByRef poGSP As GlobalSimpleProps)
    Dim dblVal1 As Double
    Dim dblVal2 As Double
    Dim dblRemainder  As Double
    
    Call poGSP.Pop(poGSP.DataStack)
    dblVal2 = CDbl(poGSP.Scratch)
    Call poGSP.Pop(poGSP.DataStack)
    dblVal1 = CDbl(poGSP.Scratch)
    dblRemainder = dblVal1 Mod dblVal2
    poGSP.Scratch = dblRemainder
    Call poGSP.Push(poGSP.DataStack)
    
    DoMod = 0
End Function

' ( val --  val val ) Duplicates the argument on top of the stack
Function DoDup(ByRef poGSP As GlobalSimpleProps)
    Call poGSP.Pop(poGSP.DataStack)
    Call poGSP.Push(poGSP.DataStack)
    Call poGSP.Push(poGSP.DataStack)
    
    DoDup = 0
End Function

' ( val1 val2 -- val2 val1 ) Swaps the positions of the top two stack arguments
Function DoSwap(ByRef poGSP As GlobalSimpleProps)
    Dim val1 As Variant
    Dim val2 As Variant
    
    Call poGSP.Pop(poGSP.DataStack)
    val2 = poGSP.Scratch
    Call poGSP.Pop(poGSP.DataStack)
    val1 = poGSP.Scratch
    Call poGSP.Push(poGSP.DataStack)
    poGSP.Scratch = val2
    Call poGSP.Push(poGSP.DataStack)
    
    DoSwap = 0
End Function

' ( val1 val2 val3 -- val2 val3 val1 ) Moves the third stack argument to the top
Function DoRot(ByRef poGSP As GlobalSimpleProps)
    Dim val1 As Variant
    Dim val2 As Variant
    Dim val3 As Variant
    
    Call poGSP.Pop(poGSP.DataStack)
    val3 = poGSP.Scratch
    Call poGSP.Pop(poGSP.DataStack)
    val2 = poGSP.Scratch
    Call poGSP.Pop(poGSP.DataStack)
    val1 = poGSP.Scratch

    poGSP.Scratch = val2
    Call poGSP.Push(poGSP.DataStack)
    poGSP.Scratch = val3
    Call poGSP.Push(poGSP.DataStack)
    poGSP.Scratch = val1
    Call poGSP.Push(poGSP.DataStack)
    
    DoRot = 0
End Function

' ( val1 val2 val3 -- val3 val1 val2 ) Moves the top stack argument to the third position
Function DoMinusRot(ByRef poGSP As GlobalSimpleProps)
    Dim val1 As Variant
    Dim val2 As Variant
    Dim val3 As Variant
    
    Call poGSP.Pop(poGSP.DataStack)
    val3 = poGSP.Scratch
    Call poGSP.Pop(poGSP.DataStack)
    val2 = poGSP.Scratch
    Call poGSP.Pop(poGSP.DataStack)
    val1 = poGSP.Scratch

    poGSP.Scratch = val3
    Call poGSP.Push(poGSP.DataStack)
    poGSP.Scratch = val1
    Call poGSP.Push(poGSP.DataStack)
    poGSP.Scratch = val2
    Call poGSP.Push(poGSP.DataStack)
    
    DoMinusRot = 0
End Function

' ( val1 val2 -- val2 ) Removes second stack argument
Function DoNip(ByRef poGSP As GlobalSimpleProps)
    Dim val1 As Variant
    
    Call poGSP.Pop(poGSP.DataStack)
    val1 = poGSP.Scratch
    Call poGSP.Pop(poGSP.DataStack)
    poGSP.Scratch = val1
    Call poGSP.Push(poGSP.DataStack)
    
    DoNip = 0
End Function

' ( val1 val2 -- val2 val1 val2 ) Copies top stack argument under second argument
Function DoTuck(ByRef poGSP As GlobalSimpleProps)
    Dim val1 As Variant
    Dim val2 As Variant
    
    Call poGSP.Pop(poGSP.DataStack)
    val2 = poGSP.Scratch
    Call poGSP.Pop(poGSP.DataStack)
    val1 = poGSP.Scratch
    poGSP.Scratch = val2
    Call poGSP.Push(poGSP.DataStack)
    poGSP.Scratch = val1
    Call poGSP.Push(poGSP.DataStack)
    poGSP.Scratch = val2
    Call poGSP.Push(poGSP.DataStack)
    
    DoTuck = 0
End Function

' ( val1 val2 -- val1 val2 val1 ) Copies second stack argument to the top of the stack
Function DoOver(ByRef poGSP As GlobalSimpleProps)
    Dim val1 As Variant
    Dim val2 As Variant
    
    Call poGSP.Pop(poGSP.DataStack)
    val2 = poGSP.Scratch
    Call poGSP.Pop(poGSP.DataStack)
    val1 = poGSP.Scratch
    poGSP.Scratch = val1
    Call poGSP.Push(poGSP.DataStack)
    poGSP.Scratch = val2
    Call poGSP.Push(poGSP.DataStack)
    poGSP.Scratch = val1
    Call poGSP.Push(poGSP.DataStack)
    
    DoOver = 0
End Function

' ( val -- ) Drops the argument at the top of the stack
Function DoDrop(ByRef poGSP As GlobalSimpleProps)
    Call poGSP.Pop(poGSP.DataStack)
    DoDrop = 0
End Function

' ( val -- ) Prints the argument at the top of the stack
Function DoDot(ByRef poGSP As GlobalSimpleProps)
    Call poGSP.Pop(poGSP.DataStack)
    MsgBox poGSP.Scratch
    DoDot = 0
End Function

' ( -- ) Nondestructive data stack print - rightmost is tos
Function DoDotDataStack(ByVal poGSP As GlobalSimpleProps)
    Dim lngNewAddress As Long
    Dim cw As CreoleWord
    Dim lngDoLitAddr As Long
    Dim litVal As Variant
    Dim i As Long
    Dim sVals As String
    
    For i = 0 To poGSP.DataStack.Count - 1
        sVals = sVals & " " & poGSP.DataStack.Items(i)
    Next i
    Debug.Print sVals
    
    DoDotDataStack = 0
End Function

' ( -- n ) Returns the stack depth
Function DoDepth(ByRef poGSP As GlobalSimpleProps)
    poGSP.Scratch = poGSP.DataStack.Count
    Call poGSP.Push(poGSP.DataStack)
    DoDepth = 0
End Function

' ( -- ) Pops up a Hello message box
Function DoHello(ByRef poGSP As GlobalSimpleProps)
    MsgBox "Hello"
    DoHello = 0
End Function

' ( -- ) Pops up a Tulip message box
Function DoTulip(ByRef poGSP As GlobalSimpleProps)
    MsgBox "Tulip"
    DoTulip = 0
End Function

' ( msg -- ) Pops up an alert saying the message
Function DoMsgBox(ByRef poGSP As GlobalSimpleProps)
    Call poGSP.Pop(poGSP.DataStack)
    MsgBox CStr(poGSP.Scratch)
    DoMsgBox = 0
End Function

' ( -- ) Pops up an alert saying the date and time
Function DoNow(ByRef poGSP As GlobalSimpleProps)
    MsgBox Now
    DoNow = 0
End Function

' ( -- ) Lists the dictionary definitions
Function DoVlist(ByRef poGSP As GlobalSimpleProps)
    Dim i As Long
    Dim j As Long
    Dim sList As String
    Dim sVals As String
    
    For i = 0 To poGSP.cfb.address.Count - 1
        sVals = ""
        For j = 0 To poGSP.cfb.address.Items(i).ParamField.Count - 1
            sVals = sVals & " " & poGSP.cfb.address.Items(i).ParamField.Items(j)
        Next j
        sVals = Trim(sVals)
        Debug.Print poGSP.cfb.address.Items(i).NameField & "," & _
        poGSP.cfb.address.Items(i).FQNameField & "," & _
        poGSP.cfb.address.Items(i).Vocabulary & "," & _
        poGSP.cfb.address.Items(i).CompileActionField & "," & _
        poGSP.cfb.address.Items(i).IndexField & "," & _
        sVals
    Next i

    DoVlist = 0
End Function










