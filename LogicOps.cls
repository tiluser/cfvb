' Creole Forth for VB
' Version 0.01
' Copyright 2021 Joseph M. O'Connor
' This Source Code Form is subject to the terms of the Mozilla Public
' License, v. 2.0. If a copy of the MPL was not distributed with this
' file, You can obtain one at http://mozilla.org/MPL/2.0/.

Option Explicit
Dim iReturnVal As Integer

Function DoEquals(poGSP As GlobalSimpleProps)
    Dim val1 As Long
    Dim val2 As Long
    
    Call poGSP.Pop(poGSP.DataStack)
    val1 = poGSP.Scratch
    Call poGSP.Pop(poGSP.DataStack)
    val2 = poGSP.Scratch
    If val1 = val2 Then
        poGSP.Scratch = -1
    Else
        poGSP.Scratch = 0
    End If
    Call poGSP.Push(poGSP.DataStack)
    DoEquals = 0
End Function

Function DoNotEquals(poGSP As GlobalSimpleProps)
    Dim val1 As Long
    Dim val2 As Long
    
    Call poGSP.Pop(poGSP.DataStack)
    val1 = poGSP.Scratch
    Call poGSP.Pop(poGSP.DataStack)
    val2 = poGSP.Scratch
    If val1 <> val2 Then
        poGSP.Scratch = -1
    Else
        poGSP.Scratch = 0
    End If
    Call poGSP.Push(poGSP.DataStack)
    DoNotEquals = 0
End Function

Function DoLessThan(poGSP As GlobalSimpleProps)
    Dim val1 As Long
    Dim val2 As Long
    
    Call poGSP.Pop(poGSP.DataStack)
    val1 = poGSP.Scratch
    Call poGSP.Pop(poGSP.DataStack)
    val2 = poGSP.Scratch
    If val2 < val1 Then
        poGSP.Scratch = -1
    Else
        poGSP.Scratch = 0
    End If
    Call poGSP.Push(poGSP.DataStack)
    DoLessThan = 0
End Function

Function DoGreaterThan(poGSP As GlobalSimpleProps)
    Dim val1 As Long
    Dim val2 As Long
    
    Call poGSP.Pop(poGSP.DataStack)
    val1 = poGSP.Scratch
    Call poGSP.Pop(poGSP.DataStack)
    val2 = poGSP.Scratch
    If val2 > val1 Then
        poGSP.Scratch = -1
    Else
        poGSP.Scratch = 0
    End If
    Call poGSP.Push(poGSP.DataStack)
    DoGreaterThan = 0
End Function

Function DoLessThanOrEquals(poGSP As GlobalSimpleProps)
    Dim val1 As Long
    Dim val2 As Long
    
    Call poGSP.Pop(poGSP.DataStack)
    val1 = poGSP.Scratch
    Call poGSP.Pop(poGSP.DataStack)
    val2 = poGSP.Scratch
    If val2 <= val1 Then
        poGSP.Scratch = -1
    Else
        poGSP.Scratch = 0
    End If
    Call poGSP.Push(poGSP.DataStack)
    DoLessThanOrEquals = 0
End Function

Function DoGreaterThanOrEquals(poGSP As GlobalSimpleProps)
    Dim val1 As Long
    Dim val2 As Long
    
    Call poGSP.Pop(poGSP.DataStack)
    val1 = poGSP.Scratch
    Call poGSP.Pop(poGSP.DataStack)
    val2 = poGSP.Scratch
    If val2 >= val1 Then
        poGSP.Scratch = -1
    Else
        poGSP.Scratch = 0
    End If
    Call poGSP.Push(poGSP.DataStack)
    DoGreaterThan = 0
End Function

Function DoNot(poGSP As GlobalSimpleProps)
    Dim val1 As Long
    
    Call poGSP.Pop(poGSP.DataStack)
    val1 = poGSP.Scratch
    If val1 = 0 Then
        poGSP.Scratch = -1
    Else
        poGSP.Scratch = 0
    End If
    Call poGSP.Push(poGSP.DataStack)
    DoNot = 0
End Function

Function DoAnd(poGSP As GlobalSimpleProps)
    Dim val1 As Long
    Dim val2 As Long
    
    Call poGSP.Pop(poGSP.DataStack)
    val1 = poGSP.Scratch
    Call poGSP.Pop(poGSP.DataStack)
    val2 = poGSP.Scratch
    Call poGSP.Push(poGSP.DataStack)
    If val1 <> 0 And val2 <> 0 Then
        poGSP.Scratch = -1
    Else
        poGSP.Scratch = 0
    End If
    DoAnd = 0
End Function

Function DoOr(poGSP As GlobalSimpleProps)
    Dim val1 As Long
    Dim val2 As Long
    
    Call poGSP.Pop(poGSP.DataStack)
    val1 = poGSP.Scratch
    Call poGSP.Pop(poGSP.DataStack)
    val2 = poGSP.Scratch
    Call poGSP.Push(poGSP.DataStack)
    If val1 <> 0 Or val2 <> 0 Then
        poGSP.Scratch = -1
    Else
        poGSP.Scratch = 0
    End If
    DoOr = 0
End Function

Function DoXor(poGSP As GlobalSimpleProps)
    Dim val1 As Long
    Dim val2 As Long
    
    Call poGSP.Pop(poGSP.DataStack)
    val1 = poGSP.Scratch
    Call poGSP.Pop(poGSP.DataStack)
    val2 = poGSP.Scratch
    Call poGSP.Push(poGSP.DataStack)
    If (val1 <> 0 Or val2 <> 0) And Not (val1 = 0 And val2 = 0) Then
        poGSP.Scratch = -1
    Else
        poGSP.Scratch = 0
    End If
    DoXor = 0
End Function

