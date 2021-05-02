VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CreoleForthBundle"
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

Dim cfbGsp As New GlobalSimpleProps
Dim cfbAddress As New ArrayList
Dim cfbDict As New Dictionary
Dim sCurrentVocab As String
Dim dctClassModules As New Dictionary
Dim coreprimsmodule As New coreprims
Dim logicOpsModule As New LogicOps
Dim interpreterModule As New interpreter
Dim compilerModule As New Compiler
Dim appSpecModule As New AppSpec

Public Property Get Gsp() As GlobalSimpleProps
    Set Gsp = cfbGsp
End Property

Public Property Set Gsp(ByRef Value As GlobalSimpleProps)
    Set cfbGsp = Value
End Property

Public Property Get address() As ArrayList
    Set address = cfbAddress
End Property

Public Property Set address(ByRef Value As ArrayList)
    Set cfbAddress = Value
End Property

Public Property Get Dict() As Dictionary
    Set Dict = cfbDict
End Property

Public Property Set Dict(ByRef Value As Dictionary)
    Set cfbDict = Value
End Property

Public Property Get CurrentVocab() As String
    CurrentVocab = sCurrentVocab
End Property

Public Property Let CurrentVocab(Value As String)
    sCurrentVocab = Value
End Property

Public Property Get ClassModules()
    Set ClassModules = dctClassModules
End Property

Public Property Set ClassModules(ByRef Value As Dictionary)
    Set dctClassModules = Value
End Property

Sub InitProps()
    Call cfbGsp.VocabStack.Clear
    Call cfbGsp.VocabStack.Add("ONLY")
    Call cfbGsp.VocabStack.Add("FORTH")
    Call cfbGsp.VocabStack.Add("APPSPEC")
    
    Call cfbGsp.cfb.ClassModules.RemoveAll
    dctClassModules.Add key:="CorePrims", Item:=coreprimsmodule
    dctClassModules.Add key:="Interpreter", Item:=interpreterModule
    dctClassModules.Add key:="Compiler", Item:=compilerModule
    dctClassModules.Add key:="LogicOps", Item:=logicOpsModule
    dctClassModules.Add key:="AppSpec", Item:=appSpecModule
End Sub

Sub BuildPrimitive(psName As String, psClassName As String, psCodeField As String, _
    psVocab As String, psCompAction As String, psHelp As String)

    Dim cw As New CreoleWord
    Dim params As New ArrayList
    Dim dataVals As New ArrayList
  
    cw.NameField = psName
    cw.ClassName = psClassName
    cw.CodeField = psCodeField
    cw.Vocabulary = psVocab
    cw.CompileActionField = psCompAction
    cw.HelpField = psHelp
    cw.PrevRowLocField = cfbAddress.Count - 1
    cw.RowLocField = cfbAddress.Count
    cw.LinkField = cfbAddress.Count - 1
    cw.IndexField = cfbAddress.Count
    Set cw.ParamField = params
    cw.ParamFieldStart = 0
    cw.ParamFieldStartFrom = 0
    Set cw.DataField = dataVals
    cfbDict.Add key:=cw.FQNameField, Item:=cw
    Call cfbAddress.Add(cw)
End Sub

Sub BuildHighLevel(psCode As String, psHelp As String)
    Dim cw As CreoleWord
    Dim k As Variant
    Dim i As Long
    Dim j As Long
    Dim sList As String
    Dim sVals As String
      
    cfbGsp.InputArea = psCode
    Set cfbGsp.cfb.ClassModules = dctClassModules
    Set cfbGsp.cfb.Dict = cfbDict
    Set cfbGsp.cfb.address = cfbAddress
    cfbGsp.cfb.CurrentVocab = sCurrentVocab

    lngReturnVal = cfbGsp.cfb.ClassModules.Item("Interpreter").DoOuter(cfbGsp)
    Set cw = cfbGsp.cfb.address.Items(cfbGsp.cfb.address.Count - 1)
    cfbGsp.cfb.address.Items(cfbAddress.Count - 1).HelpField = psHelp
    cfbGsp.cfb.Dict.Item(cw.FQNameField).HelpField = psHelp
    
    Set cfbAddress = cfbGsp.cfb.address
    Set cfbDict = cfbGsp.cfb.Dict
           
End Sub
    
Sub BuildDefinitions()
    sCurrentVocab = "FORTH"
    Call InitProps
    ' The onlies
    Call BuildPrimitive("ONLY", "Interpreter", "DoOnly", "ONLY", "EXECUTE", "( -- ) Empties the vocabulary stack, then puts ONLY on it")
    Call BuildPrimitive("FORTH", "Interpreter", "DoForth", "ONLY", "EXECUTE", "( -- ) Puts FORTH on the vocabulary stack")
    Call BuildPrimitive("APPSPEC", "Interpreter", "DoAppSpec", "ONLY", "EXECUTE", "( -- ) Puts APPSPEC on the vocabulary stack")
    Call BuildPrimitive("NOP", "CorePrims", "DoNOP", "ONLY", "COMPINPF", "( -- ) Do-nothing primitive which is surprisingly useful")
    Call BuildPrimitive("__#EOL#__", "CorePrims", "DoNOP", "ONLY", "NOP", "( -- ) EOL marker")

    ' Dialogs and help
    Call BuildPrimitive("HELLO", "CorePrims", "DoHello", "FORTH", "COMPINPF", "( -- ) prints out Hello World")
    Call BuildPrimitive("TULIP", "CorePrims", "DoTulip", "FORTH", "COMPINPF", "( -- ) Either prints Tulip or pops up a message")
    Call BuildPrimitive("MSGBOX", "CorePrims", "DoMsgBox", "FORTH", "COMPINPF", "( msg -- ) Pops up an alert saying the message")
    Call BuildPrimitive("VLIST", "CorePrims", "DoVlist", "FORTH", "COMPINPF", "( -- ) Lists the dictionary definitions")

    ' Basic math
    Call BuildPrimitive("+", "CorePrims", "DoPlus", "FORTH", "COMPINPF", "( n1 n2 -- sum ) Adds two numbers on the stack")
    Call BuildPrimitive("-", "CorePrims", "DoMinus", "FORTH", "COMPINPF", "( n1 n2 -- difference ) Subtracts two numbers on the stack")
    Call BuildPrimitive("*", "CorePrims", "DoMultiply", "FORTH", "COMPINPF", "( n1 n2 -- product ) Multiplies two numbers on the stack")
    Call BuildPrimitive("/", "CorePrims", "DoDivide", "FORTH", "COMPINPF", "( n1 n2 -- quotient ) Divides two numbers on the stack")
    Call BuildPrimitive("%", "CorePrims", "DoMod", "FORTH", "COMPINPF", "( n1 n2 -- remainder ) Returns remainder of division operation")

    ' Date/time handling
  '  Call BuildPrimitive("TODAY", "CorePrims", "DoToday", "FORTH", "COMPINPF", "( -- ) Pops up today's date")
    Call BuildPrimitive("NOW", "CorePrims", "DoNow", "FORTH", "COMPINPF", "( --  time ) Puts the time on the stack")
  '  Call BuildPrimitive(">HHMMSS", "CorePrims", "DoToHoursMinSecs", "FORTH", "COMPINPF", "( time -- ) Formats the time")

    ' Stack manipulation
    Call BuildPrimitive("DUP", "CorePrims", "DoDup", "FORTH", "COMPINPF", "( val --  val val ) Duplicates the argument on top of the stack")
    Call BuildPrimitive("SWAP", "CorePrims", "DoSwap", "FORTH", "COMPINPF", "( val1 val2 -- val2 val1 ) Swaps the positions of the top two stack arguments")
    Call BuildPrimitive("ROT", "CorePrims", "DoRot", "FORTH", "COMPINPF", "( val1 val2 val3 -- val2 val3 val1 ) Moves the third stack argument to the top")
    Call BuildPrimitive("-ROT", "CorePrims", "DoMinusRot", "FORTH", "COMPINPF", "( val1 val2 val3 -- val3 val1 val2 ) Moves the top stack argument to the third position")
    Call BuildPrimitive("NIP", "CorePrims", "DoNip", "FORTH", "COMPINPF", "( val1 val2 -- val2 ) Removes second stack argument")
    Call BuildPrimitive("TUCK", "CorePrims", "DoTuck", "FORTH", "COMPINPF", "( val1 val2 -- val2 val1 val2 ) Copies top stack argument under second argument")
    Call BuildPrimitive("OVER", "CorePrims", "DoOver", "FORTH", "COMPINPF", "( val1 val2 -- val1 val2 val1 ) Copies second stack argument to the top of the stack")
    Call BuildPrimitive("DROP", "CorePrims", "DoDrop", "FORTH", "COMPINPF", "( val -- ) Drops the argument at the top of the stack")
    Call BuildPrimitive(".", "CorePrims", "DoDot", "FORTH", "COMPINPF", "( val -- ) Prints the argument at the top of the stack")
    Call BuildPrimitive(".S", "CorePrims", "DoDotDataStack", "FORTH", "COMPINPF", "( -- ) Nondestructive data stack print - rightmost is tos")
    Call BuildPrimitive("DEPTH", "CorePrims", "doDepth", "FORTH", "COMPINPF", "( -- n ) Returns the stack depth")

    ' Logical operatives
    Call BuildPrimitive("=", "LogicOps", "DoEquals", "FORTH", "COMPINPF", "( val1 val2 -- flag ) -1 if equal, 0 otherwise")
    Call BuildPrimitive("<>", "LogicOps", "DoNotEquals", "FORTH", "COMPINPF", "( val1 val2 -- flag ) 0 if equal, -1 otherwise")
    Call BuildPrimitive("<", "LogicOps", "DoLessThan", "FORTH", "COMPINPF", "( val1 val2 -- flag ) -1 if less than, 0 otherwise")
    Call BuildPrimitive(">", "LogicOps", "DoGreaterThan", "FORTH", "COMPINPF", "( val1 val2 -- flag ) -1 if greater than, 0 otherwise")
    Call BuildPrimitive("<=", "LogicOps", "DoLessThanOrEquals", "FORTH", "COMPINPF", "( val1 val2 -- flag ) -1 if less than or equal to, 0 otherwise")
    Call BuildPrimitive(">=", "LogicOps", "DoGreaterThanOrEquals", "FORTH", "COMPINPF", "( val1 val2 -- flag ) -1 if greater than or equal to, 0 otherwise")
    Call BuildPrimitive("NOT", "LogicOps", "DoNot", "FORTH", "COMPINPF", "( val -- opval ) -1 if 0, 0 otherwise")
    Call BuildPrimitive("AND", "LogicOps", "DoAnd", "FORTH", "COMPINPF", "( val1 val2 -- flag ) -1 if both arguments are non-zero, 0 otherwise")
    Call BuildPrimitive("OR", "LogicOps", "DoOr", "FORTH", "COMPINPF", "( val1 val2 -- flag ) -1 if one or both arguments are non-zero, 0 otherwise")
    Call BuildPrimitive("XOR", "LogicOps", "DoXor", "FORTH", "COMPINPF", "( val1 val2 -- flag ) -1 if one and only one argument is non-zero, 0 otherwise")

    ' Compiler definitions
    Call BuildPrimitive(",", "Compiler", "DoComma", "FORTH", "COMPINPF", "( n --) Compiles value off the TOS into the next parameter field cell")
    Call BuildPrimitive("COMPINPF", "Compiler", "DoComma", "IMMEDIATE", "COMPINPF", "( n --) Does the same thing as , (comma) - given a different name for ease of reading")
    Call BuildPrimitive("EXECUTE", "Compiler", "DoExecute", "FORTH", "COMPINPF", "( address --) Executes the word corresponding to the address on the stack")
    Call BuildPrimitive(":", "Compiler", "CompileColon", "FORTH", "COMPINPF", "( -- ) Starts compilation of a colon definition")
    Call BuildPrimitive(";", "CorePrims", "DoNOP", "IMMEDIATE", "EXECUTE", "( -- ) Terminates compilation of a colon definition")
    Call BuildPrimitive("COMPLIT", "Compiler", "CompileLiteral", "IMMEDIATE", "EXECUTE", "( -- ) Compiles doLiteral and a literal into the parameter field")
    Call BuildPrimitive("DoLiteral", "Compiler", "DoLiteral", "IMMEDIATE", "NOP", "( -- lit ) Run-time code that pushes a literal onto the stack")
    Call BuildPrimitive("HERE", "Compiler", "DoHere", "FORTH", "COMPINPF", "( -- location ) Returns address of the next available dictionary location")
    Call BuildPrimitive("CREATE", "Compiler", "DoCreate", "FORTH", "COMPINPF", "CREATE <name>. Adds a named entry into the dictionary")
    Call BuildPrimitive("doDoes", "Compiled", "DoDoes", "IMMEDIATE", "COMPINPF", "( address -- ) Run-time code for DOES>")
    Call BuildPrimitive("DOES>", "Compiler", "CompileDoes", "FORTH", "COMPINPF", _
         "DOES> <list of runtime actions>. When defining word is created, copies code following it into the child definition")
    Call BuildPrimitive("@", "Compiler", "DoFetch", "FORTH", "COMPINPF", "( addr -- val ) Fetches the value in the param field  at addr")
    Call BuildPrimitive("!", "Compiler", "DoStore", "FORTH", "COMPINPF", "( val addr --) Stores the value in the param field  at addr")

    Call BuildPrimitive("DEFINITIONS", "Compiler", "DoSetCurrentToContext", "FORTH", "COMPINPF", _
        "(  -- ) Sets the current (compilation) vocabulary to the context vocabulary (the one on top of the vocabulary stack)")
    Call BuildPrimitive("IMMEDIATE", "Compiler", "DoImmediate", "FORTH", "COMPINPF", _
        "( -- ) Flags a word as immediate (executes instead of compiling inside a colon definition)")
    Call BuildPrimitive("DISP_PF", "Compiler", "DoDisplayParamField", "FORTH", "COMPINPF", _
        "( -- ) Displays the contents of the most recently defined definition")
    
    ' Branching compiler definitions
    Call BuildPrimitive("IF", "Compiler", "CompileIf", "IMMEDIATE", "EXECUTE", "( -- location ) Compile-time code for IF")
    Call BuildPrimitive("ELSE", "Compiler", "CompileElse", "IMMEDIATE", "EXECUTE", "( -- location ) Compile-time code for ELSE")
    Call BuildPrimitive("THEN", "Compiler", "CompileThen", "IMMEDIATE", "EXECUTE", "( -- location ) Compile-time code for THEN")
    Call BuildPrimitive("0BRANCH", "Compiler", "Do0Branch", "IMMEDIATE", "NOP", "( flag -- ) Run-time code for IF")
    Call BuildPrimitive("JUMP", "Compiler", "DoJump", "IMMEDIATE", "NOP", _
        "( -- ) Jumps unconditionally to the parameter field location next to it and is compiled by ELSE")
    Call BuildPrimitive("DoElse", "CorePrims", "DoNOP", "IMMEDIATE", "NOP", "( -- ) Run-time code for ELSE")
    Call BuildPrimitive("DoThen", "CorePrims", "DoNOP", "IMMEDIATE", "NOP", "( -- ) Run-time code for THEN")
    Call BuildPrimitive("BEGIN", "Compiler", "CompileBegin", "IMMEDIATE", "EXECUTE", "( -- beginLoc ) Compile-time code for BEGIN")
    Call BuildPrimitive("UNTIL", "Compiler", "CompileUntil", "IMMEDIATE", "EXECUTE", "( beginLoc -- ) Compile-time code for UNTIL")
    Call BuildPrimitive("DoBegin", "CorePrims", "DoNOP", "IMMEDIATE", "NOP", "( -- ) Run-time code for BEGIN")
    Call BuildPrimitive("DO", "Compiler", "CompileDo", "IMMEDIATE", "EXECUTE", "( -- beginLoc ) Compile-time code for DO")
    Call BuildPrimitive("LOOP", "Compiler", "CompileLoop", "IMMEDIATE", "EXECUTE", "( -- beginLoc ) Compile-time code for LOOP")
    Call BuildPrimitive("+LOOP", "Compiler", "CompilePlusLoop", "IMMEDIATE", "EXECUTE", "( -- beginLoc ) Compile-time code for +LOOP")
    Call BuildPrimitive("DoStartDo", "Compiler", "DoStartDo", "IMMEDIATE", "COMPINPF", "( start end -- ) Starts off the Do by getting the start and end")
    Call BuildPrimitive("DoDo", "CorePrims", "DoNOP", "IMMEDIATE", "COMPINPF", "( -- ) Marker for DoLoop to return to")
    Call BuildPrimitive("DoLoop", "Compiler", "DoLoop", "IMMEDIATE", "COMPINPF", "( -- ) Loops back to doDo until the start equals the end")
    Call BuildPrimitive("DoPlusLoop", "Compiler", "DoPlusLoop", "IMMEDIATE", "COMPINPF", "( inc -- ) Loops back to doDo until the start >= the end and increments with inc")
    Call BuildPrimitive("I", "Compiler", "DoIndexI", "FORTH", "COMPINPF", "( -- index ) Returns the index of I")
    Call BuildPrimitive("J", "Compiler", "DoIndexJ", "FORTH", "COMPINPF", "( -- index ) Returns the index of J")
    Call BuildPrimitive("K", "Compiler", "DoIndexK", "FORTH", "COMPINPF", "( -- index ) Returns the index of K")

    ' Commenting and list compiler
    Call BuildPrimitive("//", "Interpreter", "DoSingleLineCmts", "FORTH", "EXECUTE", "( -- ) Single-line comment handling")
    Call BuildPrimitive("(", "Interpreter", "DoParenCmts", "FORTH", "EXECUTE", "( -- ) Multiline comment handling")
    Call BuildPrimitive("{", "Interpreter", "CompileList", "FORTH", "EXECUTE", "( -- list ) List compiler")
 
    sCurrentVocab = "APPSPEC"
    Call BuildPrimitive("TEST", "AppSpec", "DoTest", "APPSPEC", "COMPINPF", "( -- ) Do whatever you want here")
 
 '    Call BuildHighLevel(": TEST2 IF HELLO ELSE TULIP THEN ;", "First high-level definition")
    Call BuildHighLevel(": BUTEST BEGIN 1 + DUP 10 > UNTIL ;", "( n -- ) Tests BEGIN-UNTIL looping")
    Call BuildHighLevel(": TEST3 HELLO TULIP ;", "First high-level definition")
    Call BuildHighLevel(": CONSTANT CREATE , DOES> @ ;", "Creates a constant")
    Call BuildHighLevel(": 3H 3 0 DO HELLO LOOP ;", "Three hellos")
End Sub


  

