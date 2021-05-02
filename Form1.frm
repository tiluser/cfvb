VERSION 5.00
Begin VB.Form CreoleForthForm 
   Caption         =   "Creole Forth Tester"
   ClientHeight    =   6315
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6465
   DrawMode        =   0  'Blackness
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   6315
   ScaleWidth      =   6465
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdBtn 
      Caption         =   "Execute Forth"
      Height          =   495
      Left            =   2640
      TabIndex        =   1
      Top             =   5400
      Width           =   1215
   End
   Begin VB.TextBox ForthCodeText 
      Height          =   3735
      Left            =   840
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   960
      Width           =   4815
   End
End
Attribute VB_Name = "CreoleForthForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim iReturnVal As Integer
Dim coreprims As New coreprims
Dim interpreter As New interpreter
Dim Gsp As New GlobalSimpleProps
Dim cfb1 As New CreoleForthBundle

Private Sub CmdBtn_Click()
    Gsp.OuterPtr = 0
    Gsp.InputArea = ForthCodeText.Text
    iReturnVal = interpreter.DoOuter(Gsp)
End Sub

Private Sub Form_Initialize()
    Gsp.Scratch = "ONLY"
    Call Gsp.Push(Gsp.VocabStack)
    Gsp.Scratch = "FORTH"
    Call Gsp.Push(Gsp.VocabStack)
    Gsp.Scratch = "APPSPEC"
    Call Gsp.Push(Gsp.VocabStack)

    Set cfb1.Gsp = Gsp
    Call cfb1.BuildDefinitions
    Set Gsp.cfb = cfb1
End Sub

