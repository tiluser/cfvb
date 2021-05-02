VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ArrayList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Got this one off of Stack Overflow but I'm not sure who the author is
' If anyone knows, please bring it to my attention.
Option Explicit

Private mavInternalArray() As Variant
Private mlArraySize As Long
Private mlCount As Long
Private mlGrowSize As Long
Private mfZeroIndex As Boolean

'---------------------------------------------------------------------------------------
' Procedure Clear
'---------------------------------------------------------------------------------------
Public Sub Clear()
          Dim Index As Long

        For Index = 0 To mlCount - 1
            If IsObject(mavInternalArray(Index)) Then
                Set mavInternalArray(Index) = Nothing
            End If
        Next Index
        mlCount = 0

End Sub

'---------------------------------------------------------------------------------------
' Procedure Swap
'---------------------------------------------------------------------------------------
Public Sub Swap(Index1 As Long, index2 As Long)
          Dim vTmp As Variant


        If IsObject(mavInternalArray(index2)) Then
            Set vTmp = mavInternalArray(index2)
        Else
            vTmp = mavInternalArray(index2)
        End If

        If IsObject(mavInternalArray(Index1)) Then
            Set mavInternalArray(index2) = mavInternalArray(Index1)
        Else
           mavInternalArray(index2) = mavInternalArray(Index1)
       End If

       If IsObject(vTmp) Then
           Set mavInternalArray(Index1) = vTmp
       Else
           mavInternalArray(Index1) = vTmp
       End If


End Sub

Public Property Get ZeroIndex() As Boolean
        ZeroIndex = mfZeroIndex
End Property

Public Property Let ZeroIndex(fZeroIndex As Boolean)
        mfZeroIndex = fZeroIndex
End Property

Public Property Get GrowSize() As Long
        GrowSize = mlGrowSize
End Property

Public Property Let GrowSize(lNewSize As Long)
        Debug.Assert lNewSize > 0
        mlGrowSize = lNewSize
End Property

Private Sub Class_Initialize()
        mlGrowSize = 50
        mlArraySize = mlGrowSize
        mfZeroIndex = True
        mlCount = 0


        ReDim mavInternalArray(0 To mlGrowSize - 1)

End Sub

'---------------------------------------------------------------------------------------
' Procedure Remove
'---------------------------------------------------------------------------------------
Public Sub Remove(Index As Long)
        Dim index2 As Long


        For index2 = Index To mlCount - 2
            If IsObject(mavInternalArray(index2 + 1)) Then
                Set mavInternalArray(index2) = mavInternalArray(index2 + 1)
            Else
                mavInternalArray(index2) = mavInternalArray(index2 + 1)
            End If
        Next index2
          If mlCount <= 0 Then
            Exit Sub
          End If
        mlCount = mlCount - 1
        If IsObject(mavInternalArray(mlCount)) Then
           Set mavInternalArray(mlCount) = Nothing
       Else
           mavInternalArray(mlCount) = False
       End If
End Sub

'---------------------------------------------------------------------------------------
' Procedure Items
'---------------------------------------------------------------------------------------
Public Function Items(Index As Long) As Variant
        If Not mfZeroIndex Then
            Index = Index - 1
        End If

        If Index < mlCount And Index >= 0 Then
            If IsObject(mavInternalArray(Index)) Then
                Set Items = mavInternalArray(Index)
            Else
                Items = mavInternalArray(Index)
            End If
       End If
End Function

Public Sub SetItem(Index As Long, Item As Variant)
        If Not mfZeroIndex Then
            Index = Index - 1
        End If
        If IsObject(Item) Then
            Set mavInternalArray(Index) = Item
        Else
            mavInternalArray(Index) = Item
        End If
End Sub

'---------------------------------------------------------------------------------------
' Procedure Add
'---------------------------------------------------------------------------------------
Public Function Add(vItem As Variant) As Long

        mlCount = mlCount + 1
        If mlCount > mlArraySize Then
            mlArraySize = mlArraySize + mlGrowSize
            ReDim Preserve mavInternalArray(0 To mlArraySize - 1)
        End If

        If IsObject(vItem) Then
            Set mavInternalArray(mlCount - 1) = vItem
        Else
            mavInternalArray(mlCount - 1) = vItem
       End If

       Add = mlCount - 1

End Function

'---------------------------------------------------------------------------------------
' Procedure ItemArray
'---------------------------------------------------------------------------------------
Public Function ItemArray() As Variant
        Dim vReturnArray As Variant

        vReturnArray = mavInternalArray
        ReDim Preserve vReturnArray(0 To mlCount - 1)
        ItemArray = vReturnArray
End Function

Public Function Count() As Long
        Count = mlCount
End Function


'---------------------------------------------------------------------------------------
' Procedure Insert
'---------------------------------------------------------------------------------------
Public Function Insert(Index As Long, vItem As Variant) As Long
        Dim index2 As Long

        'Make sure array is large enough for a new item
        mlCount = mlCount + 1
        If mlCount > mlArraySize Then
            mlArraySize = mlArraySize + mlGrowSize
            ReDim Preserve mavInternalArray(0 To mlArraySize - 1)
        End If

        'Bump all the items with a higher index up one spot

        If Index >= mlCount - 1 Then
            If IsObject(vItem) Then
                Set mavInternalArray(mlCount - 1) = vItem
            Else
               mavInternalArray(mlCount - 1) = vItem
           End If
       Else

           For index2 = mlCount - 1 To Index + 1 Step -1
               If IsObject(vItem) Then
                   Set mavInternalArray(index2) = mavInternalArray(index2 - 1)
               Else
                   mavInternalArray(index2) = mavInternalArray(index2 - 1)
               End If
           Next index2

           If IsObject(vItem) Then
               Set mavInternalArray(Index) = vItem
           Else
               mavInternalArray(Index) = vItem
           End If
       End If
       Insert = mlCount - 1

End Function


Public Property Get NewEnum() As IUnknown
    ''#Provides support for enumeration using For Each
    Set NewEnum = m_Customers.[_NewEnum]
End Property
