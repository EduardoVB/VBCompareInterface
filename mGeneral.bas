Attribute VB_Name = "mGeneral"
Option Explicit

Public Function AddToList(nList, nValue, Optional nOnlyIfMissing As Boolean, Optional nFirstElement As Long = 0) As Boolean
    Dim i As Long
    Dim iAdd As Boolean
    
    If Not nOnlyIfMissing Then
        iAdd = True
    Else
        iAdd = Not IsInList(nList, nValue, nFirstElement)
    End If
    If iAdd Then
        i = UBound(nList) + 1
        ReDim Preserve nList(LBound(nList) To i)
        nList(i) = nValue
        AddToList = True
    End If
End Function

Public Function AddObjectToList(nList, nObject) As Boolean
    Dim i As Long
    
    i = UBound(nList) + 1
    ReDim Preserve nList(LBound(nList) To i)
    Set nList(i) = nObject
End Function

Public Function IsInList(nList, nValue, Optional nFirstElement As Long = 0, Optional nLastElement As Long = -1) As Boolean
    Dim c As Long
    
    If nLastElement = -1 Then
        nLastElement = UBound(nList)
    Else
        If nLastElement > UBound(nList) Then
            nLastElement = UBound(nList)
        End If
    End If
    
    For c = nFirstElement To nLastElement
        If nList(c) = nValue Then
            IsInList = True
            Exit For
        End If
    Next c
End Function

Public Function IndexInList(nList, nValue) As Long
    Dim c As Long
    
    IndexInList = LBound(nList) - 1
    For c = LBound(nList) To UBound(nList)
        If nList(c) = nValue Then
            IndexInList = c
            Exit For
        End If
    Next c
End Function

