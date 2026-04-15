Attribute VB_Name = "VBASyncState"
Option Explicit

Private HashCache As Object

Private Sub EnsureCache()
    If HashCache Is Nothing Then
        Set HashCache = CreateObject("Scripting.Dictionary")
    End If
End Sub

Public Function VBASyncState_GetChangedSet(ByVal queue As Collection) As Object

    EnsureCache

    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")

    Dim i As Long
    Dim item As VBASyncImport

    For i = 1 To queue.count

        Set item = queue(i)

        If HasChanged(item.ComponentName, item.Code) Then
            result(item.ComponentName) = True
        End If

    Next i

    Set VBASyncState_GetChangedSet = result

End Function

Public Function HasChanged(ByVal name As String, ByVal Code As String) As Boolean

    EnsureCache

    Dim h As String
    h = ComputeHash(Code)

    If Not HashCache.Exists(name) Then
        HashCache(name) = h
        HasChanged = True
        Exit Function
    End If

    HasChanged = (HashCache(name) <> h)

    If HasChanged Then HashCache(name) = h

End Function

Private Function ComputeHash(ByVal s As String) As String

    Dim i As Long, h As Long
    h = 5381

    For i = 1 To Len(s)
        h = ((h * 33) Xor Asc(Mid$(s, i, 1))) And &H7FFFFFFF
    Next i

    ComputeHash = Hex$(h)

End Function

Public Sub VBASyncState_Reset()
    Set HashCache = Nothing
End Sub
