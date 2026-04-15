Attribute VB_Name = "VBASyncECS"
'---------------------------------------------------------------------------------------
' Module    : VBASyncECS.bas
' Version   : 3.4.0
' Purpose   :
'---------------------------------------------------------------------------------------
Option Explicit

Public Function VBASyncECS_BuildGraph(ByVal queue As Collection) As Object

    Dim graph As Object
    Set graph = CreateObject("Scripting.Dictionary")

    Dim item As VBASyncImport

    For Each item In queue
        graph(item.ComponentName) = ExtractDependencies(item.Code)
    Next item

    Set VBASyncECS_BuildGraph = graph

End Function

Public Function VBASyncECS_OrderSubset(ByVal queue As Collection, ByVal subset As Object) As Collection

    Dim result As New Collection
    Dim item As VBASyncImport

    For Each item In queue
        If subset.Exists(item.ComponentName) Then
            result.Add item
        End If
    Next item

    Set VBASyncECS_OrderSubset = result

End Function

Public Function VBASyncECS_OrderByDependencies(ByVal queue As Collection) As Collection

    Dim graph As Object
    Set graph = VBASyncECS_BuildGraph(queue)

    Dim result As New Collection
    Dim visited As Object
    Set visited = CreateObject("Scripting.Dictionary")

    Dim item As VBASyncImport

    For Each item In queue
        VisitNode item.ComponentName, graph, visited, result, queue
    Next item

    Set VBASyncECS_OrderByDependencies = result

End Function

' =========================
' TOPOLOGICAL VISIT
' =========================
Private Sub VisitNode(ByVal Name As String, ByVal graph As Object, _
                      ByVal visited As Object, _
                      ByRef result As Collection, _
                      ByVal queue As Collection)

    If visited.Exists(Name) Then Exit Sub
    visited(Name) = True

    If graph.Exists(Name) Then

        Dim dep As Variant
        For Each dep In graph(Name)
            If graph.Exists(dep) Then
                VisitNode dep, graph, visited, result, queue
            End If
        Next dep

    End If

    Dim item As VBASyncImport
    For Each item In queue
        If item.ComponentName = Name Then
            result.Add item
            Exit For
        End If
    Next item

End Sub

' =========================
' DEPENDENCY EXTRACTION
' =========================
Private Function ExtractDependencies(ByVal Code As String) As Collection

    Dim deps As New Collection
    Dim lines() As String
    Dim i As Long

    lines = Split(Code, vbCrLf)

    For i = LBound(lines) To UBound(lines)

        Dim l As String
        l = Trim$(lines(i))

        If InStr(1, l, " As ", vbTextCompare) > 0 Then
            SafeAdd deps, "TYPE_DEP"
        End If

        If LCase$(Left$(l, 10)) = "implements" Then
            SafeAdd deps, Trim$(Mid$(l, 11))
        End If

    Next i

    Set ExtractDependencies = deps

End Function

Private Sub SafeAdd(col As Collection, v As String)
    On Error Resume Next
    col.Add v
    On Error GoTo 0
End Sub

