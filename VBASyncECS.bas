Attribute VB_Name = "VBASyncECS"
'---------------------------------------------------------------------------------------
' Module    : VBASyncECS
' Version   : 3.4.2
' Purpose   : Dependency Graph Construction & Topological Sorting for VBA Components.
'---------------------------------------------------------------------------------------
Option Explicit

Public Function VBASyncECS_OrderByDependencies(ByVal queue As Collection) As Collection
    Dim graph As Object: Set graph = VBASyncECS_BuildGraph(queue)
    Dim result As New Collection
    Dim visited As Object: Set visited = CreateObject("Scripting.Dictionary")
    Dim item As Object
    
    For Each item In queue
        VisitNode item("ComponentName"), graph, visited, result, queue
    Next item
    
    Set VBASyncECS_OrderByDependencies = result
End Function

Private Function VBASyncECS_BuildGraph(ByVal queue As Collection) As Object
    Dim graph As Object: Set graph = CreateObject("Scripting.Dictionary")
    Dim item As Object
    For Each item In queue
        graph(item("ComponentName")) = ExtractDependencies(item("Code"))
    Next item
    Set VBASyncECS_BuildGraph = graph
End Function

Private Sub VisitNode(ByVal Name As String, ByVal graph As Object, ByVal visited As Object, ByRef result As Collection, ByVal queue As Collection)
    If visited.Exists(Name) Then Exit Sub
    visited(Name) = True
    
    If graph.Exists(Name) Then
        Dim dep As Variant
        For Each dep In graph(Name)
            VisitNode CStr(dep), graph, visited, result, queue
        Next dep
    End If
    
    Dim item As Object
    For Each item In queue
        If item("ComponentName") = Name Then
            result.Add item
            Exit For
        End If
    Next item
End Sub


' =========================================================
' DEPENDENCY EXTRACTION
' =========================================================
Private Function ExtractDependencies(ByVal Code As String) As Collection
    Dim deps As New Collection
    Dim lines() As String
    Dim i As Long
    Dim l As String

    lines = Split(Code, vbCrLf)

    For i = LBound(lines) To UBound(lines)
        l = Trim$(lines(i))

        ' 1. Check for Variable/Parameter types (e.g., "As Mesh")
        If InStr(1, l, " As ", vbTextCompare) > 0 Then
            ' This is a simplified extractor; it grabs the word after "As"
            Dim parts() As String
            parts = Split(l, " ")
            Dim j As Long
            For j = LBound(parts) To UBound(parts) - 1
                If LCase$(parts(j)) = "as" Then
                    SafeAdd deps, parts(j + 1)
                End If
            Next j
        End If

        ' 2. Check for Implements (Interfaces)
        If LCase$(Left$(l, 10)) = "implements" Then
            SafeAdd deps, Trim$(Mid$(l, 11))
        End If
        
        ' 3. Check for New keyword
        If InStr(1, l, " New ", vbTextCompare) > 0 Then
             Dim newParts() As String
             newParts = Split(l, " ")
             For j = LBound(newParts) To UBound(newParts) - 1
                If LCase$(newParts(j)) = "new" Then
                    SafeAdd deps, newParts(j + 1)
                End If
             Next j
        End If
    Next i

    Set ExtractDependencies = deps
End Function

' Helper to prevent "Key already exists" errors in the collection
Private Sub SafeAdd(ByRef col As Collection, ByVal val As String)
    On Error Resume Next
    ' Sanitize the value (remove parentheses or dots)
    val = Replace(Replace(val, "(", ""), ")", "")
    If InStr(val, ".") > 0 Then val = Left(val, InStr(val, ".") - 1)
    
    If Len(val) > 0 Then col.Add val, val
    On Error GoTo 0
End Sub
