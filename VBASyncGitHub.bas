Attribute VB_Name = "VBASyncGitHub"
'---------------------------------------------------------------------------------------
' Module    : VBASyncGitHub
' Version   : v1.4.0
' Purpose   : GitHub Tree + Diff + Future dependency registry hook
' Dependency : VBASyncState + VBASyncEngine
'---------------------------------------------------------------------------------------

Option Explicit

Private Const GITHUB_API As String = "https://api.github.com/repos/"

' =========================================================
' PUBLIC API
' =========================================================

' Returns a visual tree of repository VBA files
Public Function VBASync_GitHub_ToTree(ByVal owner As String, ByVal repo As String) As String

    Dim json As String
    json = DownloadURL(GITHUB_API & owner & "/" & repo & "/contents/")

    Dim remoteMap As Object
    Set remoteMap = ParseGitHubAPI(json)

    VBASync_GitHub_ToTree = BuildTree(remoteMap, repo)

End Function

' =========================================================
' GIT DIFF (LOCAL vs REMOTE)
' =========================================================
Public Function VBASync_RunGitDiff(ByVal owner As String, ByVal repo As String) As String

    Dim remoteMap As Object
    Dim localMap As Object
    Dim url As String

    url = GITHUB_API & owner & "/" & repo & "/contents/"
    Set remoteMap = ParseGitHubAPI(DownloadURL(url))
    Set localMap = GetLocalMap()

    If localMap Is Nothing Then
        VBASync_RunGitDiff = "ERROR: VBProject access denied."
        Exit Function
    End If

    Dim out As String
    Dim k As Variant
    Dim comp As Object

    out = "=== VBASync GIT DIFF ===" & vbCrLf & _
          "Repo: " & owner & "/" & repo & vbCrLf & _
          "Time: " & Now & vbCrLf & _
          "----------------------------------" & vbCrLf

    ' REMOTE ONLY
    For Each k In remoteMap.Keys
        If Not localMap.Exists(k) Then
            out = out & "[REMOTE ONLY] " & remoteMap(k) & vbCrLf
        End If
    Next

    ' LOCAL ONLY
    For Each k In localMap.Keys
        If Not remoteMap.Exists(k) Then

            On Error Resume Next
            Set comp = ThisWorkbook.VBProject.VBComponents(k)
            On Error GoTo 0

            out = out & "[LOCAL ONLY] " & k & GetVBAExtension(comp) & vbCrLf
        End If
    Next

    VBASync_RunGitDiff = out

End Function

' =========================================================
' TREE BUILDER
' =========================================================
Private Function BuildTree(ByVal d As Object, ByVal repoName As String) As String

    Dim k As Variant
    Dim out As String

    out = "?? " & repoName & " (GitHub)" & vbCrLf

    If d.count = 0 Then
        out = out & "+-- (no VBA files found)"
        BuildTree = out
        Exit Function
    End If

    For Each k In d.Keys
        out = out & "+-- " & d(k) & vbCrLf
    Next

    BuildTree = out

End Function

' =========================================================
' LOCAL MAP
' =========================================================
Private Function GetLocalMap() As Object

    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1

    Dim c As Object

    On Error Resume Next

    Dim test As Long
    test = ThisWorkbook.VBProject.VBComponents.count

    If Err.Number <> 0 Then
        Set GetLocalMap = Nothing
        Exit Function
    End If

    On Error GoTo 0

    For Each c In ThisWorkbook.VBProject.VBComponents

        If c.Type >= 1 And c.Type <= 3 Then
            d(c.Name) = True
        End If

    Next

    Set GetLocalMap = d

End Function

' =========================================================
' PARSER (ROBUSTED)
' =========================================================
Private Function ParseGitHubAPI(ByVal json As String) As Object

    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1

    Dim parts() As String
    parts = Split(json, """name"":""")

    Dim i As Long
    Dim fileName As String
    Dim baseName As String

    For i = 1 To UBound(parts)

        fileName = Split(parts(i), """")(0)

        If IsVBAFile(fileName) Then

            If InStr(fileName, ".") > 0 Then
                baseName = Left$(fileName, InStrRev(fileName, ".") - 1)
                d(baseName) = fileName
            End If

        End If

    Next i

    Set ParseGitHubAPI = d

End Function

Private Function IsVBAFile(ByVal f As String) As Boolean
    IsVBAFile = (InStr(1, f, ".bas", vbTextCompare) > 0 Or _
                 InStr(1, f, ".cls", vbTextCompare) > 0 Or _
                 InStr(1, f, ".frm", vbTextCompare) > 0)
End Function

' =========================================================
' EXTENSION RESOLVER
' =========================================================
Private Function GetVBAExtension(ByVal comp As Object) As String

    On Error GoTo safe

    Select Case comp.Type
        Case 1: GetVBAExtension = ".bas"
        Case 2: GetVBAExtension = ".cls"
        Case 3: GetVBAExtension = ".frm"
        Case Else: GetVBAExtension = ""
    End Select

    Exit Function

safe:
    GetVBAExtension = ""

End Function

' =========================================================
' HTTP
' =========================================================
Private Function DownloadURL(ByVal url As String) As String

    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")

    On Error GoTo fail

    http.Open "GET", url, False
    http.SetRequestHeader "User-Agent", "VBA-Sync-Engine"
    http.Send

    DownloadURL = http.ResponseText
    Exit Function

fail:
    DownloadURL = "{}"

End Function

