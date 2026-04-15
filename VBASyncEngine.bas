Attribute VB_Name = "VBASyncEngine"
'---------------------------------------------------------------------------------------
' Module    : VBASyncEngine
' Version   : 3.5.2
' Purpose   : Merged State logic, File Assembly, GitHub Integration,
'             and Legacy Split/Join Utility consolidation.
'             Fixed: Single-line If/Colon syntax errors.
'---------------------------------------------------------------------------------------
Option Explicit

Private ImportQueue As Collection
Private FailedQueue As Collection
Private HashCache As Object

' =========================================================
' MAIN MENU
' =========================================================
Public Sub VBASync_Menu()
    Dim choice As String
    Dim folderPath As String
    Dim url As String
    Dim owner As String, repo As String
    
    choice = InputBox( _
        "VBASync Engine v3.5.1" & vbCrLf & vbCrLf & _
        "1 - Full Local Sync (Export & Import)" & vbCrLf & _
        "2 - Incremental Build (Smart Import)" & vbCrLf & _
        "3 - Export Only (Clean Source)" & vbCrLf & _
        "4 - Import Only" & vbCrLf & _
        "5 - GitHub Tree" & vbCrLf & _
        "6 - GitHub Diff" & vbCrLf & _
        "7 - Assemble .bas/.cls Files to Text" & vbCrLf & _
        "8 - Export All to Text Chunks (Legacy Style)" & vbCrLf & _
        "9 - Exit", _
        "VBASync", "1")

    If choice = "" Or choice = "9" Then Exit Sub
    
    Set FailedQueue = New Collection

    Select Case choice
        Case "1"
            folderPath = PickFolder()
            If folderPath <> "" Then VBASync_RunFullSync folderPath
        Case "2"
            folderPath = PickFolder()
            If folderPath <> "" Then VBASync_IncrementalBuild folderPath
        Case "3"
            folderPath = PickFolder()
            If folderPath <> "" Then ExportCleanSource folderPath
        Case "4"
            folderPath = PickFolder()
            If folderPath <> "" Then
                Set ImportQueue = BuildQueue(folderPath)
                ProcessImportQueue ImportQueue
            End If
        Case "5"
            url = InputBox("Enter GitHub Owner/Repo (e.g. 'HowdyKeith/VBAOpenGLEngine'):")
            If url <> "" Then
                SplitGitHubURL url, owner, repo
                Dim treeResult As String
                treeResult = VBASync_GitHub_ToTree(owner, repo)
                Debug.Print treeResult
                If MsgBox("Tree output to Debug window. Copy to clipboard?", vbQuestion + vbYesNo) = vbYes Then
                    CopyTextToClipboard treeResult
                End If
            End If
        Case "6"
            url = InputBox("Enter GitHub Owner/Repo:")
            If url <> "" Then
                SplitGitHubURL url, owner, repo
                MsgBox VBASync_RunGitDiff(owner, repo)
            End If
        Case "7"
            AssembleFilesToText
        Case "8"
            folderPath = PickFolder()
            If folderPath <> "" Then VBASync_ExportAllToChunks folderPath
    End Select
End Sub

' =========================================================
' CORE ENGINE LOGIC
' =========================================================
Public Sub VBASync_RunFullSync(ByVal folderPath As String)
    ExportCleanSource folderPath
    VBASync_IncrementalBuild folderPath
End Sub

Public Sub VBASync_IncrementalBuild(ByVal folderPath As String)
    Dim queue As Collection
    Set queue = BuildQueue(folderPath)
    
    Dim item As Object
    For Each item In queue
        If HasChanged(item("ComponentName"), item("Code")) Then
            ProcessSingle item
        End If
    Next
End Sub

Private Sub ProcessImportQueue(ByVal q As Collection)
    Dim i As Long
    For i = 1 To q.count
        ProcessSingle q(i)
        DoEvents
    Next
End Sub

Private Sub ProcessSingle(ByRef item As Object)
    On Error GoTo fail
    Dim tmp As String: tmp = Environ$("TEMP") & "\" & item("ComponentName") & ".tmp"
    
    RemoveExistingSafe item("ComponentName")
    WriteFile tmp, item("Code")
    ThisWorkbook.VBProject.VBComponents.Import tmp
    Kill tmp
    Exit Sub
fail:
    Debug.Print "Failed to import: " & item("ComponentName")
End Sub

' =========================================================
' GITHUB & REMOTE LOGIC
' =========================================================
Public Function VBASync_GitHub_ToTree(ByVal owner As String, ByVal repo As String) As String
    Dim json As String: json = DownloadURL("https://api.github.com/repos/" & owner & "/" & repo & "/contents/")
    Dim remoteMap As Object: Set remoteMap = ParseGitHubAPI(json)
    Dim k As Variant, out As String
    out = "?? " & repo & " (GitHub Repository)" & vbCrLf
    For Each k In remoteMap.Keys
        out = out & "  +-- " & remoteMap(k) & vbCrLf
    Next
    VBASync_GitHub_ToTree = out
End Function

Public Function VBASync_RunGitDiff(ByVal owner As String, ByVal repo As String) As String
    Dim remoteMap As Object: Set remoteMap = ParseGitHubAPI(DownloadURL("https://api.github.com/repos/" & owner & "/" & repo & "/contents/"))
    Dim localMap As Object: Set localMap = GetLocalMap()
    Dim out As String: out = "=== GIT DIFF: " & repo & " ===" & vbCrLf
    Dim k As Variant
    For Each k In remoteMap.Keys
        If Not localMap.Exists(k) Then out = out & "[REMOTE ONLY] " & remoteMap(k) & vbCrLf
    Next
    For Each k In localMap.Keys
        If Not remoteMap.Exists(k) Then out = out & "[LOCAL ONLY]  " & k & vbCrLf
    Next
    VBASync_RunGitDiff = out
End Function

Public Function ParseGitHubAPI(ByVal json As String) As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    Dim parts() As String
    Dim i As Long, f As String
    
    parts = VBASync_SplitAndClean(json, """name"":""")
    
    For i = 1 To UBound(parts)
        f = Split(parts(i), """")(0)
        If InStr(f, ".") > 0 Then
            d(LCase(Split(f, ".")(0))) = f
        End If
    Next i
    Set ParseGitHubAPI = d
End Function

' =========================================================
' LEGACY MERGED UTILITIES (Split_Join)
' =========================================================
Public Sub VBASync_ExportAllToChunks(ByVal folderPath As String)
    Dim comp As Object
    Dim finalPath As String
    If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"
    
    For Each comp In ThisWorkbook.VBProject.VBComponents
        If IsExportable(comp) Then
            finalPath = folderPath & VBASync_SanitizeName(comp.Name) & ".txt"
            WriteFile finalPath, StripVBHeaders(comp)
        End If
    Next comp
    MsgBox "Modules exported as text chunks to: " & folderPath, vbInformation
End Sub

' =========================================================
' UPDATED FILE ASSEMBLY TOOL
' =========================================================
Public Sub AssembleFilesToText()
    Dim fd As Object, totalContent As String, fPath As Variant
    Dim fName As String, fContent As String, outputPath As String
    Dim totalCount As Long, currentCount As Long
    
    Set fd = Application.FileDialog(3) ' msoFileDialogFilePicker
    With fd
        .AllowMultiSelect = True
        .Title = "Select files to assemble for manual paste"
        .Filters.Clear
        .Filters.Add "VBA Files", "*.bas; *.cls; *.frm"
        
        If .Show <> -1 Then Exit Sub
        
        ' 1. Determine total selected
        totalCount = .SelectedItems.count
        currentCount = 0
        
        ' 2. Process each file with the X/TOTAL header
        For Each fPath In .SelectedItems
            currentCount = currentCount + 1
            fName = Mid(fPath, InStrRev(fPath, "\") + 1)
            fContent = ReadFile(CStr(fPath))
            
            ' Matches your strict format: FILE X/TOTAL: Name
            totalContent = totalContent & "FILE " & currentCount & "/" & totalCount & ": " & fName & vbCrLf & _
                           fContent & vbCrLf & vbCrLf
        Next fPath
        
        ' 3. Output and Copy
        outputPath = Left(CStr(.SelectedItems(1)), InStrRev(CStr(.SelectedItems(1)), "\")) & "Assembled_For_Paste.txt"
        WriteFile outputPath, totalContent
        
        ' Try to open in Notepad
        On Error Resume Next
        shell "notepad.exe " & outputPath, vbNormalFocus
        On Error GoTo 0

        If MsgBox("Files assembled with " & totalCount & " chunks." & vbCrLf & _
                  "Copy full text to clipboard now?", vbQuestion + vbYesNo) = vbYes Then
            CopyTextToClipboard totalContent
        End If
    End With
End Sub
Private Function VBASync_SplitAndClean(ByVal Text As String, ByVal Delimiter As String) As Variant
    Dim rawArr() As String, cleanArr() As String
    Dim i As Long, count As Long
    
    If Text = "" Then
        VBASync_SplitAndClean = Array()
        Exit Function
    End If
    
    rawArr = Split(Text, Delimiter)
    ReDim cleanArr(0 To UBound(rawArr))
    count = 0
    For i = LBound(rawArr) To UBound(rawArr)
        If Trim(rawArr(i)) <> "" Then
            cleanArr(count) = Trim(rawArr(i))
            count = count + 1
        End If
    Next i
    
    If count > 0 Then
        ReDim Preserve cleanArr(0 To count - 1)
        VBASync_SplitAndClean = cleanArr
    Else
        VBASync_SplitAndClean = Array()
    End If
End Function

Private Function VBASync_SanitizeName(ByVal Name As String) As String
    Dim badChars As Variant, c As Variant: badChars = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    VBASync_SanitizeName = Name
    For Each c In badChars
        VBASync_SanitizeName = Replace(VBASync_SanitizeName, c, "_")
    Next c
    VBASync_SanitizeName = Trim(VBASync_SanitizeName)
End Function

' =========================================================
' PRIVATE HELPERS
' =========================================================
Private Sub EnsureCache()
    If HashCache Is Nothing Then Set HashCache = CreateObject("Scripting.Dictionary")
End Sub

Public Function HasChanged(ByVal Name As String, ByVal Code As String) As Boolean
EnsureCache:     Dim h As String: h = ComputeHash(Code)
    If Not HashCache.Exists(Name) Then
        HashCache(Name) = h
        HasChanged = True
        Exit Function
    End If
    HasChanged = (HashCache(Name) <> h)
    If HasChanged Then HashCache(Name) = h
End Function

Private Function ComputeHash(ByVal s As String) As String
    Dim i As Long, h As Double
    For i = 1 To Len(s): h = h + Asc(Mid(s, i, 1)): Next i
    ComputeHash = CStr(Len(s)) & "_" & CStr(h)
End Function

Private Sub ExportCleanSource(ByVal folderPath As String)
    Dim comp As Object: If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"
    For Each comp In ThisWorkbook.VBProject.VBComponents
        If IsExportable(comp) Then WriteFile folderPath & comp.Name & GetExt(comp.Type), StripVBHeaders(comp)
    Next
End Sub

Private Sub SplitGitHubURL(ByVal url As String, ByRef outOwner As String, ByRef outRepo As String)
    Dim clean As String, p() As String
    clean = Replace(Replace(url, "https://github.com/", ""), "github.com/", "")
    p = Split(clean, "/")
    If UBound(p) >= 1 Then
        outOwner = p(0)
        outRepo = p(1)
    End If
End Sub

Private Function GetLocalMap() As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary"): d.CompareMode = 1
    Dim c As Object: For Each c In ThisWorkbook.VBProject.VBComponents
        If c.Type <= 3 Then d(LCase(c.Name)) = True
    Next
    Set GetLocalMap = d
End Function

Private Sub RemoveExistingSafe(ByVal Name As String)
    On Error Resume Next: Dim comp As Object: Set comp = ThisWorkbook.VBProject.VBComponents(Name)
    If Not comp Is Nothing Then ThisWorkbook.VBProject.VBComponents.Remove comp
End Sub

Private Function DownloadURL(ByVal url As String) As String
    Dim http As Object: Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Open "GET", url, False: http.SetRequestHeader "User-Agent", "VBA-Sync": http.Send: DownloadURL = http.ResponseText
End Function

Private Function PickFolder() As String
    Dim fd As Object: Set fd = Application.FileDialog(4): If fd.Show = -1 Then PickFolder = fd.SelectedItems(1) & "\"
End Function

Private Function IsExportable(c As Object) As Boolean: IsExportable = (c.Type >= 1 And c.Type <= 3): End Function

Private Function GetExt(t As Long) As String
    Select Case t: Case 1: GetExt = ".bas": Case 2: GetExt = ".cls": Case 3: GetExt = ".frm": End Select
End Function

Private Sub WriteFile(p As String, t As String): Dim f As Integer: f = FreeFile: Open p For Output As #f: Print #f, t: Close f: End Sub

Private Function ReadFile(p As String) As String
    Dim f As Integer: f = FreeFile: Open p For Input As #f: ReadFile = Input$(LOF(f), f): Close f
End Function

Private Function StripVBHeaders(c As Object) As String
    On Error Resume Next: StripVBHeaders = c.CodeModule.lines(1, c.CodeModule.CountOfLines)
End Function

Private Function BuildQueue(ByVal folderPath As String) As Collection
    Dim q As New Collection, f As String, item As Object
    f = Dir(folderPath & "*.*")
    Do While f <> ""
        If InStr(f, ".") > 0 Then
            Set item = CreateObject("Scripting.Dictionary")
            item("FilePath") = folderPath & f
            item("ComponentName") = Left(f, InStrRev(f, ".") - 1)
            item("Code") = ReadFile(folderPath & f)
            q.Add item
        End If
        f = Dir
    Loop
    Set BuildQueue = q
End Function

Private Sub CopyTextToClipboard(ByVal txt As String)
    On Error Resume Next: Dim DataObj As Object
    Set DataObj = CreateObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    DataObj.SetText txt: DataObj.PutInClipboard
End Sub

