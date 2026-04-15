'---------------------------------------------------------------------------------------
' Module    : VBASyncEngine
' Version   : 3.5.9
' Purpose   : Unified Orchestrator. Fixed: Single-line Sub syntax errors.
' Dependencies: VBASyncGitHub, VBASyncECS, VBASyncImport
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
        "VBASync Engine v3.5.7" & vbCrLf & vbCrLf & _
        "1 - Full Local Sync (Export & Import)" & vbCrLf & _
        "2 - Incremental Build (Smart Import)" & vbCrLf & _
        "3 - Export Only (Clean Source)" & vbCrLf & _
        "4 - Import Only (From Folder)" & vbCrLf & _
        "5 - GitHub Tree (Remote View)" & vbCrLf & _
        "6 - GitHub Diff (Remote vs Local)" & vbCrLf & _
        "7 - Assemble Files for Manual Paste" & vbCrLf & _
        "8 - Export AI-Ready Message Chunks" & vbCrLf & _
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
            url = InputBox("Enter GitHub URL or Owner/Repo:")
            If url <> "" Then
                VBASyncGitHub.SplitGitHubURL url, owner, repo
                Debug.Print VBASyncGitHub.VBASync_GitHub_ToTree(owner, repo)
                MsgBox "Tree output to Immediate Window (Ctrl+G)", vbInformation
            End If
        Case "6"
            url = InputBox("Enter GitHub URL or Owner/Repo:")
            If url <> "" Then
                VBASyncGitHub.SplitGitHubURL url, owner, repo
                MsgBox VBASyncGitHub.VBASync_RunGitDiff(owner, repo)
            End If
        Case "7"
            AssembleFilesToText
        Case "8"
            folderPath = PickFolder()
            If folderPath <> "" Then VBASync_ExportForAI folderPath
    End Select
End Sub

' =========================================================
' AI MESSAGE CHUNKING
' =========================================================
Public Sub VBASync_ExportForAI(ByVal folderPath As String)
    Dim comp As Object
    Dim fullProjectCode As String
    Dim charLimit As Long: charLimit = 4000
    Dim totalChunks As Long, i As Long, startPos As Long
    
    If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"
    
    For Each comp In ThisWorkbook.VBProject.VBComponents
        If IsExportable(comp) Then
            fullProjectCode = fullProjectCode & "FILE: " & comp.Name & vbCrLf & _
                             String(20, "-") & vbCrLf & _
                             StripVBHeaders(comp) & vbCrLf & vbCrLf & "' [EOF]" & vbCrLf
        End If
    Next comp
    
    If Len(fullProjectCode) = 0 Then Exit Sub
    
    totalChunks = Abs(Int(-Len(fullProjectCode) / charLimit))
    startPos = 1
    
    For i = 1 To totalChunks
        Dim finalOutput As String
        finalOutput = "CHUNK " & i & "/" & totalChunks & ": (Project: " & ThisWorkbook.Name & ")" & vbCrLf & _
                      String(30, "=") & vbCrLf & _
                      Mid(fullProjectCode, startPos, charLimit)
        
        WriteFile folderPath & "Message_Part_" & format(i, "00") & ".txt", finalOutput
        startPos = startPos + charLimit
    Next i
    
    On Error Resume Next
    shell "explorer.exe " & Chr(34) & folderPath & Chr(34), vbNormalFocus
    MsgBox "Created " & totalChunks & " AI-ready chunks.", vbInformation
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
    Set queue = VBASyncECS.VBASyncECS_OrderByDependencies(queue)
    ProcessImportQueue queue
End Sub

Private Sub ProcessImportQueue(ByVal q As Collection)
    Dim i As Long
    For i = 1 To q.count
        ProcessSingle q(i)
        DoEvents
    Next i
End Sub

' =========================================================
' IMPORT LOGIC (FIXED: Forces Class vs Module Folder)
' =========================================================
Private Sub ProcessSingle(ByRef item As Object)
    On Error GoTo fail
    Dim compName As String: compName = item("ComponentName")
    Dim compCode As String: compCode = item("Code")
    Dim newComp As Object
    Dim compType As Long
    
    ' 1. Determine Type based on the original file extension stored in BuildQueue
    ' We assume .cls = 2 (Class), .bas = 1 (Module)
    If item.Exists("Extension") Then
        Select Case LCase(item("Extension"))
            Case ".cls": compType = 2
            Case ".frm": compType = 3
            Case Else: compType = 1
        End Select
    Else
        ' Fallback: If extension isn't in dictionary, default to Module
        compType = 1
    End If
    
    ' 2. Remove existing to prevent "Name1" conflicts
    RemoveExistingSafe compName
    
    ' 3. Create a fresh component of the CORRECT type
    Set newComp = ThisWorkbook.VBProject.VBComponents.Add(compType)
    newComp.Name = compName
    
    ' 4. Inject the clean code
    ' Note: We skip line 1 because 'Add' might insert an 'Option Explicit' automatically
    newComp.CodeModule.DeleteLines 1, newComp.CodeModule.CountOfLines
    newComp.CodeModule.AddFromString compCode
    
    Exit Sub
fail:
    Debug.Print "Failed to import: " & compName & " - " & Err.Description
End Sub

' =========================================================
' LOCAL FILE UTILITIES
' =========================================================
Public Sub AssembleFilesToText()
    Dim fd As Object, totalContent As String, fPath As Variant
    Dim fName As String, totalCount As Long, currentCount As Long
    
    Set fd = Application.FileDialog(3)
    With fd
        .AllowMultiSelect = True
        If .Show <> -1 Then Exit Sub
        totalCount = .SelectedItems.count
        For Each fPath In .SelectedItems
            currentCount = currentCount + 1
            fName = Mid(fPath, InStrRev(fPath, "\") + 1)
            totalContent = totalContent & "FILE " & currentCount & "/" & totalCount & ": " & fName & vbCrLf & _
                           ReadFile(CStr(fPath)) & vbCrLf & vbCrLf
        Next fPath
        WriteFile .SelectedItems(1) & "_Assembled.txt", totalContent
        MsgBox "Files assembled and saved.", vbInformation
    End With
End Sub

Private Sub ExportCleanSource(ByVal folderPath As String)
    Dim comp As Object
    If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"
    
    For Each comp In ThisWorkbook.VBProject.VBComponents
        If IsExportable(comp) Then
            WriteFile folderPath & comp.Name & GetExt(comp.Type), StripVBHeaders(comp)
        End If
    Next comp
End Sub

Private Function BuildQueue(ByVal folderPath As String) As Collection
    Dim q As New Collection, f As String, item As Object
    Dim ext As String
    
    f = Dir(folderPath & "*.*")
    Do While f <> ""
        If InStr(f, ".") > 0 Then
            ext = Mid(f, InStrRev(f, "."))
            Set item = CreateObject("Scripting.Dictionary")
            item("ComponentName") = Left(f, InStrRev(f, ".") - 1)
            item("Extension") = ext ' <--- ADD THIS LINE
            item("Code") = ReadFile(folderPath & f)
            q.Add item
        End If
        f = Dir
    Loop
    Set BuildQueue = q
End Function

' =========================================================
' HELPERS (CLEANED)
' =========================================================
Private Function PickFolder() As String
    Dim fd As Object: Set fd = Application.FileDialog(4)
    If fd.Show = -1 Then PickFolder = fd.SelectedItems(1) & "\"
End Function

Private Function IsExportable(c As Object) As Boolean
    IsExportable = (c.Type >= 1 And c.Type <= 3)
End Function

Private Function GetExt(t As Long) As String
    Select Case t
        Case 1: GetExt = ".bas"
        Case 2: GetExt = ".cls"
        Case 3: GetExt = ".frm"
    End Select
End Function

Private Sub WriteFile(p As String, t As String)
    Dim f As Integer: f = FreeFile
    Open p For Output As #f: Print #f, t: Close f
End Sub

Private Function ReadFile(p As String) As String
    Dim f As Integer: f = FreeFile
    Open p For Input As #f: ReadFile = Input$(LOF(f), f): Close f
End Function

Private Function StripVBHeaders(c As Object) As String
    On Error Resume Next
    StripVBHeaders = c.CodeModule.lines(1, c.CodeModule.CountOfLines)
End Function

' FIXED: Multi-line structure for syntax safety
Private Sub RemoveExistingSafe(ByVal Name As String)
    On Error Resume Next
    Dim comp As Object
    Set comp = ThisWorkbook.VBProject.VBComponents(Name)
    If Not comp Is Nothing Then
        ThisWorkbook.VBProject.VBComponents.Remove comp
    End If
End Sub

