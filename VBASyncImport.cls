'---------------------------------------------------------------------------------------
' Class     : VBASyncImport
' Version   : 3.4.0
' Purpose   : Data container for queued VBA components during Import/Sync.
'---------------------------------------------------------------------------------------
Option Explicit

Public FilePath As String
Public ComponentName As String
Public Code As String
Public Extension As String
Public Attempt As Long
Public OriginalIndex As Long
