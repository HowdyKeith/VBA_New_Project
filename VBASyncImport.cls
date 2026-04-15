VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VBASyncImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Class    : VBASyncImport.cls
' Version   : 3.4.0
' Purpose   :
'---------------------------------------------------------------------------------------
Option Explicit

Public FilePath As String
Public ComponentName As String
Public Code As String
Public Extension As String
Public Attempt As Long
Public OriginalIndex As Long
