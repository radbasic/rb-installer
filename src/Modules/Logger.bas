Attribute VB_Name = "Logger"
' RAD Basic Installer
' Copyright (c) 2019-2025 by RAD Basic Team. All rights reserved.
' Licensed under the MIT License. See License.txt in the project root for license information.

Option Explicit

' General routine for logging errors '
Public Sub LogError(ProcName$, ErrNum&, ErrorMsg$)
  On Error GoTo ErrHandler
  Dim nUnit As Integer
  Dim logFolder As String
  nUnit = FreeFile
  
  ' Ensure path exists
  logFolder = Environ$("LOCALAPPDATA") & "\radbasic\log"
  OsUtils.EnsureFolderTree logFolder
  
  Open logFolder & "\rb-installer.log" For Append As nUnit
  Print #nUnit, Format$(Now) & " " & ProcName$ & " [Error] ErrNum: " & ErrNum & ". ErrorMessage: " & ErrorMsg
  Close nUnit
  Exit Sub

ErrHandler:
  'Failed to write log for some reason.'
  'Show MsgBox so error does not go unreported '
  MsgBox "Error in " & ProcName & vbNewLine & _
    ErrNum & ", " & ErrorMsg
End Sub

' General routine for logging errors '
Public Sub LogInfo(ProcName$, LogMsg$)
    On Error GoTo ErrHandler
    Dim nUnit As Integer
    Dim logFolder As String
    nUnit = FreeFile
    
    ' Ensure path exists
    logFolder = Environ$("LOCALAPPDATA") & "\radbasic\log"
    OsUtils.EnsureFolderTree logFolder
    
    Open logFolder & "\rb-installer.log" For Append As nUnit
    Print #nUnit, Format$(Now) & " " & ProcName$ & " [Information] " & LogMsg
    Close nUnit
    Exit Sub

ErrHandler:
    'Failed to write log for some reason.'
    'Show MsgBox so error does not go unreported '
    MsgBox "Error in " & ProcName & vbNewLine & _
    Err.Number & ", " & Err.Description
End Sub
