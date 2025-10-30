Attribute VB_Name = "ZipSupport"
' RAD Basic Installer
' Copyright (c) 2019-2025 by RAD Basic Team. All rights reserved.
' Licensed under the MIT License. See License.txt in the project root for license information.

Option Explicit

Public Sub UnzipWithPowerShell(ByVal zipPath As String, ByVal destFolder As String)
    Dim shellObj As Object
    Dim psCommand As String
    Dim ret As Long
    
    ' Object "Shell.Application" could fail in certain situations (UAC, session context, diferent privileges).
    ' So it is safer run the poweshell script.
    psCommand = "powershell -NoProfile -ExecutionPolicy Bypass -Command " & _
                 Chr$(34) & "Expand-Archive -Path '" & zipPath & "' -DestinationPath '" & destFolder & "' -Force" & Chr$(34)
    
    Debug.Print psCommand
    
    Set shellObj = CreateObject("WScript.Shell")
    ret = shellObj.Run(psCommand, 0, True) ' flags => 0=hidden, True=wait/sync
    
    If ret <> 0 Then
        MsgBox "There is an error while uncompressing the packages. Error code: " & ret, vbCritical
    End If
End Sub
