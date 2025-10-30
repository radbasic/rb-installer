Attribute VB_Name = "ReqValidator"
' RAD Basic Installer
' Copyright (c) 2019-2025 by RAD Basic Team. All rights reserved.
' Licensed under the MIT License. See License.txt in the project root for license information.

Option Explicit


Private Const REG_KEY_PATH_x86 As String = "HKEY_LOCAL_MACHINE\SOFTWARE\RAD Basic"
Private Const REG_KEY_PATH_x64 As String = "HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\RAD Basic"
Private Const REG_KEY_NAME_INSTALLDIR As String = "InstallDir"
Private Const REG_KEY_NAME_INSTALLER As String = "Installer"
Private Const REG_KEY_NAME_INSTALLER_VALUE As String = "RB-Installer"


Public Function IsOldRADBasicInstalled() As Boolean
    Dim OldInstallDir As String
    
    #If Win64 Then
        OldInstallDir = OsUtils.Registry_Read(REG_KEY_PATH_x64, REG_KEY_NAME_INSTALLDIR)
    #Else
        OldInstallDir = OsUtils.Registry_Read(REG_KEY_PATH_x86, REG_KEY_NAME_INSTALLDIR)
    #End If
    IsOldRADBasicInstalled = True

End Function


Public Function IsNewRADBasicInstalled() As Boolean
    Dim InstallerValue As String

    InstallerValue = OsUtils.Registry_Read(REG_KEY_PATH_x86, REG_KEY_NAME_INSTALLER)
    IsNewRADBasicInstalled = (InstallerValue = REG_KEY_NAME_INSTALLER_VALUE)

End Function
