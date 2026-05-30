Attribute VB_Name = "InstallEngine"
' RAD Basic Installer
' Copyright (c) 2019-2026 by RAD Basic Team. All rights reserved.
' Licensed under the MIT License. See License.txt in the project root for license information.

Option Explicit

Private Const RELATIVE_WORK_FOLDER As String = "\Temp\radbasic-installer"
Private Const RELATIVE_UNPACK_FOLDER As String = "\extract-pkg"

Private Const BASE_REGISTRY_PATH As String = "SOFTWARE\WOW6432Node\RAD Basic"
Private Const REGKEY_NAME_INSTALL_DIR As String = "InstallDir"
Private Const REGKEY_NAME_INSTALLER As String = "Installer"
Private Const REGKEY_VALUE_INSTALLER As String = "RB-Installer"

' Install the last version into the target dir as new installation
Public Function InstallVersion(targetDir As String, newInstall As Boolean) As Boolean
    Dim TmpUnpackFolder As String
    Dim zipFile As String
    Dim resultOK As Boolean
    Dim IDEExecPath As String
    Dim baseWorkFolder As String, baseUnpackFolder As String
    
    baseWorkFolder = Environ$("LOCALAPPDATA") & RELATIVE_WORK_FOLDER
    baseUnpackFolder = baseWorkFolder & RELATIVE_UNPACK_FOLDER
    
    resultOK = True
    
    ' Step 1: Check if parent tmp folder exists for download packages exists
    FrmIniSetup.SetStep "Creating tmp directories...", 10
    If Dir(baseUnpackFolder, vbDirectory) = "" Then
        OsUtils.CreateFolderTree baseUnpackFolder
    End If
    
    ' Step 2: Create a tmp dir to unpack packages into it
    TmpUnpackFolder = OsUtils.CreateTempFolder(baseUnpackFolder)
    
    ' Step 3: Download nightly
    FrmIniSetup.SetStep "Downloading packages...", 25
    ModDownloader.DownloadPkg baseWorkFolder
    zipFile = baseWorkFolder & "\radbasic-core.zip"
    
    ' Step 4: Unpack downloaded packages
    FrmIniSetup.SetStep "Unpacking files...", 65
    ZipSupport.UnzipWithPowerShell zipFile, TmpUnpackFolder
    
    ' Step 5: Ensure exists target dir
    If Dir(targetDir, vbDirectory) = "" Then
        OsUtils.CreateFolderTree targetDir
    End If
    
    ' Step 6: Copy all to dst dir
    FrmIniSetup.SetStep "Copying files...", 85
    Call CopyFolderAPI(TmpUnpackFolder & "\", targetDir)
    
    If (newInstall) Then
        ' Step 7: Create/update Registry keys
        FrmIniSetup.SetStep "Modifying Windows Registry...", 90
        resultOK = resultOK And Registry_WriteString(HKEY_LOCAL_MACHINE, BASE_REGISTRY_PATH, REGKEY_NAME_INSTALL_DIR, targetDir)
        resultOK = resultOK And Registry_WriteString(HKEY_LOCAL_MACHINE, BASE_REGISTRY_PATH, REGKEY_NAME_INSTALLER, REGKEY_VALUE_INSTALLER)
        
        ' Step 8: Start Menu & Desktop links
        FrmIniSetup.SetStep "Creating shortcuts...", 100
        IDEExecPath = targetDir & "\bin\rbide.exe"
        resultOK = resultOK And OsUtils.CreateStartMenuShortcut("RAD Basic IDE", IDEExecPath, "", "", "", True)
    Else
        FrmIniSetup.SetStep "Updating packages...", 100
    End If
    
    InstallVersion = resultOK
End Function
