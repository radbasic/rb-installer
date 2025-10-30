Attribute VB_Name = "OsUtils"
' RAD Basic Installer
' Copyright (c) 2019-2025 by RAD Basic Team. All rights reserved.
' Licensed under the MIT License. See License.txt in the project root for license information.
' Note: There are different implementation ways (as example: Windows Registry Integration)
'       as a purpose of RAD Basic feature testing.
Option Explicit

' --- WinAPI Types --- '
Private Type SHFILEOPSTRUCT
    hwnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAnyOperationsAborted As Long
    hNameMappings As Long
    lpszProgressTitle As String
End Type

' --- WinAPI Private CONSTS --- '
Private Const FO_COPY = &H2
Private Const FOF_NOCONFIRMATION = &H10
Private Const FOF_NOCONFIRMMKDIR = &H200
Private Const FOF_SILENT = &H4

Private Const REG_OPTION_NON_VOLATILE = 0
Private Const KEY_SET_VALUE = &H2
Private Const KEY_CREATE_SUB_KEY = &H4
Private Const REG_SZ = 1

Private Const ERROR_SUCCESS = 0

' --- WinAPI Public CONSTS --- '
Public Const HKEY_CLASSES_ROOT     As Long = &H80000000
Public Const HKEY_CURRENT_USER     As Long = &H80000001
Public Const HKEY_LOCAL_MACHINE    As Long = &H80000002
Public Const HKEY_USERS            As Long = &H80000003
Public Const HKEY_CURRENT_CONFIG   As Long = &H80000005

' --- WinAPI FUNCS --- '
Private Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" ( _
    ByVal hKey As Long, _
    ByVal lpSubKey As String, _
    ByVal Reserved As Long, _
    ByVal lpClass As String, _
    ByVal dwOptions As Long, _
    ByVal samDesired As Long, _
    ByVal lpSecurityAttributes As Long, _
    phkResult As Long, _
    lpdwDisposition As Long) As Long

Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" ( _
    ByVal hKey As Long, _
    ByVal lpValueName As String, _
    ByVal Reserved As Long, _
    ByVal dwType As Long, _
    ByVal lpData As String, _
    ByVal cbData As Long) As Long

Private Declare Function RegCloseKey Lib "advapi32.dll" ( _
    ByVal hKey As Long) As Long

' Implementation of registy read
' It uses "WScript.Shell"
Public Function Registry_Read(Key_Path As String, Key_Name As String) As String

    On Error Resume Next
    
    Dim Registry As Object
    Set Registry = CreateObject("WScript.Shell")
    Registry_Read = Registry.RegRead(Key_Path & "\" & Key_Name)
    
End Function

' Implementation of registy write
' It uses WinAPI calls
Public Function Registry_WriteString( _
    ByVal rootKey As Long, _
    ByVal subKey As String, _
    ByVal valueName As String, _
    ByVal valueData As String) As Boolean
    
    Dim hKey As Long
    Dim disposition As Long
    Dim ret As Long
    
    ' Open for create/update the key (not needed to find if exists)
    ret = RegCreateKeyEx(rootKey, subKey, 0, vbNullString, REG_OPTION_NON_VOLATILE, _
                         KEY_SET_VALUE Or KEY_CREATE_SUB_KEY, 0, hKey, disposition)
    
    If ret <> ERROR_SUCCESS Then
        Registry_WriteString = False
        Exit Function
    End If
    
    ' Supports only REG_SZ values
    ret = RegSetValueEx(hKey, valueName, 0, REG_SZ, valueData, Len(valueData) + 1)
    
    RegCloseKey hKey
    
    Registry_WriteString = (ret = ERROR_SUCCESS)

End Function

Public Sub CreateFolderTree(ByVal folderPath As String)
    Dim parts() As String
    Dim i As Long
    Dim currentPath As String
    
    parts = Split(folderPath, "\")
    currentPath = parts(0) & "\"   ' First parth is drive
    
    For i = 1 To UBound(parts)
        If parts(i) <> "" Then
            currentPath = currentPath & parts(i) & "\"
            If Dir(currentPath, vbDirectory) = "" Then
                MkDir currentPath
            End If
        End If
    Next i
End Sub

' Returns the path to temp folder created under parentFolder
Public Function CreateTempFolder(ByVal parentFolder As String) As String
    Dim tempFolder As String
    Dim suffix As Long
    
    ' Parent Folder must exists
    If Dir(parentFolder, vbDirectory) = "" Then
        ' Error!
        CreateTempFolder = ""
        Exit Function
        '' ERROR!!
       ' Exit Sub
    End If
    
    ' Generate a unique folder name (simple: Temp_ + número aleatorio)
    Randomize
    Do
        suffix = CLng(Rnd * 1000000)
        tempFolder = parentFolder & "\Temp_" & suffix
    Loop While Dir(tempFolder, vbDirectory) <> ""
    
    ' Create the child temp folder
    MkDir tempFolder
    
    CreateTempFolder = tempFolder
End Function

' THIS IS AN ERROR: "Only comments may appear after End Sub, End Function or End Property"
'Private Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long

'Private Type SHFILEOPSTRUCT
'    hwnd As Long
'    wFunc As Long
'    pFrom As String
'    pTo As String
'    fFlags As Integer
'    fAnyOperationsAborted As Long
'    hNameMappings As Long
'    lpszProgressTitle As String
'End Type
'
'Private Const FO_COPY = &H2
'Private Const FOF_NOCONFIRMATION = &H10
'Private Const FOF_NOCONFIRMMKDIR = &H200
'Private Const FOF_SILENT = &H4

Public Sub CopyFolderAPI(ByVal srcFolder As String, ByVal dstFolder As String)
    Dim op As SHFILEOPSTRUCT
    Dim result As Long
    
    If Right$(srcFolder, 1) <> "\" Then srcFolder = srcFolder & "\"
    If Right$(dstFolder, 1) <> "\" Then dstFolder = dstFolder & "\"
    
    op.wFunc = FO_COPY
    op.pFrom = srcFolder & "*" & vbNullChar & vbNullChar
    op.pTo = dstFolder & vbNullChar & vbNullChar
    op.fFlags = FOF_NOCONFIRMATION Or FOF_NOCONFIRMMKDIR Or FOF_SILENT
    
    result = SHFileOperation(op)
    
    If result <> 0 Then
        MsgBox "Error copiando carpeta. Código: " & result, vbCritical
    End If
End Sub

' -----------------------
' CreateStartMenuShortcut
' -----------------------
' name         -> nombre del acceso (sin .lnk)
' targetPath   -> ruta al ejecutable o archivo objetivo
' arguments    -> argumentos opcionales (puede ser "")
' iconPath     -> ruta al icono (puede ser "" para usar el icono del target)
' startIn      -> carpeta de trabajo (puede ser "")
' forAllUsers  -> True = "All Users" Start Menu, False = current user
' Returns True on success, False on failure
Public Function CreateStartMenuShortcut( _
    ByVal name As String, _
    ByVal targetPath As String, _
    ByVal arguments As String, _
    ByVal iconPath As String, _
    ByVal startIn As String, _
    ByVal forAllUsers As Boolean) As Boolean

    Dim shellObj As Object
    Dim shortcutPath As String
    Dim startMenuPrograms As String
    Dim folderPath As String

    On Error GoTo ErrHandler

    ' Validaciones básicas
    If Trim$(name) = "" Then Err.Raise vbObjectError + 1, , "Name required"
    If Dir(targetPath) = "" Then Err.Raise vbObjectError + 2, , "Target file not found: " & targetPath

    ' Seleccionar la carpeta Programs del Start Menu
    If forAllUsers Then
        ' All users Start Menu (ALLUSERSPROFILE suele apuntar a C:\ProgramData en Windows modernos)
        startMenuPrograms = Environ$("ALLUSERSPROFILE") & "\Microsoft\Windows\Start Menu\Programs"
    Else
        ' Current user Start Menu
        startMenuPrograms = Environ$("APPDATA") & "\Microsoft\Windows\Start Menu\Programs"
    End If

    ' Ruta completa del .lnk (por ejemplo: ...\Programs\My Company\MyApp.lnk)
    ' Si quieres crear en subcarpeta, modifica folderPath antes de llamar a EnsureFolderTree
    folderPath = startMenuPrograms & "\RAD Basic"
    EnsureFolderTree folderPath

    shortcutPath = folderPath & "\" & name & ".lnk"

    Set shellObj = CreateObject("WScript.Shell")
    Dim sc As Object
    Set sc = shellObj.CreateShortcut(shortcutPath)

    sc.targetPath = targetPath
    sc.arguments = arguments
    If Trim$(startIn) <> "" Then sc.WorkingDirectory = startIn Else sc.WorkingDirectory = Left$(targetPath, InStrRev(targetPath, "\") - 1)
    If Trim$(iconPath) <> "" Then
        ' IconLocation puede ser "path,to,iconindex" — con VB usamos "path,0" o solo path
        sc.IconLocation = iconPath
    Else
        sc.IconLocation = targetPath & ",0"
    End If
    sc.Description = name

    sc.Save

    CreateStartMenuShortcut = True
    Exit Function

ErrHandler:
    CreateStartMenuShortcut = False
    ' Opcional: mostrar error para depuración
    MsgBox "Error creating shortcut: " & Err.Number & " - " & Err.Description, vbExclamation
End Function


' ----------------------------------------------------------------------
' EnsureFolderTree: Ensure exists all folder path.
'                   Create folder and subfolders if they don't exists.
' ----------------------------------------------------------------------
Public Sub EnsureFolderTree(ByVal folderPath As String)
    Dim parts() As String
    Dim i As Long
    Dim currentPath As String

    If Trim$(folderPath) = "" Then Exit Sub

    ' Normalizar barras finales
    If Right$(folderPath, 1) <> "\" Then folderPath = folderPath & "\"

    parts = Split(folderPath, "\")
    ' El primer elemento suele ser la unidad (ej: "C:")
    currentPath = parts(0) & "\"

    For i = 1 To UBound(parts)
        If parts(i) <> "" Then
            currentPath = currentPath & parts(i) & "\"
            If Dir$(currentPath, vbDirectory) = "" Then
                On Error Resume Next
                MkDir currentPath
                On Error GoTo 0
            End If
        End If
    Next i
End Sub
