Attribute VB_Name = "VersionManagement"
' RAD Basic Installer
' Copyright (c) 2019-2025 by RAD Basic Team. All rights reserved.
' Licensed under the MIT License. See License.txt in the project root for license information.

Option Explicit

Public Type AppVersionInfo
    Version As String
    Build   As Long
End Type

Private Declare Function GetFileVersionInfoSize Lib "version.dll" _
    Alias "GetFileVersionInfoSizeA" ( _
    ByVal lptstrFilename As String, _
    lpdwHandle As Long) As Long

Private Declare Function GetFileVersionInfo Lib "version.dll" _
    Alias "GetFileVersionInfoA" ( _
    ByVal lptstrFilename As String, _
    ByVal dwHandle As Long, _
    ByVal dwLen As Long, _
    lpData As Any) As Long

Private Declare Function VerQueryValue Lib "version.dll" _
    Alias "VerQueryValueA" ( _
    pBlock As Any, _
    ByVal lpSubBlock As String, _
    lpBuffer As Any, _
    puLen As Long) As Long

Private Type VS_FIXEDFILEINFO
    dwSignature        As Long
    dwStrucVersion     As Long
    dwFileVersionMS    As Long
    dwFileVersionLS    As Long
    dwProductVersionMS As Long
    dwProductVersionLS As Long
    dwFileFlagsMask    As Long
    dwFileFlags        As Long
    dwFileOS           As Long
    dwFileType         As Long
    dwFileSubtype      As Long
    dwFileDateMS       As Long
    dwFileDateLS       As Long
End Type

Private Declare Sub CopyMemory Lib "kernel32" _
    Alias "RtlMoveMemory" ( _
    Destination As Any, _
    Source As Any, _
    ByVal Length As Long)

Public Function GetInstalledVersion() As AppVersionInfo
    Dim InstalledVersion As AppVersionInfo
    Dim BinPath As String
    
    BinPath = GetInstalledDir() & "\bin\radbasic.dll"
    
    InstalledVersion = GetFileVersion(BinPath)
    
    ' Return
    GetInstalledVersion = InstalledVersion
End Function

Public Function GetFileVersion(ByVal filePath As String) As AppVersionInfo
    Dim size As Long
    Dim handle As Long
    Dim buffer() As Byte
    Dim ffi As VS_FIXEDFILEINFO
    Dim ffiPtr As Long
    Dim ffiLen As Long
    Dim InstVersion As AppVersionInfo

    ' Comprobar que el fichero existe
    If Dir$(filePath) = "" Then Exit Function

    size = GetFileVersionInfoSize(filePath, handle)
    If size = 0 Then Exit Function

    ReDim buffer(0 To size - 1)

    If GetFileVersionInfo(filePath, 0, size, buffer(0)) = 0 Then Exit Function

    If VerQueryValue(buffer(0), "\", ffiPtr, ffiLen) = 0 Then Exit Function

    ' Copiar la estructura desde el puntero
    CopyMemory ffi, ByVal ffiPtr, Len(ffi)

    ' Extraer Major.Minor.Revision.Build
    InstVersion.Version = _
        CStr((ffi.dwFileVersionMS \ &H10000) And &HFFFF) & "." & _
        CStr(ffi.dwFileVersionMS And &HFFFF) & "." & _
        CStr((ffi.dwFileVersionLS \ &H10000) And &HFFFF)
    InstVersion.Build = (ffi.dwFileVersionLS And &HFFFF)
        
    GetFileVersion = InstVersion
End Function

Public Function GetOnlineVersionInfo( _
    ByVal url As String, _
    ByRef info As AppVersionInfo) As Boolean

    Dim http As Object
    Dim json As String

    On Error GoTo ErrHandler

    Set http = CreateObject("MSXML2.XMLHTTP")

    http.Open "GET", url, False
    http.Send

    If http.Status <> 200 Then Exit Function

    json = Trim$(http.responseText)
    If json = "" Then Exit Function

    info.Version = JsonGetValue(json, "version")
    info.Build = Int(JsonGetValue(json, "build"))

    If info.Version = "" And info.Build = 0 Then Exit Function

    GetOnlineVersionInfo = True
    Exit Function

ErrHandler:
    GetOnlineVersionInfo = False
End Function

Private Function JsonGetValue( _
    ByVal json As String, _
    ByVal key As String) As String

    Dim pattern As String
    Dim p1 As Long, p2 As Long

    ' Buscar: "key":"value"
    pattern = """" & key & """:"""

    p1 = InStr(1, json, pattern, vbTextCompare)
    If p1 = 0 Then Exit Function

    p1 = p1 + Len(pattern)
    p2 = InStr(p1, json, """")

    If p2 > p1 Then
        JsonGetValue = Mid$(json, p1, p2 - p1)
    End If
End Function

