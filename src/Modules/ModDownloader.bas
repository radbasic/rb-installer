Attribute VB_Name = "ModDownloader"
' RAD Basic Installer
' Copyright (c) 2019-2025 by RAD Basic Team. All rights reserved.
' Licensed under the MIT License. See License.txt in the project root for license information.

Option Explicit

' Source of downloading a file using winapi: https://stackoverflow.com/questions/1976152/download-file-vb6
Private Declare Function URLDownloadToFile Lib "urlmon" _
   Alias "URLDownloadToFileA" _
  (ByVal pCaller As Long, _
   ByVal szURL As String, _
   ByVal szFileName As String, _
   ByVal dwReserved As Long, _
   ByVal lpfnCB As Long) As Long
   
Private Declare Function DeleteUrlCacheEntry Lib "Wininet.dll" _
   Alias "DeleteUrlCacheEntryA" _
  (ByVal lpszUrlName As String) As Long
   

Private Const ERROR_SUCCESS As Long = 0
Private Const BINDF_GETNEWESTVERSION As Long = &H10
Private Const INTERNET_FLAG_RELOAD As Long = &H80000000

' Channel could be:
'   * beta
'   * nightly
'   * snapshot
Private Const CURRENT_CHANNEL As String = "beta"

Public Function DownloadFile(sSourceUrl As String, _
                             sLocalFile As String) As Boolean
                             
    DeleteUrlCacheEntry sSourceUrl

    'Download the file. BINDF_GETNEWESTVERSION forces
    'the API to download from the specified source.
    'Passing 0& as dwReserved causes the locally-cached
    'copy to be downloaded, if available. If the API
    'returns ERROR_SUCCESS (0), DownloadFile returns True.
     DownloadFile = URLDownloadToFile(0&, _
                                      sSourceUrl, _
                                      sLocalFile, _
                                      BINDF_GETNEWESTVERSION, _
                                      0&) = ERROR_SUCCESS

End Function

Public Sub DownloadPkg(downloadFolder As String)
    Dim PkgUrl As String
    Dim LocalTempPath As String
    Dim result As Boolean
    Dim version As String
    
    If (CURRENT_CHANNEL = "beta") Then
        version = "0.11.0"  ' Version must be dinamically
    End If
    
    PkgUrl = "https://downloads.radbasic.dev/channels/" & CURRENT_CHANNEL
    If (version <> "") Then
        PkgUrl = PkgUrl & "/" & version
    End If
    PkgUrl = PkgUrl & "/radbasic-core.zip"
     
    LocalTempPath = downloadFolder & "\radbasic-core.zip"
    
    result = DownloadFile(PkgUrl, LocalTempPath)
    
    ' MsgBox result
    
End Sub

