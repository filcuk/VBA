Attribute VB_Name = "modDownloadFile"
Option Explicit

Declare PtrSafe Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" ( _
    ByVal pCaller As Long, _
    ByVal szURL As String, _
    ByVal szFileName As String, _
    ByVal dwReserved As Long, _
    ByVal lpfnCB As Long _
) As Long

Function DownloadFile( _
    pURL As String, _
    pLocalPath As String, _
    Optional pOverwrite As Boolean = False _
) As Boolean

Dim WinHttpReq As Object
Dim oStream As Object
Dim Success As Boolean

    Success = False
    Set WinHttpReq = CreateObject("Microsoft.XMLHTTP")
    WinHttpReq.Open "GET", pURL, False, "", "" ', ("username", "password")
    WinHttpReq.send
    
    If WinHttpReq.Status = 200 Then
        Set oStream = CreateObject("ADODB.Stream")
        oStream.Open
        oStream.Type = 1
        oStream.Write WinHttpReq.responseBody
        oStream.SaveToFile pLocalPath, IIf(pOverwrite, 2, 1)
        oStream.Close
        Success = True
    End If

DownloadFile_Exit:
    DownloadFile = Success
    Exit Function
DownloadFile_Err:
    Success = False
    Resume DownloadFile_Exit
End Function
