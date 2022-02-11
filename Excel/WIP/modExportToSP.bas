Attribute VB_Name = "modExportToSP"
Option Explicit

Private Const mc As String = "https://cps365ltd.sharepoint.com/sites/planner2022/Export/"

Private Sub Sample()
Dim Success As Boolean

    Success = SharePoint_UploadFile("C:\Users\FilipK\Desktop\Export\test.txt", "https://cps365ltd.sharepoint.com/sites/planner2022/", "Export", "test.txt")
    MsgBox Success
End Sub


Function SharePoint_UploadFile( _
    ByRef pLocalFile As String, _
    ByRef pSPSite As String, _
    ByRef pSPLibrary As String, _
    ByRef pSPFileName As String _
)

Const spCONTENT_TYPE = "0x000000000000000000000000000000000000000"

Dim ObjectStream As Object: Set ObjectStream = CreateObject("ADODB.Stream")
Dim ObjectDOM As Object: Set ObjectDOM = CreateObject("Microsoft.XMLDOM")
Dim ObjectElement As Object: Set ObjectElement = ObjectDOM.createElement("TMP")
Dim ObjectHTTP As Object: Set ObjectHTTP = CreateObject("Microsoft.XMLHTTP")
Dim BinaryFile
Dim EncodedFile
Dim strURLService As String
Dim strSOAPAction As String
Dim strSOAPCommand As String
Dim Success As Boolean

    'Reading binary file
    ObjectStream.Open
    ObjectStream.Type = 1 'Type Binary
    ObjectStream.LoadFromFile (pLocalFile)
    BinaryFile = ObjectStream.Read()
    ObjectStream.Close
    
    'Conversion Base64
    ObjectElement.DataType = "bin.base64" 'Type Base64
    ObjectElement.nodeTypedValue = BinaryFile
    EncodedFile = ObjectElement.Text
    
    'Build request to load document
    strURLService = pSPSite + "_vti_bin/copy.asmx"
    strSOAPAction = "http://schemas.microsoft.com/sharepoint/soap/CopyIntoItems"
    strSOAPCommand = "<?xml version='1.0' encoding='utf-8'?>" & _
    "<soap:Envelope xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xmlns:xsd='http://www.w3.org/2001/XMLSchema' xmlns:soap='http://schemas.xmlsoap.org/soap/envelope/'>" & _
    "<soap:Body>" & _
    "<CopyIntoItems xmlns='http://schemas.microsoft.com/sharepoint/soap/'>" & _
    "<SourceUrl>" + pLocalFile + "</SourceUrl>" & _
    "<DestinationUrls>" & _
    "<string>" + pSPSite + pSPLibrary + "/" + pSPFileName + "</string>" & _
    "</DestinationUrls>" & _
    "<Fields>" & _
    "<FieldInformation Type='Text' InternalName='Title' DisplayName='Title' Value='this is the title value' />" & _
    "<FieldInformation Type='Choice' InternalName='Our_x0020_Status' DisplayName='Our Document Status' Value='Ready-to-distribute' />" & _
    "<FieldInformation Type='Text' InternalName='ContentTypeId' DisplayName='Content Type ID' Value='" + spCONTENT_TYPE + "' />" & _
    "</Fields>" & _
    "<Stream>" + EncodedFile + "</Stream>" & _
    "</CopyIntoItems>" & _
    "</soap:Body>" & _
    "</soap:Envelope>"
    
    ObjectHTTP.Open "Get", strURLService, False
    ObjectHTTP.setRequestHeader "Content-Type", "text/csv; charset=utf-8"
    ObjectHTTP.setRequestHeader "SOAPAction", strSOAPAction
    ObjectHTTP.send strSOAPCommand
    
    Success = True
    
SharePoint_UploadFile_Exit:
    SharePoint_UploadFile = Success
    Exit Function
SharePoint_UploadFile_Error:
    Success = False
    Resume SharePoint_UploadFile_Exit
End Function

