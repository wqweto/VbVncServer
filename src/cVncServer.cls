VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cVncServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=========================================================================
'
' VbVncServer Project (c) 2021 by wqweto@gmail.com
'
' Simple VNC Server using DXGI Desktop Duplication
'
' This project is licensed under the terms of the MIT license
' See the LICENSE file in the project root for more information
'
'=========================================================================
Option Explicit
DefObj A-Z
Private Const MODULE_NAME As String = "cVncServer"

#Const ImplZlib = (VNC_NOZLIB = 0)
#Const ImplHistograms = (VNC_NOHISTOGRAMS = 0)
#Const ImplNoIdeProtection = (MST_NO_IDE_PROTECTION <> 0)
#Const ImplUseDebugLog = (USE_DEBUG_LOG <> 0)
#Const ImplUseShared = (VNC_USESHARED <> 0)

'=========================================================================
' API
'=========================================================================

'--- for OpenInputDesktop
Private Const GENERIC_READ                                  As Long = &H80000000
'--- for SetProcessDpiAwarenessContext
Private Const DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2    As Long = -4
'--- for D3DKMTSetProcessSchedulingPriorityClass
Private Const D3DKMT_SCHEDULINGPRIORITYCLASS_REALTIME       As Long = 5
'--- for mouse_event
Private Const MOUSEEVENTF_MOVE              As Long = &H1
Private Const MOUSEEVENTF_LEFTDOWN          As Long = &H2
Private Const MOUSEEVENTF_LEFTUP            As Long = &H4
Private Const MOUSEEVENTF_RIGHTDOWN         As Long = &H8
Private Const MOUSEEVENTF_RIGHTUP           As Long = &H10
Private Const MOUSEEVENTF_MIDDLEDOWN        As Long = &H20
Private Const MOUSEEVENTF_MIDDLEUP          As Long = &H40
Private Const MOUSEEVENTF_WHEEL             As Long = &H800
Private Const MOUSEEVENTF_ABSOLUTE          As Long = &H8000&
'--- for GetSystemMetrics
Private Const SM_SWAPBUTTON                 As Long = 23
'--- for RedrawWindow
Private Const RDW_INVALIDATE                As Long = &H1
Private Const RDW_ALLCHILDREN               As Long = &H80
Private Const RDW_ERASENOW                  As Long = &H200
'--- for IStream_Seek
Private Const STREAM_SEEK_SET               As Long = 0
Private Const STREAM_SEEK_CUR               As Long = 1
Private Const STREAM_SEEK_END               As Long = 2
'--- for CryptAcquireContext
Private Const PROV_RSA_FULL                 As Long = 1
Private Const CRYPT_VERIFYCONTEXT           As Long = &HF0000000
'--- for CryptImportKey
Private Const PLAINTEXTKEYBLOB              As Long = 8
Private Const CUR_BLOB_VERSION              As Long = 2
Private Const CALG_DES                      As Long = &H6601
Private Const CRYPT_EXPORTABLE              As Long = &H1
'--- for CryptSetKeyParam
Private Const KP_MODE                       As Long = 4
Private Const CRYPT_MODE_ECB                As Long = 2
'--- for GetDriveType
Private Const DRIVE_REMOVABLE               As Long = 2
Private Const DRIVE_FIXED                   As Long = 3
Private Const DRIVE_REMOTE                  As Long = 4
Private Const DRIVE_CDROM                   As Long = 5
'--- for SHGetFolderPath
Private Const CSIDL_DESKTOP                 As Long = &H0
Private Const CSIDL_PERSONAL                As Long = &H5
Private Const CSIDL_NETHOOD                 As Long = &H13
Private Const CSIDL_FLAG_CREATE             As Long = &H8000&
'--- for FindFirstFile
Private Const MAX_PATH                      As Long = 260
Private Const INVALID_HANDLE_VALUE          As Long = -1
'--- for SetErrorMode
Private Const SEM_FAILCRITICALERRORS        As Long = &H1
'--- for SHCreateStreamOnFile
Private Const STGM_READ                     As Long = 0
Private Const STGM_WRITE                    As Long = 1
Private Const STGM_CREATE                   As Long = &H1000
'--- for WideCharToMultiByte
Private Const CP_UTF8                       As Long = 65001
'--- for SendInput
Private Const KEYEVENTF_EXTENDEDKEY         As Long = 1
Private Const KEYEVENTF_KEYUP               As Long = 2
Private Const KEYEVENTF_UNICODE             As Long = 4
Private Const INPUT_KEYBOARD                As Long = 1
'--- for thunks
Private Const MEM_COMMIT                    As Long = &H1000
Private Const PAGE_EXECUTE_READWRITE        As Long = &H40
Private Const SIGN_BIT                      As Long = &H80000000

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function IsBadReadPtr Lib "kernel32" (ByVal lp As Long, ByVal ucb As Long) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, lpBuffer As Any) As Long
Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (nDrive As Any) As Long
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileW" (ByVal lpFileName As Long, ByVal lpFindFileData As Long) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileW" (ByVal hFindFile As Long, ByVal lpFindFileData As Long) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function SetErrorMode Lib "kernel32" (ByVal wMode As Long) As Long
Private Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileW" (ByVal lpFileName As Long) As Long
Private Declare Function RemoveDirectory Lib "kernel32" Alias "RemoveDirectoryW" (ByVal lpPathName As Long) As Long
Private Declare Function CreateDirectory Lib "kernel32" Alias "CreateDirectoryW" (ByVal lpPathName As Long, ByVal lpSecurityAttributes As Long) As Long
Private Declare Function MoveFileEx Lib "kernel32" Alias "MoveFileExW" (ByVal lpExistingFileName As Long, ByVal lpNewFileName As Long, ByVal dwFlags As Long) As Long
Private Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As Any, lpSystemTime As Any) As Long
Private Declare Function SystemTimeToVariantTime Lib "oleaut32" (lpSystemTime As Any, pvTime As Date) As Long
Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesW" (ByVal lpFileName As Long) As Long
Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long
Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long
'--- user32/gdi32
Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Private Declare Function SendInput Lib "user32" (ByVal nInputs As Long, pInputs As Any, ByVal cbSize As Long) As Long
Private Declare Function MapVirtualKey Lib "user32" Alias "MapVirtualKeyA" (ByVal uCode As Long, ByVal uMapType As Long) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function RedrawWindow Lib "user32" (ByVal hWnd As Long, lprcUpdate As Any, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Private Declare Function OpenInputDesktop Lib "user32" (ByVal dwFlags As Long, ByVal fInherit As Boolean, ByVal dwDesiredAccess As Long) As Long
Private Declare Function CloseDesktop Lib "user32" (ByVal hDesktop As Long) As Long
Private Declare Function SetThreadDesktop Lib "user32" (ByVal hDesktop As Long) As Long
Private Declare Function SetProcessDpiAwarenessContext Lib "user32" (ByVal lValue As Long) As Long
Private Declare Function D3DKMTSetProcessSchedulingPriorityClass Lib "gdi32" (ByVal hProcess As Long, ByVal lPriority As Long) As Long
'--- ole/oleaut/shell
Private Declare Sub CoTaskMemFree Lib "ole32" (ByVal pv As Long)
Private Declare Function DispCallFunc Lib "oleaut32" (ByVal pvInstance As Long, ByVal oVft As Long, ByVal lCc As Long, ByVal vtReturn As VbVarType, ByVal cActuals As Long, prgVt As Any, prgpVarg As Any, pvargResult As Variant) As Long
Private Declare Function SHCreateMemStream Lib "shlwapi" Alias "#12" (pInit As Any, ByVal cbInit As Long) As stdole.IUnknown
Private Declare Function SHCreateStreamOnFile Lib "shlwapi" Alias "SHCreateStreamOnFileW" (ByVal pszFile As Long, ByVal grfMode As Long, ppstm As IUnknown) As Long
Private Declare Function SHGetFolderPath Lib "shfolder" Alias "SHGetFolderPathW" (ByVal hWnd As Long, ByVal csidl As Long, ByVal hToken As Long, ByVal dwFlags As Long, ByVal szPath As Long) As Long
'--- msvbvm60
Private Declare Function ArrPtr Lib "msvbvm60" Alias "VarPtr" (Ptr() As Any) As Long
Private Declare Function vbaObjSetAddref Lib "msvbvm60" Alias "__vbaObjSetAddref" (oDest As Any, ByVal lSrcPtr As Long) As Long
'--- winsock
Private Declare Function ws_htonl Lib "ws2_32" Alias "htonl" (ByVal hostlong As Long) As Long
Private Declare Function ws_htons Lib "ws2_32" Alias "htons" (ByVal hostshort As Long) As Integer
'--- advapi32
Private Declare Function CryptAcquireContext Lib "advapi32" Alias "CryptAcquireContextW" (phProv As Long, ByVal pszContainer As Long, ByVal pszProvider As Long, ByVal dwProvType As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptReleaseContext Lib "advapi32" (ByVal hProv As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptEncrypt Lib "advapi32" (ByVal hKey As Long, ByVal hHash As Long, ByVal Final As Long, ByVal dwFlags As Long, pbData As Any, pdwDataLen As Long, dwBufLen As Long) As Long
Private Declare Function CryptImportKey Lib "advapi32" (ByVal hProv As Long, pbData As Any, ByVal dwDataLen As Long, ByVal hPubKey As Long, ByVal dwFlags As Long, phKey As Long) As Long
Private Declare Function CryptDestroyKey Lib "advapi32" (ByVal hKey As Long) As Long
Private Declare Function CryptSetKeyParam Lib "advapi32" (ByVal hKey As Long, ByVal dwParam As Long, pbData As Any, ByVal dwFlags As Long) As Long
Private Declare Function RtlGenRandom Lib "advapi32" Alias "SystemFunction036" (RandomBuffer As Any, ByVal RandomBufferLength As Long) As Long
'--- WIC
Private Declare Function WICCreateImagingFactory_Proxy Lib "windowscodecs" (ByVal SDKVersion As Long, ppIImagingFactory As stdole.IUnknown) As Long
Private Declare Function IWICImagingFactory_CreateStream_Proxy Lib "windowscodecs" (ByVal pFactory As stdole.IUnknown, ppIWICStream As stdole.IUnknown) As Long
Private Declare Function IWICImagingFactory_CreateEncoder_Proxy Lib "windowscodecs" (ByVal pFactory As stdole.IUnknown, guidContainerFormat As Any, pguidVendor As Any, ppIEncoder As stdole.IUnknown) As Long
Private Declare Function IWICImagingFactory_CreateBitmapFromMemory_Proxy Lib "windowscodecs" (ByVal pFactory As stdole.IUnknown, ByVal uiWidth As Long, ByVal uiHeight As Long, aPixelFormat As Any, ByVal cbStride As Long, ByVal cbBufferSize As Long, pbBuffer As Any, ppIBitmap As stdole.IUnknown) As Long
Private Declare Function IWICStream_InitializeFromIStream_Proxy Lib "windowscodecs" (ByVal pIWICStream As stdole.IUnknown, ByVal pIStream As stdole.IUnknown) As Long
Private Declare Function IWICBitmapEncoder_Initialize_Proxy Lib "windowscodecs" (ByVal pIEncoder As stdole.IUnknown, ByVal pIStream As stdole.IUnknown, ByVal cacheOption As Long) As Long
Private Declare Function IWICBitmapEncoder_CreateNewFrame_Proxy Lib "windowscodecs" (ByVal pIEncoder As stdole.IUnknown, ppIFrameEncode As stdole.IUnknown, ppIEncoderOptions As stdole.IUnknown) As Long
Private Declare Function IWICBitmapEncoder_Commit_Proxy Lib "windowscodecs" (ByVal pIEncoder As stdole.IUnknown) As Long
Private Declare Function IWICBitmapFrameEncode_Initialize_Proxy Lib "windowscodecs" (ByVal pIFrameEncode As stdole.IUnknown, ByVal pIEncoderOptions As stdole.IUnknown) As Long
Private Declare Function IWICBitmapFrameEncode_SetSize_Proxy Lib "windowscodecs" (ByVal pIFrameEncode As stdole.IUnknown, ByVal uiWidth As Long, ByVal uiHeight As Long) As Long
Private Declare Function IWICBitmapFrameEncode_WriteSource_Proxy Lib "windowscodecs" (ByVal pIFrameEncode As stdole.IUnknown, ByVal pIBitmapSource As stdole.IUnknown, pWicRect As Any) As Long
Private Declare Function IWICBitmapFrameEncode_Commit_Proxy Lib "windowscodecs" (ByVal pIFrameEncode As stdole.IUnknown) As Long
Private Declare Function IPropertyBag2_Write_Proxy Lib "windowscodecs" (ByVal pPropBag As stdole.IUnknown, ByVal cProperties As Long, pBag As Any, pvarValue As Variant) As Long
#If Not ImplUseShared Then
    '--- for thunks
    Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
    Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
    Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
    Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
    #If Not ImplNoIdeProtection Then
        Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWndParent As Long, ByVal hWndChildAfter As Long, ByVal lpszClass As String, ByVal lpszWindow As String) As Long
        Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
    #End If
    Private Declare Function GetEnvironmentVariable Lib "kernel32" Alias "GetEnvironmentVariableA" (ByVal lpName As String, ByVal lpBuffer As String, ByVal nSize As Long) As Long
    Private Declare Function SetEnvironmentVariable Lib "kernel32" Alias "SetEnvironmentVariableA" (ByVal lpName As String, ByVal lpValue As String) As Long
#End If

Private Type BLOBHEADER
    bType               As Byte
    bVersion            As Byte
    reserved            As Integer
    aiKeyAlg            As Long
    cbKeySize           As Long
    Buffer(0 To 7)      As Byte
End Type
Private Const sizeof_BLOBHEADER As Long = 12

Private Type WIN32_FIND_DATA
    dwFileAttributes    As Long
    ftCreationTime      As Currency
    ftLastAccessTime    As Currency
    ftLastWriteTime     As Currency
    nFileSizeHigh       As Long
    nFileSizeLow        As Long
    dwReserved0         As Long
    dwReserved1         As Long
    cFileName           As String * MAX_PATH
    cAlternate          As String * 14
End Type
Private Const sizeof_WIN32_FIND_DATA As Long = 592

Private Type KEYBOARD_INPUT
   dwType               As Long
   wVk                  As Integer
   wScan                As Integer
   dwFlags              As Long
   dwTime               As Long
   dwExtraInfo          As Long
   dwPadding            As Currency
End Type

'=========================================================================
' Protocol enums and types
'=========================================================================

Private Enum RfbSecurityTypeEnum
    rfbSecInvalid = 0
    rfbSecNone = 1
    rfbSecVnc = 2
    rfbSecTight = 16
    rfbSecUltraVnc = 17
End Enum

Private Enum RfbMessageEnum
    '--- client-side
    rfbMsgSetPixelFormat = 0
    rfbMsgSetEncodings = 2
    rfbMsgFramebufferUpdateRequest = 3
    rfbMsgKeyEvent = 4
    rfbMsgPointerEvent = 5
    rfbMsgClientCutText = 6
    '--- server-side
    rfbMsgFramebufferUpdate = 0
    rfbMsgSetColorMapEntries = 1
    rfbMsgBell = 2
    rfbMsgServerCutText = 3
    '--- bidirectional
    rfbMsgFileTransfer = 7
    rfbMsgTightFileTransfer = 252 ' &HFC
    '--- tight file transfer
    rfbMsgTightCompressionSupportRequest = &HFC000100
    rfbMsgTightCompressionSupportReply = &HFC000101
    rfbMsgTightFileListRequest = &HFC000102
    rfbMsgTightFileListReply = &HFC000103
    rfbMsgTightMd5Request = &HFC000104
    rfbMsgTightMd5Reply = &HFC000105
    rfbMsgTightUploadStartRequest = &HFC000106
    rfbMsgTightUploadStartReply = &HFC000107
    rfbMsgTightUploadDataRequest = &HFC000108
    rfbMsgTightUploadDataReply = &HFC000109
    rfbMsgTightUploadEndRequest = &HFC00010A
    rfbMsgTightUploadEndReply = &HFC00010B
    rfbMsgTightDownloadStartRequest = &HFC00010C
    rfbMsgTightDownloadStartReply = &HFC00010D
    rfbMsgTightDownloadDataRequest = &HFC00010E
    rfbMsgTightDownloadDataReply = &HFC00010F
    rfbMsgTightDownloadEndReply = &HFC000110
    rfbMsgTightMkdirRequest = &HFC000111
    rfbMsgTightMkdirReply = &HFC000112
    rfbMsgTightRemoveRequest = &HFC000113
    rfbMsgTightRemoveReply = &HFC000114
    rfbMsgTightRenameRequest = &HFC000115
    rfbMsgTightRenameReply = &HFC000116
    rfbMsgTightDirsizeRequest = &HFC000117
    rfbMsgTightDirsizeReply = &HFC000118
    rfbMsgTightLastRequestFailedReply = &HFC000119
End Enum

Private Enum RfbEncodingEnum
    rfbEncRaw = 0
    rfbEncCopyRect = 1
    rfbEncRRE = 2
    rfbEncCoRRE = 4
    rfbEncHextile = 5
    rfbEncZlib = 6
    rfbEncTight = 7
    rfbEncZlibHextile = 8
    rfbEncUltra = 9
    rfbEncUltra2 = 10
    rfbEncTRLE = 15
    rfbEncZRLE = 16
    rfbEncHitachiZYWRLE = 17
    rfbEncJPEG = 21
    rfbEncJRLE = 22
    rfbEncJRLE2 = 24
    rfbEncTightPNG = -260
    '--- pseudo-encodings
    rfbEncPsCursorWithAlpha = -314
    rfbEncPsExtendedDesktopSize = -308
    rfbEncPsCompressionLevel = -256
    rfbEncPsCursor = -239
    rfbEncPsPointerPosition = -232
    rfbEncPsLastRect = -224
    rfbEncPsDesktopSize = -223
    rfbEncPsJpegQualityLevel = -32
    rfbEncPsExtendedClipboard = &HC0A1E5CE
    rfbEncPsFTProtocolVersion = &HFFFF8002
End Enum

Private Type RfbPixelFormat
    BitsPerPixel        As Byte
    Depth               As Byte
    BigEndianFlag       As Byte
    TrueColorFlag       As Byte
    RedMax              As Integer
    GreenMax            As Integer
    BlueMax             As Integer
    RedShift            As Byte
    GreenShift          As Byte
    BlueShift           As Byte
    Padding(0 To 2)     As Byte
End Type
Private Const sizeof_RfbPixelFormat As Long = 16

Private Type RfbServerInit
    FramebufferWidth    As Integer
    FramebufferHeight   As Integer
    ServerPixelFormat   As RfbPixelFormat
    NameLength          As Long
End Type
Private Const sizeof_RfbServerInit As Long = 4 + sizeof_RfbPixelFormat + 4

Private Type RfbClientSetPixelFormat
    MessageType         As Long
    PixelFormat         As RfbPixelFormat
End Type
Private Const UcsClientSetPixelFormatSize As Long = 4 + sizeof_RfbPixelFormat

Private Type RfbClientSetEncodings
    MessageType         As Integer
    NumberOfEncodings   As Integer
    EncodingType()      As Long
    HasEncoding         As Long
    CompressionLevel    As Long
    JpegQualityLevel    As Long
End Type
Private Const sizeof_RfbClientSetEncodings As Long = 4

Private Type RfbClientFramebufferUpdateRequest
    MessageType         As Byte
    Incremental         As Byte
    XPosition           As Integer
    YPosition           As Integer
    Width               As Integer
    Height              As Integer
End Type
Private Const sizeof_RfbClientFramebufferUpdateRequest As Long = 10

Private Type RfbClientKeyEvent
    MessageType         As Byte
    DownFlag            As Byte
    Padding(0 To 1)     As Byte
    Key                 As Long
End Type
Private Const sizeof_RfbClientKeyEvent As Long = 8

Private Type RfbClientPointerEvent
    MessageType         As Byte
    ButtonMask          As Byte
    XPosition           As Integer
    YPosition           As Integer
End Type
Private Const sizeof_RfbClientPointerEvent As Long = 6

Private Type RfbClientCutText
    MessageType         As Long
    Length              As Long
    Text()              As Byte
End Type
Private Const sizeof_RfbClientCutText As Long = 8

Private Type RfbServerRectangle
    XPosition           As Integer
    YPosition           As Integer
    Width               As Integer
    Height              As Integer
    EncodingType        As Long
End Type
Private Const sizeof_RfbServerRectangle As Long = 12

Private Type RfbServerFramebufferUpdate
    MessageType         As Integer
    NumberOfRectangles  As Integer
End Type
Private Const sizeof_RfbServerFramebufferUpdate As Long = 4

'Private Type RfbServerRgbValue
'    Red                 As Integer
'    Green               As Integer
'    Blue                As Integer
'End Type
'Private Const sizeof_RfbServerRgbValue As Long = 6
'
'Private Type RfbServerSetColorMapEntries
'    MessageType         As Integer
'    FirstColor          As Integer
'    NumberOfColors      As Integer
'End Type
'Private Const sizeof_RfbServerSetColorMapEntries As Long = 6

'Private Type RfbServerCutText
'    MessageType         As Long
'    Length              As Long
'End Type
'Private Const sizeof_RfbServerCutText As Long = 8

Private Type RfbFileTransfer
    MessageType         As Byte
    ContentType         As Byte
    ContentParam        As Integer
    Size                As Long
    Length              As Long
End Type
Private Const sizeof_RfbFileTransfer As Long = 12

Private Enum RfbFileTransferProtocolEnum
    rfbFtpProtocolVersion = 3
    rfbFtpProtocolPacketSize = 8192

    rfbFtpDirContentRequest = 1                 ' Client asks for the content of a given Server directory
    rfbFtpDirPacket = 2                         ' Full directory name or full file name. ' Null content means end of Directory
    rfbFtpFileTransferRequest = 3               ' Client asks the server for the transfer of a given file
    rfbFtpFileHeader = 4                        ' First packet of a file transfer, containing file's features
    rfbFtpFilePacket = 5                        ' One chunk of the file
    rfbFtpEndOfFile = 6                         ' End of file transfer (the file has been received or error)
    rfbFtpAbortFileTransfer = 7                 ' The file transfer must be aborted, whatever the state
    rfbFtpFileTransferOffer = 8                 ' The client offers to send a file to the server
    rfbFtpFileAcceptHeader = 9                  ' The server accepts or rejects the file
    rfbFtpCommand = 10                          ' The Client sends a simple command (File Delete, Dir create etc...)
    rfbFtpCommandReturn = 11                    ' The Client receives the server's answer about a simple command
    rfbFtpFileChecksums = 12                    ' The zipped checksums of the destination file (Delta Transfer)
    rfbFtpFileTransferAccess = 14               ' Request FileTransfer authorization
    rfbFtpFileTransferSessionStart = 15         ' indicates a client has the FT gui open
    rfbFtpFileTransferSessionEnd = 16           ' indicates a client has closed the ft gui.
    rfbFtpFileTransferProtocolVersion = 17      ' indicates ft protocol version understood by sender. contentParam is version #
    
    ' Content params -- rfbFtpDirContentRequest client Request
    rfbFtpRDirContent = 1                       ' Request a Server Directory contents
    rfbFtpRDrivesList = 2                       ' Request the server's drives list
    rfbFtpRDirRecursiveList = 3                 ' Request a server directory content recursive sorted list
    rfbFtpRDirRecursiveSize = 4                 ' Request a server directory content recursive size
    
    ' Content params -- rfbFtpDirPacket & rfbFtpCommandReturn server Answer
    rfbFtpADirectory = 1                        ' Reception of a directory name
    rfbFtpAFile = 2                             ' Reception of a file name
    rfbFtpADrivesList = 3                       ' Reception of a list of drives
    rfbFtpADirCreate = 4                        ' Response to a create dir command
    rfbFtpADirDelete = 5                        ' Response to a delete dir command
    rfbFtpAFileCreate = 6                       ' Response to a create file command
    rfbFtpAFileDelete = 7                       ' Response to a delete file command
    rfbFtpAFileRename = 8                       ' Response to a rename file command
    rfbFtpADirRename = 9                        ' Response to a rename dir command
    rfbFtpADirRecursiveListItem = 10
    rfbFtpADirRecursiveSize = 11
    
    ' Content params -- rfbFtpCommand Command
    rfbFtpCDirCreate = 1                        ' Request the server to create the given directory
    rfbFtpCDirDelete = 2                        ' Request the server to delete the given directory
    rfbFtpCFileCreate = 3                       ' Request the server to create the given file
    rfbFtpCFileDelete = 4                       ' Request the server to delete the given file
    rfbFtpCFileRename = 5                       ' Request the server to rename the given file
    rfbFtpCDirRename = 6                        ' Request the server to rename the given directory
    
    ' Errors -- content params or "size" field
    rfbFtpSErrorUnknownCmd = 1                  ' Unknown FileTransfer command.
    rfbFtpSErrorCmd = -1                        ' Error when a command fails on remote side (ret in "size" field)
    rfbFtpSProtocolEnabled = 1                  ' for rfbFtpAbortFileTransfer
    rfbFtpSProtocolDisabled = -1
    rfbFtpSPacketUncompressed = 0               ' for rfbFtpFilePacket
    rfbFtpSPacketCompressed = 1
    rfbFtpSPacketAlreadyThere = 2
End Enum

Private Type RfbTightInteractionCapabilities
    NumberOfServerMsgs  As Integer
    NumberOfClientMsgs  As Integer
    NumberOfEncodings   As Integer
    Padding             As Integer
End Type
Private Const sizeof_RfbTightInteractionCapabilities As Long = 8

Private Type RfbTightCapability
    Code                As Long
    Vendor(0 To 3)      As Byte
    Signature(0 To 7)   As Byte
End Type
Private Const sizeof_RfbTightCapability As Long = 16

'=========================================================================
' Constants and member variables
'=========================================================================

Private Enum UcsProtocolStateEnum
    uscStaNone
    uscStaExpectProtocolVersion
    uscStaExpectSecurityHandshake
    uscStaExpectSecurityTightResponse
    uscStaExpectSecurityVncResponse
    uscStaExpectClientInit
    uscStaExpectClientMessage
End Enum

Private Enum UcsSendChannelEnum
    ucsChaFrame
    ucsChaFtp
    [_ucsChaMax]
End Enum

Private Const STR_APP_NAME          As String = "VbVncServer"
Private Const STR_PROTOCOL_VERSION  As String = "RFB 003.008" & vbLf
Private Const sizeof_PROTOCOL_VERSION As Long = 12
Private Const STR_TIGHT_SERVER_MSGS As String = "&HFC000101;TGHT;FTSCSRLY|&HFC000103;TGHT;FTSFLRLY|&HFC000107;TGHT;FTSFURLY|&HFC000109;TGHT;FTSUDRLY|&HFC00010B;TGHT;FTSUERLY|&HFC00010D;TGHT;FTSFDRLY|&HFC00010F;TGHT;FTSDDRLY|&HFC000110;TGHT;FTSDERLY|&HFC000112;TGHT;FTSMDRLY|&HFC000114;TGHT;FTSFTRLY|&HFC000116;TGHT;FTSFMRLY|&HFC000118;TGHT;FTSDSRLY|&HFC000119;TGHT;FTLRFRLY"
Private Const STR_TIGHT_CLIENT_MSGS As String = "&HFC000100;TGHT;FTCCSRST|&HFC000102;TGHT;FTCFLRST|&HFC000106;TGHT;FTCFURST|&HFC000108;TGHT;FTCUDRST|&HFC00010A;TGHT;FTCUERST|&HFC00010C;TGHT;FTCFDRST|&HFC00010E;TGHT;FTCDDRST|&HFC000111;TGHT;FTCMDRST|&HFC000113;TGHT;FTCFRRST|&HFC000115;TGHT;FTCFMRST|&HFC000117;TGHT;FTCDSRST"
Private Const STR_TIGHT_ENCODINGS   As String = "1;STDV;COPYRECT|6;TRDV;ZLIB____|7;TGHT;TIGHT___|-32;TGHT;JPEGQLVL|-223;TGHT;NEWFBSIZ|-239;TGHT;RCHCURSR|-256;TGHT;COMPRLVL"
Private Const STR_TIGHT_AUTH_TYPES  As String = "2;STDV;VNCAUTH_"
Private Const STR_DIR_PREFIX        As String = "!UVNCDIR-"
Private Const sizeof_DIR_PREFIX     As Long = 9
Private Const STR_ZIP_SUFFIX        As String = ".zip"
Private Const sizeof_ZIP_SUFFIX     As Long = 4
Private Const LNG_TILE_SIZE         As Long = 256
Private Const LNG_CHALLENGE_SIZE    As Long = 16

Private m_cConnections          As Collection
Private m_sPassword             As String
'--- client
Private m_lParentWeakRef        As Long
Private m_eState                As UcsProtocolStateEnum
Private WithEvents m_oSocket    As cAsyncSocket
Attribute m_oSocket.VB_VarHelpID = -1
Private m_uRecvBuffer           As UcsBuffer
Private m_uSendBuffer(0 To [_ucsChaMax] - 1) As UcsBuffer
Private m_eSecurityType         As RfbSecurityTypeEnum
Private m_baSecurityChallenge() As Byte
Private m_uClientPixelFormat    As RfbClientSetPixelFormat
Private m_uClientEncodings      As RfbClientSetEncodings
Private m_uClientFramebufferUpdate As RfbClientFramebufferUpdateRequest
Private m_uClientKeyEvent       As RfbClientKeyEvent
Private m_uClientPointerEvent   As RfbClientPointerEvent
Private m_uClientCutText        As RfbClientCutText
Private m_uCaptureCtx           As UcsCaptureContext
Private m_uCaptureFrame         As UcsCaptureFrame
Private m_pTimer                As IUnknown
#If ImplZlib Then
    Private m_oZipArchive       As cZipArchive
    Private m_hDeflateCtx       As Long
#End If
Private m_pWicFactory           As stdole.IUnknown
Private m_aWicFormatJpeg()      As Long
Private m_aWicFormatPng()       As Long
Private m_aWicFormat32bppPBGRA() As Long
Private m_pImageStream          As stdole.IUnknown
Private m_bFtpSendRunning       As Boolean
Private m_pFtpSendStream        As stdole.IUnknown
Private m_pFtpSendChecksums     As stdole.IUnknown
Private m_sFtpSendArchiveFile   As String
Private m_sFtpRecvFileName      As String
Private m_pFtpRecvStream        As stdole.IUnknown

Private Type UcsBuffer
    Data()              As Byte
    Pos                 As Long
    Size                As Long
End Type

Private Type UcsCaptureContext
    DeviceName          As String
    Width               As Long
    Height              As Long
    Timeout             As Long
    Context             As ID3D11DeviceContext
    Duplication         As IDXGIOutputDuplication
    StageTexture        As ID3D11Texture2D
    DesktopResource     As ID3D11Resource
    InSystemMemory      As Boolean
    Pitch               As Long
    Stride              As Long
End Type

Private Type UcsCaptureTile
    SrcIndex            As Long
    Left                As Long
    Top                 As Long
    Width               As Long
    Height              As Long
End Type

Private Type UcsCaptureFrame
    NumMoveRects        As Long
    MoveRects()         As DXGI_OUTDUPL_MOVE_RECT
    NumDirtyRects       As Long
    DirtyRects()        As D3D11_RECT
    DirtyBuffers()      As UcsBuffer
    PointerSize         As Long
    PointerShape()      As Byte
    PointerVisible      As Boolean
    PointerInfo         As DXGI_OUTDUPL_POINTER_SHAPE_INFO
    PointerShapeBufferSize As Long
    NumTiles            As Long
    Tiles()             As UcsCaptureTile
End Type

Private Type UcsTightTransfer
    MessageType         As RfbMessageEnum
    Path                As String
    NewFileName         As String
    CompressionLevel    As Long
    CompressedSize      As Long
    UncompressedSize    As Long
    Offset              As Currency
    DataSize            As Long
    Flags               As Long
    ModificationTime    As Currency
End Type

'=========================================================================
' Error handling
'=========================================================================

Private Sub PrintError(sFunction As String)
    #If ImplUseDebugLog Then
        DebugLog MODULE_NAME, sFunction & "(" & Erl & ")", Err.Description & " &H" & Hex$(Err.Number), vbLogEventTypeError
    #Else
        Debug.Print "Critical error: " & Err.Description & " [" & MODULE_NAME & "." & sFunction & "]"
    #End If
End Sub

'=========================================================================
' Properties
'=========================================================================

Public Property Get Socket() As cAsyncSocket
    Set Socket = m_oSocket
End Property

Public Property Get LastError() As String
    If Not m_oSocket Is Nothing Then
        LastError = m_oSocket.GetErrorDescription(m_oSocket.LastError)
    End If
End Property

Public Property Get Password() As String
    Password = m_sPassword
End Property

Public Property Get CaptureWidth() As Long
    CaptureWidth = m_uCaptureCtx.Width
End Property

Public Property Get CaptureHeight() As Long
    CaptureHeight = m_uCaptureCtx.Height
End Property

Private Property Get pvAddressOfTimerProc() As cVncServer
    Set pvAddressOfTimerProc = InitAddressOfMethod(Me, 0)
End Property

Private Property Get pvParent() As cVncServer
    Call vbaObjSetAddref(pvParent, m_lParentWeakRef)
End Property

Private Property Get pvFirstPendingChannel() As UcsSendChannelEnum
    Dim eChannel        As UcsSendChannelEnum
    
    pvFirstPendingChannel = -1
    For eChannel = 0 To [_ucsChaMax] - 1
        If m_uSendBuffer(eChannel).Size > m_uSendBuffer(eChannel).Pos Then
            pvFirstPendingChannel = eChannel
            Exit For
        End If
    Next
End Property

Private Property Get pvHasSupport(ByVal eEncoding As RfbEncodingEnum) As Boolean
    pvHasSupport = (m_uClientEncodings.HasEncoding And pvGetEncodingBit(eEncoding)) <> 0
End Property

'=========================================================================
' Methods
'=========================================================================

Public Function Init(Optional ListenAddress As String, Optional ByVal ListenPort As Long = 5900, Optional Password As String) As Boolean
    Set m_cConnections = New Collection
    m_sPassword = Password
    Set m_oSocket = New cAsyncSocket
    If Not m_oSocket.Create(SocketAddress:=ListenAddress, SocketPort:=ListenPort) Then
        GoTo QH
    End If
    If Not m_oSocket.Listen() Then
        GoTo QH
    End If
    '--- success
    Init = True
QH:
End Function

Friend Function frNewConnection(oSocket As cAsyncSocket, oParent As cVncServer) As Boolean
    Const WINCODEC_SDK_VERSION1   As Long = &H236&
    Const WINCODEC_SDK_VERSION2   As Long = &H237&
    Const IPPROTO_TCP             As Long = 6
    Const TCP_NODELAY             As Long = 1
    
    Set m_oSocket = oSocket
    m_lParentWeakRef = ObjPtr(oParent)
    m_sPassword = oParent.Password
    m_oSocket.SockOpt(TCP_NODELAY, IPPROTO_TCP) = 1 '--- disable Nagle algorithm
    If Not pvSendProtocolVersion(m_uSendBuffer(ucsChaFrame)) Then
        GoTo QH
    End If
    m_eState = uscStaExpectProtocolVersion
    '--- init WIC state
    If WICCreateImagingFactory_Proxy(WINCODEC_SDK_VERSION2, m_pWicFactory) < 0 Then
        Call WICCreateImagingFactory_Proxy(WINCODEC_SDK_VERSION1, m_pWicFactory)
    End If
    ReDim m_aWicFormatJpeg(0 To 3) As Long
    m_aWicFormatJpeg(0) = &H19E4A5AA
    m_aWicFormatJpeg(1) = &H4FC55662
    m_aWicFormatJpeg(2) = &H5817C0A0
    m_aWicFormatJpeg(3) = &H57108E02
    ReDim m_aWicFormatPng(0 To 3) As Long
    m_aWicFormatPng(0) = &H1B7CFAF4
    m_aWicFormatPng(1) = &H473C713F
    m_aWicFormatPng(2) = &H3761CDBB
    m_aWicFormatPng(3) = &HAFAE5F42
    ReDim m_aWicFormat32bppPBGRA(0 To 3) As Long
    m_aWicFormat32bppPBGRA(0) = &H6FDDC324
    m_aWicFormat32bppPBGRA(1) = &H4BFE4E03
    m_aWicFormat32bppPBGRA(2) = &H773D85B1
    m_aWicFormat32bppPBGRA(3) = &H10C98D76
    Set m_pImageStream = SHCreateMemStream(ByVal 0, 0)
    '--- success
    frNewConnection = True
QH:
End Function

Friend Sub frCloseConnection(oConn As cVncServer)
    Dim uEmpty          As UcsCaptureContext
    
    If Not oConn Is Nothing Then
        On Error Resume Next
        m_cConnections.Remove "#" & ObjPtr(oConn)
        On Error GoTo 0
    End If
    If m_cConnections.Count = 0 Then
        m_uCaptureCtx = uEmpty
        Set m_pTimer = Nothing
    End If
End Sub

Private Function pvSendProtocolVersion(uOutput As UcsBuffer) As Boolean
    pvBufferWriteArray uOutput, StrConv(STR_PROTOCOL_VERSION, vbFromUnicode)
    '--- success (or failure)
    pvSendProtocolVersion = pvBufferSendAsync(uOutput)
End Function

Private Function pvSendSecurityTypes(uOutput As UcsBuffer) As Boolean
    pvBufferWriteLong uOutput, 3
    pvBufferWriteLong uOutput, IIf(LenB(m_sPassword) <> 0, rfbSecVnc, rfbSecNone)
    pvBufferWriteLong uOutput, rfbSecUltraVnc
    pvBufferWriteLong uOutput, rfbSecTight
    '--- success (or failure)
    pvSendSecurityTypes = pvBufferSendAsync(uOutput)
End Function

Private Function pvSendSingleSecurityType(uOutput As UcsBuffer, ByVal eType As RfbSecurityTypeEnum) As Boolean
    pvBufferWriteLong uOutput, 1
    pvBufferWriteLong uOutput, eType
    '--- success (or failure)
    pvSendSingleSecurityType = pvBufferSendAsync(uOutput)
End Function

Private Function pvSendSecurityChallenge(uOutput As UcsBuffer, baChallenge() As Byte, baMod() As Byte, baPub() As Byte) As Boolean
    pvBufferWriteArray uOutput, baChallenge
    If pvArraySize(baMod) > 0 Then
        pvBufferWriteArray uOutput, baMod
    End If
    If pvArraySize(baPub) > 0 Then
        pvBufferWriteArray uOutput, baPub
    End If
    '--- success (or failure)
    pvSendSecurityChallenge = pvBufferSendAsync(uOutput)
End Function

Private Function pvSendSecurityResult(uOutput As UcsBuffer, ByVal lResult As Long, Optional sReason As String) As Boolean
    pvBufferWriteLong uOutput, lResult, Size:=4
    If LenB(sReason) <> 0 Then
        pvBufferWriteLong uOutput, Len(sReason), Size:=4
        pvBufferWriteArray uOutput, StrConv(sReason, vbFromUnicode)
    End If
    '--- success (or failure)
    pvSendSecurityResult = pvBufferSendAsync(uOutput)
End Function

Private Function pvSendTightSecurityType(uOutput As UcsBuffer) As Boolean
    Const NumberOfTunnels As Long = 0
    Const NumberOfAuths As Long = 0
    Dim uAuthTypes()    As RfbTightCapability
    
    pvBufferWriteLong uOutput, NumberOfTunnels, Size:=4
    If LenB(m_sPassword) = 0 Then
        pvBufferWriteLong uOutput, NumberOfAuths, Size:=4
    Else
        pvInitTightCapabilities uAuthTypes, STR_TIGHT_AUTH_TYPES
        pvBufferWriteLong uOutput, UBound(uAuthTypes) + 1, Size:=4
        pvBufferWriteBlob uOutput, VarPtr(uAuthTypes(0)), (UBound(uAuthTypes) + 1) * sizeof_RfbTightCapability
    End If
    '--- success (or failure)
    pvSendTightSecurityType = pvBufferSendAsync(uOutput)
End Function

Private Function pvSendServerInit(uOutput As UcsBuffer, ByVal lWidth As Long, ByVal lHeight As Long, sName As String) As Boolean
    Dim uInit           As RfbServerInit
    Dim uInteraction    As RfbTightInteractionCapabilities
    Dim uServerMsgs()   As RfbTightCapability
    Dim uClientMsgs()   As RfbTightCapability
    Dim uEncodings()    As RfbTightCapability
    
    With uInit
        .FramebufferWidth = pvNetworkShort(lWidth)
        .FramebufferHeight = pvNetworkShort(lHeight)
        With .ServerPixelFormat
            .BitsPerPixel = 32
            .Depth = 24
            .BigEndianFlag = 0
            .TrueColorFlag = 1
            .RedMax = pvNetworkShort(255)
            .GreenMax = pvNetworkShort(255)
            .BlueMax = pvNetworkShort(255)
            .RedShift = 16
            .GreenShift = 8
            .BlueShift = 0
        End With
        '--- save as default pixel format
        m_uClientPixelFormat.PixelFormat = .ServerPixelFormat
        .NameLength = pvNetworkLong(Len(sName))
    End With
    pvBufferWriteBlob uOutput, VarPtr(uInit), sizeof_RfbServerInit
    pvBufferWriteArray uOutput, StrConv(sName, vbFromUnicode)
    If m_eSecurityType = rfbSecTight Then
        pvInitTightCapabilities uServerMsgs, STR_TIGHT_SERVER_MSGS
        pvInitTightCapabilities uClientMsgs, STR_TIGHT_CLIENT_MSGS
        pvInitTightCapabilities uEncodings, STR_TIGHT_ENCODINGS
        With uInteraction
            .NumberOfServerMsgs = pvNetworkShort(UBound(uServerMsgs) + 1)
            .NumberOfClientMsgs = pvNetworkShort(UBound(uClientMsgs) + 1)
            .NumberOfEncodings = pvNetworkShort(UBound(uEncodings) + 1)
        End With
        pvBufferWriteBlob uOutput, VarPtr(uInteraction), sizeof_RfbTightInteractionCapabilities
        pvBufferWriteBlob uOutput, VarPtr(uServerMsgs(0)), (UBound(uServerMsgs) + 1) * sizeof_RfbTightCapability
        pvBufferWriteBlob uOutput, VarPtr(uClientMsgs(0)), (UBound(uClientMsgs) + 1) * sizeof_RfbTightCapability
        pvBufferWriteBlob uOutput, VarPtr(uEncodings(0)), (UBound(uEncodings) + 1) * sizeof_RfbTightCapability
    End If
    '--- succees (or failure)
    pvSendServerInit = pvBufferSendAsync(uOutput)
End Function

Private Function pvSendEmptyFramebufferUpdate(uOutput As UcsBuffer) As Boolean
    Dim uUpdate         As RfbServerFramebufferUpdate
    Dim uRectangle      As RfbServerRectangle
    
    With uUpdate
        .MessageType = rfbMsgFramebufferUpdate
        .NumberOfRectangles = pvNetworkShort(1)
    End With
    pvBufferWriteBlob uOutput, VarPtr(uUpdate), sizeof_RfbServerFramebufferUpdate
    pvBufferWriteBlob uOutput, VarPtr(uRectangle), sizeof_RfbServerRectangle
    '--- success (or failure)
    pvSendEmptyFramebufferUpdate = pvBufferSendAsync(uOutput)
End Function

Private Sub pvHandleReceive()
    Const FUNC_NAME     As String = "pvHandleReceive"
    Const SecurityResultOk As Long = 0
    Const SecurityResultFailed As Long = 1
    Const SecurityResultContinue As Long = -1
    Dim baBuffer()      As Byte
    Dim baPassword()    As Byte
    Dim eSecurityType   As RfbSecurityTypeEnum
    
    On Error GoTo EH
    Do
        If Not m_oSocket.ReceiveArray(baBuffer) Then
            Exit Do
        End If
        If UBound(baBuffer) < 0 Then
            Exit Do
        End If
        pvBufferWriteArray m_uRecvBuffer, baBuffer
        With m_uRecvBuffer
            Do While .Pos < .Size
                Select Case m_eState
                Case uscStaExpectClientMessage
                    If Not pvHandleClientMessage(m_uRecvBuffer) Then
                        Exit Do
                    End If
                Case uscStaExpectProtocolVersion
                    If .Size - .Pos < sizeof_PROTOCOL_VERSION Then
                        Exit Do
                    End If
                    '--- ignore client protocol version
                    .Pos = .Pos + sizeof_PROTOCOL_VERSION
                    If Not pvSendSecurityTypes(m_uSendBuffer(ucsChaFrame)) Then
                        GoTo QH
                    End If
                    m_eState = uscStaExpectSecurityHandshake
                Case uscStaExpectSecurityHandshake
                    eSecurityType = .Data(.Pos)
                    .Pos = .Pos + 1
                    m_eSecurityType = eSecurityType
DoAuth:
                    Select Case eSecurityType
                    Case rfbSecNone
                        If LenB(m_sPassword) <> 0 Then
                            pvSendSecurityResult m_uSendBuffer(ucsChaFrame), SecurityResultFailed, "Password required"
                            GoTo QH
                        End If
                        If Not pvSendSecurityResult(m_uSendBuffer(ucsChaFrame), SecurityResultOk) Then
                            GoTo QH
                        End If
                        m_eState = uscStaExpectClientInit
                    Case rfbSecVnc
                        m_baSecurityChallenge = pvCryptoRandomBytes(LNG_CHALLENGE_SIZE)
                        Erase baBuffer
                        If Not pvSendSecurityChallenge(m_uSendBuffer(ucsChaFrame), m_baSecurityChallenge, baBuffer, baBuffer) Then
                            GoTo QH
                        End If
                        m_eState = uscStaExpectSecurityVncResponse
                    Case rfbSecTight
                        If Not pvSendTightSecurityType(m_uSendBuffer(ucsChaFrame)) Then
                            GoTo QH
                        End If
                        If LenB(m_sPassword) = 0 Then
                            If Not pvSendSecurityResult(m_uSendBuffer(ucsChaFrame), SecurityResultOk) Then
                                GoTo QH
                            End If
                            m_eState = uscStaExpectClientInit
                        Else
                            m_eState = uscStaExpectSecurityTightResponse
                        End If
                    Case rfbSecUltraVnc
                        If LenB(m_sPassword) <> 0 Then
                            If Not pvSendSecurityResult(m_uSendBuffer(ucsChaFrame), SecurityResultContinue) Then
                                GoTo QH
                            End If
                            If Not pvSendSingleSecurityType(m_uSendBuffer(ucsChaFrame), rfbSecVnc) Then
                                GoTo QH
                            End If
                        Else
                            If Not pvSendSecurityResult(m_uSendBuffer(ucsChaFrame), SecurityResultOk) Then
                                GoTo QH
                            End If
                            m_eState = uscStaExpectClientInit
                        End If
                    Case Else
                        #If ImplUseDebugLog Then
                            DebugLog MODULE_NAME, FUNC_NAME, "Unknown security type (" & eSecurityType & ")"
                        #End If
                        GoTo QH
                    End Select
                Case uscStaExpectSecurityTightResponse
                    If .Size - .Pos < 4 Then
                        Exit Do
                    End If
                    Call CopyMemory(eSecurityType, .Data(.Pos), 4)
                    .Pos = .Pos + 4
                    eSecurityType = pvNetworkLong(eSecurityType)
                    GoTo DoAuth
                Case uscStaExpectSecurityVncResponse
                    If .Size - .Pos < LNG_CHALLENGE_SIZE Then
                        Exit Do
                    End If
                    baPassword = m_baSecurityChallenge
                    If Not pvCryptoDesEncrypt(baPassword, StrConv(m_sPassword, vbFromUnicode)) Then
                        GoTo QH
                    End If
                    If Not pvArrayEqual(.Data, .Pos, LNG_CHALLENGE_SIZE, baPassword) Then
                        pvSendSecurityResult m_uSendBuffer(ucsChaFrame), SecurityResultFailed, "Wrong password"
                        GoTo QH
                    End If
                    .Pos = .Pos + LNG_CHALLENGE_SIZE
                    If Not pvSendSecurityResult(m_uSendBuffer(ucsChaFrame), SecurityResultOk) Then
                        GoTo QH
                    End If
                    m_eState = uscStaExpectClientInit
                Case uscStaExpectClientInit
                    '--- ignore shared flag
                    .Pos = .Pos + 1
                    If Not pvSendServerInit(m_uSendBuffer(ucsChaFrame), pvParent.CaptureWidth, pvParent.CaptureHeight, STR_APP_NAME) Then
                        GoTo QH
                    End If
                    m_eState = uscStaExpectClientMessage
                Case Else
                    Exit Do
                End Select
            Loop
            If .Pos > 0 Then
                If .Size > .Pos Then
                    Call CopyMemory(.Data(0), .Data(.Pos), .Size - .Pos)
                End If
                .Size = .Size - .Pos
                .Pos = 0
            End If
        End With
    Loop
    Exit Sub
QH:
    m_eState = uscStaNone
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Function pvHandleClientMessage(uInput As UcsBuffer) As Boolean
    Const FUNC_NAME     As String = "pvHandleClientMessage"
    Dim lIdx            As Long
    Dim uPointerEvent   As RfbClientPointerEvent
    Dim uFileTransfer   As RfbFileTransfer
    Dim baData()        As Byte
    Dim uTightTransfer  As UcsTightTransfer
    Dim lSize           As Long
    
    On Error GoTo EH
    With uInput
'        DebugLog MODULE_NAME, FUNC_NAME, "Message=" & .Data(.Pos)
        Select Case .Data(.Pos)
        Case rfbMsgSetPixelFormat
            If .Size - .Pos < UcsClientSetPixelFormatSize Then
                GoTo QH
            End If
            Call CopyMemory(m_uClientPixelFormat, .Data(.Pos), UcsClientSetPixelFormatSize)
            With m_uClientPixelFormat.PixelFormat
                .RedMax = pvNetworkShort(.RedMax)
                .GreenMax = pvNetworkShort(.GreenMax)
                .BlueMax = pvNetworkShort(.BlueMax)
            End With
            .Pos = .Pos + UcsClientSetPixelFormatSize
            Debug.Assert m_uClientPixelFormat.PixelFormat.BitsPerPixel <= 32
        Case rfbMsgSetEncodings
            If .Size - .Pos < sizeof_RfbClientSetEncodings Then
                GoTo QH
            End If
            Call CopyMemory(m_uClientEncodings, .Data(.Pos), sizeof_RfbClientSetEncodings)
            With m_uClientEncodings
                .NumberOfEncodings = pvNetworkShort(.NumberOfEncodings)
            End With
            If m_uClientEncodings.NumberOfEncodings > 0 Then
                If .Size - .Pos < sizeof_RfbClientSetEncodings + m_uClientEncodings.NumberOfEncodings * 4 Then
                    GoTo QH
                End If
                ReDim m_uClientEncodings.EncodingType(0 To m_uClientEncodings.NumberOfEncodings - 1) As Long
                Call CopyMemory(m_uClientEncodings.EncodingType(0), .Data(.Pos + sizeof_RfbClientSetEncodings), m_uClientEncodings.NumberOfEncodings * 4)
                With m_uClientEncodings
                    .HasEncoding = pvGetEncodingBit(rfbEncRaw)
                    .CompressionLevel = 0
                    .JpegQualityLevel = 0
                    For lIdx = 0 To .NumberOfEncodings - 1
                        .EncodingType(lIdx) = pvNetworkLong(.EncodingType(lIdx))
                        .HasEncoding = .HasEncoding Or pvGetEncodingBit(.EncodingType(lIdx))
                        Select Case .EncodingType(lIdx)
                        Case rfbEncPsCompressionLevel To rfbEncPsCompressionLevel + 9
                            .CompressionLevel = .EncodingType(lIdx) - rfbEncPsCompressionLevel + 1
                        Case rfbEncPsJpegQualityLevel To rfbEncPsJpegQualityLevel + 9
                            .JpegQualityLevel = .EncodingType(lIdx) - rfbEncPsJpegQualityLevel + 1
                        Case rfbEncPsFTProtocolVersion
                            If Not pvSendFileTransfer(m_uSendBuffer(ucsChaFtp), rfbFtpFileTransferProtocolVersion, rfbFtpProtocolVersion) Then
                                GoTo QH
                            End If
                        Case rfbEncRaw, rfbEncRRE, rfbEncCoRRE, rfbEncHextile, rfbEncZlibHextile, rfbEncUltra, rfbEncUltra2, rfbEncTRLE, rfbEncZRLE, _
                                rfbEncHitachiZYWRLE, rfbEncJPEG, rfbEncJRLE, rfbEncJRLE2, rfbEncTightPNG, _
                                rfbEncPsExtendedDesktopSize, rfbEncPsPointerPosition, rfbEncPsLastRect, rfbEncPsDesktopSize, rfbEncPsExtendedClipboard
                            '--- do nothing
                        Case Else
                            #If ImplUseDebugLog Then
'                                If pvGetEncodingBit(.EncodingType(lIdx)) = 0 Then
'                                    DebugLog MODULE_NAME, FUNC_NAME, "Unknown encoding type " & .EncodingType(lIdx) & " (&H" & Hex$(.EncodingType(lIdx)) & ")"
'                                End If
                            #End If
                        End Select
                    Next
                End With
            End If
            .Pos = .Pos + sizeof_RfbClientSetEncodings + m_uClientEncodings.NumberOfEncodings * 4
        Case rfbMsgFramebufferUpdateRequest
            If .Size - .Pos < sizeof_RfbClientFramebufferUpdateRequest Then
                GoTo QH
            End If
            Call CopyMemory(m_uClientFramebufferUpdate, .Data(.Pos), sizeof_RfbClientFramebufferUpdateRequest)
            With m_uClientFramebufferUpdate
                .XPosition = pvNetworkShort(.XPosition)
                .YPosition = pvNetworkShort(.YPosition)
                .Width = pvNetworkShort(.Width)
                .Height = pvNetworkShort(.Height)
            End With
            .Pos = .Pos + sizeof_RfbClientFramebufferUpdateRequest
            If m_uClientFramebufferUpdate.Incremental = 0 Then
                Call RedrawWindow(0, ByVal 0, 0, RDW_INVALIDATE Or RDW_ALLCHILDREN Or RDW_ERASENOW)
            End If
        Case rfbMsgKeyEvent
            If .Size - .Pos < sizeof_RfbClientKeyEvent Then
                GoTo QH
            End If
            Call CopyMemory(m_uClientKeyEvent, .Data(.Pos), sizeof_RfbClientKeyEvent)
            m_uClientKeyEvent.Key = pvNetworkLong(m_uClientKeyEvent.Key)
            .Pos = .Pos + sizeof_RfbClientKeyEvent
            If Not pvHandleKeyEvent(m_uClientKeyEvent) Then
                GoTo QH
            End If
        Case rfbMsgPointerEvent
            If .Size - .Pos < sizeof_RfbClientPointerEvent Then
                GoTo QH
            End If
            Call CopyMemory(uPointerEvent, .Data(.Pos), sizeof_RfbClientPointerEvent)
            With uPointerEvent
                .XPosition = pvNetworkShort(.XPosition)
                .YPosition = pvNetworkShort(.YPosition)
            End With
            .Pos = .Pos + sizeof_RfbClientPointerEvent
            If Not pvHandlePointerEvent(pvParent.CaptureWidth, pvParent.CaptureHeight, uPointerEvent) Then
                GoTo QH
            End If
        Case rfbMsgClientCutText
            If .Size - .Pos < sizeof_RfbClientCutText Then
                GoTo QH
            End If
            Call CopyMemory(m_uClientCutText, .Data(.Pos), sizeof_RfbClientCutText)
            With m_uClientCutText
                .Length = pvNetworkLong(.Length)
            End With
            If m_uClientCutText.Length > 0 Then
                If .Size - .Pos < sizeof_RfbClientCutText + m_uClientCutText.Length Then
                    GoTo QH
                End If
                pvArrayAllocate m_uClientCutText.Text, m_uClientCutText.Length, FUNC_NAME
                Call CopyMemory(m_uClientCutText.Text(0), .Data(.Pos + sizeof_RfbClientCutText), m_uClientCutText.Length)
            End If
            .Pos = .Pos + sizeof_RfbClientCutText + m_uClientCutText.Length
            If Not pvHandleCutEvent(m_uClientCutText) Then
                GoTo QH
            End If
        Case rfbMsgFileTransfer
            If .Size - .Pos < sizeof_RfbFileTransfer Then
                GoTo QH
            End If
            Call CopyMemory(uFileTransfer, .Data(.Pos), sizeof_RfbFileTransfer)
            With uFileTransfer
                .Size = pvNetworkLong(.Size)
                .Length = pvNetworkLong(.Length)
            End With
            lIdx = uFileTransfer.Length
            If uFileTransfer.ContentType = rfbFtpFileTransferOffer Then
                lIdx = lIdx + 4
            ElseIf uFileTransfer.ContentType = rfbFtpFilePacket And uFileTransfer.Size = rfbFtpSPacketAlreadyThere Then
                lIdx = 0
            End If
            If .Size - .Pos < sizeof_RfbFileTransfer + lIdx Then
                GoTo QH
            End If
            .Pos = .Pos + sizeof_RfbFileTransfer
            If uFileTransfer.Length > 0 And lIdx > 0 Then
                pvArrayAllocate baData, uFileTransfer.Length, FUNC_NAME & ".baData"
                Call CopyMemory(ByVal pvArrayPtr(baData), .Data(.Pos), uFileTransfer.Length)
                .Pos = .Pos + uFileTransfer.Length
            Else
                baData = vbNullString
            End If
            '--- retrieve high part of content size (ignored)
            If uFileTransfer.ContentType = rfbFtpFileTransferOffer Then
                Call CopyMemory(lIdx, .Data(.Pos), 4)
                .Pos = .Pos + 4
            Else
                lIdx = 0
            End If
            If Not pvHandleFileTransfer(m_uSendBuffer(ucsChaFtp), uFileTransfer, baData) Then
                GoTo QH
            End If
        Case rfbMsgTightFileTransfer
            If .Size - .Pos < 4 Then
                GoTo QH
            End If
            Call CopyMemory(uTightTransfer.MessageType, .Data(.Pos), 4)
            uTightTransfer.MessageType = pvNetworkLong(uTightTransfer.MessageType)
            Select Case uTightTransfer.MessageType
            Case rfbMsgTightCompressionSupportRequest
                .Pos = .Pos + 4
            Case rfbMsgTightFileListRequest
                If .Size - .Pos < 4 + 1 + 4 Then
                    GoTo QH
                End If
                Call CopyMemory(lIdx, .Data(.Pos + 5), 4)
                lIdx = pvNetworkLong(lIdx)
                If .Size - .Pos < 4 + 1 + 4 + lIdx Then
                    GoTo QH
                End If
                .Pos = .Pos + 4
                With uTightTransfer
                    pvBufferReadLong uInput, .CompressionLevel
                    pvBufferReadUtf8 uInput, .Path
                End With
            Case rfbMsgTightUploadStartRequest
                If .Size - .Pos < 4 + 4 + 1 + 8 Then
                    GoTo QH
                End If
                Call CopyMemory(lIdx, .Data(.Pos + 4), 4)
                lIdx = pvNetworkLong(lIdx)
                If .Size - .Pos < 4 + 4 + lIdx + 1 + 8 Then
                    GoTo QH
                End If
                .Pos = .Pos + 4
                With uTightTransfer
                    pvBufferReadUtf8 uInput, .Path
                    pvBufferReadLong uInput, .Flags
                    pvBufferReadCurrency uInput, .Offset
                End With
            Case rfbMsgTightUploadDataRequest
                If .Size - .Pos < 4 + 1 + 4 + 4 Then
                    GoTo QH
                End If
                Call CopyMemory(lIdx, .Data(.Pos + 5), 4)
                lIdx = pvNetworkLong(lIdx)
                If .Size - .Pos < 4 + 1 + 4 + 4 + lIdx Then
                    GoTo QH
                End If
                .Pos = .Pos + 4
                With uTightTransfer
                    pvBufferReadLong uInput, .CompressionLevel
                    pvBufferReadLong uInput, .CompressedSize, Size:=4
                    pvBufferReadLong uInput, .UncompressedSize, Size:=4
                End With
                pvArrayAllocate baData, uTightTransfer.CompressedSize, FUNC_NAME & ".baData"
                If uTightTransfer.CompressedSize > 0 Then
                    Call CopyMemory(ByVal pvArrayPtr(baData), .Data(.Pos), pvArraySize(baData))
                    .Pos = .Pos + uTightTransfer.CompressedSize
                End If
            Case rfbMsgTightUploadEndRequest
                If .Size - .Pos < 4 + 2 + 8 Then
                    GoTo QH
                End If
                .Pos = .Pos + 4
                With uTightTransfer
                    pvBufferReadLong uInput, .Flags, Size:=2
                    pvBufferReadCurrency uInput, .ModificationTime
                End With
            Case rfbMsgTightDownloadStartRequest
                If .Size - .Pos < 4 + 4 + 8 Then
                    GoTo QH
                End If
                Call CopyMemory(lIdx, .Data(.Pos + 4), 4)
                lIdx = pvNetworkLong(lIdx)
                If .Size - .Pos < 4 + 4 + lIdx + 8 Then
                    GoTo QH
                End If
                .Pos = .Pos + 4
                With uTightTransfer
                    pvBufferReadUtf8 uInput, .Path
                    pvBufferReadCurrency uInput, .Offset
                End With
            Case rfbMsgTightDownloadDataRequest
                If .Size - .Pos < 4 + 1 + 4 Then
                    GoTo QH
                End If
                .Pos = .Pos + 4
                With uTightTransfer
                    pvBufferReadLong uInput, .CompressionLevel
                    pvBufferReadLong uInput, .DataSize, Size:=4
                End With
            Case rfbMsgTightMkdirRequest, rfbMsgTightRemoveRequest, rfbMsgTightDirsizeRequest
                If .Size - .Pos < 4 + 4 Then
                    GoTo QH
                End If
                Call CopyMemory(lIdx, .Data(.Pos + 4), 4)
                lIdx = pvNetworkLong(lIdx)
                If .Size - .Pos < 4 + 4 + lIdx Then
                    GoTo QH
                End If
                .Pos = .Pos + 4
                pvBufferReadUtf8 uInput, uTightTransfer.Path
            Case rfbMsgTightRenameRequest
                If .Size - .Pos < 4 + 4 + 4 Then
                    GoTo QH
                End If
                Call CopyMemory(lIdx, .Data(.Pos + 4), 4)
                lIdx = pvNetworkLong(lIdx)
                If .Size - .Pos < 4 + 4 + lIdx + 4 Then
                    GoTo QH
                End If
                Call CopyMemory(lSize, .Data(.Pos + 4 + 4 + lIdx), 4)
                lSize = pvNetworkLong(lSize)
                If .Size - .Pos < 4 + 4 + lIdx + 4 + lSize Then
                    GoTo QH
                End If
                .Pos = .Pos + 4
                With uTightTransfer
                    pvBufferReadUtf8 uInput, .Path
                    pvBufferReadUtf8 uInput, .NewFileName
                End With
            Case Else
                #If ImplUseDebugLog Then
                    DebugLog MODULE_NAME, FUNC_NAME, "Unknown tight file transfer message (&H" & Hex$(uTightTransfer.MessageType) & ")"
                #End If
                uTightTransfer.MessageType = 0
                '--- flush
                .Pos = .Size
            End Select
            If Not pvHandleTightTransfer(m_uSendBuffer(ucsChaFtp), uTightTransfer, baData) Then
                GoTo QH
            End If
        Case Else
            #If ImplUseDebugLog Then
                DebugLog MODULE_NAME, FUNC_NAME, "Unknown client message (" & .Data(.Pos) & ")"
            #End If
            '--- flush
            .Pos = .Size
        End Select
    End With
    '--- succes
    pvHandleClientMessage = True
QH:
    Exit Function
EH:
    PrintError FUNC_NAME
End Function

Private Function pvHandleKeyEvent(uEvent As RfbClientKeyEvent) As Boolean
    Const FUNC_NAME     As String = "pvHandleKeyEvent"
    Const UNICODE_FLAG  As Long = &H1000000
    Dim uInput          As KEYBOARD_INPUT
    Dim bExtended       As Boolean
    
    On Error GoTo EH
    #If ImplUseDebugLog Then
'        DebugLog MODULE_NAME, FUNC_NAME, "uEvent.Key=&H" & Hex$(uEvent.Key) & ", uEvent.DownFlag=" & uEvent.DownFlag
    #End If
    uInput.dwType = INPUT_KEYBOARD
    If uEvent.DownFlag = 0 Then
        uInput.dwFlags = KEYEVENTF_KEYUP
    End If
    If pvMapToVirtualKey(uEvent.Key, uInput.wVk, bExtended) Then
        uInput.wScan = MapVirtualKey(uInput.wVk, 0)
        If bExtended Then
            uInput.dwFlags = uInput.dwFlags Or KEYEVENTF_EXTENDEDKEY
        End If
    ElseIf (uEvent.Key And UNICODE_FLAG) <> 0 Then
        uInput.wScan = uEvent.Key And Not UNICODE_FLAG
        uInput.dwFlags = uInput.dwFlags Or KEYEVENTF_UNICODE
    Else
        Select Case uEvent.Key
        Case 32 To 126, 160 To 255
            uInput.wScan = uEvent.Key
            uInput.dwFlags = uInput.dwFlags Or KEYEVENTF_UNICODE
        End Select
    End If
    Call SendInput(1, uInput, LenB(uInput))
    '--- success
    pvHandleKeyEvent = True
    Exit Function
EH:
    PrintError FUNC_NAME
End Function

Private Function pvMapToVirtualKey(ByVal lKey As Long, nVk As Integer, bExtended As Boolean) As Boolean
    Const STR_VK_MAP    As String = "&HFFE1;160|&HFFE2;161|&HFFE3;162|&HFFE4;163;1|&HFFE9;164|&HFFEA;165;1|&HFF51;37;1|&HFF52;38;1|&HFF53;39;1|&HFF54;40;1|&HFF09;9|&HFF08;8|&HFF0D;13|&HFF62;43|&HFF13;19|&HFF1B;27|&HFF0B;12|&HFFFF;46;1|&HFF50;36;1|&HFF55;33;1|&HFF56;34;1|&HFF57;35;1|&HFF60;41|&HFF61;44|&HFF63;45;1|&HFF6A;47|&HFF6B;3|&HFF80;32|&HFF89;9|&HFF8D;13;1|&HFF91;112|&HFF92;113|&HFF93;114|&HFF94;115|&HFF95;36|&HFF96;37|&HFF97;38|&HFF98;39|&HFF99;40|&HFF9A;33|&HFF9B;34|&HFF9C;35|&HFF9D;12|&HFF9E;45|&HFF9F;46|&HFFBD;187|&HFFAA;106|&HFFAB;107|&HFFAC;108|&HFFAD;109|&HFFAE;110|&HFFAF;111;1|&HFFB0;96|&HFFB1;97|&HFFB2;98|&HFFB3;99|&HFFB4;100|&HFFB5;101|&HFFB6;102|&HFFB7;103|&HFFB8;104|&HFFB9;105|&HFFBE;112|&HFFBF;113|&HFFC0;114|&HFFC1;115|&HFFC2;116|&HFFC3;117|&HFFC4;118|&HFFC5;119|&HFFC6;120|&HFFC7;121|&HFFC8;122|&HFFC9;123|&HFFCA;124|&HFFCB;125|&HFFCC;126|&HFFCD;127|&HFFCE;128|&HFFCF;129|&HFFD0;130|&HFFD1;131|&HFFD2;132|&HFFD3;133|&HFFD4;134|&HFFD5;135|&HFFEB;91|&HFFEC;92|&HFF67;93"
    Static vMap         As Variant
    Dim vElem           As Variant
    Dim vSplit          As Variant
    Dim lIdx            As Long
    
    If Not IsArray(vMap) Then
        vSplit = Split(STR_VK_MAP, "|")
        ReDim vMap(0 To UBound(vSplit)) As Variant
        For lIdx = 0 To UBound(vSplit)
            vElem = Split(vSplit(lIdx), ";")
            If UBound(vElem) >= 2 Then
                vMap(lIdx) = Array(CLng(vElem(0)), CInt(vElem(1)), CBool(vElem(2)))
            Else
                vMap(lIdx) = Array(CLng(vElem(0)), CInt(vElem(1)), False)
            End If
        Next
    End If
    For Each vElem In vMap
        If lKey = vElem(0) Then
            nVk = vElem(1)
            bExtended = vElem(2)
            pvMapToVirtualKey = True
            Exit For
        End If
    Next
End Function

Private Function pvHandlePointerEvent(ByVal lWidth As Long, ByVal lHeight As Long, uEvent As RfbClientPointerEvent) As Boolean
    Dim LNG_BUTTON1     As Long
    Const LNG_BUTTON2   As Long = 2 ^ 1
    Dim LNG_BUTTON3     As Long
    Const LNG_BUTTON4   As Long = 2 ^ 3
    Const LNG_BUTTON5   As Long = 2 ^ 4
    Dim lFlags          As Long
    Dim lWheel          As Long
    
    If GetSystemMetrics(SM_SWAPBUTTON) = 0 Then
        LNG_BUTTON1 = 2 ^ 0
        LNG_BUTTON3 = 2 ^ 2
    Else
        LNG_BUTTON1 = 2 ^ 2
        LNG_BUTTON3 = 2 ^ 0
    End If
    With uEvent
        lFlags = MOUSEEVENTF_ABSOLUTE Or MOUSEEVENTF_MOVE
        If (.ButtonMask And LNG_BUTTON1) <> (m_uClientPointerEvent.ButtonMask And LNG_BUTTON1) Then
            lFlags = lFlags Or IIf((.ButtonMask And LNG_BUTTON1) <> 0, MOUSEEVENTF_LEFTDOWN, MOUSEEVENTF_LEFTUP)
        End If
        If (.ButtonMask And LNG_BUTTON2) <> (m_uClientPointerEvent.ButtonMask And LNG_BUTTON2) Then
            lFlags = lFlags Or IIf((.ButtonMask And LNG_BUTTON2) <> 0, MOUSEEVENTF_MIDDLEDOWN, MOUSEEVENTF_MIDDLEUP)
        End If
        If (.ButtonMask And LNG_BUTTON3) <> (m_uClientPointerEvent.ButtonMask And LNG_BUTTON3) Then
            lFlags = lFlags Or IIf((.ButtonMask And LNG_BUTTON3) <> 0, MOUSEEVENTF_RIGHTDOWN, MOUSEEVENTF_RIGHTUP)
        End If
        If (.ButtonMask And LNG_BUTTON4) <> (m_uClientPointerEvent.ButtonMask And LNG_BUTTON4) Then
            lFlags = lFlags Or MOUSEEVENTF_WHEEL
            lWheel = 120
        End If
        If (.ButtonMask And LNG_BUTTON5) <> (m_uClientPointerEvent.ButtonMask And LNG_BUTTON5) Then
            lFlags = lFlags Or MOUSEEVENTF_WHEEL
            lWheel = -120
        End If
        Call mouse_event(lFlags, .XPosition * 65535 \ (lWidth - 1), .YPosition * 65535 \ (lHeight - 1), lWheel, 0)
    End With
    m_uClientPointerEvent = uEvent
    '--- success
    pvHandlePointerEvent = True
End Function

Private Function pvHandleCutEvent(uEvent As RfbClientCutText) As Boolean
    Const FUNC_NAME     As String = "pvHandleCutEvent"
    
    On Error GoTo EH
    Clipboard.Clear
    Clipboard.SetText Replace(StrConv(uEvent.Text, vbUnicode), vbLf, vbCrLf)
    '--- success
    pvHandleCutEvent = True
    Exit Function
EH:
    PrintError FUNC_NAME
End Function

'= file transfer =========================================================

Private Function pvSendFileTransfer( _
            uOutput As UcsBuffer, _
            ByVal ContentType As RfbFileTransferProtocolEnum, _
            Optional ByVal ContentParam As RfbFileTransferProtocolEnum, _
            Optional ByVal ContentSize As Long, _
            Optional ByVal HighContentSize As Variant, _
            Optional ByVal Ptr As Long, _
            Optional ByVal Size As Long, _
            Optional ByVal DelaySend As Boolean) As Boolean
    Dim uFileTransfer   As RfbFileTransfer
        
    With uFileTransfer
        .MessageType = rfbMsgFileTransfer
        .ContentType = ContentType
        .ContentParam = ContentParam
        .Size = pvNetworkLong(ContentSize)
        .Length = pvNetworkLong(Size)
    End With
    pvBufferWriteBlob uOutput, VarPtr(uFileTransfer), sizeof_RfbFileTransfer
    If Ptr <> 0 Then
        pvBufferWriteBlob uOutput, Ptr, Size
    End If
    If Not IsMissing(HighContentSize) Then
        pvBufferWriteLong uOutput, HighContentSize, Size:=4
    End If
    If DelaySend Then
        '--- success
        pvSendFileTransfer = True
    Else
        '--- succees (or failure)
        pvSendFileTransfer = pvBufferSendAsync(uOutput)
    End If
End Function

Private Function pvSendFileTransferFileContent(uOutput As UcsBuffer, pStream As stdole.IUnknown, sArchiveFile As String, pChecksums As stdole.IUnknown) As Boolean
    Const FUNC_NAME     As String = "pvSendFileTransferFileContent"
    Dim baBuffer()      As Byte
    Dim lRead           As Long
    Dim lChecksum       As Long
    
    On Error GoTo EH
    ReDim baBuffer(0 To rfbFtpProtocolPacketSize - 1) As Byte
    Do While Not m_oSocket.HasPendingEvent
        If pvCheckHResult(IStream_Read(pStream, pvArrayPtr(baBuffer), pvArraySize(baBuffer), lRead)) < 0 Then
            If Not pvSendFileTransfer(uOutput, rfbFtpAbortFileTransfer) Then
                GoTo QH
            End If
Cleanup:
            Set pStream = Nothing
            Call DeleteFile(StrPtr(sArchiveFile))
            Set pChecksums = Nothing
            Exit Do
        End If
        If lRead = 0 Then
            If Not pvSendFileTransfer(uOutput, rfbFtpEndOfFile) Then
                GoTo QH
            End If
            GoTo Cleanup
        End If
        If pChecksums Is Nothing Then
            lChecksum = 0
        ElseIf IStream_Read(pChecksums, VarPtr(lChecksum), 4) < 0 Then
            lChecksum = 0
            Set pChecksums = Nothing
        ElseIf lChecksum <> pvCalcAdler32(1, baBuffer, 0, lRead) Then
            lChecksum = 0
        Else
            lChecksum = 1
        End If
        If lChecksum <> 0 Then
            If Not pvSendFileTransfer(uOutput, rfbFtpFilePacket, 0, ContentSize:=rfbFtpSPacketAlreadyThere, _
                    Ptr:=0, Size:=lRead) Then
                GoTo QH
            End If
        Else
            If Not pvSendFileTransfer(uOutput, rfbFtpFilePacket, 0, ContentSize:=rfbFtpSPacketUncompressed, _
                    Ptr:=pvArrayPtr(baBuffer), Size:=lRead) Then
                GoTo QH
            End If
        End If
    Loop
    '--- success
    pvSendFileTransferFileContent = True
QH:
    Exit Function
EH:
    PrintError FUNC_NAME
End Function

Private Function pvHandleFileTransfer(uOutput As UcsBuffer, uMsg As RfbFileTransfer, baData() As Byte) As Boolean
    Const FUNC_NAME     As String = "pvHandleFileTransfer"
    Const STR_TYPES     As String = "0|DirContentRequest|DirPacket|FileTransferRequest|FileHeader|FilePacket|EndOfFile|AbortFileTransfer|FileTransferOffer|FileAcceptHeader|Command|CommandReturn|FileChecksums|13|FileTransferAccess|FileTransferSessionStart|FileTransferSessionEnd|FileTransferProtocolVersion"
    Const FORMAT_FILEDATETIME As String = "mm\/dd\/yyyy hh:nn"
    Const ERROR_VALUE   As Currency = -0.0001
    Const offsetof_cFileName As Long = sizeof_WIN32_FIND_DATA - 2 * MAX_PATH - 2 * 14
    Dim sPath           As String
    Dim baBuffer()      As Byte
    Dim lSize           As Long
    Dim lIdx            As Long
    Dim eStatus         As RfbFileTransferProtocolEnum
    Dim vSplit          As Variant
    Dim cStreamSize     As Currency
    Dim uStreamSize     As LARGE_INTEGER
    Dim sTemp           As String
    Dim hFind           As Long
    Dim uData           As WIN32_FIND_DATA
    Dim sFile           As String
    Dim lBufferSize     As Long
    Dim bDelaySend      As Boolean
    
    On Error GoTo EH
    If uMsg.ContentType <> rfbFtpFilePacket And uMsg.ContentType <> rfbFtpFileChecksums Then
        If UBound(baData) >= 0 Then
            sPath = StrConv(baData, vbUnicode)
        End If
    End If
    #If ImplUseDebugLog Then
        DebugLog MODULE_NAME, FUNC_NAME, At(Split(STR_TYPES, "|"), uMsg.ContentType) & _
            IIf(uMsg.ContentParam <> 0, ", ContentParam=" & uMsg.ContentParam, vbNullString) & _
            IIf(uMsg.Size <> 0, ", Size=" & uMsg.Size, vbNullString) & _
            IIf(uMsg.Length <> 0, ", Length=" & uMsg.Length, vbNullString) & _
            IIf(LenB(sPath) <> 0, ", sPath=" & sPath, vbNullString)
    #End If
    Select Case uMsg.ContentType
    Case rfbFtpDirContentRequest
        Select Case uMsg.ContentParam
        Case rfbFtpRDrivesList
            ReDim baBuffer(0 To 256) As Byte
            lSize = GetLogicalDriveStrings(UBound(baBuffer), baBuffer(0))
            For lIdx = 0 To lSize - 3 Step 4
                Select Case GetDriveType(baBuffer(lIdx))
                Case DRIVE_REMOVABLE
                    baBuffer(lIdx + 2) = 102    ' Asc("f")
                Case DRIVE_FIXED
                    baBuffer(lIdx + 2) = 108    ' Asc("l")
                Case DRIVE_REMOTE
                    baBuffer(lIdx + 2) = 110    ' Asc("n")
                Case DRIVE_CDROM
                    baBuffer(lIdx + 2) = 99     ' Asc("c")
                End Select
            Next
            If Not pvSendFileTransfer(uOutput, rfbFtpDirPacket, rfbFtpADrivesList, _
                    Ptr:=pvArrayPtr(baBuffer), Size:=lSize) Then
                GoTo QH
            End If
        Case rfbFtpRDirContent
            Select Case sPath
            Case "My Documents\"
                sPath = GetSpecialFolder(CSIDL_PERSONAL) & "\"
            Case "Desktop\"
                sPath = GetSpecialFolder(CSIDL_DESKTOP) & "\"
            Case "Network Favorites\"
                sPath = GetSpecialFolder(CSIDL_NETHOOD) & "\"
            End Select
            hFind = pvFindFirstFile(PathCombine(sPath, "*.*"), uData)
            If hFind = INVALID_HANDLE_VALUE Then
                If Not pvSendFileTransfer(uOutput, rfbFtpDirPacket, rfbFtpADirectory) Then
                    GoTo QH
                End If
            Else
                lBufferSize = m_oSocket.SockOpt(ucsSsoSendBuffer)
                baBuffer = StrConv(sPath, vbFromUnicode)
                If Not pvSendFileTransfer(uOutput, rfbFtpDirPacket, rfbFtpADirectory, Ptr:=pvArrayPtr(baBuffer), Size:=pvArraySize(baBuffer)) Then
                    GoTo QH
                End If
                ReDim baBuffer(0 To sizeof_WIN32_FIND_DATA) As Byte
                Do
                    sFile = Left$(uData.cFileName, InStr(uData.cFileName, Chr$(0)) - 1)
                    If sFile <> "." And sFile <> ".." Then
                        Call CopyMemory(baBuffer(0), uData, offsetof_cFileName)
                        Call CopyMemory(baBuffer(offsetof_cFileName), ByVal sFile, Len(sFile))
                        If uOutput.Size > uOutput.Pos + lBufferSize Then
                            bDelaySend = m_oSocket.HasPendingEvent
                        Else
                            bDelaySend = True
                        End If
                        If Not pvSendFileTransfer(uOutput, rfbFtpDirPacket, rfbFtpADirectory, _
                                Ptr:=pvArrayPtr(baBuffer), Size:=offsetof_cFileName + Len(sFile), DelaySend:=bDelaySend) Then
                            GoTo QH
                        End If
                    End If
                Loop While FindNextFile(hFind, VarPtr(uData)) <> 0
                '--- end of transfer
                If Not pvSendFileTransfer(uOutput, rfbFtpDirPacket) Then
                    GoTo QH
                End If
            End If
        Case Else
            #If ImplUseDebugLog Then
                DebugLog MODULE_NAME, FUNC_NAME, "Unknown DirContentRequest content param (" & uMsg.ContentParam & ")"
            #End If
        End Select
    Case rfbFtpCommand
        Select Case uMsg.ContentParam
        Case rfbFtpCDirCreate
            uMsg.ContentParam = rfbFtpADirCreate
            If CreateDirectory(StrPtr(sPath), 0) = 0 Then
                eStatus = rfbFtpSErrorCmd
            End If
        Case rfbFtpCFileDelete
            uMsg.ContentParam = rfbFtpAFileDelete
            If DeleteFile(StrPtr(sPath)) = 0 Then
                If RemoveDirectory(StrPtr(sPath)) = 0 Then
                    eStatus = rfbFtpSErrorCmd
                End If
            End If
        Case rfbFtpCFileRename
            uMsg.ContentParam = rfbFtpAFileRename
            vSplit = Split(sPath, "*", Limit:=2)
            If UBound(vSplit) < 1 Then
                eStatus = rfbFtpSErrorCmd
            ElseIf MoveFileEx(StrPtr(vSplit(0)), StrPtr(vSplit(1)), 0) = 0 Then
                eStatus = rfbFtpSErrorCmd
            End If
        Case Else
            eStatus = rfbFtpSErrorCmd
            sPath = vbNullString
        End Select
        baBuffer = StrConv(sPath, vbFromUnicode)
        If Not pvSendFileTransfer(uOutput, rfbFtpCommandReturn, uMsg.ContentParam, ContentSize:=eStatus, _
                Ptr:=pvArrayPtr(baBuffer), Size:=pvArraySize(baBuffer)) Then
            GoTo QH
        End If
    Case rfbFtpFileTransferRequest
        m_bFtpSendRunning = False
        Set m_pFtpSendStream = Nothing
        m_sFtpSendArchiveFile = vbNullString
        #If ImplZlib Then
            If GetFileAttributes(StrPtr(sPath)) = -1 Then
                sTemp = Mid$(sPath, InStrRev(sPath, "\") + 1)
                If Left(sTemp, 2) = "[ " And Right$(sTemp, 2) = " ]" Then
                    m_sFtpSendArchiveFile = Left$(sPath, Len(sPath) - Len(sTemp)) & STR_DIR_PREFIX & Mid$(sTemp, 3, Len(sTemp) - 4) & STR_ZIP_SUFFIX
                    sTemp = Left$(sPath, Len(sPath) - Len(sTemp)) & Mid$(sTemp, 3, Len(sTemp) - 4)
                    If (GetFileAttributes(StrPtr(sTemp)) And vbDirectory + vbVolume) = vbDirectory Then
                        With New cZipArchive
                            .AddFromFolder PathCombine(sTemp, "*.*"), Recursive:=True, IncludeEmptyFolders:=True
                            If .CompressArchive(m_sFtpSendArchiveFile) Then
                                sPath = m_sFtpSendArchiveFile
                            End If
                        End With
                    End If
                End If
            End If
        #End If
        If SHCreateStreamOnFile(StrPtr(sPath), STGM_READ, m_pFtpSendStream) < 0 Then
            cStreamSize = ERROR_VALUE
        ElseIf pvCheckHResult(IStream_Seek(m_pFtpSendStream, 0, STREAM_SEEK_END, cStreamSize)) < 0 Then
            cStreamSize = ERROR_VALUE
        ElseIf pvCheckHResult(IStream_Seek(m_pFtpSendStream, 0, STREAM_SEEK_SET)) < 0 Then
            cStreamSize = ERROR_VALUE
        Else
            sPath = sPath & "," & Format$(pvGetFileDateTime(sPath), FORMAT_FILEDATETIME)
        End If
        Call CopyMemory(uStreamSize, cStreamSize, 8)
        baBuffer = StrConv(sPath, vbFromUnicode)
        If Not pvSendFileTransfer(uOutput, rfbFtpFileHeader, 0, _
                ContentSize:=uStreamSize.LowPart, HighContentSize:=uStreamSize.HighPart, _
                Ptr:=pvArrayPtr(baBuffer), Size:=pvArraySize(baBuffer), DelaySend:=True) Then
            GoTo QH
        End If
    Case rfbFtpFileChecksums
        '--- note: used when resuming force stopped or connection broken transfers
        Set m_pFtpSendChecksums = SHCreateMemStream(ByVal pvArrayPtr(baData), pvArraySize(baData))
    Case rfbFtpFileHeader
        If uMsg.Size = rfbFtpSErrorCmd Then
            Set m_pFtpSendStream = Nothing
            Call DeleteFile(StrPtr(m_sFtpSendArchiveFile))
        Else
            m_bFtpSendRunning = True
        End If
    Case rfbFtpFileTransferOffer
        m_sFtpRecvFileName = Split(sPath, ",", Limit:=2)(0)
        If pvCheckHResult(pvCreateStreamOnFile(m_sFtpRecvFileName, STGM_WRITE Or STGM_CREATE, m_pFtpRecvStream)) < 0 Then
            eStatus = rfbFtpSErrorCmd
        End If
        baBuffer = StrConv(m_sFtpRecvFileName, vbFromUnicode)
        If Not pvSendFileTransfer(uOutput, rfbFtpFileAcceptHeader, ContentSize:=eStatus, _
                Ptr:=pvArrayPtr(baBuffer), Size:=pvArraySize(baBuffer)) Then
            GoTo QH
        End If
    Case rfbFtpFilePacket
        If IStream_Write(m_pFtpRecvStream, pvArrayPtr(baData), pvArraySize(baData)) < 0 Then
            If Not pvSendFileTransfer(uOutput, rfbFtpAbortFileTransfer) Then
                GoTo QH
            End If
        End If
    Case rfbFtpEndOfFile
        Set m_pFtpRecvStream = Nothing
        #If ImplZlib Then
            If Right$(m_sFtpRecvFileName, sizeof_ZIP_SUFFIX) = STR_ZIP_SUFFIX Then
                sPath = Mid$(m_sFtpRecvFileName, InStrRev(m_sFtpRecvFileName, "\") + 1)
                If Left$(sPath, sizeof_DIR_PREFIX) = STR_DIR_PREFIX Then
                    sPath = Left$(m_sFtpRecvFileName, Len(m_sFtpRecvFileName) - Len(sPath)) & _
                        Mid$(sPath, sizeof_DIR_PREFIX + 1, Len(sPath) - sizeof_DIR_PREFIX - sizeof_ZIP_SUFFIX)
                    With New cZipArchive
                        .OpenArchive m_sFtpRecvFileName
                        .Extract sPath
                    End With
                    Call DeleteFile(StrPtr(m_sFtpRecvFileName))
                End If
            End If
        #End If
    Case rfbFtpAbortFileTransfer
        If Not m_pFtpSendStream Is Nothing Then
            Set m_pFtpSendStream = Nothing
            Call DeleteFile(StrPtr(m_sFtpSendArchiveFile))
            If Not pvSendFileTransfer(uOutput, rfbFtpAbortFileTransfer) Then
                GoTo QH
            End If
        ElseIf Not m_pFtpRecvStream Is Nothing Then
            Set m_pFtpRecvStream = Nothing
        ElseIf Not pvSendFileTransfer(uOutput, IIf(uMsg.ContentParam = 0, rfbFtpAbortFileTransfer, rfbFtpFileTransferAccess), _
                ContentParam:=rfbFtpProtocolVersion, ContentSize:=rfbFtpSProtocolEnabled) Then
            GoTo QH
        End If
    End Select
    '--- success
    pvHandleFileTransfer = True
QH:
    Exit Function
EH:
    PrintError FUNC_NAME
End Function

Private Function pvHandleTightTransfer(uOutput As UcsBuffer, uMsg As UcsTightTransfer, baData() As Byte) As Boolean
    Const FUNC_NAME     As String = "pvHandleTightTransfer"
    Const STR_TYPES     As String = "CompressionSupportRequest|CompressionSupportReply|FileListRequest|FileListReply|Md5Request|Md5Reply|UploadStartRequest|UploadStartReply|UploadDataRequest|UploadDataReply|UploadEndRequest|UploadEndReply|DownloadStartRequest|DownloadStartReply|DownloadDataRequest|DownloadDataReply|DownloadEndReply|MkdirRequest|MkdirReply|RemoveRequest|RemoveReply|RenameRequest|RenameReply|DirsizeRequest|DirsizeReply|LastRequestFailedReply"
    Const FLAG_FILE     As Long = 0
    Const FLAG_DIR      As Long = 1
    Const NO_COMPRESSION As Long = 0
    Dim baBuffer()      As Byte
    Dim lSize           As Long
    Dim lIdx            As Long
    Dim sFolder         As String
    Dim uList           As UcsBuffer
    Dim sFile           As String
    Dim hFind           As Long
    Dim uData           As WIN32_FIND_DATA
    Dim lCount          As Long
    Dim cFileTime       As Currency
    
    On Error GoTo EH
    #If ImplUseDebugLog Then
        DebugLog MODULE_NAME, FUNC_NAME, At(Split(STR_TYPES, "|"), uMsg.MessageType And &HFF) & _
            IIf(uMsg.UncompressedSize <> 0, ", UncompressedSize=" & uMsg.UncompressedSize, vbNullString) & _
            IIf(uMsg.DataSize <> 0, ", DataSize=" & uMsg.DataSize, vbNullString) & _
            IIf(uMsg.Offset <> 0, ", Offset=" & uMsg.Offset, vbNullString) & _
            IIf(LenB(uMsg.Path) <> 0, ", Path=" & uMsg.Path, vbNullString)
    #End If
    Select Case uMsg.MessageType
    Case rfbMsgTightCompressionSupportRequest
        pvBufferWriteLong uOutput, rfbMsgTightCompressionSupportReply, Size:=4
        pvBufferWriteLong uOutput, NO_COMPRESSION
    Case rfbMsgTightFileListRequest
        sFolder = Mid$(uMsg.Path, 2)
        If LenB(sFolder) = 0 Then
            ReDim baBuffer(0 To 256) As Byte
            lSize = GetLogicalDriveStrings(UBound(baBuffer), baBuffer(0))
            pvBufferWriteLong uList, lSize \ 4, Size:=4
            For lIdx = 0 To lSize - 3 Step 4
                sFile = Space$(2)
                Call CopyMemory(ByVal sFile, baBuffer(lIdx), Len(sFile))
                pvBufferWriteCurrency uList, 0 '--- size
                pvBufferWriteCurrency uList, 0 '--- last modified
                pvBufferWriteLong uList, FLAG_DIR, Size:=2
                pvBufferWriteUtf8 uList, sFile
            Next
        Else
            '--- count will be fixed later
            pvBufferWriteLong uList, 0, Size:=4
            hFind = pvFindFirstFile(PathCombine(sFolder, "*.*"), uData)
            If hFind <> INVALID_HANDLE_VALUE Then
                Do
                    sFile = Left$(uData.cFileName, InStr(uData.cFileName, Chr$(0)) - 1)
                    If sFile <> "." And sFile <> ".." Then
                        lCount = lCount + 1
                        pvBufferWriteLong uList, uData.nFileSizeHigh, Size:=4
                        pvBufferWriteLong uList, uData.nFileSizeLow, Size:=4
                        pvBufferWriteCurrency uList, pvToUnixTime(uData.ftLastWriteTime)
                        pvBufferWriteLong uList, IIf((uData.dwFileAttributes And vbDirectory) <> 0, FLAG_DIR, FLAG_FILE), Size:=2
                        pvBufferWriteUtf8 uList, sFile
                    End If
                Loop While FindNextFile(hFind, VarPtr(uData)) <> 0
                '--- fix count of entries
                Call CopyMemory(ByVal pvArrayPtr(uList.Data), pvNetworkLong(lCount), 4)
            End If
        End If
        pvBufferWriteLong uOutput, rfbMsgTightFileListReply, Size:=4
        pvBufferWriteLong uOutput, NO_COMPRESSION
        pvBufferWriteLong uOutput, uList.Size, Size:=4 '-- uncompressed size
        pvBufferWriteLong uOutput, uList.Size, Size:=4 '-- compressed size
        pvBufferWriteBlob uOutput, pvArrayPtr(uList.Data), uList.Size
    Case rfbMsgTightUploadStartRequest
        sFile = Mid$(uMsg.Path, 2)
        If pvCheckHResult(pvCreateStreamOnFile(sFile, STGM_WRITE Or IIf((uMsg.Flags And 1) <> 0, STGM_CREATE, 0), m_pFtpRecvStream)) < 0 Then
            GoTo Failed
        End If
        If pvCheckHResult(IStream_Seek(m_pFtpRecvStream, uMsg.Offset, STREAM_SEEK_SET)) < 0 Then
            GoTo Failed
        End If
        pvBufferWriteLong uOutput, rfbMsgTightUploadStartReply, Size:=4
    Case rfbMsgTightUploadDataRequest
        If pvCheckHResult(IStream_Write(m_pFtpRecvStream, pvArrayPtr(baData), pvArraySize(baData))) < 0 Then
            GoTo Failed
        End If
        pvBufferWriteLong uOutput, rfbMsgTightUploadDataReply, Size:=4
    Case rfbMsgTightUploadEndRequest
        Set m_pFtpRecvStream = Nothing
        pvBufferWriteLong uOutput, rfbMsgTightUploadEndReply, Size:=4
    Case rfbMsgTightDownloadStartRequest
        m_sFtpRecvFileName = Mid(uMsg.Path, 2)
        If pvCheckHResult(pvCreateStreamOnFile(m_sFtpRecvFileName, STGM_READ, m_pFtpSendStream)) < 0 Then
            GoTo Failed
        End If
        If pvCheckHResult(IStream_Seek(m_pFtpSendStream, uMsg.Offset, STREAM_SEEK_SET)) < 0 Then
            GoTo Failed
        End If
        pvBufferWriteLong uOutput, rfbMsgTightDownloadStartReply, Size:=4
    Case rfbMsgTightDownloadDataRequest
        pvArrayAllocate baBuffer, Clamp(uMsg.DataSize, , 2 ^ 20), FUNC_NAME & ".baBuffer"
        If pvCheckHResult(IStream_Read(m_pFtpSendStream, pvArrayPtr(baBuffer), pvArraySize(baBuffer), lSize)) < 0 Then
            GoTo Failed
        End If
        If lSize = 0 Then
            pvBufferWriteLong uOutput, rfbMsgTightDownloadEndReply, Size:=4
            pvBufferWriteLong uOutput, FLAG_DIR
            pvGetFileDateTime m_sFtpRecvFileName, cFileTime
            pvBufferWriteCurrency uOutput, pvToUnixTime(cFileTime)
            Set m_pFtpSendStream = Nothing
        Else
            pvBufferWriteLong uOutput, rfbMsgTightDownloadDataReply, Size:=4
            pvBufferWriteLong uOutput, NO_COMPRESSION
            pvBufferWriteLong uOutput, lSize, Size:=4 '-- uncompressed size
            pvBufferWriteLong uOutput, lSize, Size:=4 '-- compressed size
            pvBufferWriteBlob uOutput, pvArrayPtr(baBuffer), lSize
        End If
    Case rfbMsgTightMkdirRequest
        sFolder = Mid$(uMsg.Path, 2)
        If CreateDirectory(StrPtr(sFolder), 0) = 0 Then
            GoTo Failed
        End If
        pvBufferWriteLong uOutput, rfbMsgTightMkdirReply, Size:=4
    Case rfbMsgTightRemoveRequest
        sFile = Mid$(uMsg.Path, 2)
        If DeleteFile(StrPtr(sFile)) = 0 Then
            If RemoveDirectory(StrPtr(sFile)) = 0 Then
                GoTo Failed
            End If
        End If
        pvBufferWriteLong uOutput, rfbMsgTightRemoveReply, Size:=4
    Case rfbMsgTightRenameRequest
        sFolder = Mid$(uMsg.Path, 2)
        sFile = Mid$(uMsg.NewFileName, 2)
        If MoveFileEx(StrPtr(sFolder), StrPtr(sFile), 0) = 0 Then
            GoTo Failed
        End If
        pvBufferWriteLong uOutput, rfbMsgTightRenameReply, Size:=4
    Case rfbMsgTightDirsizeRequest
        sFolder = Mid$(uMsg.Path, 2)
        pvBufferWriteLong uOutput, rfbMsgTightDirsizeReply, Size:=4
        pvBufferWriteCurrency uOutput, pvGetFolderSize(sFolder)
    Case 0
Failed:
        pvBufferWriteLong uOutput, rfbMsgTightLastRequestFailedReply, Size:=4
    End Select
    '--- success (or failure)
    pvHandleTightTransfer = pvBufferSendAsync(uOutput)
QH:
    If hFind <> INVALID_HANDLE_VALUE And hFind <> 0 Then
        Call FindClose(hFind)
    End If
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume QH
End Function

'= desktop duplication ===================================================

Public Function TimerProc() As Long
Attribute TimerProc.VB_MemberFlags = "40"
    Const FUNC_NAME     As String = "TimerProc"
    Dim oConn           As cVncServer
    
    On Error GoTo EH
    Set m_pTimer = InitFireOnceTimerThunk(Me, pvAddressOfTimerProc.TimerProc)
    For Each oConn In m_cConnections
        If Not oConn.frNotifyBeforeSendFrame() Then
            GoTo QH
        End If
    Next
    If Not m_uCaptureCtx.Duplication Is Nothing Then
        m_uCaptureFrame.NumTiles = 0
        m_uCaptureFrame.NumMoveRects = 0
        m_uCaptureFrame.PointerShapeBufferSize = 0
        If Not pvCaptureFrame(m_uCaptureCtx, m_uCaptureFrame) Then
            GoTo QH
        End If
    End If
    For Each oConn In m_cConnections
        oConn.frNotifySendFrame m_uCaptureFrame
    Next
    For Each oConn In m_cConnections
        oConn.frNotifyAfterSendFrame
    Next
QH:
    Exit Function
EH:
    PrintError FUNC_NAME
End Function

Friend Function frNotifyBeforeSendFrame() As Boolean
    Dim eChannel        As UcsSendChannelEnum
    
    eChannel = pvFirstPendingChannel
    If eChannel >= 0 Then
        If Not pvBufferSendAsync(m_uSendBuffer(eChannel)) Then
            GoTo QH
        End If
    Else
        '--- success
        frNotifyBeforeSendFrame = True
    End If
QH:
End Function

Friend Function frNotifySendFrame(uCaptureFrame As UcsCaptureFrame) As Boolean
    If m_uClientFramebufferUpdate.MessageType = 0 Then
        GoTo QH
    ElseIf Not pvHasSupport(rfbEncRaw) Then
        If Not pvSendEmptyFramebufferUpdate(m_uSendBuffer(ucsChaFrame)) Then
            GoTo QH
        End If
    ElseIf uCaptureFrame.NumTiles + uCaptureFrame.NumMoveRects + uCaptureFrame.PointerShapeBufferSize = 0 Then
        If Not pvSendEmptyFramebufferUpdate(m_uSendBuffer(ucsChaFrame)) Then
            GoTo QH
        End If
    Else
        If Not pvCaptureSend(pvParent.CaptureWidth, uCaptureFrame, m_uSendBuffer(ucsChaFrame)) Then
            GoTo QH
        End If
    End If
    '--- success
    frNotifySendFrame = True
QH:
End Function

Friend Function frNotifyAfterSendFrame() As Boolean
    If Not m_pFtpSendStream Is Nothing And m_bFtpSendRunning Then
        If m_uSendBuffer(ucsChaFrame).Size <= m_uSendBuffer(ucsChaFrame).Pos Then
            If Not pvSendFileTransferFileContent(m_uSendBuffer(ucsChaFtp), m_pFtpSendStream, m_sFtpSendArchiveFile, m_pFtpSendChecksums) Then
                GoTo QH
            End If
        End If
    End If
    '--- success
    frNotifyAfterSendFrame = True
QH:
End Function

Private Function pvCaptureEnumDevices() As Collection
    Dim pFactory        As IDXGIFactory1
    Dim lIdx            As Long
    Dim lJdx            As Long
    Dim pAdapter        As IDXGIAdapter1
    Dim pOutput         As IDXGIOutput1
    Dim uAdapterDesc    As DXGI_ADAPTER_DESC
    Dim uOutputDesc     As DXGI_OUTPUT_DESC
    
    Set pvCaptureEnumDevices = New Collection
    Set pFactory = CreateDXGIFactory1(IIDFromString(szIID_IDXGIFactory1))
    For lIdx = 0 To 100
        If pFactory.EnumAdapters1(lIdx, pAdapter) < 0 Then
            Exit For
        End If
        pAdapter.GetDesc uAdapterDesc
        For lJdx = 0 To 100
            If pAdapter.EnumOutputs(lJdx, pOutput) < 0 Then
                Exit For
            End If
            pOutput.GetDesc uOutputDesc
            pvCaptureEnumDevices.Add Array(Replace(uOutputDesc.DeviceName, vbNullChar, vbNullString), _
                Replace(uAdapterDesc.Description, vbNullChar, vbNullString))
        Next
    Next
End Function

Private Function pvCaptureInit(uCtx As UcsCaptureContext, ByVal sDeviceName As String, ByVal lTimeout As Long) As Boolean
    Const FUNC_NAME     As String = "pvCaptureInit"
    Dim hDesktop        As Long
    Dim pFactory        As IDXGIFactory1
    Dim lIdx            As Long
    Dim lJdx            As Long
    Dim pAdapter        As IDXGIAdapter1
    Dim pOutput         As IDXGIOutput1
    Dim uOutputDesc     As DXGI_OUTPUT_DESC
    Dim hResult         As Long
    Dim pD3D11Device    As ID3D11Device
    Dim pDXGIDevice     As IDXGIDevice1
    Dim uTextureDesc    As D3D11_TEXTURE2D_DESC
    Dim uDuplDesc       As DXGI_OUTDUPL_DESC
    Dim uResource       As D3D11_MAPPED_SUBRESOURCE
    
    On Error GoTo EH
    '--- allow capture the secure desktop
    hDesktop = OpenInputDesktop(0, 0, GENERIC_READ)
    If hDesktop <> 0 Then
        Call SetThreadDesktop(hDesktop)
        Call CloseDesktop(hDesktop)
    End If
    On Error Resume Next '--- Windows 10, version 1703 and above
    Call SetProcessDpiAwarenessContext(DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2)
    On Error GoTo EH
    With uCtx
        .DeviceName = vbNullString
        Set .DesktopResource = Nothing
        Set .Duplication = Nothing
        Set .StageTexture = Nothing
        Set .Context = Nothing
        Set pFactory = CreateDXGIFactory1(IIDFromString(szIID_IDXGIFactory1))
        For lIdx = 0 To 100
            hResult = pFactory.EnumAdapters1(lIdx, pAdapter)
            If hResult = DXGI_ERROR_NOT_FOUND Then
                Exit For
            End If
            If hResult < 0 Then
                Err.Raise hResult, "IDXGIFactory1.EnumAdapters1"
            End If
            For lJdx = 0 To 100
                hResult = pAdapter.EnumOutputs(lJdx, pOutput)
                If hResult = DXGI_ERROR_NOT_FOUND Then
                    Exit For
                End If
                If hResult < 0 Then
                    Err.Raise hResult, "IDXGIAdapter1.EnumOutputs"
                End If
                pOutput.GetDesc uOutputDesc
                If LenB(sDeviceName) <> 0 And Not Replace(uOutputDesc.DeviceName, vbNullChar, vbNullString) Like sDeviceName Then
                    GoTo Continue
                End If
                If uOutputDesc.AttachedToDesktop <> 0 Then
                    lIdx = 100
                    Exit For
                End If
Continue:
            Next
        Next
        If pOutput Is Nothing Then
            GoTo QH
        End If
        .DeviceName = Replace(uOutputDesc.DeviceName, vbNullChar, vbNullString)
        .Width = uOutputDesc.DesktopCoordinates.Right - uOutputDesc.DesktopCoordinates.Left
        .Height = uOutputDesc.DesktopCoordinates.Bottom - uOutputDesc.DesktopCoordinates.Top
        .Timeout = lTimeout
        hResult = D3D11CreateDevice(pAdapter, D3D_DRIVER_TYPE_UNKNOWN, 0, D3D11_CREATE_DEVICE_VIDEO_SUPPORT, ByVal 0, 0, D3D11_SDK_VERSION, pD3D11Device, 0, .Context)
        If hResult < 0 Then
            Err.Raise hResult, "D3D11CreateDevice"
        End If
        Call D3DKMTSetProcessSchedulingPriorityClass(GetCurrentProcess(), D3DKMT_SCHEDULINGPRIORITYCLASS_REALTIME)
        Set pDXGIDevice = pD3D11Device
        pDXGIDevice.SetGPUThreadPriority 7
        pDXGIDevice.SetMaximumFrameLatency 1
        hResult = pOutput.DuplicateOutput(pD3D11Device, .Duplication)
        If hResult >= 0 Then
            .Duplication.GetDesc uDuplDesc
            .InSystemMemory = (uDuplDesc.DesktopImageInSystemMemory <> 0)
            Debug.Assert uDuplDesc.ModeDesc.Format = DXGI_FORMAT_B8G8R8A8_UNORM
            With uTextureDesc
                .Width = uCtx.Width
                .Height = uCtx.Height
                .MipLevels = 1
                .ArraySize = 1
                .Format = uDuplDesc.ModeDesc.Format
                .SampleDesc.Count = 1
                .SampleDesc.Quality = 0
                .Usage = D3D11_USAGE_STAGING
                .BindFlags = 0
                .CPUAccessFlags = D3D11_CPU_ACCESS_READ
                .MiscFlags = 0
            End With
            Set .StageTexture = pD3D11Device.CreateTexture2D(uTextureDesc, ByVal 0)
            hResult = .Context.Map(.StageTexture, 0, D3D11_MAP_READ, 0, uResource)
            If hResult < 0 Then
                Err.Raise hResult, "ID3D11DeviceContext.Map"
            End If
            .Pitch = uResource.RowPitch
            .Stride = uResource.RowPitch / IIf(uDuplDesc.ModeDesc.Format = DXGI_FORMAT_R16G16B16A16_FLOAT, 8, 4)
            .Context.Unmap .StageTexture, 0
        End If
    End With
    '--- success
    pvCaptureInit = True
QH:
    Exit Function
EH:
    PrintError FUNC_NAME
End Function

Private Function pvCaptureFrame(uCtx As UcsCaptureContext, uFrame As UcsCaptureFrame) As Boolean
    Const FUNC_NAME     As String = "pvCaptureFrame"
    Const sizeof_OUTDUPL_MOVE_RECT As Long = 24
    Const sizeof_D3D11_RECT As Long = 16
    Dim hResult         As Long
    Dim lIdx            As Long
    Dim lY              As Long
    Dim lLeft           As Long
    Dim lSize           As Long
    Dim uFrameInfo      As DXGI_OUTDUPL_FRAME_INFO
    Dim pTex            As ID3D11Texture2D
    Dim uResource       As D3D11_MAPPED_SUBRESOURCE
    Dim uMapRect        As DXGI_MAPPED_RECT

    On Error GoTo EH
    With uCtx
        If .Duplication Is Nothing Then
            GoTo QH
        End If
        If Not .DesktopResource Is Nothing Then
            .Duplication.ReleaseFrame
            Set .DesktopResource = Nothing
        End If
        hResult = .Duplication.AcquireNextFrame(1, uFrameInfo, .DesktopResource)
        If hResult = DXGI_ERROR_WAIT_TIMEOUT Then
            GoTo QH
        End If
        If hResult < 0 Then
            GoTo QH
        End If
        If uFrameInfo.LastPresentTime.LowPart <> 0 Or uFrameInfo.LastPresentTime.HighPart <> 0 Then
            Set pTex = .StageTexture
        End If
        Call CopyMemory(lIdx, ByVal ArrPtr(uFrame.MoveRects), 4)
        If lIdx = 0 Then
            With uFrame
                ReDim .MoveRects(0 To 0) As DXGI_OUTDUPL_MOVE_RECT
                ReDim .DirtyRects(0 To 0) As D3D11_RECT
                ReDim .PointerShape(0 To 4095) As Byte
                ReDim .Tiles(0 To 127) As UcsCaptureTile
            End With
        End If
        '--- capture frame
        hResult = .Duplication.GetFrameMoveRects((UBound(uFrame.MoveRects) + 1) * sizeof_OUTDUPL_MOVE_RECT, uFrame.MoveRects(0), lSize)
        If hResult = DXGI_ERROR_MORE_DATA Then
            ReDim uFrame.MoveRects(0 To lSize \ sizeof_OUTDUPL_MOVE_RECT - 1) As DXGI_OUTDUPL_MOVE_RECT
            hResult = .Duplication.GetFrameMoveRects((UBound(uFrame.MoveRects) + 1) * sizeof_OUTDUPL_MOVE_RECT, uFrame.MoveRects(0), lSize)
        End If
        If hResult < 0 Then
            GoTo QH
        End If
        uFrame.NumMoveRects = lSize \ sizeof_OUTDUPL_MOVE_RECT
        hResult = .Duplication.GetFrameDirtyRects((UBound(uFrame.DirtyRects) + 1) * sizeof_D3D11_RECT, uFrame.DirtyRects(0), lSize)
        If hResult = DXGI_ERROR_MORE_DATA Then
            ReDim uFrame.DirtyRects(0 To lSize \ sizeof_D3D11_RECT - 1) As D3D11_RECT
            hResult = .Duplication.GetFrameDirtyRects((UBound(uFrame.DirtyRects) + 1) * sizeof_D3D11_RECT, uFrame.DirtyRects(0), lSize)
        End If
        If hResult < 0 Then
            GoTo QH
        End If
        uFrame.NumDirtyRects = lSize \ sizeof_D3D11_RECT
        If uFrame.NumDirtyRects > 0 Then
            ReDim uFrame.DirtyBuffers(0 To uFrame.NumDirtyRects - 1) As UcsBuffer
        End If
        If uFrameInfo.PointerShapeBufferSize > 0 Then
            hResult = .Duplication.GetFramePointerShape((UBound(uFrame.PointerShape) + 1), uFrame.PointerShape(0), uFrame.PointerSize, uFrame.PointerInfo)
            If hResult = DXGI_ERROR_MORE_DATA Then
                pvArrayAllocate uFrame.PointerShape, uFrame.PointerSize, FUNC_NAME
                hResult = .Duplication.GetFramePointerShape((UBound(uFrame.PointerShape) + 1), uFrame.PointerShape(0), uFrame.PointerSize, uFrame.PointerInfo)
            End If
            If hResult < 0 Then
                GoTo QH
            End If
            With uFrame
                .PointerShapeBufferSize = uFrameInfo.PointerShapeBufferSize
                If .PointerInfo.Type = DXGI_OUTDUPL_POINTER_SHAPE_TYPE_COLOR Then
                    For lIdx = 0 To UBound(.PointerShape) Step 4
                        '--- swap BGRA to RGBA quads
                        lLeft = .PointerShape(lIdx)
                        .PointerShape(lIdx) = .PointerShape(lIdx + 2)
                        .PointerShape(lIdx + 2) = lLeft
                    Next
                Else
                    .PointerInfo.Height = .PointerInfo.Height \ 2
                End If
            End With
        End If
        If uFrameInfo.LastMouseUpdateTime.LowPart <> 0 Or uFrameInfo.LastMouseUpdateTime.HighPart <> 0 Then
            uFrame.PointerVisible = (uFrameInfo.PointerPosition.Visible <> 0)
        End If
        If .InSystemMemory Then
            .Duplication.MapDesktopSurface uMapRect
            With uFrame
                For lIdx = 0 To .NumDirtyRects - 1
                    lLeft = .DirtyRects(lIdx).Left
                    lSize = (.DirtyRects(lIdx).Right - lLeft) * 4
                    '--- pre-allocate buffer
                    pvBufferWriteBlob .DirtyBuffers(lIdx), 0, (.DirtyRects(lIdx).Bottom - .DirtyRects(lIdx).Top) * lSize
                    .DirtyBuffers(lIdx).Size = 0
                    '--- copy rows
                    For lY = .DirtyRects(lIdx).Top To .DirtyRects(lIdx).Bottom - 1
                        pvBufferWriteBlob .DirtyBuffers(lIdx), uMapRect.pBits + lY * uMapRect.Pitch + lLeft * 4, lSize
                    Next
                Next
            End With
            .Duplication.UnMapDesktopSurface
        ElseIf Not pTex Is Nothing Then
            .Context.CopyResource pTex, .DesktopResource
            hResult = .Context.Map(pTex, 0, D3D11_MAP_READ, 0, uResource)
            If hResult < 0 Then
                Err.Raise hResult, "ID3D11DeviceContext.Map"
            End If
            With uFrame
                For lIdx = 0 To .NumDirtyRects - 1
                    lLeft = .DirtyRects(lIdx).Left
                    lSize = (.DirtyRects(lIdx).Right - lLeft) * 4
                    '--- pre-allocate buffer
                    pvBufferWriteBlob .DirtyBuffers(lIdx), 0, (.DirtyRects(lIdx).Bottom - .DirtyRects(lIdx).Top) * lSize
                    .DirtyBuffers(lIdx).Size = 0
                    '--- copy rows
                    For lY = .DirtyRects(lIdx).Top To .DirtyRects(lIdx).Bottom - 1
                        pvBufferWriteBlob .DirtyBuffers(lIdx), uResource.pData + lY * uResource.RowPitch + lLeft * 4, lSize
                    Next
                Next
            End With
            .Context.Unmap pTex, 0
            uResource.pData = 0
        End If
    End With
    '--- success (or failure)
    pvCaptureFrame = pvCapturePrepareTiles(uFrame)
QH:
    If uResource.pData <> 0 Then
        uCtx.Context.Unmap pTex, 0
    End If
    Exit Function
EH:
    PrintError FUNC_NAME
End Function

Private Function pvCaptureSend(ByVal lWidth As Long, uFrame As UcsCaptureFrame, uOutput As UcsBuffer) As Boolean
    Const FUNC_NAME     As String = "pvCaptureSend"
    Const LEVEL_CUTOFF  As Long = 85
    Dim lIdx            As Long
    Dim lX              As Long
    Dim lY              As Long
    Dim lPitch          As Long
    Dim lBytesPerPixel  As Long
    Dim lBufferSize     As Long
    Dim baRow()         As Byte
    Dim aRedLut()       As Long
    Dim aGreenLut()     As Long
    Dim aBlueLut()      As Long
    Dim aRedHist()      As Long
    Dim aGreenHist()    As Long
    Dim aBlueHist()     As Long
    Dim uUpdate         As RfbServerFramebufferUpdate
    Dim uRectangle      As RfbServerRectangle
    Dim lSrcIndex       As Long
    Dim lPtr            As Long
    Dim uTile           As UcsBuffer
    Dim baBit(0 To 7)   As Byte
    Dim baMask()        As Byte
    Dim bAndFlag        As Boolean
    Dim bXorFlag        As Boolean
    Dim cTemp           As Currency
    Dim lSize           As Long
    Dim lQualityLevel   As Long
    Dim lColor1         As Long
#If ImplHistograms Then
    Dim aHistogram()    As Long
    Dim lColor2         As Long
#End If
#If ImplUseDebugLog Then
    Dim sDebugInfo      As String
#End If
    
    On Error GoTo EH
    With uFrame
        lBytesPerPixel = Clamp(m_uClientPixelFormat.PixelFormat.BitsPerPixel \ 8, 1, 4)
        If lBytesPerPixel = 1 Or lBytesPerPixel = 2 Then
            pvCapturePrepareLut aRedLut, m_uClientPixelFormat.PixelFormat.RedMax, 2 ^ m_uClientPixelFormat.PixelFormat.RedShift
            pvCapturePrepareLut aGreenLut, m_uClientPixelFormat.PixelFormat.GreenMax, 2 ^ m_uClientPixelFormat.PixelFormat.GreenShift
            pvCapturePrepareLut aBlueLut, m_uClientPixelFormat.PixelFormat.BlueMax, 2 ^ m_uClientPixelFormat.PixelFormat.BlueShift
        End If
        #If ImplHistograms Then
            pvCapturePrepareLut aRedHist, 7, 2 ^ 5
            pvCapturePrepareLut aGreenHist, 7, 2 ^ 2
            pvCapturePrepareLut aBlueHist, 3, 2 ^ 0
        #End If
        pvArrayAllocate baRow, lWidth * 4, FUNC_NAME & ".baRow"
        '--- send pointer
        If .PointerShapeBufferSize > 0 And pvHasSupport(rfbEncPsCursor) Then
            Select Case .PointerInfo.Type
            Case DXGI_OUTDUPL_POINTER_SHAPE_TYPE_COLOR, DXGI_OUTDUPL_POINTER_SHAPE_TYPE_MONOCHROME
                lPitch = ((.PointerInfo.Width + 7) \ 8)
                pvArrayAllocate baMask, .PointerInfo.Height * lPitch, FUNC_NAME & ".baMask"
                For lIdx = 0 To 7
                    baBit(lIdx) = &H80 \ 2 ^ (lIdx Mod 8)
                Next
                With uUpdate
                    .MessageType = rfbMsgFramebufferUpdate
                    .NumberOfRectangles = pvNetworkShort(1)
                End With
                pvBufferWriteBlob uOutput, VarPtr(uUpdate), sizeof_RfbServerFramebufferUpdate
                With uRectangle
                    .XPosition = pvNetworkShort(uFrame.PointerInfo.HotSpot.X)
                    .YPosition = pvNetworkShort(uFrame.PointerInfo.HotSpot.Y)
                    .Width = pvNetworkShort(uFrame.PointerInfo.Width)
                    .Height = pvNetworkShort(uFrame.PointerInfo.Height)
                    If uFrame.PointerInfo.Type = DXGI_OUTDUPL_POINTER_SHAPE_TYPE_COLOR And pvHasSupport(rfbEncPsCursorWithAlpha) Then
                        .EncodingType = pvNetworkLong(rfbEncPsCursorWithAlpha)
                    Else
                        .EncodingType = pvNetworkLong(rfbEncPsCursor)
                    End If
                End With
                pvBufferWriteBlob uOutput, VarPtr(uRectangle), sizeof_RfbServerRectangle
                If .PointerInfo.Type = DXGI_OUTDUPL_POINTER_SHAPE_TYPE_COLOR And pvHasSupport(rfbEncPsCursorWithAlpha) Then
                    pvBufferWriteLong uOutput, pvNetworkLong(rfbEncRaw), Size:=4
                    pvBufferWriteBlob uOutput, VarPtr(.PointerShape(0)), .PointerInfo.Height * .PointerInfo.Width * 4
                ElseIf .PointerInfo.Type = DXGI_OUTDUPL_POINTER_SHAPE_TYPE_COLOR Then
                    Select Case lBytesPerPixel
                    Case 1, 2, 4
                        For lY = 0 To .PointerInfo.Height - 1
                            Debug.Assert IsBadReadPtr(VarPtr(.PointerShape(lY * .PointerInfo.Pitch)), .PointerInfo.Width * 4) = 0
                            Call CopyMemory(baRow(0), .PointerShape(lY * .PointerInfo.Pitch), .PointerInfo.Width * 4)
                            For lX = 0 To .PointerInfo.Width - 1
                                If (baRow(lX * 4 + 3)) > &H80 Then
                                    lIdx = lY * lPitch + lX \ 8
                                    baMask(lIdx) = baMask(lIdx) Or baBit(lX Mod 8)
                                End If
                                If lBytesPerPixel <> 4 Then
                                    lColor1 = aRedLut(baRow(lX * 4 + 2)) Or aGreenLut(baRow(lX * 4 + 1)) Or aBlueLut(baRow(lX * 4 + 0))
                                    If lBytesPerPixel = 1 Then
                                        baRow(lX) = lColor1
                                    Else
                                        baRow(lX * lBytesPerPixel + 0) = lColor1 And &HFF&
                                        baRow(lX * lBytesPerPixel + 1) = lColor1 \ &H100&
                                    End If
                                End If
                            Next
                            pvBufferWriteBlob uOutput, pvArrayPtr(baRow), .PointerInfo.Width * lBytesPerPixel
                        Next
                    Case Else
                        pvBufferWriteBlob uOutput, 0, .PointerInfo.Height * .PointerInfo.Width * lBytesPerPixel
                    End Select
                    pvBufferWriteArray uOutput, baMask
                Else
                    lPtr = .PointerInfo.Pitch * .PointerInfo.Height
                    For lY = 0 To .PointerInfo.Height - 1
                        For lX = 0 To .PointerInfo.Width - 1
                            bAndFlag = (.PointerShape(lY * .PointerInfo.Pitch + lX \ 8) And baBit(lX Mod 8)) <> 0
                            bXorFlag = (.PointerShape(lPtr + lY * .PointerInfo.Pitch + lX \ 8) And baBit(lX Mod 8)) <> 0
                            If Not bAndFlag Or bXorFlag Then
                                lIdx = lY * lPitch + lX \ 8
                                baMask(lIdx) = baMask(lIdx) Or baBit(lX Mod 8)
                            End If
                            pvBufferWriteLong uOutput, IIf(bAndFlag Xor bXorFlag, -1, 0), Size:=lBytesPerPixel
                        Next
                    Next
                    pvBufferWriteArray uOutput, baMask
                End If
            End Select
        End If
        '--- send desktop
        If pvHasSupport(rfbEncCopyRect) Then
            lSize = uFrame.NumMoveRects
        End If
        lBufferSize = m_oSocket.SockOpt(ucsSsoSendBuffer)
        With uUpdate
            .MessageType = rfbMsgFramebufferUpdate
            .NumberOfRectangles = pvNetworkShort(lSize + uFrame.NumTiles)
        End With
        pvBufferWriteBlob uOutput, VarPtr(uUpdate), sizeof_RfbServerFramebufferUpdate
        For lIdx = 0 To lSize - 1
            With .MoveRects(lIdx).DestinationRect
                uRectangle.XPosition = pvNetworkShort(.Left)
                uRectangle.YPosition = pvNetworkShort(.Top)
                uRectangle.Width = pvNetworkShort(.Right - .Left)
                uRectangle.Height = pvNetworkShort(.Bottom - .Top)
            End With
            uRectangle.EncodingType = pvNetworkLong(rfbEncCopyRect)
            pvBufferWriteBlob uOutput, VarPtr(uRectangle), sizeof_RfbServerRectangle
            pvBufferWriteLong uOutput, .MoveRects(lIdx).SourcePoint.X, Size:=2
            pvBufferWriteLong uOutput, .MoveRects(lIdx).SourcePoint.Y, Size:=2
        Next
        For lIdx = 0 To .NumTiles - 1
            #If ImplUseDebugLog Then
                sDebugInfo = lIdx & "/" & .NumTiles & ": " & .Tiles(lIdx).Left & "," & .Tiles(lIdx).Top
                sDebugInfo = sDebugInfo & Space$(18 - Len(sDebugInfo))
            #End If
            lSrcIndex = .Tiles(lIdx).SrcIndex
            uRectangle.XPosition = pvNetworkShort(.DirtyRects(lSrcIndex).Left + .Tiles(lIdx).Left)
            uRectangle.YPosition = pvNetworkShort(.DirtyRects(lSrcIndex).Top + .Tiles(lIdx).Top)
            uRectangle.Width = pvNetworkShort(.Tiles(lIdx).Width)
            uRectangle.Height = pvNetworkShort(.Tiles(lIdx).Height)
            '--- delay sending encoding
            pvBufferWriteBlob uOutput, VarPtr(uRectangle), sizeof_RfbServerRectangle - 4
            lPitch = (.DirtyRects(lSrcIndex).Right - .DirtyRects(lSrcIndex).Left) * 4
            lPtr = VarPtr(.DirtyBuffers(lSrcIndex).Data(0))
            uTile.Size = 0
            If pvHasSupport(rfbEncTight) And m_uClientEncodings.JpegQualityLevel > 0 And lBytesPerPixel > 1 Then
                #If ImplHistograms Then
                    ReDim aHistogram(0 To 255) As Long
                    For lY = .Tiles(lIdx).Top To .Tiles(lIdx).Top + .Tiles(lIdx).Height - 1
                        lSize = uTile.Size
                        pvBufferWriteBlob uTile, lPtr + lY * lPitch + .Tiles(lIdx).Left * 4, .Tiles(lIdx).Width * 4
                        For lX = 0 To .Tiles(lIdx).Width - 1
                            lColor1 = aRedHist(uTile.Data(lSize + lX * 4 + 2)) Or aGreenHist(uTile.Data(lSize + lX * 4 + 1)) Or aBlueHist(uTile.Data(lSize + lX * 4 + 0))
                            aHistogram(lColor1) = aHistogram(lColor1) + 1
                        Next
                    Next
                    lColor1 = 0
                    lColor2 = 0
                    For lX = 0 To 255
                        If aHistogram(lX) > lColor1 Then
                            lColor1 = aHistogram(lX)
                        ElseIf aHistogram(lX) > lColor2 Then
                            lColor2 = aHistogram(lX)
                        End If
                    Next
                    cTemp = (lColor1 + lColor2) * 100 / (.Tiles(lIdx).Width * .Tiles(lIdx).Height)
                    lQualityLevel = (m_uClientEncodings.JpegQualityLevel - 1) * 10
                    If cTemp > LEVEL_CUTOFF Then
                        lQualityLevel = lQualityLevel + (100 - lQualityLevel) * (cTemp - LEVEL_CUTOFF) / (100 - LEVEL_CUTOFF)
                    Else
                        lQualityLevel = lQualityLevel * cTemp / LEVEL_CUTOFF
                    End If
                #Else
                    For lY = .Tiles(lIdx).Top To .Tiles(lIdx).Top + .Tiles(lIdx).Height - 1
                        pvBufferWriteBlob uTile, lPtr + lY * lPitch + .Tiles(lIdx).Left * 4, .Tiles(lIdx).Width * 4
                    Next
                    lQualityLevel = (m_uClientEncodings.JpegQualityLevel - 1) * 10
                #End If
                #If ImplUseDebugLog Then
                    sDebugInfo = sDebugInfo & " H=" & Format$(cTemp, "0.00") & " QL=" & lQualityLevel
                #End If
                If pvCheckHResult(IStream_Seek(m_pImageStream, 0, STREAM_SEEK_SET)) < 0 Then
                    GoTo QH
                End If
                If Not pvWicConvertImage(m_pImageStream, m_aWicFormatJpeg, Clamp(lQualityLevel, 1, 100), _
                        uTile.Data, .Tiles(lIdx).Width, .Tiles(lIdx).Height, m_aWicFormat32bppPBGRA) Then
                    GoTo QH
                End If
                If pvCheckHResult(IStream_Seek(m_pImageStream, 0, STREAM_SEEK_CUR, cTemp)) < 0 Then
                    GoTo QH
                End If
                lSize = cTemp * 10000
                #If ImplUseDebugLog Then
                    cTemp = lSize * 100 / (.Tiles(lIdx).Width * .Tiles(lIdx).Height * 4)
                    sDebugInfo = sDebugInfo & " JPG=" & Format$(cTemp, "0.00") & "%"
                #End If
                pvBufferWriteLong uOutput, rfbEncTight, Size:=4
                '--- compression control: &H90 = no reset, JpegCompression
                pvBufferWriteLong uOutput, &H90
                pvBufferWriteCompactLong uOutput, lSize
                pvBufferWriteStream uOutput, m_pImageStream
            Else
                #If ImplZlib Then
                    If pvHasSupport(rfbEncTight) And m_uClientEncodings.CompressionLevel > 0 And lBytesPerPixel = 4 Then
                        lBytesPerPixel = 3
                    End If
                #End If
                Select Case lBytesPerPixel
                Case 1, 2, 3, 4
                    For lY = .Tiles(lIdx).Top To .Tiles(lIdx).Top + .Tiles(lIdx).Height - 1
                        If lBytesPerPixel = 1 Or lBytesPerPixel = 2 Then
                            Debug.Assert IsBadReadPtr(lPtr + lY * lPitch + .Tiles(lIdx).Left * 4, .Tiles(lIdx).Width * 4) = 0
                            Call CopyMemory(baRow(0), ByVal lPtr + lY * lPitch + .Tiles(lIdx).Left * 4, .Tiles(lIdx).Width * 4)
                            For lX = 0 To .Tiles(lIdx).Width - 1
                                lColor1 = aRedLut(baRow(lX * 4 + 2)) Or aGreenLut(baRow(lX * 4 + 1)) Or aBlueLut(baRow(lX * 4 + 0))
                                If lBytesPerPixel = 1 Then
                                    baRow(lX) = lColor1
                                Else
                                    baRow(lX * 2 + 0) = lColor1 And &HFF&
                                    baRow(lX * 2 + 1) = lColor1 \ &H100&
                                End If
                            Next
                            pvBufferWriteBlob uTile, pvArrayPtr(baRow), .Tiles(lIdx).Width
                        ElseIf lBytesPerPixel = 3 Then
                            Debug.Assert IsBadReadPtr(lPtr + lY * lPitch + .Tiles(lIdx).Left * 4, .Tiles(lIdx).Width * 4) = 0
                            Call CopyMemory(baRow(4), ByVal lPtr + lY * lPitch + .Tiles(lIdx).Left * 4, .Tiles(lIdx).Width * 4)
                            For lX = 0 To .Tiles(lIdx).Width - 1
                                baRow(lX * 3 + 0) = baRow(4 + lX * 4 + 2)
                                baRow(lX * 3 + 1) = baRow(4 + lX * 4 + 1)
                                baRow(lX * 3 + 2) = baRow(4 + lX * 4 + 0)
                            Next
                            pvBufferWriteBlob uTile, pvArrayPtr(baRow), .Tiles(lIdx).Width * 3
                        Else
                            pvBufferWriteBlob uTile, lPtr + lY * lPitch + .Tiles(lIdx).Left * 4, .Tiles(lIdx).Width * 4
                        End If
                    Next
                Case Else
                    pvBufferWriteBlob uTile, 0, .Tiles(lIdx).Height * .Tiles(lIdx).Width * lBytesPerPixel
                End Select
                #If ImplZlib Then
                    If pvHasSupport(rfbEncTight) And m_uClientEncodings.CompressionLevel > 0 Then
                        pvBufferWriteLong uOutput, rfbEncTight, Size:=4
                        lSize = uOutput.Size
                        '--- compression control: &H0 = no reset, use stream 0, no filter
                        pvBufferWriteLong uOutput, &H0
                        pvBufferWriteZlibBlob uOutput, VarPtr(uTile.Data(0)), uTile.Size, CompactSize:=True
                        #If ImplUseDebugLog Then
                            cTemp = (uOutput.Size - lSize) * 100 / (.Tiles(lIdx).Width * .Tiles(lIdx).Height * 4)
                            sDebugInfo = sDebugInfo & " ZLib=" & Format$(cTemp, "0.00") & "%"
                        #End If
                    ElseIf pvHasSupport(rfbEncZlib) Then
                        pvBufferWriteLong uOutput, rfbEncZlib, Size:=4
                        pvBufferWriteZlibBlob uOutput, VarPtr(uTile.Data(0)), uTile.Size
                    Else
                        pvBufferWriteLong uOutput, rfbEncRaw, Size:=4
                        pvBufferWriteBlob uOutput, VarPtr(uTile.Data(0)), uTile.Size
                    End If
                #Else
                    pvBufferWriteLong uOutput, rfbEncRaw, Size:=4
                    pvBufferWriteBlob uOutput, VarPtr(uTile.Data(0)), uTile.Size
                #End If
            End If
            #If ImplUseDebugLog Then
                If lIdx <= Form1.labDebug.UBound Then
                    Form1.labDebug(lIdx).Caption = sDebugInfo
                End If
            #End If
            If uOutput.Size > uOutput.Pos + lBufferSize Then
                If Not m_oSocket.HasPendingEvent Then
                    If Not pvBufferSendAsync(uOutput) Then
                        GoTo QH
                    End If
                End If
            End If
            If m_oSocket.AvailableBytes > 0 Then
                pvHandleReceive
                If lBytesPerPixel <> m_uClientPixelFormat.PixelFormat.BitsPerPixel \ 8 Then
                    lBytesPerPixel = Clamp(m_uClientPixelFormat.PixelFormat.BitsPerPixel \ 8, 1, 4)
                    If lBytesPerPixel = 1 Or lBytesPerPixel = 2 Then
                        pvCapturePrepareLut aRedLut, m_uClientPixelFormat.PixelFormat.RedMax, 2 ^ m_uClientPixelFormat.PixelFormat.RedShift
                        pvCapturePrepareLut aGreenLut, m_uClientPixelFormat.PixelFormat.GreenMax, 2 ^ m_uClientPixelFormat.PixelFormat.GreenShift
                        pvCapturePrepareLut aBlueLut, m_uClientPixelFormat.PixelFormat.BlueMax, 2 ^ m_uClientPixelFormat.PixelFormat.BlueShift
                    End If
                End If
            End If
        Next
        #If ImplUseDebugLog Then
            For lIdx = .NumTiles To Form1.labDebug.UBound
                Form1.labDebug(lIdx).Caption = vbNullString
            Next
        #End If
    End With
    If Not m_oSocket.HasPendingEvent Then
        If Not pvBufferSendAsync(uOutput) Then
            GoTo QH
        End If
    End If
    '--- success
    pvCaptureSend = True
QH:
    Exit Function
EH:
    PrintError FUNC_NAME
End Function

Private Function pvCapturePrepareTiles(uFrame As UcsCaptureFrame) As Boolean
    Const FUNC_NAME     As String = "pvCapturePrepareTiles"
    Dim lIdx            As Long
    Dim lWidth          As Long
    Dim lHeight         As Long
    Dim lX              As Long
    Dim lY              As Long
    
    On Error GoTo EH
    uFrame.NumTiles = 0
    For lIdx = 0 To uFrame.NumDirtyRects - 1
        With uFrame.DirtyRects(lIdx)
            lWidth = .Right - .Left
            lHeight = .Bottom - .Top
        End With
        For lY = 0 To lHeight - 1 Step LNG_TILE_SIZE
            For lX = 0 To lWidth - 1 Step LNG_TILE_SIZE
                If uFrame.NumTiles > UBound(uFrame.Tiles) Then
                    ReDim Preserve uFrame.Tiles(0 To UBound(uFrame.Tiles) * 2 + 1) As UcsCaptureTile
                End If
                With uFrame.Tiles(uFrame.NumTiles)
                    .SrcIndex = lIdx
                    .Left = lX
                    .Top = lY
                    .Width = Clamp(lX + LNG_TILE_SIZE, , lWidth) - lX
                    .Height = Clamp(lY + LNG_TILE_SIZE, , lHeight) - lY
                End With
                uFrame.NumTiles = uFrame.NumTiles + 1
            Next
        Next
    Next
    '--- success
    pvCapturePrepareTiles = True
    Exit Function
EH:
    PrintError FUNC_NAME
End Function

Private Sub pvCapturePrepareLut(aLut() As Long, ByVal lMax As Long, ByVal lShift As Long)
    Const FUNC_NAME     As String = "pvCapturePrepareLut"
    Dim lIdx            As Long
    
    On Error GoTo EH
    ReDim aLut(0 To 255) As Long
    For lIdx = 0 To 255
        aLut(lIdx) = Int(lIdx * lMax / 255 + 0.5) * lShift
    Next
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

'= helpers ===============================================================

Private Sub pvInitTightCapabilities(uCaps() As RfbTightCapability, sInit As String)
    Dim vSplit          As Variant
    Dim vElem           As Variant
    Dim lIdx            As Long
    Dim sBuffer         As String
    
    vSplit = Split(sInit, "|")
    ReDim uCaps(0 To UBound(vSplit)) As RfbTightCapability
    For Each vElem In vSplit
        vElem = Split(vElem, ";")
        With uCaps(lIdx)
            .Code = pvNetworkLong(vElem(0))
            sBuffer = vElem(1)
            Debug.Assert Len(sBuffer) = 4
            Call CopyMemory(.Vendor(0), ByVal sBuffer, Len(sBuffer))
            sBuffer = vElem(2)
            Debug.Assert Len(sBuffer) = 8
            Call CopyMemory(.Signature(0), ByVal sBuffer, Len(sBuffer))
        End With
        lIdx = lIdx + 1
    Next
End Sub

Private Function pvWicConvertImage( _
            pOutput As stdole.IUnknown, _
            aContainerFormat() As Long, _
            ByVal lImageQuality As Long, _
            baInput() As Byte, _
            ByVal lWidth As Long, _
            ByVal lHeight As Long, _
            aPixelFormat() As Long) As Boolean
    Const FUNC_NAME     As String = "pvWicConvertImage"
    Const WICBitmapEncoderNoCache As Long = 2
    Dim pBitmap         As stdole.IUnknown
    Dim pWicStream      As stdole.IUnknown
    Dim pEncoder        As stdole.IUnknown
    Dim pFrame          As stdole.IUnknown
    Dim pPropBag        As stdole.IUnknown
    Dim aBag(0 To 7)    As Long
    
    On Error GoTo EH
    If pvCheckHResult(IWICImagingFactory_CreateBitmapFromMemory_Proxy(m_pWicFactory, lWidth, lHeight, aPixelFormat(0), lWidth * 4, UBound(baInput) + 1, baInput(0), pBitmap)) < 0 Then
        GoTo QH
    End If
    If pvCheckHResult(IWICImagingFactory_CreateStream_Proxy(m_pWicFactory, pWicStream)) < 0 Then
        GoTo QH
    End If
    If pvCheckHResult(IWICStream_InitializeFromIStream_Proxy(pWicStream, pOutput)) < 0 Then
        GoTo QH
    End If
    If pvCheckHResult(IWICImagingFactory_CreateEncoder_Proxy(m_pWicFactory, aContainerFormat(0), ByVal 0, pEncoder)) < 0 Then
        GoTo QH
    End If
    If pvCheckHResult(IWICBitmapEncoder_Initialize_Proxy(pEncoder, pWicStream, WICBitmapEncoderNoCache)) < 0 Then
        GoTo QH
    End If
    If pvCheckHResult(IWICBitmapEncoder_CreateNewFrame_Proxy(pEncoder, pFrame, pPropBag)) < 0 Then
        GoTo QH
    End If
    If lImageQuality <> 0 Then
        aBag(3) = StrPtr("ImageQuality")
        If pvCheckHResult(IPropertyBag2_Write_Proxy(pPropBag, 1, aBag(0), CSng(lImageQuality) / 100!)) < 0 Then
            GoTo QH
        End If
    End If
    If pvCheckHResult(IWICBitmapFrameEncode_Initialize_Proxy(pFrame, pPropBag)) < 0 Then
        GoTo QH
    End If
    If pvCheckHResult(IWICBitmapFrameEncode_SetSize_Proxy(pFrame, lWidth, lHeight)) < 0 Then
        GoTo QH
    End If
    If pvCheckHResult(IWICBitmapFrameEncode_WriteSource_Proxy(pFrame, pBitmap, ByVal 0)) < 0 Then
        GoTo QH
    End If
    If pvCheckHResult(IWICBitmapFrameEncode_Commit_Proxy(pFrame)) < 0 Then
        GoTo QH
    End If
    If pvCheckHResult(IWICBitmapEncoder_Commit_Proxy(pEncoder)) < 0 Then
        GoTo QH
    End If
    '--- success
    pvWicConvertImage = True
QH:
    Exit Function
EH:
    PrintError FUNC_NAME
End Function

Private Function pvCheckHResult(ByVal hResult As Long, Optional vErrSource As Variant) As Long
    Const FUNC_NAME     As String = "pvCheckHResult"
    
    On Error GoTo EH
    If hResult < 0 Then
        Err.Raise hResult, vErrSource
    End If
QH:
    pvCheckHResult = pvCheckHResult
    Exit Function
EH:
    PrintError IIf(IsMissing(vErrSource), FUNC_NAME, vErrSource)
    Resume QH
End Function

Private Function pvGetEncodingBit(ByVal eEncoding As RfbEncodingEnum) As Long
    Select Case eEncoding
    Case rfbEncRaw
        pvGetEncodingBit = 2 ^ 0
    Case rfbEncCopyRect
        pvGetEncodingBit = 2 ^ 1
    Case rfbEncZlib
        pvGetEncodingBit = 2 ^ 2
    Case rfbEncTight
        pvGetEncodingBit = 2 ^ 3
    Case rfbEncPsCursorWithAlpha
        pvGetEncodingBit = 2 ^ 4
    Case rfbEncPsCursor
        pvGetEncodingBit = 2 ^ 5
    End Select
End Function

Private Function pvFindFirstFile(sPath As String, uData As WIN32_FIND_DATA) As Long
    Dim lMode           As Long
    
    lMode = SetErrorMode(SEM_FAILCRITICALERRORS)
    pvFindFirstFile = FindFirstFile(StrPtr(sPath), VarPtr(uData))
    Call SetErrorMode(lMode)
End Function

Private Function pvCreateStreamOnFile(sFile As String, ByVal lFileMode As Long, pStream As stdole.IUnknown) As Long
    Dim lMode           As Long
    
    lMode = SetErrorMode(SEM_FAILCRITICALERRORS)
    pvCreateStreamOnFile = SHCreateStreamOnFile(StrPtr(sFile), lFileMode, pStream)
    Call SetErrorMode(lMode)
End Function

'= stream helpers ========================================================

Private Function IStream_Read(pstm As stdole.IUnknown, ByVal lPtr As Long, ByVal lSize As Long, Optional BytesRead As Long) As Long
    If Not pstm Is Nothing Then
        IStream_Read = DispCallByVtbl(pstm, 3, lPtr, lSize, VarPtr(BytesRead))
    Else
        IStream_Read = -1
    End If
End Function

Private Function IStream_Write(pstm As IUnknown, ByVal lPtr As Long, ByVal lSize As Long, Optional BytesWritten As Long) As Long
    If Not pstm Is Nothing Then
        IStream_Write = DispCallByVtbl(pstm, 4, lPtr, lSize, VarPtr(BytesWritten))
    Else
        IStream_Write = -1
    End If
End Function

Private Function IStream_Seek(pstm As stdole.IUnknown, ByVal cMove As Currency, ByVal dwOrigin As Long, Optional NewPosition As Currency) As Long
    If Not pstm Is Nothing Then
        IStream_Seek = DispCallByVtbl(pstm, 5, cMove, dwOrigin, VarPtr(NewPosition))
    Else
        IStream_Seek = -1
    End If
End Function

Private Function DispCallByVtbl(pUnk As stdole.IUnknown, ByVal lIndex As Long, ParamArray A() As Variant) As Variant
    Const CC_STDCALL    As Long = 4
    Dim lIdx            As Long
    Dim vParam()        As Variant
    Dim vType(0 To 63)  As Integer
    Dim vPtr(0 To 63)   As Long
    Dim hResult         As Long
    
    vParam = A
    For lIdx = 0 To UBound(vParam)
        vType(lIdx) = VarType(vParam(lIdx))
        vPtr(lIdx) = VarPtr(vParam(lIdx))
    Next
    hResult = DispCallFunc(ObjPtr(pUnk), lIndex * 4, CC_STDCALL, vbLong, lIdx, vType(0), vPtr(0), DispCallByVtbl)
    If hResult < 0 Then
        Err.Raise hResult, "DispCallFunc"
    End If
End Function

'= buffer helpers ========================================================

Private Sub pvBufferWriteArray(uOutput As UcsBuffer, baSrc() As Byte)
    Dim lSize       As Long
    
    With uOutput
        lSize = pvArraySize(baSrc)
        If lSize > 0 Then
            .Size = pvArrayWriteBlob(.Data, .Size, pvArrayPtr(baSrc), lSize)
        End If
    End With
End Sub

Private Sub pvBufferWriteBlob(uOutput As UcsBuffer, ByVal lPtr As Long, ByVal lSize As Long)
    uOutput.Size = pvArrayWriteBlob(uOutput.Data, uOutput.Size, lPtr, lSize)
End Sub

#If ImplZlib Then
Private Sub pvBufferWriteZlibBlob(uOutput As UcsBuffer, ByVal lPtr As Long, ByVal lSize As Long, Optional ByVal CompactSize As Boolean)
    Dim bFirst          As Boolean
    Dim lOutputPtr      As Long
    Dim lOutputSize     As Long
    
    If m_oZipArchive Is Nothing Then
        Set m_oZipArchive = New cZipArchive
        m_hDeflateCtx = m_oZipArchive.DeflateInit
        bFirst = True
    End If
    If m_oZipArchive.DeflateBlob(m_hDeflateCtx, lPtr, lSize, lOutputPtr, lOutputSize, Level:=IIf(m_uClientEncodings.CompressionLevel > 0, m_uClientEncodings.CompressionLevel - 1, 6)) Then
        If CompactSize Then
            pvBufferWriteCompactLong uOutput, lOutputSize + IIf(bFirst, 2, 0)
        Else
            pvBufferWriteLong uOutput, lOutputSize + IIf(bFirst, 2, 0), Size:=4
        End If
        If bFirst Then
            pvBufferWriteLong uOutput, &H789C&, Size:=2
        End If
        pvBufferWriteBlob uOutput, lOutputPtr, lOutputSize
        Call CoTaskMemFree(lOutputPtr)
    End If
End Sub
#End If

Private Function pvArrayWriteBlob(baBuffer() As Byte, ByVal lPos As Long, ByVal lPtr As Long, ByVal lSize As Long) As Long
    Const FUNC_NAME     As String = "pvArrayWriteBlob"
    Dim lBufPtr         As Long
    
    Call CopyMemory(lBufPtr, ByVal ArrPtr(baBuffer), 4)
    If lBufPtr = 0 Then
        pvArrayAllocate baBuffer, Clamp(lPos + lSize, 256), FUNC_NAME & ".baBuffer"
    ElseIf UBound(baBuffer) < lPos + lSize - 1 Then
        pvArrayReallocate baBuffer, lPos + lSize, FUNC_NAME & ".baBuffer"
    End If
    If lSize > 0 And lPtr <> 0 Then
        Debug.Assert IsBadReadPtr(lPtr, lSize) = 0
        Call CopyMemory(baBuffer(lPos), ByVal lPtr, lSize)
    End If
    pvArrayWriteBlob = lPos + lSize
End Function

Private Sub pvBufferWriteLong(uOutput As UcsBuffer, ByVal lValue As Long, Optional ByVal Size As Long = 1)
    Static baTemp(0 To 3) As Byte
    Dim lPos            As Long

    With uOutput
        If Size <= 1 Then
            pvBufferWriteBlob uOutput, VarPtr(lValue), Size
        Else
            lPos = .Size
            pvBufferWriteBlob uOutput, 0, Size
            Call CopyMemory(baTemp(0), lValue, 4)
            .Data(lPos + 0) = baTemp(Size - 1)
            .Data(lPos + 1) = baTemp(Size - 2)
            If Size >= 3 Then .Data(lPos + 2) = baTemp(Size - 3)
            If Size >= 4 Then .Data(lPos + 3) = baTemp(Size - 4)
        End If
    End With
End Sub

Private Sub pvBufferWriteCompactLong(uOutput As UcsBuffer, ByVal lValue As Long)
    Dim lNext           As Long
    
    lNext = lValue \ &H80
    pvBufferWriteLong uOutput, (lValue And &H7F) Or IIf(lNext <> 0, &H80, 0)
    If lNext <> 0 Then
        lValue = lNext
        lNext = lValue \ &H80
        pvBufferWriteLong uOutput, (lValue And &H7F) Or IIf(lNext <> 0, &H80, 0)
        If lNext <> 0 Then
            pvBufferWriteLong uOutput, lNext
        End If
    End If
End Sub

Private Sub pvBufferWriteCurrency(uOutput As UcsBuffer, cValue As Currency)
    Dim uValue          As LARGE_INTEGER
    
    Call CopyMemory(uValue, cValue, 8)
    pvBufferWriteLong uOutput, uValue.HighPart, Size:=4
    pvBufferWriteLong uOutput, uValue.LowPart, Size:=4
End Sub

Private Sub pvBufferWriteStream(uOutput As UcsBuffer, pStream As stdole.IUnknown)
    Dim cTemp           As Currency
    Dim lSize           As Long
    
    If pvCheckHResult(IStream_Seek(pStream, 0, STREAM_SEEK_CUR, cTemp)) < 0 Then
        GoTo QH
    End If
    lSize = cTemp * 10000
    If lSize > 0 Then
        pvBufferWriteBlob uOutput, 0, lSize
        If pvCheckHResult(IStream_Seek(pStream, 0, STREAM_SEEK_SET)) < 0 Then
            GoTo QH
        End If
        If pvCheckHResult(IStream_Read(pStream, VarPtr(uOutput.Data(uOutput.Size - lSize)), lSize)) < 0 Then
            GoTo QH
        End If
    End If
QH:
End Sub

Private Sub pvBufferWriteUtf8(uOutput As UcsBuffer, sText As String)
    Dim lSize           As Long
    Dim lPos            As Long
    
    lSize = WideCharToMultiByte(CP_UTF8, 0, StrPtr(sText), Len(sText), ByVal 0, 0, 0, 0)
    If lSize > 0 Then
        pvBufferWriteLong uOutput, lSize + 1, Size:=4
        lPos = uOutput.Size
        pvBufferWriteBlob uOutput, 0, lSize + 1
        Call WideCharToMultiByte(CP_UTF8, 0, StrPtr(sText), Len(sText), uOutput.Data(lPos), lSize, 0, 0)
    Else
        pvBufferWriteLong uOutput, 0, Size:=4
    End If
End Sub

Private Function pvBufferSendAsync(uOutput As UcsBuffer) As Boolean
    Const FUNC_NAME     As String = "pvBufferSendAsync"
    Dim eChannel        As UcsSendChannelEnum
    Dim lBytes          As Long
    Dim lBufferSize     As Long
    
    On Error GoTo EH
    For eChannel = 0 To [_ucsChaMax] - 1
        If VarPtr(m_uSendBuffer(eChannel)) = VarPtr(uOutput) Then
            Exit For
        End If
    Next
    If eChannel = pvFirstPendingChannel Then
        lBufferSize = m_oSocket.SockOpt(ucsSsoSendBuffer)
        Do While uOutput.Pos < uOutput.Size
            lBytes = IIf(uOutput.Size - uOutput.Pos > lBufferSize, lBufferSize, uOutput.Size - uOutput.Pos)
            lBytes = m_oSocket.Send(VarPtr(uOutput.Data(uOutput.Pos)), lBytes)
            If m_oSocket.HasPendingEvent Then
                Exit Do
            ElseIf lBytes < 0 Then
                GoTo QH
            Else
                uOutput.Pos = uOutput.Pos + lBytes
            End If
        Loop
        If uOutput.Pos > 0 Then
            If uOutput.Pos >= uOutput.Size Then
                uOutput.Pos = 0
                uOutput.Size = 0
            End If
        End If
    End If
    '--- success
    pvBufferSendAsync = True
QH:
    Exit Function
EH:
    PrintError FUNC_NAME
End Function

Private Sub pvBufferReadLong(uInput As UcsBuffer, lValue As Long, Optional ByVal Size As Long = 1)
    lValue = 0
    Call CopyMemory(lValue, uInput.Data(uInput.Pos), Size)
    Select Case Size
    Case 2
        lValue = pvNetworkShort(lValue)
    Case 3
        lValue = pvNetworkLong(lValue * &H100)
    Case 4
        lValue = pvNetworkLong(lValue)
    End Select
    uInput.Pos = uInput.Pos + Size
End Sub

Private Sub pvBufferReadCurrency(uInput As UcsBuffer, cValue As Currency)
    Dim uValue          As LARGE_INTEGER
    
    pvBufferReadLong uInput, uValue.HighPart, Size:=4
    pvBufferReadLong uInput, uValue.LowPart, Size:=4
    Call CopyMemory(cValue, uValue, 8)
End Sub

Private Sub pvBufferReadUtf8(uInput As UcsBuffer, sValue As String)
    Dim lBytes          As Long
    Dim lSize           As Long
    
    Call CopyMemory(lBytes, uInput.Data(uInput.Pos), 4)
    lBytes = pvNetworkLong(lBytes)
    If lBytes > 0 Then
        sValue = String$(2 * lBytes, 0)
        lSize = MultiByteToWideChar(CP_UTF8, 0, uInput.Data(uInput.Pos + 4), lBytes, StrPtr(sValue), Len(sValue))
        sValue = Left$(sValue, lSize)
        If Right$(sValue, 1) = vbNullChar Then
            sValue = Left$(sValue, Len(sValue) - 1)
        End If
    End If
    uInput.Pos = uInput.Pos + 4 + lBytes
End Sub

'= arrays helpers ========================================================

Private Sub pvArrayAllocate(baRetVal() As Byte, ByVal lSize As Long, sFuncName As String)
    On Error GoTo EH
    If lSize > 0 Then
        ReDim baRetVal(0 To lSize - 1) As Byte
    Else
        baRetVal = vbNullString
    End If
    Exit Sub
EH:
    PrintError sFuncName
End Sub

Private Sub pvArrayReallocate(baArray() As Byte, ByVal lSize As Long, sFuncName As String)
    On Error GoTo EH
    If lSize > 0 Then
        ReDim Preserve baArray(0 To lSize - 1) As Byte
    Else
        baArray = vbNullString
    End If
    Exit Sub
EH:
    PrintError sFuncName
End Sub

Private Property Get pvArrayPtr(baArray() As Byte) As Long
    Dim lPtr            As Long
    
    Call CopyMemory(lPtr, ByVal ArrPtr(baArray), 4)
    If lPtr <> 0 Then
        Call CopyMemory(pvArrayPtr, ByVal lPtr + 12, 4)
    End If
End Property

Private Property Get pvArraySize(baArray() As Byte) As Long
    Dim lPtr            As Long
    
    Call CopyMemory(lPtr, ByVal ArrPtr(baArray), 4)
    If lPtr <> 0 Then
        pvArraySize = UBound(baArray) + 1
    End If
End Property

Private Function pvArrayEqual(baFirst() As Byte, ByVal lFirstPos As Long, ByVal lSize As Long, baSecond() As Byte, Optional ByVal lSecondPos As Long) As Boolean
    Dim lIdx            As Long
    Dim lResult         As Long
    
    For lIdx = 0 To lSize - 1
        lResult = baFirst(lFirstPos + lIdx) Xor baSecond(lSecondPos + lIdx) Or lResult
    Next
    pvArrayEqual = (lResult = 0)
End Function

'= crypto helpers ========================================================

Private Function pvCryptoRandomBytes(ByVal lSize As Long) As Byte()
    Const FUNC_NAME     As String = "pvCryptoRandomBytes"
    Dim baRetVal()      As Byte
    
    pvArrayAllocate baRetVal, lSize, FUNC_NAME & ".baRetVal"
    Call RtlGenRandom(baRetVal(0), lSize)
    pvCryptoRandomBytes = baRetVal
End Function

Private Function pvCryptoDesEncrypt(baInput() As Byte, baKey() As Byte) As Boolean
    Const FUNC_NAME     As String = "pvCryptoDesEncrypt"
    Const LNG_FACILITY_WIN32 As Long = &H80070000
    Dim hProv           As Long
    Dim uBlob           As BLOBHEADER
    Dim lIdx            As Long
    Dim hKey            As Long
    Dim lSize           As Long
    Dim sApiName        As String
    
    If CryptAcquireContext(hProv, 0, 0, PROV_RSA_FULL, CRYPT_VERIFYCONTEXT) = 0 Then
        sApiName = "CryptAcquireContext"
        GoTo QH
    End If
    uBlob.bType = PLAINTEXTKEYBLOB
    uBlob.bVersion = CUR_BLOB_VERSION
    uBlob.aiKeyAlg = CALG_DES
    uBlob.cbKeySize = UBound(uBlob.Buffer) + 1
    For lIdx = 0 To Clamp(UBound(baKey), , UBound(uBlob.Buffer))
        uBlob.Buffer(lIdx) = pvCryptoReverseBits(baKey(lIdx)) '--- VNC impl idiosyncrasy
    Next
    If CryptImportKey(hProv, uBlob, sizeof_BLOBHEADER + uBlob.cbKeySize, 0, CRYPT_EXPORTABLE, hKey) = 0 Then
        sApiName = "CryptImportKey"
        GoTo QH
    End If
    If CryptSetKeyParam(hKey, KP_MODE, CRYPT_MODE_ECB, 0) = 0 Then
        sApiName = "CryptSetKeyParam(KP_MODE)"
        GoTo QH
    End If
    lSize = UBound(baInput) + 1
    If CryptEncrypt(hKey, 0, 0, 0, baInput(0), lSize, lSize) = 0 Then
        sApiName = "CryptEncrypt"
        GoTo QH
    End If
    '--- success
    pvCryptoDesEncrypt = True
QH:
    If LenB(sApiName) <> 0 Then
        pvCheckHResult Err.LastDllError Or -(Err.LastDllError > 0) * LNG_FACILITY_WIN32, FUNC_NAME & "." & sApiName
    End If
    If hKey <> 0 Then
        Call CryptDestroyKey(hKey)
    End If
    If hProv <> 0 Then
        Call CryptReleaseContext(hProv, 0)
    End If
End Function

Private Function pvCryptoReverseBits(ByVal lByte As Long) As Byte
    Const POW2_4 As Long = 2 ^ 4
    Const POW2_2 As Long = 2 ^ 2
    Const POW2_1 As Long = 2 ^ 1
    
    lByte = ((lByte And &HF0) \ POW2_4) Or ((lByte And &HF) * POW2_4) And &HFF
    lByte = ((lByte And &HCC) \ POW2_2) Or ((lByte And &H33) * POW2_2) And &HFF
    lByte = ((lByte And &HAA) \ POW2_1) Or ((lByte And &H55) * POW2_1) And &HFF
    pvCryptoReverseBits = lByte
End Function

'= common helpers ========================================================

Private Function pvNetworkShort(ByVal lValue As Long) As Integer
    pvNetworkShort = ws_htons(lValue)
End Function

Private Function pvNetworkLong(ByVal lValue As Long) As Long
    pvNetworkLong = ws_htonl(lValue)
End Function

Private Function pvCalcAdler32(ByVal lAdler As Long, baData() As Byte, ByVal lPos As Long, ByVal lSize As Long) As Long
    Const LNG_BASE      As Long = 65521
    Dim lA              As Long
    Dim lB              As Long
    Dim lIdx            As Long

    lA = lAdler And &HFFFF&
    Call CopyMemory(lB, ByVal VarPtr(lAdler) + 2, 2)
    For lIdx = lPos To lPos + lSize - 1
        lA = (lA + baData(lIdx)) Mod LNG_BASE
        lB = (lB + lA) Mod LNG_BASE
    Next
    pvCalcAdler32 = lA
    Call CopyMemory(ByVal VarPtr(pvCalcAdler32) + 2, lB, 2)
End Function

Private Function pvGetFileDateTime(sFile As String, Optional cLastWriteTime As Currency) As Date
    Dim hFind           As Long
    Dim uData           As WIN32_FIND_DATA
    Dim aTime(0 To 7)   As Integer
    
    hFind = pvFindFirstFile(sFile, uData)
    If hFind <> INVALID_HANDLE_VALUE Then
        Call FindClose(hFind)
        Call FileTimeToSystemTime(uData.ftLastWriteTime, aTime(0))
        Call SystemTimeToVariantTime(aTime(0), pvGetFileDateTime)
        cLastWriteTime = uData.ftLastWriteTime
    End If
End Function

Private Function pvToUnixTime(ByVal cFileTime As Currency) As Currency
    pvToUnixTime = cFileTime / 10000 - 1164447360#
    If pvToUnixTime < 0 Then
        pvToUnixTime = 0
    End If
End Function

Private Function pvGetFolderSize(sFolder As String) As Currency
    Dim lMode           As Long
    Dim hFind           As Long
    Dim uData           As WIN32_FIND_DATA
    Dim sFile           As String
    Dim cTemp           As Currency
    
    hFind = pvFindFirstFile(PathCombine(sFolder, "*.*"), uData)
    If hFind <> INVALID_HANDLE_VALUE Then
        Do
            sFile = Left$(uData.cFileName, InStr(uData.cFileName, Chr$(0)) - 1)
            If sFile <> "." And sFile <> ".." Then
                If (uData.dwFileAttributes And vbDirectory) <> 0 Then
                    pvGetFolderSize = pvGetFolderSize + pvGetFolderSize(PathCombine(sFolder, sFile))
                Else
                    lMode = uData.nFileSizeLow
                    uData.nFileSizeLow = uData.nFileSizeHigh
                    uData.nFileSizeHigh = lMode
                    Call CopyMemory(cTemp, uData.nFileSizeHigh, 8)
                    pvGetFolderSize = pvGetFolderSize + cTemp
                End If
            End If
        Loop While FindNextFile(hFind, VarPtr(uData)) <> 0
    End If
QH:
    If hFind <> INVALID_HANDLE_VALUE And hFind <> 0 Then
        Call FindClose(hFind)
    End If
End Function

#If Not ImplUseShared Then
Private Function Clamp( _
            ByVal lValue As Long, _
            Optional ByVal lMin As Long = -2147483647, _
            Optional ByVal lMax As Long = 2147483647) As Long
    Select Case lValue
    Case lMin To lMax
        Clamp = lValue
    Case Is < lMin
        Clamp = lMin
    Case Is > lMax
        Clamp = lMax
    End Select
End Function

Private Function GetSpecialFolder(ByVal lType As Long) As String
    GetSpecialFolder = String$(1000, 0)
    If SHGetFolderPath(0, lType Or CSIDL_FLAG_CREATE, 0, 0, StrPtr(GetSpecialFolder)) < 0 Then
        Call SHGetFolderPath(0, lType, 0, 0, GetSpecialFolder)
    End If
    GetSpecialFolder = Left$(GetSpecialFolder, InStr(GetSpecialFolder, vbNullChar) - 1)
End Function

Private Function PathCombine(sPath As String, sFile As String) As String
    PathCombine = sPath & IIf(LenB(sPath) <> 0 And Right$(sPath, 1) <> "\" And LenB(sFile) <> 0, "\", vbNullString) & sFile
End Function

Public Property Get At(vData As Variant, ByVal lIdx As Long, Optional sDefault As String) As String
    On Error GoTo QH
    At = sDefault
    If IsArray(vData) Then
        If LBound(vData) <= lIdx And lIdx <= UBound(vData) Then
            At = vData(lIdx)
        End If
    End If
QH:
End Property
#End If

'=========================================================================
' Event impl
'=========================================================================

Private Sub m_oSocket_OnAccept()
    Const FUNC_NAME     As String = "m_oSocket_OnAccept"
    Dim oSocket         As cAsyncSocket
    Dim vDevice         As Variant
    Dim oConn           As cVncServer
    
    On Error GoTo EH
    If Not m_oSocket.Accept(oSocket) Then
        GoTo QH
    End If
    If m_cConnections.Count = 0 Then
        vDevice = pvCaptureEnumDevices().Item(1)
        If Not pvCaptureInit(m_uCaptureCtx, vDevice(0), 1) Then
            GoTo QH
        End If
    End If
    Set oConn = New cVncServer
    m_cConnections.Add oConn, "#" & ObjPtr(oConn)
    If Not oConn.frNewConnection(oSocket, Me) Then
        frCloseConnection oConn
        GoTo QH
    End If
    Set m_pTimer = InitFireOnceTimerThunk(Me, pvAddressOfTimerProc.TimerProc)
QH:
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub m_oSocket_OnReceive()
    pvHandleReceive
End Sub

Private Sub m_oSocket_OnSend()
    Const FUNC_NAME     As String = "m_oSocket_OnSend"
    Dim eChannel        As UcsSendChannelEnum
    
    On Error GoTo EH
    eChannel = pvFirstPendingChannel
    If eChannel >= 0 Then
        If Not pvBufferSendAsync(m_uSendBuffer(eChannel)) Then
            GoTo QH
        End If
    End If
QH:
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub m_oSocket_OnClose()
    Const FUNC_NAME     As String = "m_oSocket_OnClose"
    
    On Error GoTo EH
    pvParent.frCloseConnection Me
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub m_oSocket_OnError(ByVal ErrorCode As Long, ByVal EventMask As UcsAsyncSocketEventMaskEnum)
    Const FUNC_NAME     As String = "m_oSocket_OnError"
    Const WSAECONNABORTED As Long = 10053
    
    On Error GoTo EH
    #If ImplUseDebugLog Then
        If ErrorCode <> WSAECONNABORTED Then
            DebugLog MODULE_NAME, FUNC_NAME, m_oSocket.GetErrorDescription(ErrorCode) & " (" & ErrorCode & ")"
        End If
    #End If
    pvParent.frCloseConnection Me
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

'=========================================================================
' The Modern Subclassing Thunk (MST)
'=========================================================================

#If Not ImplUseShared Then
Public Function InitAddressOfMethod(pObj As Object, ByVal MethodParamCount As Long) As Object
    Dim STR_THUNK       As String: STR_THUNK = "6AAAAABag+oFV4v6ge9QEMEAgcekEcEAuP9EJAS5+QcAAPOri8LB4AgFuQAAAKuLwsHoGAUAjYEAq7gIAAArq7hEJASLq7hJCIsEq7iBi1Qkq4tEJAzB4AIFCIkCM6uLRCQMweASBcDCCACriTrHQgQBAAAAi0QkCIsAiUIIi0QkEIlCDIHqUBDBAIvCBTwRwQCri8IFUBHBAKuLwgVgEcEAq4vCBYQRwQCri8IFjBHBAKuLwgWUEcEAq4vCBZwRwQCri8IFpBHBALn5BwAAq4PABOL6i8dfgcJQEMEAi0wkEIkRK8LCEAAPHwCLVCQE/0IEi0QkDIkQM8DCDABmkItUJAT/QgSLQgTCBAAPHwCLVCQE/0oEi0IEg/gAfgPCBABZWotCDGgAgAAAagBSUf/gZpC4AUAAgMIIALgBQACAwhAAuAFAAIDCGAC4AUAAgMIkAA==" ' 25.3.2019 14:01:08
    Const THUNK_SIZE    As Long = 16728
    Dim hThunk          As Long
    Dim lSize           As Long
    
    hThunk = pvThunkAllocate(STR_THUNK, THUNK_SIZE)
    If hThunk = 0 Then
        Exit Function
    End If
    lSize = CallWindowProc(hThunk, ObjPtr(pObj), MethodParamCount, GetProcAddress(GetModuleHandle("kernel32"), "VirtualFree"), VarPtr(InitAddressOfMethod))
    Debug.Assert lSize = THUNK_SIZE
End Function

Public Function InitFireOnceTimerThunk(pObj As Object, ByVal pfnCallback As Long, Optional Delay As Long) As IUnknown
    Dim STR_THUNK       As String: STR_THUNK = "6AAAAABag+oFgeogEQUAV1aLdCQUg8YIgz4AdCqL+oHHDBMFAIvCBSgSBQCri8IFZBIFAKuLwgV0EgUAqzPAq7kIAAAA86WBwgwTBQBSahj/UhBai/iLwqu4AQAAAKszwKuri3QkFKWlg+8Yi0IMSCX/AAAAUItKDDsMJHULWIsPV/9RFDP/62P/QgyBYgz/AAAAjQTKjQTIjUyIMIB5EwB101jHAf80JLiJeQTHQQiJRCQEi8ItDBMFAAWgEgUAUMHgCAW4AAAAiUEMWMHoGAUA/+CQiUEQiU8MUf90JBRqAGoAiw//URiJRwiLRCQYiTheX7g8EwUALSARBQAFABQAAMIQAGaQi0QkCIM4AHUqg3gEAHUkgXgIwAAAAHUbgXgMAAAARnUSi1QkBP9CBItEJAyJEDPAwgwAuAJAAIDCDACQi1QkBP9CBItCBMIEAA8fAItUJAT/SgSLQgR1HYtCDMZAEwCLCv9yCGoA/1Eci1QkBIsKUv9RFDPAwgQAi1QkBItCEIXAdFuLCotBKIXAdCdS/9Bag/gBd0mLClL/USxahcB1PosKUmrw/3Eg/1EkWqkAAAAIdSuLClL/cghqAP9RHFr/QgQzwFBU/3IQ/1IUi1QkCMdCCAAAAABS6G////9YwhQADx8AjURAAQ==" ' 13.5.2020 18:59:12
    Const THUNK_SIZE    As Long = 5660
    Static hThunk       As Long
    Dim aParams(0 To 9) As Long
    Dim lSize           As Long
    
    aParams(0) = ObjPtr(pObj)
    aParams(1) = pfnCallback
    #If ImplSelfContained Then
        If hThunk = 0 Then
            hThunk = pvThunkGlobalData("InitFireOnceTimerThunk")
        End If
    #End If
    If hThunk = 0 Then
        hThunk = pvThunkAllocate(STR_THUNK, THUNK_SIZE)
        If hThunk = 0 Then
            Exit Function
        End If
        aParams(2) = GetProcAddress(GetModuleHandle("ole32"), "CoTaskMemAlloc")
        aParams(3) = GetProcAddress(GetModuleHandle("ole32"), "CoTaskMemFree")
        aParams(4) = GetProcAddress(GetModuleHandle("user32"), "SetTimer")
        aParams(5) = GetProcAddress(GetModuleHandle("user32"), "KillTimer")
        '--- for IDE protection
        Debug.Assert pvThunkIdeOwner(aParams(6))
        If aParams(6) <> 0 Then
            aParams(7) = GetProcAddress(GetModuleHandle("user32"), "GetWindowLongA")
            aParams(8) = GetProcAddress(GetModuleHandle("vba6"), "EbMode")
            aParams(9) = GetProcAddress(GetModuleHandle("vba6"), "EbIsResetting")
        End If
        #If ImplSelfContained Then
            pvThunkGlobalData("InitFireOnceTimerThunk") = hThunk
        #End If
    End If
    lSize = CallWindowProc(hThunk, 0, Delay, VarPtr(aParams(0)), VarPtr(InitFireOnceTimerThunk))
    Debug.Assert lSize = THUNK_SIZE
End Function

Private Function pvThunkIdeOwner(hIdeOwner As Long) As Boolean
    #If Not ImplNoIdeProtection Then
        Dim lProcessId      As Long
        
        Do
            hIdeOwner = FindWindowEx(0, hIdeOwner, "IDEOwner", vbNullString)
            Call GetWindowThreadProcessId(hIdeOwner, lProcessId)
        Loop While hIdeOwner <> 0 And lProcessId <> GetCurrentProcessId()
    #End If
    pvThunkIdeOwner = True
End Function

Private Function pvThunkAllocate(sText As String, Optional ByVal Size As Long) As Long
    Static Map(0 To &H3FF) As Long
    Dim baInput()       As Byte
    Dim lIdx            As Long
    Dim lChar           As Long
    Dim lPtr            As Long
    
    pvThunkAllocate = VirtualAlloc(0, IIf(Size > 0, Size, (Len(sText) \ 4) * 3), MEM_COMMIT, PAGE_EXECUTE_READWRITE)
    If pvThunkAllocate = 0 Then
        Exit Function
    End If
    '--- init decoding maps
    If Map(65) = 0 Then
        baInput = StrConv("ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/", vbFromUnicode)
        For lIdx = 0 To UBound(baInput)
            lChar = baInput(lIdx)
            Map(&H0 + lChar) = lIdx * (2 ^ 2)
            Map(&H100 + lChar) = (lIdx And &H30) \ (2 ^ 4) Or (lIdx And &HF) * (2 ^ 12)
            Map(&H200 + lChar) = (lIdx And &H3) * (2 ^ 22) Or (lIdx And &H3C) * (2 ^ 6)
            Map(&H300 + lChar) = lIdx * (2 ^ 16)
        Next
    End If
    '--- base64 decode loop
    baInput = StrConv(Replace(Replace(sText, vbCr, vbNullString), vbLf, vbNullString), vbFromUnicode)
    lPtr = pvThunkAllocate
    For lIdx = 0 To UBound(baInput) - 3 Step 4
        lChar = Map(baInput(lIdx + 0)) Or Map(&H100 + baInput(lIdx + 1)) Or Map(&H200 + baInput(lIdx + 2)) Or Map(&H300 + baInput(lIdx + 3))
        Call CopyMemory(ByVal lPtr, lChar, 3)
        lPtr = (lPtr Xor SIGN_BIT) + 3 Xor SIGN_BIT
    Next
End Function

Private Property Get pvThunkGlobalData(sKey As String) As Long
    Dim sBuffer     As String
    
    sBuffer = String$(50, 0)
    Call GetEnvironmentVariable("_MST_GLOBAL" & GetCurrentProcessId() & "_" & sKey, sBuffer, Len(sBuffer) - 1)
    pvThunkGlobalData = Val(Left$(sBuffer, InStr(sBuffer, vbNullChar) - 1))
End Property

Private Property Let pvThunkGlobalData(sKey As String, ByVal lValue As Long)
    Call SetEnvironmentVariable("_MST_GLOBAL" & GetCurrentProcessId() & "_" & sKey, lValue)
End Property
#End If

'=========================================================================
' Base class events
'=========================================================================

#If ImplZlib Then
Private Sub Class_Terminate()
    If Not m_oZipArchive Is Nothing Then
        m_oZipArchive.DeflateEnd m_hDeflateCtx
    End If
End Sub
#End If
