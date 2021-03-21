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

#Const ImplZlib = VNC_NOZLIB = 0

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
'--- for thunks
Private Const MEM_COMMIT                    As Long = &H1000
Private Const PAGE_EXECUTE_READWRITE        As Long = &H40
Private Const SIGN_BIT                      As Long = &H80000000

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function ArrPtr Lib "msvbvm60" Alias "VarPtr" (Ptr() As Any) As Long
Private Declare Function IsBadReadPtr Lib "kernel32" (ByVal lp As Long, ByVal ucb As Long) As Long
Private Declare Function ws_htonl Lib "ws2_32" Alias "htonl" (ByVal hostlong As Long) As Long
Private Declare Function ws_htons Lib "ws2_32" Alias "htons" (ByVal hostshort As Long) As Integer
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Private Declare Function OpenInputDesktop Lib "user32" (ByVal dwFlags As Long, ByVal fInherit As Boolean, ByVal dwDesiredAccess As Long) As Long
Private Declare Function CloseDesktop Lib "user32" (ByVal hDesktop As Long) As Long
Private Declare Function SetThreadDesktop Lib "user32" (ByVal hDesktop As Long) As Long
Private Declare Function SetProcessDpiAwarenessContext Lib "user32" (ByVal lValue As Long) As Long
Private Declare Function D3DKMTSetProcessSchedulingPriorityClass Lib "gdi32" (ByVal hProcess As Long, ByVal lPriority As Long) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Sub CoTaskMemFree Lib "ole32" (ByVal pv As Long)
Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function RedrawWindow Lib "user32" (ByVal hWnd As Long, lprcUpdate As Any, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
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

'=========================================================================
' Protocol enums and types
'=========================================================================

Private Const STR_CLIENT_MSG_CAPABILITIES As String = "130;TGHT;FTC_LSRQ|131;TGHT;FTC_DNRQ|132;TGHT;FTC_UPRQ|133;TGHT;FTC_UPDT|134;TGHT;FTC_DNCN|135;TGHT;FTC_UPFL|136;TGHT;FTC_FCDR"
Private Const STR_SERVER_MSG_CAPABILITIES As String = "130;TGHT;FTS_LSDT|131;TGHT;FTS_DNDT|132;TGHT;FTS_UPCN|133;TGHT;FTS_DNFL"

Private Enum UcsRfbProtocolStateEnum
    ucsStaNone
    ucsStaExpectProtocolVersion
    ucsStaExpectSecurityHandshake
    ucsStaExpectClientInit
    ucsStaExpectClientMessage
End Enum

Private Enum UcsRfbAutheticationEnum
    ucsAuthInvalid = 0
    ucsAuthNone = 1
    ucsAuthVnc = 2
    ucsAuthTight = 16
End Enum

Private Enum UcsRfbMessageEnum
    '--- client-side
    ucsMsgSetPixelFormat = 0
    ucsMsgSetEncodings = 2
    ucsMsgFramebufferUpdateRequest = 3
    ucsMsgKeyEvent = 4
    ucsMsgPointerEvent = 5
    ucsMsgClientCutText = 6
    ucsMsgFileListRequest = 130
    ucsMsgFileDownloadRequest = 131
    ucsMsgFileUploadRequest = 132
    ucsMsgFileUploadData = 133
    ucsMsgFileDownloadCancel = 134
    ucsMsgFileUploadFailed = 135
    ucsMsgFileCreateDirectoryRequest = 136
    '--- server-side
    ucsMsgFramebufferUpdate = 0
    ucsMsgSetColorMapEntries = 1
    ucsMsgBell = 2
    ucsMsgServerCutText = 3
    ucsMsgFileListData = 130
    ucsMsgFileDownloadData = 131
    ucsMsgFileUploadCancel = 132
    ucsMsgFileDownloadFailed = 133
End Enum

Private Enum UcsRdfEncodingEnum
    ucsEncRaw = 0
    ucsEncCopyRect = 1
    ucsEncRRE = 2
    ucsEncHextile = 5
    ucsEncZlib = 6
    ucsEncTight = 7
    ucsEncTRLE = 15
    ucsEncZRLE = 16
    ucsEncCursorWithAlphaPseudoEncoding = -314
    ucsEncCursorPseudoEncoding = -239
    ucsEncDesktopSizePseudoEncoding = -223
End Enum

Private Type UcsPixelFormatType
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
Private Const UcsPixelFormatSize As Long = 16

Private Type UcsServerInitType
    FramebufferWidth    As Integer
    FramebufferHeight   As Integer
    ServerPixelFormat   As UcsPixelFormatType
    NameLength          As Long
'    NameString()        As Byte
End Type
Private Const UcsServerInitSize As Long = 4 + UcsPixelFormatSize + 4

Private Type UcsClientSetPixelFormatType
    MessageType         As Long             ' ucsMsgSetPixelFormat = 0
    PixelFormat         As UcsPixelFormatType
End Type
Private Const UcsClientSetPixelFormatSize As Long = 4 + UcsPixelFormatSize

Private Type UcsClientSetEncodingsType
    MessageType         As Integer          ' ucsMsgSetEncodings = 2
    NumberOfEncodings   As Integer
    EncodingType()      As Long
    HasZlib             As Boolean
    HasTight            As Boolean
    HasCursorWithAlphaPseudoEncoding As Boolean
    HasCursorPseudoEncoding As Boolean
End Type
Private Const UcsClientSetEncodingsSize As Long = 4

Private Type UcsClientFramebufferUpdateRequestType
    MessageType         As Byte             ' ucsMsgFramebufferUpdateRequest = 3
    Incremental         As Byte
    XPosition           As Integer
    YPosition           As Integer
    Width               As Integer
    Height              As Integer
End Type
Private Const UcsClientFramebufferUpdateRequestSize As Long = 10

Private Type UcsClientKeyEventType
    MessageType         As Byte             ' ucsMsgKeyEvent = 4
    DownFlag            As Byte
    Padding(0 To 1)     As Byte
    Key                 As Long
End Type
Private Const UcsClientKeyEventSize As Long = 8

Private Type UcsClientPointerEventType
    MessageType         As Byte             ' ucsMsgPointerEvent = 5
    ButtonMask          As Byte
    XPosition           As Integer
    YPosition           As Integer
End Type
Private Const UcsClientPointerEventSize As Long = 6

Private Type UcsClientCutTextType
    MessageType         As Long             ' ucsMsgClientCutText = 6
    Length              As Long
    Text()              As Byte
End Type
Private Const UcsClientCutTextSize As Long = 8

Private Type UcsServerRectangleType
    XPosition           As Integer
    YPosition           As Integer
    Width               As Integer
    Height              As Integer
    EncodingType        As Long
'    Data()              As Byte
End Type
Private Const UcsServerRectangleSize As Long = 12

Private Type UcsServerFramebufferUpdateType
    MessageType         As Integer          ' ucsMsgFramebufferUpdate = 0
    NumberOfRectangles  As Integer
'    Rectangle()         As UcsServerRectangleType
End Type
Private Const UcsServerFramebufferUpdateSize As Long = 4

Private Type UcsServerRgbValueType
    Red                 As Integer
    Green               As Integer
    Blue                As Integer
End Type
Private Const UcsServerRgbValueSize As Long = 6

Private Type UcsServerSetColorMapEntriesType
    MessageType         As Integer          ' ucsMsgSetColorMapEntries = 1
    FirstColor          As Integer
    NumberOfColors      As Integer
'    Color()             As UcsServerRgbValueType
End Type
Private Const UcsServerSetColorMapEntriesSize As Long = 6

Private Type UcsServerCutTextType
    MessageType         As Long             ' ucsMsgServerCutText = 3
    Length              As Long
'    Text()              As Byte
End Type
Private Const UcsServerCutTextSize As Long = 8

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

'=========================================================================
' Constants and member variables
'=========================================================================

Private Const STR_APP_NAME          As String = "VbVncServer"
Private Const STR_PROTOCOL_VERSION  As String = "RFB 003.008" & vbLf

Private m_eState                As UcsRfbProtocolStateEnum
Private WithEvents m_oSocket    As cAsyncSocket
Attribute m_oSocket.VB_VarHelpID = -1
Private m_uRecvBuffer           As UcsBuffer
Private m_uSendBuffer           As UcsBuffer
Private m_cClients              As Object ' Collection
Private m_uClientPixelFormat    As UcsClientSetPixelFormatType
Private m_uClientEncodings      As UcsClientSetEncodingsType
Private m_uClientFramebufferUpdate As UcsClientFramebufferUpdateRequestType
Private m_uClientKey            As UcsClientKeyEventType
Private m_uClientPointer        As UcsClientPointerEventType
Private m_uClientCutText        As UcsClientCutTextType
Private m_uCaptureCtx           As UcsCaptureContext
Private m_uCaptureFrame         As UcsCaptureFrame
Private m_pTimer                As IUnknown
#If ImplZlib Then
    Private m_oZipArchive       As cZipArchive
    Private m_hDeflateCtx       As Long
#End If

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

Private Property Get pvAddressOfTimerProc() As cVncServer
    Set pvAddressOfTimerProc = InitAddressOfMethod(Me, 0)
End Property

'=========================================================================
' Methods
'=========================================================================

Public Function Init(Optional sAddress As String, Optional ByVal lPort As Long = 5900) As Boolean
    Set m_oSocket = New cAsyncSocket
    If Not m_oSocket.Create(SocketAddress:=sAddress, SocketPort:=lPort) Then
        GoTo QH
    End If
    If Not m_oSocket.Listen() Then
        GoTo QH
    End If
    '--- success
    Init = True
QH:
End Function

Public Function Accept(oSocket As cAsyncSocket) As Boolean
    Set m_oSocket = oSocket
    If Not pvSendProtocolVersion(m_uSendBuffer, m_oSocket) Then
        GoTo QH
    End If
    m_eState = ucsStaExpectProtocolVersion
    '--- success
    Accept = True
QH:
End Function

Private Function pvSendProtocolVersion(uOutput As UcsBuffer, oSocket As cAsyncSocket) As Boolean
    pvBufferWriteArray uOutput, oSocket.ToTextArray(STR_PROTOCOL_VERSION)
    If Not pvBufferSendAsync(uOutput, oSocket) Then
        GoTo QH
    End If
    '--- succees
    pvSendProtocolVersion = True
QH:
End Function

Private Function pvSendSecurityType(uOutput As UcsBuffer, oSocket As cAsyncSocket) As Boolean
    Const SecurityResultOk As Long = 0
    
    pvBufferWriteLong uOutput, 1
    pvBufferWriteLong uOutput, ucsAuthNone
    pvBufferWriteLong uOutput, SecurityResultOk, Size:=4
    If Not pvBufferSendAsync(uOutput, oSocket) Then
        GoTo QH
    End If
    '--- succees
    pvSendSecurityType = True
QH:
End Function

Private Function pvSendServerInit(uCtx As UcsCaptureContext, uOutput As UcsBuffer, oSocket As cAsyncSocket) As Boolean
    Dim uInit           As UcsServerInitType
    
    With uInit
        .FramebufferWidth = pvNetworkShort(uCtx.Width)
        .FramebufferHeight = pvNetworkShort(uCtx.Height)
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
        .NameLength = pvNetworkLong(Len(STR_APP_NAME))
    End With
    pvBufferWriteBlob uOutput, VarPtr(uInit), UcsServerInitSize
    pvBufferWriteArray uOutput, StrConv(STR_APP_NAME, vbFromUnicode)
    If Not pvBufferSendAsync(uOutput, oSocket) Then
        GoTo QH
    End If
    '--- save as default pixel format
    m_uClientPixelFormat.PixelFormat = uInit.ServerPixelFormat
    '--- succees
    pvSendServerInit = True
QH:
End Function

Private Function pvSendEmptyFramebufferUpdate(uOutput As UcsBuffer, oSocket As cAsyncSocket) As Boolean
    Dim uUpdate         As UcsServerFramebufferUpdateType
    Dim uRectangle      As UcsServerRectangleType
    
    With uUpdate
        .MessageType = ucsMsgFramebufferUpdate
        .NumberOfRectangles = pvNetworkShort(1)
    End With
    pvBufferWriteBlob uOutput, VarPtr(uUpdate), UcsServerFramebufferUpdateSize
    pvBufferWriteBlob uOutput, VarPtr(uRectangle), UcsServerRectangleSize
    If Not pvBufferSendAsync(uOutput, oSocket) Then
        GoTo QH
    End If
    '--- success
    pvSendEmptyFramebufferUpdate = True
QH:
End Function

Private Function pvParseClientMessage(uInput As UcsBuffer) As Boolean
    Const FUNC_NAME     As String = "pvParseClientMessage"
    Dim lIdx            As Long
    Dim uPointerEvent   As UcsClientPointerEventType
    
    On Error GoTo EH
    With uInput
        Select Case .Data(.Pos)
        Case ucsMsgSetPixelFormat
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
        Case ucsMsgSetEncodings
            If .Size - .Pos < UcsClientSetEncodingsSize Then
                GoTo QH
            End If
            Call CopyMemory(m_uClientEncodings, .Data(.Pos), UcsClientSetEncodingsSize)
            With m_uClientEncodings
                .NumberOfEncodings = pvNetworkShort(.NumberOfEncodings)
            End With
            If m_uClientEncodings.NumberOfEncodings > 0 Then
                If .Size - .Pos < UcsClientSetEncodingsSize + m_uClientEncodings.NumberOfEncodings * 4 Then
                    GoTo QH
                End If
                ReDim m_uClientEncodings.EncodingType(0 To m_uClientEncodings.NumberOfEncodings - 1) As Long
                Call CopyMemory(m_uClientEncodings.EncodingType(0), .Data(.Pos + UcsClientSetEncodingsSize), m_uClientEncodings.NumberOfEncodings * 4)
                With m_uClientEncodings
                    For lIdx = 0 To .NumberOfEncodings - 1
                        .EncodingType(lIdx) = pvNetworkLong(.EncodingType(lIdx))
                        Select Case .EncodingType(lIdx)
                        Case ucsEncZlib
                            .HasZlib = True
                        Case ucsEncTight
                            .HasTight = True
                        Case ucsEncCursorWithAlphaPseudoEncoding
                            .HasCursorWithAlphaPseudoEncoding = True
                        Case ucsEncCursorPseudoEncoding
                            .HasCursorPseudoEncoding = True
                        End Select
                    Next
                End With
            End If
            .Pos = .Pos + UcsClientSetEncodingsSize + m_uClientEncodings.NumberOfEncodings * 4
            Set m_pTimer = Nothing
            If Not pvSendEmptyFramebufferUpdate(m_uSendBuffer, m_oSocket) Then
                GoTo QH
            End If
        Case ucsMsgFramebufferUpdateRequest
            If .Size - .Pos < UcsClientFramebufferUpdateRequestSize Then
                GoTo QH
            End If
            Call CopyMemory(m_uClientFramebufferUpdate, .Data(.Pos), UcsClientFramebufferUpdateRequestSize)
            With m_uClientFramebufferUpdate
                .XPosition = pvNetworkShort(.XPosition)
                .YPosition = pvNetworkShort(.YPosition)
                .Width = pvNetworkShort(.Width)
                .Height = pvNetworkShort(.Height)
            End With
            .Pos = .Pos + UcsClientFramebufferUpdateRequestSize
            If m_uClientFramebufferUpdate.Incremental = 0 Then
                Call RedrawWindow(0, ByVal 0, 0, RDW_INVALIDATE Or RDW_ALLCHILDREN Or RDW_ERASENOW)
            End If
            Set m_pTimer = InitFireOnceTimerThunk(Me, pvAddressOfTimerProc.TimerProc)
        Case ucsMsgKeyEvent
            If .Size - .Pos < UcsClientKeyEventSize Then
                GoTo QH
            End If
            Call CopyMemory(m_uClientKey, .Data(.Pos), UcsClientKeyEventSize)
            m_uClientKey.Key = pvNetworkLong(m_uClientKey.Key)
            .Pos = .Pos + UcsClientKeyEventSize
        Case ucsMsgPointerEvent
            If .Size - .Pos < UcsClientPointerEventSize Then
                GoTo QH
            End If
            Call CopyMemory(uPointerEvent, .Data(.Pos), UcsClientPointerEventSize)
            With uPointerEvent
                .XPosition = pvNetworkShort(.XPosition)
                .YPosition = pvNetworkShort(.YPosition)
            End With
            .Pos = .Pos + UcsClientPointerEventSize
            If Not pvHandlePointerEvent(m_uCaptureCtx, uPointerEvent) Then
                GoTo QH
            End If
        Case ucsMsgClientCutText
            If .Size - .Pos < UcsClientCutTextSize Then
                GoTo QH
            End If
            Call CopyMemory(m_uClientCutText, .Data(.Pos), UcsClientCutTextSize)
            With m_uClientCutText
                .Length = pvNetworkLong(.Length)
            End With
            If m_uClientCutText.Length > 0 Then
                If .Size - .Pos < UcsClientCutTextSize + m_uClientCutText.Length Then
                    GoTo QH
                End If
                pvArrayAllocate m_uClientCutText.Text, m_uClientCutText.Length, FUNC_NAME
                Call CopyMemory(m_uClientCutText.Text(0), .Data(.Pos + UcsClientCutTextSize), m_uClientCutText.Length)
            End If
            .Pos = .Pos + UcsClientCutTextSize + m_uClientCutText.Length
        Case Else
            Debug.Print "Unknown client message (" & .Data(.Pos) & ")", Timer
            '--- flush
            .Pos = .Size
        End Select
    End With
    '--- succes
    pvParseClientMessage = True
QH:
    Exit Function
EH:
    PrintError FUNC_NAME
End Function

Private Function pvHandlePointerEvent(uCtx As UcsCaptureContext, uEvent As UcsClientPointerEventType) As Boolean
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
        If (.ButtonMask And LNG_BUTTON1) <> (m_uClientPointer.ButtonMask And LNG_BUTTON1) Then
            lFlags = lFlags Or IIf((.ButtonMask And LNG_BUTTON1) <> 0, MOUSEEVENTF_LEFTDOWN, MOUSEEVENTF_LEFTUP)
        End If
        If (.ButtonMask And LNG_BUTTON2) <> (m_uClientPointer.ButtonMask And LNG_BUTTON2) Then
            lFlags = lFlags Or IIf((.ButtonMask And LNG_BUTTON2) <> 0, MOUSEEVENTF_MIDDLEDOWN, MOUSEEVENTF_MIDDLEUP)
        End If
        If (.ButtonMask And LNG_BUTTON3) <> (m_uClientPointer.ButtonMask And LNG_BUTTON3) Then
            lFlags = lFlags Or IIf((.ButtonMask And LNG_BUTTON3) <> 0, MOUSEEVENTF_RIGHTDOWN, MOUSEEVENTF_RIGHTUP)
        End If
        If (.ButtonMask And LNG_BUTTON4) <> (m_uClientPointer.ButtonMask And LNG_BUTTON4) Then
            lFlags = lFlags Or MOUSEEVENTF_WHEEL
            lWheel = 120
        End If
        If (.ButtonMask And LNG_BUTTON5) <> (m_uClientPointer.ButtonMask And LNG_BUTTON5) Then
            lFlags = lFlags Or MOUSEEVENTF_WHEEL
            lWheel = -120
        End If
        Call mouse_event(lFlags, .XPosition * 65535 \ (uCtx.Width - 1), .YPosition * 65535 \ (uCtx.Height - 1), lWheel, 0)
    End With
    m_uClientPointer = uEvent
    '--- success
    pvHandlePointerEvent = True
End Function

'= desktop duplication ===================================================

Public Function TimerProc() As Long
Attribute TimerProc.VB_MemberFlags = "40"
    If m_uSendBuffer.Size = 0 And m_uCaptureFrame.NumTiles = 0 Then
        If pvCaptureFrame(m_uCaptureCtx, m_uCaptureFrame) Then
            If pvCaptureSend(m_uCaptureCtx, m_uCaptureFrame, m_uSendBuffer, m_oSocket) Then
                pvBufferSendAsync m_uSendBuffer, m_oSocket
            End If
        End If
    End If
    Set m_pTimer = InitFireOnceTimerThunk(Me, pvAddressOfTimerProc.TimerProc)
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
    Dim pOutput5        As IDXGIOutput5
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
        If TypeOf pOutput Is IDXGIOutput5 Then
            Set pOutput5 = pOutput
            Dim aFormats(0 To 3) As DXGI_FORMAT
            aFormats(0) = DXGI_FORMAT_B8G8R8A8_UNORM
            aFormats(1) = DXGI_FORMAT_R8G8B8A8_UNORM
            aFormats(2) = DXGI_FORMAT_R10G10B10A2_UNORM
            aFormats(3) = DXGI_FORMAT_R16G16B16A16_FLOAT
            hResult = pOutput5.DuplicateOutput1(pD3D11Device, 0, UBound(aFormats) + 1, aFormats(0), .Duplication)
            If hResult < 0 Then
                Err.Raise hResult, "IDXGIOutput5.DuplicateOutput1"
            End If
        Else
            hResult = pOutput.DuplicateOutput(pD3D11Device, .Duplication)
            If hResult < 0 Then
                Err.Raise hResult, "IDXGIOutput1.DuplicateOutput"
            End If
        End If
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
    Const SIZE_OUTDUPL_MOVE_RECT As Long = 24
    Const SIZE_RECT     As Long = 16
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
        '--- capture frame
        hResult = .Duplication.GetFrameMoveRects((UBound(uFrame.MoveRects) + 1) * SIZE_OUTDUPL_MOVE_RECT, uFrame.MoveRects(0), lSize)
        If hResult = DXGI_ERROR_MORE_DATA Then
            ReDim uFrame.MoveRects(0 To lSize \ SIZE_OUTDUPL_MOVE_RECT - 1) As DXGI_OUTDUPL_MOVE_RECT
            hResult = .Duplication.GetFrameMoveRects((UBound(uFrame.MoveRects) + 1) * SIZE_OUTDUPL_MOVE_RECT, uFrame.MoveRects(0), lSize)
        End If
        If hResult < 0 Then
            GoTo QH
        End If
        uFrame.NumMoveRects = lSize \ SIZE_OUTDUPL_MOVE_RECT
        hResult = .Duplication.GetFrameDirtyRects((UBound(uFrame.DirtyRects) + 1) * SIZE_RECT, uFrame.DirtyRects(0), lSize)
        If hResult = DXGI_ERROR_MORE_DATA Then
            ReDim uFrame.DirtyRects(0 To lSize \ SIZE_RECT - 1) As D3D11_RECT
            hResult = .Duplication.GetFrameDirtyRects((UBound(uFrame.DirtyRects) + 1) * SIZE_RECT, uFrame.DirtyRects(0), lSize)
        End If
        If hResult < 0 Then
            GoTo QH
        End If
        uFrame.NumDirtyRects = lSize \ SIZE_RECT
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
            uFrame.PointerShapeBufferSize = uFrameInfo.PointerShapeBufferSize
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
    '--- success
    pvCaptureFrame = True
QH:
    If uResource.pData <> 0 Then
        uCtx.Context.Unmap pTex, 0
    End If
    Exit Function
EH:
    PrintError FUNC_NAME
End Function

Private Function pvCaptureSend(uCtx As UcsCaptureContext, uFrame As UcsCaptureFrame, uOutput As UcsBuffer, oSocket As cAsyncSocket) As Boolean
    Const FUNC_NAME     As String = "pvCaptureSend"
    Dim lIdx            As Long
    Dim lX              As Long
    Dim lY              As Long
    Dim lPitch          As Long
    Dim lBytesPerPixel  As Long
    Dim lBufferSize     As Long
    Dim baRow()         As Byte
    Dim baRedLut()      As Byte
    Dim baGreenLut()    As Byte
    Dim baBlueLut()     As Byte
    Dim uUpdate         As UcsServerFramebufferUpdateType
    Dim uRectangle      As UcsServerRectangleType
    Dim lSrcIndex       As Long
    Dim lPtr            As Long
    Dim uTile           As UcsBuffer
    Dim baBit(0 To 7)   As Byte
    Dim baMask()        As Byte
    Dim bAndFlag        As Boolean
    Dim bXorFlag        As Boolean
    
    On Error GoTo EH
    If Not pvCapturePrepareTiles(uFrame) Then
        GoTo QH
    End If
    With uFrame
        lBytesPerPixel = m_uClientPixelFormat.PixelFormat.BitsPerPixel \ 8
        If lBytesPerPixel = 1 Then
            pvCapturePrepareLut baRedLut, m_uClientPixelFormat.PixelFormat.RedMax, 2 ^ m_uClientPixelFormat.PixelFormat.RedShift
            pvCapturePrepareLut baGreenLut, m_uClientPixelFormat.PixelFormat.GreenMax, 2 ^ m_uClientPixelFormat.PixelFormat.GreenShift
            pvCapturePrepareLut baBlueLut, m_uClientPixelFormat.PixelFormat.BlueMax, 2 ^ m_uClientPixelFormat.PixelFormat.BlueShift
        End If
        pvArrayAllocate baRow, uCtx.Width * 4, FUNC_NAME & ".baRow"
        '--- send pointer
        If .PointerShapeBufferSize > 0 And m_uClientEncodings.HasCursorPseudoEncoding Then
            Select Case .PointerInfo.Type
            Case DXGI_OUTDUPL_POINTER_SHAPE_TYPE_COLOR, DXGI_OUTDUPL_POINTER_SHAPE_TYPE_MONOCHROME
                If .PointerInfo.Type = DXGI_OUTDUPL_POINTER_SHAPE_TYPE_MONOCHROME Then
                    .PointerInfo.Height = .PointerInfo.Height \ 2
                End If
                lPitch = ((.PointerInfo.Width + 7) \ 8)
                pvArrayAllocate baMask, .PointerInfo.Height * lPitch, FUNC_NAME & ".baMask"
                For lIdx = 0 To 7
                    baBit(lIdx) = &H80 \ 2 ^ (lIdx Mod 8)
                Next
                With uUpdate
                    .MessageType = ucsMsgFramebufferUpdate
                    .NumberOfRectangles = pvNetworkShort(1)
                End With
                pvBufferWriteBlob uOutput, VarPtr(uUpdate), UcsServerFramebufferUpdateSize
                With uRectangle
                    .XPosition = pvNetworkShort(uFrame.PointerInfo.HotSpot.X)
                    .YPosition = pvNetworkShort(uFrame.PointerInfo.HotSpot.Y)
                    .Width = pvNetworkShort(uFrame.PointerInfo.Width)
                    .Height = pvNetworkShort(uFrame.PointerInfo.Height)
                    If uFrame.PointerInfo.Type = DXGI_OUTDUPL_POINTER_SHAPE_TYPE_COLOR And m_uClientEncodings.HasCursorWithAlphaPseudoEncoding Then
                        .EncodingType = pvNetworkLong(ucsEncCursorWithAlphaPseudoEncoding)
                    Else
                        .EncodingType = pvNetworkLong(ucsEncCursorPseudoEncoding)
                    End If
                End With
                pvBufferWriteBlob uOutput, VarPtr(uRectangle), UcsServerRectangleSize
                If .PointerInfo.Type = DXGI_OUTDUPL_POINTER_SHAPE_TYPE_COLOR And m_uClientEncodings.HasCursorWithAlphaPseudoEncoding Then
                    pvBufferWriteLong uOutput, pvNetworkLong(ucsEncRaw), Size:=4
                    pvBufferWriteBlob uOutput, VarPtr(.PointerShape(0)), .PointerInfo.Height * .PointerInfo.Width * 4
                ElseIf .PointerInfo.Type = DXGI_OUTDUPL_POINTER_SHAPE_TYPE_COLOR Then
                    Select Case lBytesPerPixel
                    Case 1, 4
                        For lY = 0 To .PointerInfo.Height - 1
                            Debug.Assert IsBadReadPtr(VarPtr(.PointerShape(lY * .PointerInfo.Pitch)), .PointerInfo.Width * 4) = 0
                            Call CopyMemory(baRow(0), .PointerShape(lY * .PointerInfo.Pitch), .PointerInfo.Width * 4)
                            For lX = 0 To .PointerInfo.Width - 1
                                If (baRow(lX * 4 + 3)) > &H80 Then
                                    lIdx = lY * lPitch + lX \ 8
                                    baMask(lIdx) = baMask(lIdx) Or baBit(lX Mod 8)
                                End If
                                If lBytesPerPixel = 1 Then
                                    baRow(lX) = baRedLut(baRow(lX * 4 + 2)) Or baGreenLut(baRow(lX * 4 + 1)) Or baBlueLut(baRow(lX * 4 + 0))
                                End If
                            Next
                            pvBufferWriteBlob uOutput, VarPtr(baRow(0)), .PointerInfo.Width * lBytesPerPixel
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
            .PointerShapeBufferSize = 0
        End If
        '--- send desktop
        lBufferSize = oSocket.SockOpt(ucsSsoSendBuffer)
        With uUpdate
            .MessageType = ucsMsgFramebufferUpdate
            .NumberOfRectangles = pvNetworkShort(uFrame.NumTiles)
        End With
        pvBufferWriteBlob uOutput, VarPtr(uUpdate), UcsServerFramebufferUpdateSize
        For lIdx = 0 To .NumTiles - 1
            lSrcIndex = .Tiles(lIdx).SrcIndex
            uRectangle.XPosition = pvNetworkShort(.DirtyRects(lSrcIndex).Left + .Tiles(lIdx).Left)
            uRectangle.YPosition = pvNetworkShort(.DirtyRects(lSrcIndex).Top + .Tiles(lIdx).Top)
            uRectangle.Width = pvNetworkShort(.Tiles(lIdx).Width)
            uRectangle.Height = pvNetworkShort(.Tiles(lIdx).Height)
            #If ImplZlib Then
                If m_uClientEncodings.HasTight Then
                    uRectangle.EncodingType = pvNetworkLong(ucsEncTight)
                    If lBytesPerPixel = 4 Then
                        lBytesPerPixel = 3
                    End If
                ElseIf m_uClientEncodings.HasZlib Then
                    uRectangle.EncodingType = pvNetworkLong(ucsEncZlib)
                Else
                    uRectangle.EncodingType = pvNetworkLong(ucsEncRaw)
                End If
            #Else
                uRectangle.EncodingType = pvNetworkLong(ucsEncRaw)
            #End If
            pvBufferWriteBlob uOutput, VarPtr(uRectangle), UcsServerRectangleSize
            lPitch = (.DirtyRects(lSrcIndex).Right - .DirtyRects(lSrcIndex).Left) * 4
            lPtr = VarPtr(.DirtyBuffers(lSrcIndex).Data(0))
            uTile.Size = 0
            Select Case lBytesPerPixel
            Case 1, 3, 4
                For lY = .Tiles(lIdx).Top To .Tiles(lIdx).Top + .Tiles(lIdx).Height - 1
                    If lBytesPerPixel = 1 Then
                        Debug.Assert IsBadReadPtr(lPtr + lY * lPitch + .Tiles(lIdx).Left * 4, .Tiles(lIdx).Width * 4) = 0
                        Call CopyMemory(baRow(0), ByVal lPtr + lY * lPitch + .Tiles(lIdx).Left * 4, .Tiles(lIdx).Width * 4)
                        For lX = 0 To .Tiles(lIdx).Width - 1
                            baRow(lX) = baRedLut(baRow(lX * 4 + 2)) Or baGreenLut(baRow(lX * 4 + 1)) Or baBlueLut(baRow(lX * 4 + 0))
                        Next
                        pvBufferWriteBlob uTile, VarPtr(baRow(0)), .Tiles(lIdx).Width
                    ElseIf lBytesPerPixel = 3 Then
                        Debug.Assert IsBadReadPtr(lPtr + lY * lPitch + .Tiles(lIdx).Left * 4, .Tiles(lIdx).Width * 4) = 0
                        Call CopyMemory(baRow(4), ByVal lPtr + lY * lPitch + .Tiles(lIdx).Left * 4, .Tiles(lIdx).Width * 4)
                        For lX = 0 To .Tiles(lIdx).Width - 1
                            baRow(lX * 3 + 0) = baRow(4 + lX * 4 + 2)
                            baRow(lX * 3 + 1) = baRow(4 + lX * 4 + 1)
                            baRow(lX * 3 + 2) = baRow(4 + lX * 4 + 0)
                        Next
                        pvBufferWriteBlob uTile, VarPtr(baRow(0)), .Tiles(lIdx).Width * 3
                    Else
                        pvBufferWriteBlob uTile, lPtr + lY * lPitch + .Tiles(lIdx).Left * 4, .Tiles(lIdx).Width * 4
                    End If
                Next
            Case Else
                pvBufferWriteBlob uTile, 0, .Tiles(lIdx).Height * .Tiles(lIdx).Width * lBytesPerPixel
            End Select
            #If ImplZlib Then
                If m_uClientEncodings.HasTight Then
                    '--- compression control: no reset, use stream 0, no filter
                    pvBufferWriteLong uOutput, 0
                    pvBufferWriteZlibBlob uOutput, VarPtr(uTile.Data(0)), uTile.Size, CompactSize:=True
                ElseIf m_uClientEncodings.HasZlib Then
                    pvBufferWriteZlibBlob uOutput, VarPtr(uTile.Data(0)), uTile.Size
                Else
                    pvBufferWriteBlob uOutput, VarPtr(uTile.Data(0)), uTile.Size
                End If
            #Else
                pvBufferWriteBlob uOutput, VarPtr(uTile.Data(0)), uTile.Size
            #End If
            If uOutput.Size >= uOutput.Pos + lBufferSize Then
                If Not oSocket.HasPendingEvent Then
                    If Not pvBufferSendAsync(uOutput, oSocket) Then
                        GoTo QH
                    End If
                End If
            End If
            If oSocket.AvailableBytes > 0 Then
                m_oSocket_OnReceive
                If lBytesPerPixel <> m_uClientPixelFormat.PixelFormat.BitsPerPixel \ 8 Then
                    lBytesPerPixel = m_uClientPixelFormat.PixelFormat.BitsPerPixel \ 8
                    If lBytesPerPixel = 1 Then
                        pvCapturePrepareLut baRedLut, m_uClientPixelFormat.PixelFormat.RedMax, 2 ^ m_uClientPixelFormat.PixelFormat.RedShift
                        pvCapturePrepareLut baGreenLut, m_uClientPixelFormat.PixelFormat.GreenMax, 2 ^ m_uClientPixelFormat.PixelFormat.GreenShift
                        pvCapturePrepareLut baBlueLut, m_uClientPixelFormat.PixelFormat.BlueMax, 2 ^ m_uClientPixelFormat.PixelFormat.BlueShift
                    End If
                End If
            End If
        Next
        .NumTiles = 0
    End With
    If uOutput.Size = 0 Then
        GoTo QH
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
    Const TILE_SIZE     As Long = 256
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
        For lY = 0 To lHeight - 1 Step TILE_SIZE
            For lX = 0 To lWidth - 1 Step TILE_SIZE
                If uFrame.NumTiles > UBound(uFrame.Tiles) Then
                    ReDim Preserve uFrame.Tiles(0 To UBound(uFrame.Tiles) * 2 + 1) As UcsCaptureTile
                End If
                With uFrame.Tiles(uFrame.NumTiles)
                    .SrcIndex = lIdx
                    .Left = lX
                    .Top = lY
                    .Width = Clamp(lX + TILE_SIZE, , lWidth) - lX
                    .Height = Clamp(lY + TILE_SIZE, , lHeight) - lY
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

Private Sub pvCapturePrepareLut(baLut() As Byte, ByVal lMax As Long, ByVal lShift As Long)
    Const FUNC_NAME     As String = "pvCapturePrepareLut"
    Dim lIdx            As Long
    
    On Error GoTo EH
    pvArrayAllocate baLut, 256, FUNC_NAME & ".baLut"
    For lIdx = 0 To 255
        baLut(lIdx) = Int(lIdx * lMax / 255 + 0.5) * lShift
    Next
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

'= buffer helpers ========================================================

Private Sub pvBufferWriteArray(uOutput As UcsBuffer, baSrc() As Byte)
    Dim lSize       As Long
    
    With uOutput
        lSize = pvArraySize(baSrc)
        If lSize > 0 Then
            .Size = pvArrayWriteBlob(.Data, .Size, VarPtr(baSrc(0)), lSize)
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
    If m_oZipArchive.DeflateBlob(m_hDeflateCtx, lPtr, lSize, lOutputPtr, lOutputSize, Level:=6) Then
        If CompactSize Then
            pvBufferWriteCompactLong uOutput, lOutputSize + IIf(bFirst, 2, 0)
        Else
            pvBufferWriteLong uOutput, lOutputSize + IIf(bFirst, 2, 0), Size:=4
        End If
        If bFirst Then
            pvBufferWriteLong uOutput, &H789C&, Size:=2
        End If
        pvBufferWriteBlob uOutput, lOutputPtr, lOutputSize
'        pvBufferWriteLong uOutput, pvNetworkLong(pvCalcAdler32(lPtr, lSize)), Size:=4
        Call CoTaskMemFree(lOutputPtr)
    End If
End Sub
#End If

Private Function pvArrayWriteBlob(baBuffer() As Byte, ByVal lPos As Long, ByVal lPtr As Long, ByVal lSize As Long) As Long
    Const FUNC_NAME     As String = "pvArrayWriteBlob"
    Dim lBufPtr         As Long
    
    '--- peek long at ArrPtr(baBuffer)
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

Private Function pvBufferSendAsync(uOutput As UcsBuffer, oSocket As cAsyncSocket) As Boolean
    Const FUNC_NAME     As String = "pvBufferSendAsync"
    Dim lBytes          As Long
    Dim lBufferSize     As Long
    
    On Error GoTo EH
    lBufferSize = oSocket.SockOpt(ucsSsoSendBuffer)
    Do While uOutput.Pos < uOutput.Size
        lBytes = IIf(uOutput.Size - uOutput.Pos > lBufferSize, lBufferSize, uOutput.Size - uOutput.Pos)
        lBytes = oSocket.Send(VarPtr(uOutput.Data(uOutput.Pos)), lBytes)
        If oSocket.HasPendingEvent Then
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
    '--- success
    pvBufferSendAsync = True
QH:
    Exit Function
EH:
    PrintError FUNC_NAME
End Function

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

Private Property Get pvArraySize(baArray() As Byte) As Long
    Dim lPtr            As Long
    
    '--- peek long at ArrPtr(baArray)
    Call CopyMemory(lPtr, ByVal ArrPtr(baArray), 4)
    If lPtr <> 0 Then
        pvArraySize = UBound(baArray) + 1
    End If
End Property

'= common helpers ========================================================

Private Function pvNetworkShort(ByVal lValue As Long) As Integer
    pvNetworkShort = ws_htons(lValue)
End Function

Private Function pvNetworkLong(ByVal lValue As Long) As Long
    pvNetworkLong = ws_htonl(lValue)
End Function

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

Private Property Get TimerEx() As Double
    Dim cFreq           As Currency
    Dim cValue          As Currency

    Call QueryPerformanceFrequency(cFreq)
    Call QueryPerformanceCounter(cValue)
    TimerEx = cValue / cFreq
End Property

'Private Function pvCalcAdler32(ByVal lPtr As Long, ByVal lSize As Long) As Long
'    Const LNG_BASE      As Long = 65521
'    Dim baData()        As Byte
'    Dim lA              As Long
'    Dim lB              As Long
'    Dim lIdx            As Long
'
'    If lSize > 0 Then
'        ReDim baData(0 To lSize - 1) As Byte
'        Call CopyMemory(baData(0), ByVal lPtr, lSize)
'        lA = 1
'        For lIdx = 0 To lSize - 1
'            lA = (lA + baData(lIdx)) Mod LNG_BASE
'            lB = (lB + lA) Mod LNG_BASE
'        Next
'        pvCalcAdler32 = lB * &H1000& Or lA
'    End If
'End Function

'=========================================================================
' Event impl
'=========================================================================

Private Sub m_oSocket_OnAccept()
    Dim oSocket         As cAsyncSocket
    
    If Not m_oSocket.Accept(oSocket) Then
        GoTo QH
    End If
    Set m_cClients = New cVncServer
    m_cClients.Accept oSocket
QH:
End Sub

Private Sub m_oSocket_OnReceive()
    Const FUNC_NAME     As String = "m_oSocket_OnReceive"
    Dim baBuffer()      As Byte
    Dim vElem           As Variant
    
    On Error GoTo EH
    Do
        If Not m_oSocket.ReceiveArray(baBuffer) Then
            GoTo QH
        End If
        If UBound(baBuffer) < 0 Then
            GoTo QH
        End If
        pvBufferWriteArray m_uRecvBuffer, baBuffer
        With m_uRecvBuffer
            Do While .Pos < .Size
                Select Case m_eState
                Case ucsStaExpectClientMessage
                    If Not pvParseClientMessage(m_uRecvBuffer) Then
                        Exit Do
                    End If
                Case ucsStaExpectProtocolVersion
                    If .Size - .Pos < Len(STR_PROTOCOL_VERSION) Then
                        Exit Do
                    End If
                    '--- ignore client protocol version
                    .Pos = .Pos + Len(STR_PROTOCOL_VERSION)
                    If Not pvSendSecurityType(m_uSendBuffer, m_oSocket) Then
                        m_eState = ucsStaNone
                        GoTo QH
                    End If
                    m_eState = ucsStaExpectSecurityHandshake
                Case ucsStaExpectSecurityHandshake
                    '--- ignore security type
                    Debug.Assert .Data(.Pos) = ucsAuthNone
                    .Pos = .Pos + 1
                    m_eState = ucsStaExpectClientInit
                Case ucsStaExpectClientInit
                    '--- ignore shared flag
                    .Pos = .Pos + 1
                    vElem = pvCaptureEnumDevices().Item(1)
                    If Not pvCaptureInit(m_uCaptureCtx, vElem(0), 1) Then
                        GoTo QH
                    End If
                    With m_uCaptureFrame
                        ReDim .MoveRects(0 To 0) As DXGI_OUTDUPL_MOVE_RECT
                        ReDim .DirtyRects(0 To 0) As D3D11_RECT
                        ReDim .PointerShape(0 To 0) As Byte
                        ReDim .Tiles(0 To 31) As UcsCaptureTile
                    End With
                    If Not pvSendServerInit(m_uCaptureCtx, m_uSendBuffer, m_oSocket) Then
                        m_eState = ucsStaNone
                        GoTo QH
                    End If
                    m_eState = ucsStaExpectClientMessage
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
QH:
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub m_oSocket_OnSend()
    Const FUNC_NAME     As String = "m_oSocket_OnSend"
    
    On Error GoTo EH
    If m_uSendBuffer.Size > m_uSendBuffer.Pos Then
        If Not pvBufferSendAsync(m_uSendBuffer, m_oSocket) Then
            GoTo QH
        End If
    End If
QH:
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

'=========================================================================
' The Modern Subclassing Thunk (MST)
'=========================================================================

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


