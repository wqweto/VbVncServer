## VB VNC Server

Simple VNC Server in VB6 using DXGI Desktop Duplication

### Description

 - [ ] Supports password based authentication 
     - Plain, VNC and Tight security types (DES in ECB mode)
     - UltraVNC and UltraVNC_MsLogonIIAuth (Diffie-Hellman KX for the shared key and DES in CBC mode)
 - [ ] Supports Raw, CopyRect, Zlib and Tight encodings
     - Windows Imaging Component (WIC) for fast JPEG compression in Tight encoding
 - [ ] Uses Desktop Duplication from DirectX 11 for incremental screen updates
     - Chunks output in 256x256 tiles
     - Multiple clients per desktop duplication source
 - [ ] Supports cursor pseudo-encodings
     - Including cursor with alpha
 - [ ] Supports file transfer protocol extensions
     - UltraVNC (including ZIP compression on folder downloads)
     - TightVNC (latest version of protocol extension only)
 - [ ] Single-thread asynchronous implementation
 - [ ] Conditional compilation to reduce footprint on final executable
     - Optional ZLib support
     - Optional histograms for JPEG quality estimation

### How to use

Copy `VBD3D11.tlb` type library from `lib\VBD3D11\typelib` to your local disk and register it before adding a project reference to this newly registered "DirectX 11 for VB6 1.0" from Project->References menu in VB IDE.

Add `cVncServer.cls`, `cAsyncSocket.cls` (from `lib\VbAsyncSocket\src`) and `cZipArchive.cls` (from `lib\ZipArchive\src`) to your project and use `cVncServer.Init` method to start a new VNC server like this:

```
Option Explicit

Private m_oServer As cVncServer

Private Sub Form_Load()
    Dim lPort           As Long
    
    lPort = Val(Environ$("MYLOBAPP_VNC_PORT"))
    If lPort <> 0 Then
        Set m_oServer = New cVncServer
        m_oServer.Init "0.0.0.0", lPort, Environ$("MYLOBAPP_VNC_PASSWORD")
    End If
End Sub
```

The snippet above allows optionally starting the built-in in your LOB application VNC server by configuring a listening port through the `MYLOBAPP_VNC_PORT` environment variable and optionally securing it with a password through the `MYLOBAPP_VNC_PASSWORD` environment variable.

Please when copy/pasting the sample code  use common sense and change the prefix of these environment variables to include your LOB application name.

To reduce footprint on final executable you can optionally exclude `cZipArchive.cls` from your project by using `VNC_NOZLIB = 1` in conditional compilation settings which will remove (or reduce) some of the server functionalities.
