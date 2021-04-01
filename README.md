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
     - Inclusing cursor with alpha
 - [ ] Supports file transfer protocol extensions
     - UltraVNC (incl. ZIP compression on folder downloads)
     - TightVNC (latest implementation only)
 - [ ] Single-thread asynchronous implementation
 - [ ] Conditional compilation to reduce footprint on final executable
     - Optional ZLib support
     - Optional histograms for JPEG quality estimation
