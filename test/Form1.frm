VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2316
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   3624
   LinkTopic       =   "Form1"
   ScaleHeight     =   2316
   ScaleWidth      =   3624
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   192
      Left            =   252
      TabIndex        =   0
      Top             =   504
      UseMnemonic     =   0   'False
      Width           =   492
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_oServer As cVncServer

Private Sub Form_Load()
    Dim sAddress        As String
    Dim lPort           As Long
    
    Set m_oServer = New cVncServer
    If Not m_oServer.Init("0.0.0.0", 5900) Then
        MsgBox m_oServer.LastError, vbExclamation
        Unload Me
    Else
        m_oServer.Socket.GetSockName sAddress, lPort
        Label1.Caption = "Waiting for connection on " & sAddress & ":" & lPort
    End If
End Sub
