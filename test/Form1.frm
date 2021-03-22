VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8232
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4260
   LinkTopic       =   "Form1"
   ScaleHeight     =   8232
   ScaleWidth      =   4260
   StartUpPosition =   3  'Windows Default
   Begin VB.Label labDebug 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   7.8
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   0
      Left            =   168
      TabIndex        =   1
      Top             =   336
      UseMnemonic     =   0   'False
      Width           =   84
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   192
      Left            =   168
      TabIndex        =   0
      Top             =   84
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
    Dim lIdx            As Long
    
    For lIdx = 1 To 100
        Load labDebug(lIdx)
        labDebug(lIdx).Move labDebug(lIdx - 1).Left, labDebug(lIdx - 1).Top + 240
        labDebug(lIdx).Visible = True
    Next
    
    Set m_oServer = New cVncServer
    If Not m_oServer.Init("0.0.0.0", 5900) Then
        MsgBox m_oServer.LastError, vbExclamation
        Unload Me
    Else
        m_oServer.Socket.GetSockName sAddress, lPort
        Label1.Caption = "Waiting for connection on " & sAddress & ":" & lPort
    End If
End Sub
