VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Text Chat"
   ClientHeight    =   4428
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   6828
   LinkTopic       =   "Form2"
   ScaleHeight     =   4428
   ScaleWidth      =   6828
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   432
      Left            =   504
      TabIndex        =   0
      Top             =   3864
      Width           =   5220
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2952
      Left            =   168
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   168
      Width           =   5472
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents m_oServer    As cVncServer
Attribute m_oServer.VB_VarHelpID = -1
Private m_lConnID               As Long

Property Get ConnID() As Long
    ConnID = m_lConnID
End Property

Public Function Init(oServer As cVncServer, ByVal ConnID As Long) As Boolean
    Set m_oServer = oServer
    m_lConnID = ConnID
    Show
    '--- success
    Init = True
End Function

Private Sub pvAppendText(ByVal sText As String)
    If Right$(sText, 2) <> vbCrLf Then
        sText = sText & vbCrLf
    End If
    With Text1
        .SelStart = &H7FFF
        If .SelStart + Len(sText) > &H7FFF& Then
            sText = .Text & sText
            .Text = Mid$(sText, InStr(Len(sText) - &H8001&, sText, vbCrLf) + 2)
        Else
            .SelText = sText
        End If
        .SelStart = &H7FFF
    End With
End Sub

Private Sub m_oServer_OnTextChatMsg(ByVal ConnID As Long, ByVal MsgType As Long, ByVal MsgText As String)
    If ConnID = m_lConnID And MsgType = 0 Then
        pvAppendText m_lConnID & ": " & MsgText
    End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If m_oServer.TextChatMsg(m_lConnID, 0, Text2.Text & vbCrLf) Then
            pvAppendText "Me: " & Text2.Text
            Text2.Text = vbNullString
        End If
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Resize()
    If WindowState <> vbMinimized Then
        Text2.Move 30, ScaleHeight - Text2.Height - 30, ScaleWidth - 60
        Text1.Move 0, 0, ScaleWidth, Text2.Top - 60
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    m_oServer.TextChatMsg m_lConnID, 2
    ChatWindows.Remove "#" & m_lConnID
End Sub
