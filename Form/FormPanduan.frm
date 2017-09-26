VERSION 5.00
Begin VB.Form FormPanduan 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Panduan"
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6750
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormPanduan.frx":0000
   LinkTopic       =   "Panduan"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   6750
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CmdExit 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Keluar"
      Height          =   375
      Left            =   5160
      TabIndex        =   2
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Frame FramePanduan 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Panduan "
      Height          =   2895
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   6495
      Begin VB.TextBox TextPanduan 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2535
         Left            =   2160
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Text            =   "FormPanduan.frx":0CCA
         Top             =   240
         Width           =   4215
      End
      Begin VB.Image ImageLogo 
         Height          =   2535
         Left            =   120
         Picture         =   "FormPanduan.frx":0F70
         Stretch         =   -1  'True
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.CommandButton CmdPanduanManual 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Panduan Manual"
      Height          =   375
      Left            =   3600
      TabIndex        =   0
      Top             =   3120
      Width           =   1455
   End
End
Attribute VB_Name = "FormPanduan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''********************************************************************''
''                                                                    ''
'' \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\' ''
'' \\ Project Name       :   Panda Pustaka                        \\' ''
'' \\ Project Version    :   1.0                                  \\' ''
'' \\ Project Author     :   Raka Suryaardi Widjaja               \\' ''
'' \\ Project Home Page  :   www.Gurisa.Com                       \\' ''
'' \\ Project License    :   All Right Reserved Gurisa © 2015     \\' ''
'' \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\' ''
''                                                                    ''
''********************************************************************''

Private Sub CmdExit_Click()
Unload Me
End Sub

Private Sub CmdPanduanManual_Click()
If Dir$(App.Path & "\Panduan\Panduan.pdf") <> "" Then
    OpenFile (App.Path & "\Panduan\Panduan.pdf")
Else
    MsgBox "File Panduan Tidak Tersedia", vbInformation, "Hubungi Pengembang Aplikasi"
End If
End Sub

