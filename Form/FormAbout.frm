VERSION 5.00
Begin VB.Form FormAbout 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tentang"
   ClientHeight    =   2955
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7110
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "FormAbout.frx":0000
   LinkTopic       =   "About Author"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2955
   ScaleWidth      =   7110
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox CheckTampil 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Jangan Tampilkan Lagi"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2880
      TabIndex        =   2
      Top             =   2520
      Width           =   2295
   End
   Begin VB.TextBox TextKeterangan 
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
      ForeColor       =   &H00000000&
      Height          =   2175
      Left            =   2880
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "FormAbout.frx":0CCA
      Top             =   240
      Width           =   4095
   End
   Begin VB.CommandButton CmdClose 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Tutup"
      Height          =   375
      Left            =   5400
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   0
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Image ImageLogo 
      Height          =   2625
      Left            =   120
      Picture         =   "FormAbout.frx":0E3F
      Stretch         =   -1  'True
      Top             =   120
      Width           =   2625
   End
End
Attribute VB_Name = "FormAbout"
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

Private Sub CheckTampil_Click()
Koneksi "SELECT status_setting_enum, status_setting_text FROM tb_setting WHERE id_setting='3'"

If DB!status_setting_enum = "Ya" Then
    Koneksi "SELECT status_setting_enum, status_setting_text FROM tb_setting WHERE id_setting='1'"
    If DB!status_setting_enum = "Ya" Then
        CheckTampil.Value = Checked
        Conn.Execute "UPDATE tb_setting SET status_setting_enum='Tidak', status_setting_text='Tidak' WHERE id_setting='1'"
    ElseIf DB!status_setting_enum = "Tidak" Then
        CheckTampil.Value = Unchecked
        Conn.Execute "UPDATE tb_setting SET status_setting_enum='Ya', status_setting_text='Ya' WHERE id_setting='1'"
    End If
ElseIf DB!status_setting_enum = "Tidak" Then
    CheckTampil.Enabled = False
End If
End Sub

Private Sub CmdClose_Click()
Unload Me
End Sub

Private Sub Form_Load()
Koneksi "SELECT status_setting_enum, status_setting_text FROM tb_setting WHERE id_setting='3'"
If DB!status_setting_enum = "Ya" Then
    CheckTampil.Enabled = True
    Koneksi "SELECT status_setting_enum, status_setting_text FROM tb_setting WHERE id_setting='1'"
    If DB!status_setting_enum = "Ya" Then
        CheckTampil.Value = Unchecked
    ElseIf DB!status_setting_enum = "Tidak" Then
        CheckTampil.Value = Checked
    End If
ElseIf DB!status_setting_enum = "Tidak" Then
    CheckTampil.Enabled = False
    CheckTampil.Value = Unchecked
End If
End Sub
