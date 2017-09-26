VERSION 5.00
Begin VB.Form FormLogin 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Panda Pustaka"
   ClientHeight    =   3405
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   Icon            =   "FormLogin.frx":0000
   LinkTopic       =   "Login"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3405
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdLogin 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Login"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   4
      Top             =   2520
      Width           =   1335
   End
   Begin VB.TextBox TextPassword 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1440
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   2040
      Width           =   2895
   End
   Begin VB.TextBox TextID 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1440
      MaxLength       =   5
      TabIndex        =   0
      Top             =   1560
      Width           =   2895
   End
   Begin VB.Timer TimerExit 
      Left            =   120
      Top             =   2640
   End
   Begin VB.Label LabelPassword 
      BackColor       =   &H00000000&
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label LabelID 
      BackColor       =   &H00000000&
      Caption         =   "Identifier"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Image ImageBackgroundLogin 
      Height          =   1935
      Left            =   0
      Picture         =   "FormLogin.frx":0CCA
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   4575
   End
   Begin VB.Image ImageLogo 
      Height          =   1095
      Left            =   1680
      Picture         =   "FormLogin.frx":0E0D
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "FormLogin"
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
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" ( _
ByVal hwnd As Long, _
ByVal nIndex As Long) As Long
     
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" ( _
ByVal hwnd As Long, _
ByVal nIndex As Long, _
ByVal dwNewLong As Long) As Long
                   
Private Declare Function SetLayeredWindowAttributes Lib "user32" ( _
ByVal hwnd As Long, _
ByVal crKey As Long, _
ByVal bAlpha As Byte, _
ByVal dwFlags As Long) As Long
     
Private Const GWL_STYLE = (-16)
Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_LAYERED = &H80000
Private Const LWA_COLORKEY = &H1
Private Const LWA_ALPHA = &H2

Private Sub CmdLogin_Click()
TextID = FilterInjeksi(TextID.Text)
TextPassword = FilterInjeksi(TextPassword.Text)

If TextID.Text = "" Or TextPassword.Text = "" Then
    MsgBox "Silahkan Isi ID Dan Password Dengan Benar", vbCritical, "ID Atau Password Masing Kosong"
    Exit Sub
ElseIf TemukanError(TextID.Text + TextPassword.Text) = True Then
    Exit Sub
End If

Koneksi "SELECT * FROM tb_petugas WHERE nama_petugas='" & TextID & "' AND password_petugas='" & TextPassword & "'"
If DB.RecordCount > 0 Then
    MsgBox "Selamat Datang Kembali " & DB!nama_petugas & " ^_^", vbInformation, "Panda Pustaka"
    FormUtama.StatusBarUtama.Panels(4) = FormLogin.TextID.Text
    Koneksi "SELECT * FROM tb_petugas WHERE nama_petugas='" & TextID & "'"
    FormUtama.StatusBarUtama.Panels(5) = DB!id_petugas
    FormUtama.LabelProfileIdentifier = DB!id_petugas
    FormUtama.LabelProfileNama = DB!nama_petugas
    FormUtama.LabelProfileAlamat = DB!alamat_petugas
    FormUtama.Show
    Koneksi "SELECT status_setting_enum, status_setting_text FROM tb_setting WHERE id_setting='3'"
    If DB!status_setting_enum = "Ya" Then
        Koneksi "SELECT status_setting_enum, status_setting_text FROM tb_setting WHERE id_setting='1'"
        If DB!status_setting_enum = "Ya" Then
            FormAbout.Show
        End If
        
        Koneksi "SELECT status_setting_enum, status_setting_text FROM tb_setting WHERE id_setting='2'"
        If DB!status_setting_enum = "Ya" Then
            FormPopUp.Show
            FormPopUp.LabelNama.Caption = TextID.Text
            Koneksi "SELECT COUNT(status_request) AS jumlah_pesan FROM tb_request WHERE nama_penerima='" & TextID & "' AND status_request='Belum Di Baca'"
            FormPopUp.LabelJumlah.Caption = DB!jumlah_pesan & " Pesan Baru"
        End If
    Else
        FormAbout.Show
        FormPopUp.Show
        FormPopUp.LabelNama.Caption = TextID.Text
        Koneksi "SELECT COUNT(status_request) AS jumlah_pesan FROM tb_request WHERE nama_penerima='" & TextID & "' AND status_request='Belum Di Baca'"
        FormPopUp.LabelJumlah.Caption = DB!jumlah_pesan & " Pesan Baru"
    End If
    Unload Me
Else
    Koneksi "SELECT * FROM tb_anggota WHERE id_anggota='" & TextID & "' AND password_anggota='" & TextPassword & "'"
    If DB.RecordCount > 0 Then
        MsgBox "Selamat Datang Kembali " & DB!nama_anggota & " ^_^", vbInformation, "Panda Pustaka"
        FormClient.StatusBarClient.Panels(1) = "ID :"
        FormClient.StatusBarClient.Panels(2) = FormLogin.TextID.Text
        FormClient.Show
        FormPopUp.Show
        Koneksi "SELECT nama_anggota FROM tb_anggota WHERE id_anggota='" & TextID & "'"
        FormPopUp.LabelNama.Caption = DB!nama_anggota
        Koneksi "SELECT COUNT(status_request) AS jumlah_pesan FROM tb_request WHERE id_anggota_tujuan='" & TextID & "' AND status_request='Belum Di Baca'"
        FormPopUp.LabelJumlah.Caption = DB!jumlah_pesan & " Pesan Baru"
        Unload Me
    Else
        MsgBox "Silahkan Periksa ID Dan Password", vbCritical, "ID Atau Password Salah"
    End If
End If
End Sub

Private Sub Form_Load()
Koneksi "SHOW DATABASES"
TextID.Text = ""
TextPassword.Text = ""
TimerExit.Enabled = True
TimerExit.Interval = 60000

Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
ImageLogo.Stretch = True
Me.BackColor = vbCyan
SetWindowLong Me.hwnd, GWL_EXSTYLE, GetWindowLong(Me.hwnd, GWL_EXSTYLE) Or WS_EX_LAYERED
SetLayeredWindowAttributes Me.hwnd, vbCyan, 0&, LWA_COLORKEY
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    TextID.SetFocus
End If
End Sub

Private Sub TextID_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
TextPassword.SetFocus
End If
End Sub

Private Sub TextPassword_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
CmdLogin.SetFocus
End If
End Sub

Private Sub TimerExit_Timer()
Unload Me
End Sub
