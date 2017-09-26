VERSION 5.00
Begin VB.Form FormPanel 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Panel"
   ClientHeight    =   4530
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8310
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormPanel.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4530
   ScaleWidth      =   8310
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FramePanel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Panel"
      ForeColor       =   &H80000008&
      Height          =   2775
      Left            =   120
      TabIndex        =   15
      Top             =   120
      Width           =   8055
      Begin VB.CommandButton CmdRestoreDefault 
         Caption         =   "Restore Default"
         Height          =   375
         Left            =   5280
         TabIndex        =   32
         Top             =   1800
         Width           =   2535
      End
      Begin VB.TextBox TextServer 
         Height          =   375
         Left            =   1320
         TabIndex        =   25
         Top             =   840
         Width           =   3135
      End
      Begin VB.TextBox TextUserID 
         Height          =   375
         Left            =   1320
         TabIndex        =   24
         Top             =   1320
         Width           =   3735
      End
      Begin VB.TextBox TextPassword 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1320
         PasswordChar    =   "*"
         TabIndex        =   23
         Top             =   1800
         Width           =   3735
      End
      Begin VB.TextBox TextDriver 
         Height          =   375
         Left            =   1320
         TabIndex        =   22
         Top             =   360
         Width           =   3135
      End
      Begin VB.TextBox TextPort 
         Height          =   375
         Left            =   5400
         TabIndex        =   21
         Top             =   360
         Width           =   2415
      End
      Begin VB.ComboBox ComboDatabase 
         Height          =   330
         Left            =   1320
         TabIndex        =   20
         Text            =   "ComboDatabase"
         Top             =   2280
         Width           =   3735
      End
      Begin VB.CommandButton CmdSaveSetting 
         Caption         =   "Simpan"
         Height          =   375
         Left            =   5280
         TabIndex        =   19
         Top             =   2280
         Width           =   2535
      End
      Begin VB.CommandButton CmdTestSetting 
         Caption         =   "Test"
         Height          =   375
         Left            =   5280
         TabIndex        =   18
         Top             =   1320
         Width           =   2535
      End
      Begin VB.CheckBox CheckPopUpMessage 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Pop Up Message"
         Height          =   255
         Left            =   4680
         TabIndex        =   17
         Top             =   840
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.CheckBox CheckPopUpAbout 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Pop Up About"
         Height          =   255
         Left            =   6480
         TabIndex        =   16
         Top             =   840
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin VB.Label LabelServer 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Server"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   31
         Top             =   840
         Width           =   855
      End
      Begin VB.Label LabelPort 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Port"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4680
         TabIndex        =   30
         Top             =   360
         Width           =   735
      End
      Begin VB.Label LabelUserID 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "User ID"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   29
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label LabelPassword 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   28
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label LabelDriver 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Driver"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   27
         Top             =   360
         Width           =   855
      End
      Begin VB.Label LabelDatabase 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Database"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   26
         Top             =   2280
         Width           =   855
      End
   End
   Begin VB.Frame FrameMenu 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   5520
      TabIndex        =   14
      Top             =   3000
      Width           =   2655
      Begin VB.Image ImagePanel 
         Height          =   840
         Left            =   120
         Picture         =   "FormPanel.frx":0CCA
         Stretch         =   -1  'True
         Top             =   360
         Width           =   720
      End
      Begin VB.Image ImageInternet 
         Height          =   840
         Left            =   1800
         Picture         =   "FormPanel.frx":42B8
         Stretch         =   -1  'True
         Top             =   360
         Width           =   840
      End
      Begin VB.Image ImageDonate 
         Height          =   840
         Left            =   960
         Picture         =   "FormPanel.frx":87D8
         Stretch         =   -1  'True
         Top             =   360
         Width           =   840
      End
   End
   Begin VB.Frame FrameAktivasi 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Aktivasi"
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   120
      TabIndex        =   7
      Top             =   3000
      Width           =   5295
      Begin VB.CommandButton CmdGenerateGet 
         Caption         =   "Generate"
         Height          =   735
         Left            =   3240
         TabIndex        =   11
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton CmdAktivasi 
         Caption         =   "Aktivasi"
         Height          =   735
         Left            =   4200
         TabIndex        =   10
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox TextCodeSend 
         Height          =   315
         Left            =   1320
         TabIndex        =   9
         Top             =   840
         Width           =   1815
      End
      Begin VB.TextBox TextCodeGet 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label LabelCodeSend 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Code Send"
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
         Left            =   240
         TabIndex        =   13
         Top             =   840
         Width           =   975
      End
      Begin VB.Label LabelCodeGet 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Code Get"
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
         Left            =   240
         TabIndex        =   12
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame FrameDonate 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Donate"
      ForeColor       =   &H80000008&
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8055
      Begin VB.Frame FrameBCA 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   2175
         Left            =   4560
         TabIndex        =   2
         Top             =   480
         Width           =   3375
         Begin VB.TextBox TextANBCA 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   5
            Text            =   "Raka Suryaardi Widjaja"
            Top             =   1680
            Width           =   1935
         End
         Begin VB.TextBox TextBCA 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   4
            Text            =   "0080898061"
            Top             =   1680
            Width           =   1095
         End
         Begin VB.Image ImageBCA 
            Height          =   1335
            Left            =   120
            Picture         =   "FormPanel.frx":D873
            Stretch         =   -1  'True
            Top             =   240
            Width           =   3135
         End
      End
      Begin VB.Frame FrameBitCoin 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   2175
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   4335
         Begin VB.TextBox TextBitCoin 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
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
            Height          =   375
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   3
            Text            =   "1QLP8hWB49Pvxg4uhPcKKqWqN2ihwQ5H29"
            Top             =   1680
            Width           =   4095
         End
         Begin VB.Image ImageQRBitCoin 
            Height          =   1335
            Left            =   2520
            Picture         =   "FormPanel.frx":461B9
            Stretch         =   -1  'True
            Top             =   240
            Width           =   1695
         End
         Begin VB.Image ImageBitCoin 
            Height          =   1335
            Left            =   120
            Picture         =   "FormPanel.frx":49E9D
            Stretch         =   -1  'True
            Top             =   240
            Width           =   2295
         End
      End
      Begin VB.Label LabelDonate 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Bring Me Some Coffe And Donuts!"
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
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   2895
      End
   End
End
Attribute VB_Name = "FormPanel"
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

Private Sub CheckPopUpAbout_Click()
If CheckPopUpAbout.Value = Checked Then
    Conn.Execute "UPDATE tb_setting SET status_setting_enum='Ya', status_setting_text='Ya' WHERE id_setting='1'"
ElseIf CheckPopUpAbout.Value = Unchecked Then
    Conn.Execute "UPDATE tb_setting SET status_setting_enum='Tidak', status_setting_text='Tidak' WHERE id_setting='1'"
End If
End Sub

Private Sub CheckPopUpMessage_Click()
If CheckPopUpMessage.Value = Checked Then
    Conn.Execute "UPDATE tb_setting SET status_setting_enum='Ya', status_setting_text='Ya' WHERE id_setting='2'"
ElseIf CheckPopUpMessage.Value = Unchecked Then
    Conn.Execute "UPDATE tb_setting SET status_setting_enum='Tidak', status_setting_text='Tidak' WHERE id_setting='2'"
End If
End Sub

Private Sub CmdAktivasi_Click()
TextCodeGet = FilterInjeksi(TextCodeGet.Text)

If CmdAktivasi.Caption = "Aktivasi" Then
    If TextCodeGet.Text = "" Or TextCodeSend.Text = "" Then
        MsgBox "Isi Code Aktivasi", vbExclamation, "Code Aktivasi Masih Kosong"
    ElseIf Val(TextCodeSend.Text) = Val(TextCodeGet.Text) + 12111997 Then
        Conn.Execute "UPDATE tb_setting SET status_setting_enum='Ya', status_setting_text='Ya', code_get='" & TextCodeGet & "', code_send='" & TextCodeSend.Text & "' WHERE id_setting='3'"
        TextCodeSend.Enabled = False
        CmdGenerateGet.Enabled = False
        CmdAktivasi.Caption = "Batal"
        CheckPopUpAbout.Enabled = True
        CheckPopUpMessage.Enabled = True
        LabelDonate.Caption = "Enjoy Your Day!"
        LabelDonate.ForeColor = &H8000&
        MsgBox "Aktivasi Berhasil Di Lakukan", vbInformation, "Terima Kasih Kopi Dan Donat Nya (^_^)/"
    Else
        MsgBox "Code Aktivasi Salah", vbCritical, "Code Aktivasi Salah"
    End If
ElseIf CmdAktivasi.Caption = "Batal" Then
    If MsgBox("Batalkan Status Aktivasi Panda Pustaka?", vbInformation + vbYesNo, "Batalkan Status Aktivasi") = vbYes Then
        Conn.Execute "UPDATE tb_setting SET status_setting_enum='Tidak', status_setting_text='Tidak', code_get='', code_send='' WHERE id_setting='3'"
        CmdAktivasi.Enabled = True
        CmdAktivasi.Caption = "Aktivasi"
        CmdGenerateGet.Enabled = True
        TextCodeSend.Enabled = True
        CheckPopUpAbout.Enabled = False
        CheckPopUpMessage.Enabled = False
        Koneksi "SELECT code_get, code_send FROM tb_setting WHERE id_setting='3'"
        TextCodeGet.Text = DB!code_get
        TextCodeSend.Text = DB!code_send
        LabelDonate.Caption = "Bring Me Some Coffe And Donuts!"
        LabelDonate.ForeColor = &HFF&
        Call CmdGenerateGet_Click
        Conn.Execute "UPDATE tb_setting SET status_setting_enum='Ya', status_setting_text='Ya' WHERE id_setting='1'"
        CheckPopUpAbout.Value = Checked
        Conn.Execute "UPDATE tb_setting SET status_setting_enum='Ya', status_setting_text='Ya' WHERE id_setting='2'"
        CheckPopUpMessage.Value = Checked
        MsgBox "Berhasil Membatalkan Aktivasi", vbExclamation, "Terima Kasih Sudah Memberikan Kopi Dan Donat Nya (^_^)/"
    Else
        Exit Sub
    End If
End If
End Sub

Private Sub CmdGenerateGet_Click()
TextCodeGet.Text = Format(Date, "DDYYYYMM") & Format(Time, "HHSSMM")
End Sub

Private Sub CmdRestoreDefault_Click()
If MsgBox("Ubah Ke Pengaturan Bawaan ?", vbInformation + vbYesNo, "Pengaturan Bawaan") = vbYes Then
    Call WriteToPanelDefault
    Call ReadFromPanel
    MsgBox "Berhasil Mengubah Pengaturan Bawaan" & vbNewLine & "Membuka Ulang Aplikasi", vbInformation, "Buka Ulang Aplikasi"
    Unload FormAbout
    Unload FormAnggota
    Unload FormClient
    Unload FormDenda
    Unload FormLog
    Unload FormLogin
    Unload FormPanduan
    Unload FormPanel
    Unload FormPopUp
    Unload FormReport
    Unload FormSplash
    Unload FormUtama
    FormSplash.Show
Else
    Exit Sub
End If
End Sub

Private Sub CmdSaveSetting_Click()
If CmdSaveSetting.Caption = "Ubah" Then
    If MsgBox("Mengubah Pengaturan Dapat Menyebabkan Aplikasi Tidak Berfungsi" & vbNewLine & "Tetap Ubah Pengaturan ?", vbInformation + vbYesNo, "Ubah Pengaturan") = vbYes Then
        CmdSaveSetting.Caption = "Simpan"
        TextDriver.Enabled = True
        TextServer.Enabled = True
        TextUserID.Enabled = True
        TextPassword.Enabled = True
        TextPort.Enabled = True
        CmdTestSetting.Enabled = False
        CmdRestoreDefault.Enabled = False
        ComboDatabase.Enabled = True
        Call ReadFromPanel
    Else
        Exit Sub
    End If
ElseIf CmdSaveSetting.Caption = "Simpan" Then
    If TextDriver.Text = "" Or TextServer.Text = "" Or TextUserID.Text = "" Or TextPassword.Text = "" Or TextPort.Text = "" Then
        MsgBox "Isi Pengaturan Dengan Lengkap", vbExclamation, "Pengaturan"
    Else
    If MsgBox("Pastikan Pengaturan Sudah Benar" & vbNewLine & "Simpan Pengaturan ?", vbInformation + vbYesNo, "Simpan Pengaturan") = vbYes Then
        CmdSaveSetting.Caption = "Ubah"
        TextDriver.Enabled = False
        TextServer.Enabled = False
        TextUserID.Enabled = False
        TextPassword.Enabled = False
        TextPort.Enabled = False
        CmdTestSetting.Enabled = True
        CmdRestoreDefault.Enabled = True
        ComboDatabase.Enabled = False
        Call WriteToPanel
        MsgBox "Berhasil Mengubah Pengaturan" & vbNewLine & "Membuka Ulang Aplikasi", vbInformation, "Buka Ulang Aplikasi"
        Unload FormAbout
        Unload FormAnggota
        Unload FormClient
        Unload FormDenda
        Unload FormLog
        Unload FormLogin
        Unload FormPanduan
        Unload FormPanel
        Unload FormPopUp
        Unload FormReport
        Unload FormSplash
        Unload FormUtama
        FormSplash.Show
    Else
        Exit Sub
    End If
    End If
Else
    MsgBox "Aktivitas Hacking Terdeteksi", vbCritical, "Hacking Activity"
    Unload Me
End If
End Sub

Private Sub CmdTestSetting_Click()
TestKoneksi "SHOW DATABASES"
End Sub

Private Sub Form_Load()
If App.PrevInstance = True Then
    Unload Me
End If
Call CmdGenerateGet_Click
Call ImagePanel_Click
End Sub

Private Sub ImageDonate_Click()
FramePanel.Visible = False
FrameDonate.Visible = True

Koneksi "SELECT code_get, code_send FROM tb_setting WHERE id_setting='3'"
    TextCodeGet.Text = DB!code_get
    TextCodeSend.Text = DB!code_send
    If Val(TextCodeSend.Text) = Val(TextCodeGet.Text) + 12111997 Then
        CmdAktivasi.Caption = "Batal"
        CmdGenerateGet.Enabled = False
        TextCodeSend.Enabled = False
        CheckPopUpAbout.Enabled = True
        CheckPopUpMessage.Enabled = True
        LabelDonate.Caption = "Enjoy Your Day!"
        LabelDonate.ForeColor = &H8000&
            Koneksi "SELECT status_setting_enum FROM tb_setting WHERE id_setting='1'"
            If DB!status_setting_enum = "Ya" Then
                CheckPopUpAbout.Value = Checked
            ElseIf DB!status_setting_enum = "Tidak" Then
                CheckPopUpAbout.Value = Unchecked
            End If
            
            Koneksi "SELECT status_setting_enum FROM tb_setting WHERE id_setting='2'"
            If DB!status_setting_enum = "Ya" Then
                CheckPopUpMessage.Value = Checked
            ElseIf DB!status_setting_enum = "Tidak" Then
                CheckPopUpMessage.Value = Unchecked
            End If
    Else
        CmdAktivasi.Caption = "Aktivasi"
        CmdGenerateGet.Enabled = True
        Call CmdGenerateGet_Click
        CheckPopUpAbout.Enabled = False
        CheckPopUpMessage.Enabled = False
        LabelDonate.Caption = "Bring Me Some Coffe And Donuts!"
        LabelDonate.ForeColor = &HFF&
    End If
End Sub

Private Sub ImagePanel_Click()
FramePanel.Visible = True
FrameDonate.Visible = False

If Dir$(App.Path & "\Panel.panpus") = "" Then
        MsgBox "Konfigurasi File Tidak Di Temukan" & vbNewLine & "Aplikasi Akan Membuat Konfigurasi Bawaan", vbExclamation, "Konfigurasi Bawaan"
        Call WriteToPanelDefault
        Call ReadFromPanel
        TextDriver.Enabled = False
        TextServer.Enabled = False
        TextUserID.Enabled = False
        TextPassword.Enabled = False
        TextPort.Enabled = False
        ComboDatabase.Enabled = False
        CmdSaveSetting.Caption = "Ubah"
    Else
        Call ReadFromPanel
        TextDriver.Enabled = False
        TextServer.Enabled = False
        TextUserID.Enabled = False
        TextPassword.Enabled = False
        TextPort.Enabled = False
        ComboDatabase.Enabled = False
        CmdSaveSetting.Caption = "Ubah"
    End If
End Sub
