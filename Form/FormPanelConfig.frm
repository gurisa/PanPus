VERSION 5.00
Begin VB.Form FormPanelConfig 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Panel"
   ClientHeight    =   3495
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6390
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormPanelConfig.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   6390
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FramePanel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Panel"
      ForeColor       =   &H80000008&
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6135
      Begin VB.CommandButton CmdTestSetting 
         Caption         =   "Test"
         Height          =   375
         Left            =   4560
         TabIndex        =   9
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton CmdSaveSetting 
         Caption         =   "Simpan"
         Height          =   375
         Left            =   4560
         TabIndex        =   8
         Top             =   1320
         Width           =   1335
      End
      Begin VB.ComboBox ComboDatabase 
         Height          =   330
         Left            =   1320
         TabIndex        =   7
         Text            =   "ComboDatabase"
         Top             =   2280
         Width           =   3135
      End
      Begin VB.TextBox TextPort 
         Height          =   375
         Left            =   1320
         TabIndex        =   6
         Top             =   2760
         Width           =   3135
      End
      Begin VB.TextBox TextDriver 
         Height          =   375
         Left            =   1320
         TabIndex        =   5
         Top             =   360
         Width           =   3135
      End
      Begin VB.TextBox TextPassword 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1320
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   1800
         Width           =   3135
      End
      Begin VB.TextBox TextUserID 
         Height          =   375
         Left            =   1320
         TabIndex        =   3
         Top             =   1320
         Width           =   3135
      End
      Begin VB.TextBox TextServer 
         Height          =   375
         Left            =   1320
         TabIndex        =   2
         Top             =   840
         Width           =   3135
      End
      Begin VB.CommandButton CmdRestoreDefault 
         Caption         =   "Restore Default"
         Height          =   375
         Left            =   4560
         TabIndex        =   1
         Top             =   840
         Width           =   1335
      End
      Begin VB.Image ImagePanel 
         Height          =   1095
         Left            =   4680
         Picture         =   "FormPanelConfig.frx":0CCA
         Stretch         =   -1  'True
         Top             =   1920
         Width           =   1215
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
         TabIndex        =   15
         Top             =   2280
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
         TabIndex        =   14
         Top             =   360
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
         TabIndex        =   13
         Top             =   1800
         Width           =   855
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
         TabIndex        =   12
         Top             =   1320
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
         Left            =   240
         TabIndex        =   11
         Top             =   2760
         Width           =   735
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
         TabIndex        =   10
         Top             =   840
         Width           =   855
      End
   End
End
Attribute VB_Name = "FormPanelConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdRestoreDefault_Click()
If MsgBox("Ubah Ke Pengaturan Bawaan ?", vbInformation + vbYesNo, "Pengaturan Bawaan") = vbYes Then
    Call WriteToPanelDefault
    Call ReadFromPanel
    MsgBox "Berhasil Mengubah Pengaturan Bawaan" & vbNewLine & "Membuka Ulang Aplikasi", vbInformation, "Buka Ulang Aplikasi"
    Unload FormPanelConfig
    Call EksekusiPandaPustaka
    End
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
        Unload FormPanelConfig
        Call EksekusiPandaPustaka
        End
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
