VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FormUtama 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Panda Pustaka"
   ClientHeight    =   8835
   ClientLeft      =   60
   ClientTop       =   705
   ClientWidth     =   12510
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormUtama.frx":0000
   LinkTopic       =   "Perpustakaan"
   MaxButton       =   0   'False
   ScaleHeight     =   589
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   834
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameHead 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   8295
      Left            =   1920
      TabIndex        =   44
      Top             =   120
      Width           =   10455
      Begin VB.Frame FrameDataPinjam 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Peminjam "
         Height          =   1575
         Left            =   120
         TabIndex        =   54
         Top             =   240
         Width           =   10215
         Begin VB.TextBox ComboNamaAnggota 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   240
            Locked          =   -1  'True
            TabIndex        =   57
            Top             =   480
            Width           =   9735
         End
         Begin VB.ComboBox ComboIDAnggota 
            Height          =   360
            Left            =   240
            Style           =   2  'Dropdown List
            TabIndex        =   56
            Top             =   960
            Width           =   2175
         End
         Begin VB.CommandButton CmdPinjamkan 
            Caption         =   "&Pinjamkan"
            Height          =   375
            Left            =   8640
            TabIndex        =   55
            Top             =   960
            Width           =   1335
         End
      End
      Begin VB.Frame FrameDataPinjamDetail 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Detail Pinjam "
         Height          =   6375
         Left            =   120
         TabIndex        =   45
         Top             =   1800
         Width           =   10215
         Begin VB.TextBox TextNamaBuku 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   240
            Locked          =   -1  'True
            TabIndex        =   52
            Top             =   5160
            Width           =   7815
         End
         Begin VB.CommandButton CmdPinjam 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Pinjam"
            Height          =   375
            Left            =   6840
            TabIndex        =   51
            Top             =   5640
            Width           =   1215
         End
         Begin VB.ComboBox ComboIDBuku 
            Height          =   360
            Left            =   2760
            Style           =   2  'Dropdown List
            TabIndex        =   50
            Top             =   4680
            Width           =   2415
         End
         Begin VB.CommandButton CmdHapusListData 
            Caption         =   "&Hapus"
            Height          =   375
            Left            =   5520
            TabIndex        =   49
            Top             =   5640
            Width           =   1215
         End
         Begin VB.ComboBox ComboIDPinjamDetail 
            BackColor       =   &H00FFFFFF&
            Height          =   360
            Left            =   240
            Style           =   2  'Dropdown List
            TabIndex        =   48
            Top             =   4680
            Width           =   2415
         End
         Begin VB.TextBox TextJumlahBuku 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   5280
            TabIndex        =   47
            Text            =   "Jumlah"
            Top             =   4680
            Width           =   1455
         End
         Begin VB.CommandButton CmdTambah 
            Caption         =   "&Tambah"
            Height          =   375
            Left            =   6840
            TabIndex        =   46
            Top             =   4680
            Width           =   1215
         End
         Begin MSComctlLib.ListView ListData 
            Height          =   4095
            Left            =   240
            TabIndex        =   53
            Top             =   480
            Width           =   9735
            _ExtentX        =   17171
            _ExtentY        =   7223
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            _Version        =   393217
            ForeColor       =   0
            BackColor       =   16777215
            Appearance      =   1
            NumItems        =   0
         End
         Begin VB.Image ImageHiasanPinjam 
            Height          =   1335
            Left            =   8400
            Picture         =   "FormUtama.frx":0CCA
            Stretch         =   -1  'True
            Top             =   4680
            Width           =   1335
         End
      End
   End
   Begin VB.Frame FrameDetailPinjam 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   8295
      Left            =   1920
      TabIndex        =   34
      Top             =   120
      Width           =   10455
      Begin VB.Frame FramePinjam 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Daftar Pinjam "
         ForeColor       =   &H00000000&
         Height          =   7935
         Left            =   120
         TabIndex        =   35
         Top             =   240
         Width           =   10215
         Begin VB.CommandButton CmdHapusPinjaman 
            Caption         =   "&Hapus"
            Height          =   495
            Left            =   240
            TabIndex        =   41
            Top             =   3960
            Width           =   735
         End
         Begin VB.CommandButton CmdPinjamDetail 
            Caption         =   "&Detail"
            Height          =   495
            Left            =   240
            TabIndex        =   40
            Top             =   4560
            Width           =   735
         End
         Begin VB.CommandButton CmdKembalikan 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Kembalikan"
            Height          =   375
            Left            =   7200
            TabIndex        =   39
            Top             =   7200
            Width           =   1335
         End
         Begin VB.CommandButton CmdHapusDetailPinjaman 
            Caption         =   "&Hapus Detail"
            Height          =   375
            Left            =   5760
            TabIndex        =   38
            Top             =   7200
            Width           =   1335
         End
         Begin VB.CommandButton CmdSelesaiPinjam 
            Caption         =   "&Selesai"
            Height          =   375
            Left            =   8640
            TabIndex        =   37
            Top             =   7200
            Width           =   1335
         End
         Begin VB.CommandButton CmdCariDetailPinjam 
            Caption         =   "&Cari"
            Height          =   375
            Left            =   1080
            TabIndex        =   36
            Top             =   7200
            Width           =   1335
         End
         Begin MSComctlLib.ListView ListPinjamDetail 
            Height          =   3255
            Left            =   1080
            TabIndex        =   42
            Top             =   3840
            Width           =   8895
            _ExtentX        =   15690
            _ExtentY        =   5741
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            _Version        =   393217
            ForeColor       =   0
            BackColor       =   16777215
            Appearance      =   1
            NumItems        =   0
         End
         Begin MSComctlLib.ListView ListPinjam 
            Height          =   3375
            Left            =   240
            TabIndex        =   43
            Top             =   480
            Width           =   9735
            _ExtentX        =   17171
            _ExtentY        =   5953
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            _Version        =   393217
            ForeColor       =   0
            BackColor       =   16777215
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   0
         End
      End
   End
   Begin VB.Frame FrameDetailKembali 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   8295
      Left            =   1920
      TabIndex        =   28
      Top             =   120
      Width           =   10455
      Begin VB.Frame FrameKembali 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Daftar Kembali "
         ForeColor       =   &H00000000&
         Height          =   7935
         Left            =   120
         TabIndex        =   29
         Top             =   240
         Width           =   10215
         Begin VB.CommandButton CmdHapusKembali 
            Caption         =   "&Hapus"
            Height          =   375
            Left            =   7200
            TabIndex        =   32
            Top             =   7200
            Width           =   1335
         End
         Begin VB.CommandButton CmdDenda 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Denda"
            Height          =   375
            Left            =   8640
            TabIndex        =   31
            Top             =   7200
            Width           =   1335
         End
         Begin VB.CommandButton CmdCariDetailKembali 
            Caption         =   "&Cari"
            Height          =   375
            Left            =   240
            TabIndex        =   30
            Top             =   7200
            Width           =   1335
         End
         Begin MSComctlLib.ListView ListKembali 
            Height          =   6615
            Left            =   240
            TabIndex        =   33
            Top             =   480
            Width           =   9735
            _ExtentX        =   17171
            _ExtentY        =   11668
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            _Version        =   393217
            ForeColor       =   0
            BackColor       =   16777215
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   0
         End
      End
   End
   Begin VB.Frame FrameProfilePetugas 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3255
      Left            =   1920
      TabIndex        =   16
      Top             =   120
      Width           =   10455
      Begin VB.Frame FrameViewProfile 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   2895
         Left            =   2160
         TabIndex        =   58
         Top             =   240
         Width           =   8175
         Begin VB.CommandButton CmdUbahProfile 
            Caption         =   "&Ubah"
            Height          =   375
            Left            =   6360
            TabIndex        =   59
            Top             =   360
            Width           =   975
         End
         Begin VB.Label LabelProfilePassword 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "*****"
            Height          =   375
            Left            =   1320
            TabIndex        =   67
            Top             =   1320
            Width           =   4815
         End
         Begin VB.Label LabelPasswordProfile 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Password"
            Height          =   375
            Left            =   240
            TabIndex        =   66
            Top             =   1320
            Width           =   975
         End
         Begin VB.Label LabelNamaProfile 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Username"
            Height          =   375
            Left            =   240
            TabIndex        =   65
            Top             =   840
            Width           =   1095
         End
         Begin VB.Label LabelIdentifierProfile 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Identitas"
            Height          =   375
            Left            =   240
            TabIndex        =   64
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label LabelAlamatProfile 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Alamat"
            Height          =   375
            Left            =   240
            TabIndex        =   63
            Top             =   1800
            Width           =   975
         End
         Begin VB.Label LabelProfileNama 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Username"
            Height          =   375
            Left            =   1320
            TabIndex        =   62
            Top             =   840
            Width           =   4815
         End
         Begin VB.Label LabelProfileAlamat 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Alamat"
            Height          =   615
            Left            =   1320
            TabIndex        =   61
            Top             =   1800
            Width           =   4815
         End
         Begin VB.Label LabelProfileIdentifier 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Identifier"
            Height          =   375
            Left            =   1320
            TabIndex        =   60
            Top             =   360
            Width           =   4815
         End
      End
      Begin VB.Frame FrameChangeProfile 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   2895
         Left            =   2160
         TabIndex        =   17
         Top             =   240
         Width           =   8175
         Begin VB.CommandButton CmdKembaliProfile 
            Caption         =   "&Kembali"
            Height          =   375
            Left            =   1560
            TabIndex        =   23
            Top             =   2280
            Width           =   1215
         End
         Begin VB.CommandButton CmdSimpanProfile 
            Caption         =   "&Simpan"
            Height          =   375
            Left            =   6720
            TabIndex        =   22
            Top             =   2280
            Width           =   1215
         End
         Begin VB.TextBox TextSetNamaPetugas 
            Height          =   375
            IMEMode         =   3  'DISABLE
            Left            =   1560
            MaxLength       =   5
            TabIndex        =   21
            Text            =   "Username"
            Top             =   840
            Width           =   6375
         End
         Begin VB.TextBox TextSetIdentifier 
            Enabled         =   0   'False
            Height          =   375
            IMEMode         =   3  'DISABLE
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   20
            Text            =   "Identifier"
            Top             =   360
            Width           =   6375
         End
         Begin VB.TextBox TextSetPassword 
            Height          =   375
            IMEMode         =   3  'DISABLE
            Left            =   1560
            PasswordChar    =   "*"
            TabIndex        =   19
            Text            =   "Password"
            Top             =   1320
            Width           =   6375
         End
         Begin VB.TextBox TextSetAlamat 
            Height          =   375
            IMEMode         =   3  'DISABLE
            Left            =   1560
            TabIndex        =   18
            Text            =   "Alamat"
            Top             =   1800
            Width           =   6375
         End
         Begin VB.Label LabelUbahNama 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Username"
            Height          =   375
            Left            =   240
            TabIndex        =   27
            Top             =   840
            Width           =   1095
         End
         Begin VB.Label LabelUbahIdenfier 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Identitas"
            Height          =   375
            Left            =   240
            TabIndex        =   26
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label LabelUbahPassword 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Password"
            Height          =   375
            Left            =   240
            TabIndex        =   25
            Top             =   1320
            Width           =   975
         End
         Begin VB.Label LabelUbahAlamat 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Alamat"
            Height          =   375
            Left            =   240
            TabIndex        =   24
            Top             =   1800
            Width           =   975
         End
      End
      Begin VB.Image ImageProfilePetugas 
         Height          =   1935
         Left            =   240
         Picture         =   "FormUtama.frx":39F0
         Stretch         =   -1  'True
         Top             =   720
         Width           =   1695
      End
   End
   Begin VB.Frame FrameAdministrator 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5055
      Left            =   1920
      TabIndex        =   2
      Top             =   3360
      Width           =   10455
      Begin VB.Frame FramePetugas 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Petugas"
         Height          =   4695
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   10215
         Begin VB.TextBox TextIDPetugas 
            Height          =   375
            Left            =   7440
            MaxLength       =   3
            TabIndex        =   10
            Top             =   480
            Width           =   2535
         End
         Begin VB.TextBox TextNamaPetugas 
            Height          =   375
            Left            =   7440
            MaxLength       =   5
            TabIndex        =   9
            Top             =   960
            Width           =   2535
         End
         Begin VB.TextBox TextPasswordPetugas 
            Height          =   375
            IMEMode         =   3  'DISABLE
            Left            =   7440
            PasswordChar    =   "*"
            TabIndex        =   8
            Top             =   1920
            Width           =   2535
         End
         Begin VB.TextBox TextAlamatPetugas 
            Height          =   375
            Left            =   7440
            TabIndex        =   7
            Top             =   1440
            Width           =   2535
         End
         Begin VB.CommandButton CmdTambahPetugas 
            Caption         =   "&Tambah"
            Height          =   375
            Left            =   8760
            TabIndex        =   6
            Top             =   2400
            Width           =   1215
         End
         Begin VB.CommandButton CmdUbahPetugas 
            Caption         =   "&Ubah"
            Height          =   375
            Left            =   4800
            TabIndex        =   5
            Top             =   4080
            Width           =   1215
         End
         Begin VB.CommandButton CmdHapusPetugas 
            Caption         =   "&Hapus"
            Height          =   375
            Left            =   240
            TabIndex        =   4
            Top             =   4080
            Width           =   1215
         End
         Begin MSComctlLib.ListView ListAdministrator 
            Height          =   3495
            Left            =   240
            TabIndex        =   11
            Top             =   480
            Width           =   5775
            _ExtentX        =   10186
            _ExtentY        =   6165
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   0
            BackColor       =   16777215
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   0
         End
         Begin VB.Label LabelIDPetugas 
            BackStyle       =   0  'Transparent
            Caption         =   "Identitas"
            Height          =   375
            Left            =   6240
            TabIndex        =   15
            Top             =   480
            Width           =   1095
         End
         Begin VB.Label LabelNamaPetugas 
            BackStyle       =   0  'Transparent
            Caption         =   "Username"
            Height          =   375
            Left            =   6240
            TabIndex        =   14
            Top             =   960
            Width           =   1095
         End
         Begin VB.Label LabelPasswordPetugas 
            BackStyle       =   0  'Transparent
            Caption         =   "Password"
            Height          =   375
            Left            =   6240
            TabIndex        =   13
            Top             =   1920
            Width           =   1095
         End
         Begin VB.Label LabelAlamatPetugas 
            BackStyle       =   0  'Transparent
            Caption         =   "Alamat"
            Height          =   375
            Left            =   6240
            TabIndex        =   12
            Top             =   1440
            Width           =   1095
         End
      End
   End
   Begin VB.Frame FrameMenu 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   8295
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1695
      Begin VB.Timer TimerUtama 
         Left            =   120
         Top             =   7800
      End
      Begin VB.Image ImageAbout 
         Height          =   1215
         Left            =   240
         Picture         =   "FormUtama.frx":741B
         Stretch         =   -1  'True
         Top             =   6720
         Width           =   1215
      End
      Begin VB.Image ImageAdministrator 
         Height          =   1215
         Left            =   240
         Picture         =   "FormUtama.frx":B16E
         Stretch         =   -1  'True
         Top             =   5160
         Width           =   1215
      End
      Begin VB.Image ImageDetailKembali 
         Height          =   1215
         Left            =   240
         Picture         =   "FormUtama.frx":EC8B
         Stretch         =   -1  'True
         Top             =   3480
         Width           =   1215
      End
      Begin VB.Image ImageDetailPinjam 
         Height          =   1215
         Left            =   240
         Picture         =   "FormUtama.frx":11147
         Stretch         =   -1  'True
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Image ImageHead 
         Height          =   1215
         Left            =   240
         Picture         =   "FormUtama.frx":13884
         Stretch         =   -1  'True
         Top             =   360
         Width           =   1215
      End
   End
   Begin MSComctlLib.StatusBar StatusBarUtama 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   8565
      Width           =   12510
      _ExtentX        =   22066
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2822
            MinWidth        =   2822
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Enabled         =   0   'False
            Object.Width           =   12462
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1058
            MinWidth        =   1058
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu Menu 
      Caption         =   "Menu"
      Begin VB.Menu Tools 
         Caption         =   "Tools"
         Begin VB.Menu Refresh 
            Caption         =   "Refresh"
            Shortcut        =   {F5}
         End
         Begin VB.Menu Buku 
            Caption         =   "Buku"
            Shortcut        =   {F4}
         End
         Begin VB.Menu Anggota 
            Caption         =   "Anggota"
            Shortcut        =   {F3}
         End
         Begin VB.Menu Denda 
            Caption         =   "Denda"
            Shortcut        =   {F2}
         End
         Begin VB.Menu Laporan 
            Caption         =   "Laporan"
            Shortcut        =   {F7}
         End
         Begin VB.Menu Log 
            Caption         =   "Log"
            Shortcut        =   {F11}
         End
      End
      Begin VB.Menu Logout 
         Caption         =   "Logout"
      End
      Begin VB.Menu Exit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu Help 
      Caption         =   "Help"
      Begin VB.Menu Guide 
         Caption         =   "Guide"
         Shortcut        =   {F1}
      End
      Begin VB.Menu Panel 
         Caption         =   "Panel"
         Shortcut        =   {F9}
      End
      Begin VB.Menu About 
         Caption         =   "About"
         Shortcut        =   {F12}
      End
   End
End
Attribute VB_Name = "FormUtama"
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

Option Explicit
Dim ID_Pinjam As Integer
Dim ID_PinjamDetail As Integer
Dim EksekusiData As Integer
Dim EksekusiKembali As Integer
Dim DataOtomatis As ListItem

Private Sub About_Click()
FormAbout.Show
End Sub

Private Sub Anggota_Click()
FormAnggota.Show
End Sub

Private Sub CmdCariDetailKembali_Click()
FormLog.Show
FormLog.FramePesanBaru.Visible = False
FormLog.FrameMessage.Visible = False
FormLog.FrameDenda.Visible = False
FormLog.FramePinjam.Visible = False
FormLog.FrameKembali.Visible = True
End Sub

Private Sub CmdCariDetailPinjam_Click()
FormLog.Show
FormLog.FramePesanBaru.Visible = False
FormLog.FrameMessage.Visible = False
FormLog.FrameKembali.Visible = False
FormLog.FrameDenda.Visible = False
FormLog.FramePinjam.Visible = True
End Sub

Private Sub CmdHapusDetailPinjaman_Click()
If ListPinjamDetail.ListItems.Count > 0 Then
    If MsgBox("Hapus Detail Pinjaman Buku Dengan Nomer Detail " & ListPinjamDetail.SelectedItem & " ?", vbQuestion + vbYesNo) = vbYes Then
        Conn.Execute "DELETE FROM tb_pinjam_detail WHERE id_pinjam_detail='" & ListPinjamDetail.SelectedItem.Text & "'"
        Call CmdPinjamDetail_Click
        MsgBox "Data Berhasil Di Hapus", vbInformation, "Data Berhasil Di Hapus"
    End If
Else
    MsgBox "Silahkan Tampilkan Data Detail Terlebih Dahulu", vbInformation, "Tidak Ada Data Detail Di Temukan"
End If
End Sub

Private Sub CmdHapusKembali_Click()
If ListKembali.ListItems.Count > 0 Then
    If MsgBox("Hapus Pengembalian Nomor " & ListKembali.SelectedItem.Text & " ?", vbQuestion + vbYesNo, "Hapus Pengembalian") = vbYes Then
        Conn.Execute "DELETE FROM tb_kembali WHERE id_kembali='" & ListKembali.SelectedItem.Text & "'"
        Call MasukanDataKembali
        IsiKembali ListKembali
        MsgBox "Data Berhasil Di Hapus", vbInformation, "Data Berhasil Di Hapus"
    End If
Else
    MsgBox "Data Kembali Tidak Tersedia", vbInformation, "Data Tidak Tersedia"
End If
End Sub

Private Sub CmdHapusListData_Click()
ListData.ListItems.Clear
TextJumlahBuku.Text = "Jumlah"

Koneksi "SELECT nama_buku FROM tb_buku WHERE id_buku='" & ComboIDBuku.Text & "'"
TextNamaBuku.Text = DB!nama_Buku
End Sub

Private Sub CmdHapusPetugas_Click()
Dim EksekusiPanpus As String

Koneksi "SELECT * FROM tb_petugas"
If DB.RecordCount = 0 Then
    MsgBox "Tidak Ada Petugas" & vbNewLine & "Aplikasi Akan Membuat Petugas Baru Secara Otomatis", vbInformation, "Petugas"
    Conn.Execute "INSERT INTO tb_petugas(id_petugas,nama_petugas,alamat_petugas,'password_petugas') VALUES('999','root','Aktivitas Hacking Terdeteksi','toor')"
ElseIf DB.RecordCount <= 2 Then
    MsgBox "Harus Terdapat Minimal 2 Petugas", vbExclamation, "Petugas"
Else
    If ListAdministrator.SelectedItem.Text = "1" Or ListAdministrator.SelectedItem.Text = "999" Or ListAdministrator.SelectedItem.Text = "7" Then
        MsgBox "Petugas Dengan Identifier Khusus Tidak Dapat Di Hapus", vbInformation, "Tidak Dapat Menghapus"
    ElseIf ListAdministrator.SelectedItem.SubItems(1) = "admin" Or ListAdministrator.SelectedItem.SubItems(1) = "root" Then
        MsgBox "Petugas Dengan Username Khusus Tidak Dapat Di Hapus", vbInformation, "Tidak Dapat Menghapus"
    Else
        If MsgBox("Hapus Petugas " & ListAdministrator.SelectedItem.SubItems(1) & " ?", vbExclamation + vbYesNo, "Hapus Petugas") = vbYes Then
            Conn.Execute "DELETE FROM tb_petugas WHERE id_petugas='" & ListAdministrator.SelectedItem.Text & "'"
            Call MasukanAdministrator
            MsgBox "Berhasil Menghapus Petugas", vbInformation, "Berhasil Menghapus"
            If Dir$(App.Path & "\Perpustakaan.exe") <> "" Then
                EksekusiPanpus = Shell(App.Path & "\Perpustakaan.exe", vbNormalFocus)
            Else
                MsgBox "Panda Pustaka Tidak Tersedia" & vbNewLine & "Hubungi Administrator", vbCritical, "Panpus Tidak Tersedia"
                End
            End If
            End
        Else
            Exit Sub
        End If
    End If
End If
End Sub

Private Sub CmdHapusPinjaman_Click()
If ListPinjam.ListItems.Count > 0 Then
    If MsgBox("Hapus Data Pinjam Dengan Nomer Pinjam " & ListPinjam.SelectedItem.Text & " ?", vbQuestion + vbYesNo, "Hapus Data Pinjam") = vbYes Then
        Conn.Execute "DELETE FROM tb_pinjam WHERE id_pinjam='" & ListPinjam.SelectedItem.Text & "'"
        MasukanDataPinjam
        ListPinjamDetail.ListItems.Clear
        Call IDPinjam
        Call IDPinjamDetail
        Call IsiPinjamDetail
        MsgBox "Data Pinjam Berhasil Di Hapus", vbInformation, "Data Pinjam Berhasil Di Hapus"
    End If
Else
    MsgBox "Tidak Ada Data Pinjam Yang Bisa Di Hapus", vbCritical, "Data Pinjam Tidak Tersedia"
End If
End Sub

Private Sub CmdKembaliProfile_Click()
FrameChangeProfile.Visible = False
FrameViewProfile.Visible = True
End Sub

Private Sub CmdPinjam_Click()
If ListData.ListItems.Count = 0 Then
    MsgBox "Pilih Data Buku Yang Akan Di Pinjam", vbCritical, "Masukan Data Buku"
ElseIf ListData.ListItems.Count > 0 Then
    For EksekusiData = 1 To ListData.ListItems.Count
        Conn.Execute "INSERT INTO tb_pinjam_detail(id_pinjam, id_buku, nama_buku, jumlah_buku, tanggal_pinjam) VALUES('" & ListData.ListItems.Item(EksekusiData).SubItems(1) & "','" & ListData.ListItems.Item(EksekusiData).SubItems(2) & "','" & ListData.ListItems.Item(EksekusiData).SubItems(3) & "','" & ListData.ListItems.Item(EksekusiData).SubItems(4) & "','" & StatusBarUtama.Panels(1) & "')"
        Conn.Execute "UPDATE tb_buku SET tb_buku.jumlah_buku=tb_buku.jumlah_buku - '" & ListData.ListItems.Item(EksekusiData).SubItems(4) & "' WHERE tb_buku.id_buku='" & ListData.ListItems.Item(EksekusiData).SubItems(2) & "'"
    Next EksekusiData
    Call Form_Load
    ListData.ListItems.Clear
    TextNamaBuku.Text = ""
    TextJumlahBuku.Text = ""
    MsgBox "Buku Berhasil Di Pinjamkan", vbInformation, "Data Berhasil Di Massukan"
Else
    MsgBox "Silahkan Tambahkan Data Buku Yang Akan Di Pinjam", vbInformation, "Tambahkan Data Buku"
End If
End Sub

Private Sub CmdPinjamDetail_Click()
If ListPinjam.ListItems.Count > 0 Then
    If DB.RecordCount > 0 Then
        Koneksi "SELECT id_pinjam_detail, id_buku, nama_buku, jumlah_buku, tanggal_pinjam, status_pinjam_detail FROM tb_pinjam_detail  WHERE id_pinjam='" & ListPinjam.SelectedItem.Text & "' AND status_pinjam_detail='Pinjam'"
        Call IsiPinjamDetail
    Else
        ListPinjamDetail.ListItems.Clear
    End If
Else
    MsgBox "Tidak Terdapat Data Pinjam", vbInformation, "Data Pinjam Tidak Tersedia"
End If
End Sub

Private Sub CmdSelesaiPinjam_Click()
Call CmdPinjamDetail_Click
If ListPinjam.ListItems.Count > 0 Then
    If MsgBox("Apakah Semua Buku Di Pinjaman Nomer " & ListPinjam.SelectedItem.Text & " Sudah Di Kembalikan ?", vbQuestion + vbYesNo) = vbYes Then
        If ListPinjamDetail.ListItems.Count = 0 Then
            Conn.Execute "UPDATE tb_pinjam SET status_pinjam='Kembali' WHERE id_pinjam='" & ListPinjam.SelectedItem.Text & "'"
            MasukanDataPinjam
            ListPinjamDetail.ListItems.Clear
            MsgBox "Buku Berhasil Di Kembalikan", vbInformation, "Buku Berhasil Di Kembalikan"
        Else
            MsgBox "Masih Ada Buku Yang Belum Di Kembalikan", vbCritical, "Buku Belum Di Kembalikan"
        End If
    End If
Else
    MsgBox "Tidak Ada Data Pinjam Tersedia", vbInformation, "Tidak Ada Data Pinjam Tersedia"
End If
End Sub

Private Sub CmdSimpanProfile_Click()
Dim EksekusiPanpus As String

TextSetIdentifier = FilterInjeksi(TextSetIdentifier.Text)
TextSetNamaPetugas = FilterInjeksi(TextSetNamaPetugas.Text)
TextSetAlamat = FilterInjeksi(TextSetAlamat.Text)
TextSetPassword = FilterInjeksi(TextSetPassword.Text)

If TextSetIdentifier.Text = "" And TextSetNamaPetugas.Text = "" And TextSetAlamat.Text = "" And TextSetPassword.Text = "" Then
    MsgBox "Masukan Data Dengan Benar", vbExclamation, "Masukan Data Dengan Benar"
Else
    If MsgBox("Ubah Data Petugas ?" & vbNewLine & "Pastikan Data Petugas Benar" & vbNewLine & "Kesalahan Data Petugas Dapat Menyebabkan Kegagalan Login", vbExclamation + vbYesNo, "Ubah Data Petugas") = vbYes Then
        Conn.Execute "UPDATE tb_petugas SET nama_petugas='" & TextSetNamaPetugas & "', alamat_petugas='" & TextSetAlamat & "', password_petugas='" & TextSetPassword & "' WHERE id_petugas='" & LabelProfileIdentifier.Caption & "'"
        MsgBox "Berhasil Mengubah Data Petugas " & TextNamaPetugas.Text & "", vbInformation, "Berhasil Mengubah Data Petugas"
        If Dir$(App.Path & "\Perpustakaan.exe") <> "" Then
            EksekusiPanpus = Shell(App.Path & "\Perpustakaan.exe", vbNormalFocus)
        Else
            MsgBox "Panda Pustaka Tidak Tersedia" & vbNewLine & "Hubungi Administrator", vbCritical, "Panpus Tidak Tersedia"
            End
        End If
        End
    Else
        Call CmdKembaliProfile_Click
    End If
End If
End Sub

Private Sub CmdTambahPetugas_Click()
TextIDPetugas = FilterInjeksi(TextIDPetugas.Text)
TextNamaPetugas = FilterInjeksi(TextNamaPetugas.Text)
TextAlamatPetugas = FilterInjeksi(TextAlamatPetugas.Text)
TextPasswordPetugas = FilterInjeksi(TextPasswordPetugas.Text)

If TextIDPetugas.Text = "" Or 0 Or TextNamaPetugas.Text = "" Or TextAlamatPetugas.Text = "" Or TextPasswordPetugas.Text = "" Then
    MsgBox "Masukan Data Petugas Dengan Benar", vbExclamation, "Masukan Data Dengan Benar"
Else
    Koneksi "SELECT * FROM tb_petugas WHERE id_petugas='" & TextIDPetugas & "' OR nama_petugas='" & TextNamaPetugas & "'"
    If DB.RecordCount = 0 Then
        Conn.Execute "INSERT INTO tb_petugas(id_petugas,nama_petugas,alamat_petugas,password_petugas) VALUES('" & TextIDPetugas.Text & "','" & TextNamaPetugas.Text & "','" & TextAlamatPetugas.Text & "','" & TextPasswordPetugas.Text & "')"
        Call MasukanAdministrator
        TextIDPetugas.Text = ""
        TextNamaPetugas.Text = ""
        TextAlamatPetugas.Text = ""
        TextPasswordPetugas.Text = ""
        MsgBox "Petugas Berhasil Di Tambahkan", vbInformation, "Berhasil Di Tambahkan"
    Else
        MsgBox "ID Atau Username Sudah Di Gunakan", vbExclamation, "Sudah Di Gunakan"
    End If
End If
End Sub

Private Sub CmdUbahPetugas_Click()
Dim EksekusiPanpus As String

If CmdUbahPetugas.Caption = "&Ubah" Then
    CmdUbahPetugas.Caption = "&Simpan"
    Koneksi "SELECT * FROM tb_petugas WHERE id_petugas='" & ListAdministrator.SelectedItem.Text & "'"
    TextIDPetugas.Text = DB!id_petugas
    TextNamaPetugas.Text = DB!nama_petugas
    TextAlamatPetugas.Text = DB!alamat_petugas
    TextPasswordPetugas.Text = DB!password_petugas
    TextIDPetugas.Enabled = False
    CmdHapusPetugas.Enabled = False
    CmdTambahPetugas.Enabled = False
    TextNamaPetugas.SetFocus
ElseIf CmdUbahPetugas.Caption = "&Simpan" Then
    CmdUbahPetugas.Caption = "&Ubah"
    TextIDPetugas.Enabled = True
    CmdHapusPetugas.Enabled = True
    CmdTambahPetugas.Enabled = True
    If MsgBox("Ubah Data Petugas ?", vbInformation + vbYesNo, "Ubah Data Petugas") = vbYes Then
        TextIDPetugas = FilterInjeksi(TextIDPetugas.Text)
        TextNamaPetugas = FilterInjeksi(TextNamaPetugas.Text)
        TextAlamatPetugas = FilterInjeksi(TextAlamatPetugas.Text)
        TextPasswordPetugas = FilterInjeksi(TextPasswordPetugas.Text)
        Conn.Execute "UPDATE tb_petugas SET nama_petugas='" & TextNamaPetugas & "',alamat_petugas='" & TextAlamatPetugas & "',password_petugas='" & TextPasswordPetugas & "' WHERE id_petugas='" & TextIDPetugas.Text & "'"
        TextIDPetugas.Text = ""
        TextNamaPetugas.Text = ""
        TextAlamatPetugas.Text = ""
        TextPasswordPetugas.Text = ""
        Call MasukanAdministrator
        MsgBox "Data Petugas Berhasil Di Ubah", vbInformation, "Data Berhasil Di Ubah"
        If Dir$(App.Path & "\Perpustakaan.exe") <> "" Then
            EksekusiPanpus = Shell(App.Path & "\Perpustakaan.exe", vbNormalFocus)
        Else
            MsgBox "Panda Pustaka Tidak Tersedia" & vbNewLine & "Hubungi Administrator", vbCritical, "Panpus Tidak Tersedia"
            End
        End If
        End
    Else
        Exit Sub
    End If
Else
    MsgBox "Aktivitas Hacking Terdeteksi", vbExclamation, "Aktivitas Hacking"
    End
End If
End Sub

Private Sub CmdUbahProfile_Click()
FrameViewProfile.Visible = False
FrameChangeProfile.Visible = True
Call LoadProfilePetugas
End Sub

Private Sub LoadProfilePetugas()
Koneksi "SELECT * FROM tb_petugas WHERE id_petugas='" & StatusBarUtama.Panels(5).Text & "'"

TextSetIdentifier.Text = DB!id_petugas
TextSetNamaPetugas = DB!nama_petugas
TextSetAlamat = DB!alamat_petugas
TextSetPassword = DB!password_petugas
End Sub

Private Sub ComboIDAnggota_Click()
Koneksi "SELECT * FROM tb_anggota WHERE id_anggota='" & ComboIDAnggota.Text & "'"
Do While Not DB.EOF
    ComboNamaAnggota.Text = DB!nama_anggota
    DB.MoveNext
Loop
End Sub

Private Sub ComboIDBuku_Click()
Koneksi "SELECT nama_buku FROM tb_buku WHERE id_buku='" & ComboIDBuku.Text & "'"
Do While Not DB.EOF
    ComboIDBuku.ToolTipText = DB!nama_Buku
    TextNamaBuku.Text = DB!nama_Buku
    DB.MoveNext
Loop
End Sub

Private Sub ComboIDPinjamDetail_Click()
Koneksi "SELECT nama_anggota FROM tb_pinjam WHERE id_pinjam='" & ComboIDPinjamDetail.Text & "'"
Do While Not DB.EOF
    ComboIDPinjamDetail.ToolTipText = "Pinjaman " & DB!nama_anggota & ""
    DB.MoveNext
Loop
End Sub

Private Sub Form_Resize()
On Error Resume Next
If Me.WindowState = 1 Then Exit Sub

With FormUtama
    .Height = 9700
    .ScaleHeight = 589
    .ScaleWidth = 834
    .Width = 12700
    .ScaleMode = 3
End With
End Sub

Private Sub ImageAbout_Click()
FormAbout.Show
End Sub

Private Sub ImageAdministrator_Click()
FrameHead.Visible = False
FrameDetailKembali.Visible = False
FrameDetailPinjam.Visible = False
FrameProfilePetugas.Visible = True
FrameAdministrator.Visible = True
End Sub

Private Sub ImageDetailKembali_Click()
FrameAdministrator.Visible = False
FrameHead.Visible = False
FrameDetailPinjam.Visible = False
FrameProfilePetugas.Visible = False
FrameDetailKembali.Visible = True
End Sub

Private Sub ImageDetailPinjam_Click()
FrameAdministrator.Visible = False
FrameHead.Visible = False
FrameDetailKembali.Visible = False
FrameProfilePetugas.Visible = False
FrameDetailPinjam.Visible = True
End Sub

Private Sub ImageHead_Click()
FrameAdministrator.Visible = False
FrameDetailKembali.Visible = False
FrameDetailPinjam.Visible = False
FrameProfilePetugas.Visible = False
FrameHead.Visible = True
End Sub


Private Sub Laporan_Click()
FormReport.Show
End Sub

Private Sub Log_Click()
FormLog.Show
End Sub

Private Sub Panel_Click()
FormPanel.Show
End Sub

Private Sub Refresh_Click()
Call Form_Load
End Sub

Private Sub Buku_Click()
FormBuku.Show
End Sub

Private Sub CmdDenda_Click()
Koneksi "SELECT tanggal_kembali, tanggal_pinjam FROM tb_kembali"
If ListKembali.ListItems.Count = 0 Then
    MsgBox "Denda Tidak Dapat Di Berikan", vbInformation, "Tidak Ada Data Yang Di Kembalikan"
ElseIf DB!tanggal_kembali - 3 < DB!tanggal_pinjam Then
    MsgBox "Peminjaman Buku Berjalan Dengan Baik", vbInformation, "Tidak Terdeteksi Keterlambatan Peminjaman"
Else
    FormDenda.Show
End If
End Sub

Private Sub CmdKembalikan_Click()
If ListPinjam.ListItems.Count = 0 Then
    MsgBox "Tidak Ada Data Tersedia", vbInformation, "Tidak Ada Data Tersedia"
ElseIf ListPinjamDetail.ListItems.Count > 0 Then
    If MsgBox("Kembalikan Buku " & "Dengan Nomer Detail " & ListPinjamDetail.SelectedItem.Text & " ?", vbQuestion + vbYesNo, "Kembalikan Buku?") = vbYes Then
        ID_PinjamDetail = ListPinjamDetail.SelectedItem.Text
        Conn.Execute "INSERT INTO tb_kembali(id_pinjam_detail, id_anggota, nama_petugas, nama_anggota, nama_buku, jumlah_buku, tanggal_pinjam, tanggal_kembali) VALUES('" & ListPinjamDetail.SelectedItem.Text & "','" & ListPinjam.SelectedItem.SubItems(1) & "','" & StatusBarUtama.Panels(4) & "','" & ListPinjam.SelectedItem.SubItems(2) & "','" & ListPinjamDetail.SelectedItem.SubItems(2) & "','" & ListPinjamDetail.SelectedItem.SubItems(3) & "','" & ListPinjamDetail.SelectedItem.SubItems(4) & "','" & StatusBarUtama.Panels(1) & "')"
        Conn.Execute "UPDATE tb_pinjam_detail SET status_pinjam_detail='Kembali' WHERE id_pinjam_detail='" & ID_PinjamDetail & "'"
        Conn.Execute "UPDATE tb_buku SET jumlah_buku=jumlah_buku + '" & ListPinjamDetail.SelectedItem.SubItems(3) & "' WHERE id_buku='" & ListPinjamDetail.SelectedItem.SubItems(1) & "'"
        MasukanDataPinjam
        MasukanDataKembali
        Call CmdPinjamDetail_Click
        MsgBox "Data Buku Berhasil Di Kembalikan", vbInformation, "Data Buku Berhasil Di Kembalikan"
    End If
Else
    MsgBox "Silahkan Pilih Detail Buku Yang Akan Di Kembalikan", vbInformation, "Pilih Detail Buku"
End If
End Sub


Private Sub CmdPinjamkan_Click()
If ComboIDAnggota.ListCount = 0 Then
    MsgBox "Silahkan Tambahkan Anggota Panda Pustaka Terlebih Dahulu", vbInformation, "Tidak Ada Anggota Di Temukan"
    FormAnggota.Show
Else
    If ComboNamaAnggota.Text = "" Or ComboIDAnggota.Text = "" Then
        MsgBox "Silahkan Isi Keterangan Pinjaman Dengan Benar", vbCritical, "Data Belum Di Isi"
    ElseIf ComboNamaAnggota.Text = "Nama Anggota" Or ComboIDAnggota.Text = "ID Anggota" Then
        MsgBox "Silahkan Isi Keterangan Pinjaman", vbInformation, "Data Belum Di Isi"
    Else
        Conn.Execute "INSERT INTO tb_pinjam (id_anggota, tanggal_pinjam, nama_petugas, nama_anggota, status_pinjam) VALUES('" & ComboIDAnggota.Text & "','" & StatusBarUtama.Panels(1).Text & "','" & StatusBarUtama.Panels(4) & "','" & ComboNamaAnggota.Text & "','Pinjam')"
        MasukanDataPinjam
        Call IDPinjam
        Call IDPinjamDetail
        ListPinjamDetail.ListItems.Clear
        MsgBox "Data Berhasil Di Masukan", vbInformation, "Data Berhasil Di Masukan"
    End If
End If
End Sub

Private Sub CmdTambah_Click()
Dim JumlahBukuDiPinjam As Integer

If ComboIDPinjamDetail.ListCount = 0 Then
    MsgBox "Silahkan Tambahkan Keterangan Peminjaman Terlebih Dahulu", vbInformation, "Tidak Ada Pinjaman Di Lakukan"
    ComboIDAnggota.SetFocus
Else
    If TextNamaBuku.Text = "" Or TextJumlahBuku.Text = "" Then
        MsgBox "Silahkan Isikan Data", vbCritical, "Data Masih Kosong"
    ElseIf TextNamaBuku.Text = "Nama Buku" Or TextJumlahBuku.Text = "Jumlah" Or TextJumlahBuku.Text = "0" Then
        MsgBox "Silahkan Isikan Data", vbCritical, "Pilih Data Yang Tersedia"
    Else
        If ListData.ListItems.Count = 0 Then
            Koneksi "SELECT jumlah_buku FROM tb_buku WHERE id_buku='" & ComboIDBuku.Text & "'"
            If DB!jumlah_Buku < Val(TextJumlahBuku.Text) Then
                MsgBox "Stok Buku Tidak Mencukupi Untuk Di Pinjam", vbExclamation, "Tidak Dapat Di Pinjamkan"
            Else
                Set DataOtomatis = ListData.ListItems.Add
                    DataOtomatis.SubItems(1) = ComboIDPinjamDetail.Text
                    DataOtomatis.SubItems(2) = ComboIDBuku.Text
                    DataOtomatis.SubItems(3) = TextNamaBuku.Text
                    DataOtomatis.SubItems(4) = TextJumlahBuku.Text
                    MsgBox "Data Berhasil Di Tambahkan", vbInformation, "Data Berhasil Di Tambahkan"
                    CmdHapusListData.Enabled = True
            End If
        Else
        For EksekusiData = 1 To ListData.ListItems.Count
        If ListData.ListItems.Count > 0 Then
            If ComboIDPinjamDetail.Text <> ListData.ListItems.Item(EksekusiData).SubItems(1) Then
                MsgBox "Setiap Transaksi Hanya Dapat Di Lakukan Untuk 1 ID", vbExclamation, "Tidak Boleh Beda ID"
            Else
                If ComboIDBuku.Text = ListData.ListItems.Item(EksekusiData).SubItems(2) Then
                    If MsgBox("Terdeteksi Duplikasi Data" & vbNewLine & "Hapus Data Sebelumnya ?", vbExclamation + vbYesNo, "Terdeteksi Duplikasi Data") = vbYes Then
                        ListData.ListItems.Clear
                    Else
                        Exit Sub
                    End If
                Else
                    Koneksi "SELECT jumlah_buku FROM tb_buku WHERE id_buku='" & ComboIDBuku.Text & "'"
                    If DB!jumlah_Buku < Val(TextJumlahBuku.Text) Then
                        MsgBox "Stok Buku Tidak Mencukupi Untuk Di Pinjam", vbExclamation, "Tidak Dapat Di Pinjamkan"
                    Else
                    Set DataOtomatis = ListData.ListItems.Add
                        DataOtomatis.SubItems(1) = ComboIDPinjamDetail.Text
                        DataOtomatis.SubItems(2) = ComboIDBuku.Text
                        DataOtomatis.SubItems(3) = TextNamaBuku.Text
                        DataOtomatis.SubItems(4) = TextJumlahBuku.Text
                        MsgBox "Data Berhasil Di Tambahkan", vbInformation, "Data Berhasil Di Tambahkan"
                        CmdHapusListData.Enabled = True
                    End If
                End If
            End If
        End If
        Next EksekusiData
        End If
    End If
End If
End Sub

Private Sub Denda_Click()
Call CmdDenda_Click
End Sub

Private Sub Exit_Click()
If MsgBox("" & StatusBarUtama.Panels(4) & ", Apakah Anda Yakin Akan Keluar Dari Panda Pustaka ?", vbExclamation + vbYesNo, "Keluar Dari Panda Pustaka?") = vbYes Then
    Unload FormAbout
    Unload FormAnggota
    Unload FormBuku
    Unload FormDenda
    Unload FormLog
    Unload FormPanduan
    Unload FormPanel
    Unload FormPopUp
    Unload FormReport
    End
End If
End Sub

Private Sub Form_Load()
Call Default
Call IDBuku
Call MasukanDataPinjam
Call MasukanDataKembali
Call MasukanIDAnggota
Call IDPinjamDetail
Call MasukanAdministrator
Call ImageHead_Click

ListPinjam.FullRowSelect = True
ListPinjam.GridLines = True
ListPinjam.LabelEdit = lvwManual

ListKembali.FullRowSelect = True
ListKembali.GridLines = True
ListKembali.LabelEdit = lvwManual

ListData.FullRowSelect = True
ListData.GridLines = True
ListData.LabelEdit = lvwManual

ListPinjamDetail.FullRowSelect = True
ListPinjamDetail.GridLines = True
ListPinjamDetail.LabelEdit = lvwManual

FormUtama.TimerUtama.Enabled = True
FormUtama.TimerUtama.Interval = 100
End Sub

Private Sub TextIDPetugas_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= Asc("0") & Chr(13) _
     And KeyAscii <= Asc("9") & Chr(13) _
     Or KeyAscii = vbKeyBack _
     Or KeyAscii = vbKeyDelete _
     Or KeyAscii = vbKeySpace) Then
    TextIDPetugas.Text = "0"
    KeyAscii = 0
End If
End Sub

Private Sub Form_Terminate()
Call Exit_Click
End Sub

Private Sub MasukanAdministrator()
Koneksi "SELECT id_petugas, nama_petugas, alamat_petugas FROM tb_petugas ORDER BY id_petugas ASC"
IsiPetugas ListAdministrator
End Sub

Private Sub Form_Unload(Cancel As Integer)
If MsgBox("" & StatusBarUtama.Panels(4) & ", Apakah Anda Yakin Akan Keluar Dari Panda Pustaka ?", vbExclamation + vbYesNo, "Keluar Dari Panda Pustaka?") = vbYes Then
    Unload FormAbout
    Unload FormAnggota
    Unload FormBuku
    Unload FormDenda
    Unload FormLog
    Unload FormPanduan
    Unload FormPanel
    Unload FormPopUp
    Unload FormReport
    End
Else
    Cancel = 1
End If
End Sub

Private Sub Guide_Click()
FormPanduan.Show
End Sub

Private Sub Logout_Click()
If MsgBox("" & StatusBarUtama.Panels(4) & ", Apakah Anda Yakin Akan Logout Dari Panda Pustaka ?", vbQuestion + vbYesNo, "Logout Dari Panda Pustaka?") = vbYes Then
    FormUtama.Hide
    Unload FormAbout
    Unload FormAnggota
    Unload FormBuku
    Unload FormDenda
    Unload FormLog
    Unload FormPanduan
    Unload FormPanel
    Unload FormPopUp
    Unload FormReport
    FormLogin.Show
End If
End Sub

Private Sub TextJumlahBuku_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= Asc("0") & Chr(13) _
     And KeyAscii <= Asc("9") & Chr(13) _
     Or KeyAscii = vbKeyBack _
     Or KeyAscii = vbKeyDelete _
     Or KeyAscii = vbKeySpace) Then
    TextJumlahBuku.Text = "0"
    KeyAscii = 0
End If
End Sub

Private Sub TextJumlahBuku_Click()
If TextJumlahBuku.Text = "Jumlah" Then
    TextJumlahBuku.Text = ""
ElseIf TextJumlahBuku.Text = "" Then
    TextJumlahBuku.Text = "Jumlah"
End If
End Sub

Private Sub TimerUtama_Timer()
StatusBarUtama.Panels(1) = Format(Date, "YYYY-MM-DD")
StatusBarUtama.Panels(2) = Format(Time, "HH:MM:SS")
End Sub
