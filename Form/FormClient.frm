VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FormClient 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Panda Pustaka"
   ClientHeight    =   7230
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10935
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormClient.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7230
   ScaleMode       =   0  'User
   ScaleWidth      =   10935
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdHistory 
      Caption         =   "&History"
      Height          =   375
      Left            =   5160
      TabIndex        =   60
      Top             =   6480
      Width           =   975
   End
   Begin VB.Frame FrameHistory 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&History"
      Height          =   4455
      Left            =   120
      TabIndex        =   58
      Top             =   1920
      Width           =   10695
      Begin VB.Frame FramePemberitahuanPeminjaman 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Pemberitahuan Peminjaman"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1815
         Left            =   7560
         TabIndex        =   63
         Top             =   240
         Width           =   3015
         Begin VB.TextBox TextIsiPemberitahuan 
            Height          =   1455
            Left            =   120
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   64
            Text            =   "FormClient.frx":0CCA
            Top             =   240
            Width           =   2775
         End
      End
      Begin VB.TextBox TextSearchHistory 
         Height          =   315
         Left            =   7560
         MaxLength       =   251
         ScrollBars      =   1  'Horizontal
         TabIndex        =   62
         Top             =   2520
         Width           =   3015
      End
      Begin VB.ComboBox ComboKategoriHistory 
         Height          =   330
         Left            =   7560
         Style           =   2  'Dropdown List
         TabIndex        =   61
         Top             =   2160
         Width           =   3015
      End
      Begin MSComctlLib.ListView ListHistory 
         Height          =   3855
         Left            =   120
         TabIndex        =   59
         Top             =   360
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   6800
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Image ImageHistoryClient 
         Height          =   1215
         Left            =   8520
         Picture         =   "FormClient.frx":0CDC
         Stretch         =   -1  'True
         Top             =   3000
         Width           =   1815
      End
   End
   Begin VB.CommandButton CmdPanduan 
      Caption         =   "&Panduan"
      Height          =   375
      Left            =   6360
      TabIndex        =   57
      Top             =   6480
      Width           =   975
   End
   Begin VB.Frame FrameAnggota 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Anggota"
      Height          =   4455
      Left            =   120
      TabIndex        =   51
      Top             =   1920
      Width           =   10695
      Begin VB.ComboBox ComboPetugasAnggota 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   120
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   55
         Top             =   360
         Width           =   2535
      End
      Begin VB.Frame FrameCariPetugasAnggota 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Cari "
         Height          =   1335
         Left            =   120
         TabIndex        =   52
         Top             =   720
         Width           =   2535
         Begin VB.ComboBox ComboCariPetugasAnggota 
            Height          =   330
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   54
            Top             =   360
            Width           =   2295
         End
         Begin VB.TextBox TextCariPetugasAnggota 
            Height          =   315
            Left            =   120
            TabIndex        =   53
            Top             =   840
            Width           =   2295
         End
      End
      Begin MSComctlLib.ListView ListPetugasAnggota 
         Height          =   3855
         Left            =   2760
         TabIndex        =   56
         Top             =   360
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   6800
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
      Begin VB.Image ImageCariAnggota 
         Height          =   1335
         Left            =   360
         Picture         =   "FormClient.frx":3F37
         Stretch         =   -1  'True
         Top             =   2520
         Width           =   2175
      End
   End
   Begin VB.CommandButton CmdAnggota 
      Caption         =   "&Anggota"
      Height          =   375
      Left            =   1440
      TabIndex        =   50
      Top             =   6480
      Width           =   1095
   End
   Begin VB.CommandButton CmdLogout 
      Caption         =   "&Logout"
      Height          =   375
      Left            =   9840
      TabIndex        =   47
      Top             =   6480
      Width           =   975
   End
   Begin VB.CommandButton CmdMessage 
      Caption         =   "&Message"
      Height          =   375
      Left            =   3960
      TabIndex        =   35
      Top             =   6480
      Width           =   975
   End
   Begin VB.CommandButton CmdRequest 
      Caption         =   "&Request"
      Height          =   375
      Left            =   2760
      TabIndex        =   34
      Top             =   6480
      Width           =   975
   End
   Begin VB.CommandButton CmdBuku 
      Caption         =   "&Buku"
      Height          =   375
      Left            =   120
      TabIndex        =   33
      Top             =   6480
      Width           =   1095
   End
   Begin VB.Frame FrameMessage 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Message "
      Height          =   4455
      Left            =   120
      TabIndex        =   25
      Top             =   1920
      Width           =   10695
      Begin VB.CommandButton CmdRefreshMessage 
         Caption         =   "&Refresh"
         Height          =   375
         Left            =   9600
         TabIndex        =   37
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton CmdHapus 
         Caption         =   "&Hapus"
         Height          =   375
         Left            =   9600
         TabIndex        =   32
         Top             =   1200
         Width           =   975
      End
      Begin VB.CommandButton CmdBalas 
         Caption         =   "&Balas"
         Height          =   375
         Left            =   9600
         TabIndex        =   31
         Top             =   3960
         Width           =   975
      End
      Begin VB.TextBox TextMessage 
         Height          =   1695
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   30
         Top             =   2160
         Width           =   10455
      End
      Begin VB.CommandButton CmdLihat 
         Caption         =   "&Lihat"
         Height          =   375
         Left            =   9600
         TabIndex        =   29
         Top             =   720
         Width           =   975
      End
      Begin MSComctlLib.ListView ListMessage 
         Height          =   1455
         Left            =   120
         TabIndex        =   28
         Top             =   240
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   2566
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Label LabelKategori 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         ForeColor       =   &H000040C0&
         Height          =   255
         Left            =   1080
         TabIndex        =   49
         Top             =   4080
         Width           =   1455
      End
      Begin VB.Label LabelJudulKategori 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Kategori"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   48
         Top             =   4080
         Width           =   855
      End
      Begin VB.Label LabelJudulWaktu 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Waktu"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5520
         TabIndex        =   46
         Top             =   3840
         Width           =   615
      End
      Begin VB.Label LabelJudulTanggal 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   45
         Top             =   3840
         Width           =   735
      End
      Begin VB.Label LabelJudulPenerima 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Penerima"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5520
         TabIndex        =   44
         Top             =   1920
         Width           =   855
      End
      Begin VB.Label LabelJudulPengirim 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Pengirim"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   43
         Top             =   1920
         Width           =   855
      End
      Begin VB.Label LabelWaktu 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "-"
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
         Left            =   6600
         TabIndex        =   42
         Top             =   3840
         Width           =   2655
      End
      Begin VB.Label LabelPenerima 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   255
         Left            =   6600
         TabIndex        =   41
         Top             =   1920
         Width           =   3975
      End
      Begin VB.Label LabelPengirim 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   1080
         TabIndex        =   40
         Top             =   1920
         Width           =   4215
      End
      Begin VB.Label LabelTanggal 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "-"
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
         Left            =   1080
         TabIndex        =   39
         Top             =   3840
         Width           =   2175
      End
      Begin VB.Label LabelHeaderMessage 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   1680
         Width           =   10455
      End
   End
   Begin VB.Frame FrameRequest 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Request "
      Height          =   4455
      Left            =   120
      TabIndex        =   3
      Top             =   1920
      Width           =   10695
      Begin VB.ComboBox ComboTujuan 
         Height          =   330
         Left            =   7800
         Style           =   2  'Dropdown List
         TabIndex        =   36
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox TextTujuan 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   9360
         TabIndex        =   27
         Top             =   360
         Width           =   1215
      End
      Begin VB.Timer TimerClient 
         Left            =   9120
         Top             =   3960
      End
      Begin VB.CommandButton CmdKirim 
         Caption         =   "&Kirim"
         Height          =   375
         Left            =   9600
         TabIndex        =   24
         Top             =   3960
         Width           =   975
      End
      Begin VB.TextBox TextKonten 
         Height          =   3015
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   22
         Top             =   840
         Width           =   10455
      End
      Begin VB.TextBox TextJudul 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   960
         TabIndex        =   21
         Top             =   360
         Width           =   3975
      End
      Begin VB.Label LabelTujuan 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Tujuan  :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6960
         TabIndex        =   26
         Top             =   360
         Width           =   855
      End
      Begin VB.Label LabelPerihal 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Perihal  :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame FrameBuku 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Buku "
      Height          =   4455
      Left            =   120
      TabIndex        =   2
      Top             =   1920
      Width           =   10695
      Begin VB.TextBox TextCariBuku 
         Height          =   315
         Left            =   3600
         MaxLength       =   250
         TabIndex        =   66
         Top             =   240
         Width           =   6855
      End
      Begin VB.ComboBox ComboCariBuku 
         Height          =   330
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   65
         Top             =   240
         Width           =   3375
      End
      Begin MSComctlLib.ListView ListBuku 
         Height          =   3615
         Left            =   120
         TabIndex        =   20
         Top             =   720
         Width           =   10455
         _ExtentX        =   18441
         _ExtentY        =   6376
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
   Begin VB.Frame FrameProfileAnggota 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Profile "
      Height          =   1695
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   10695
      Begin VB.Label LabelDendaAnggota 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Contoh"
         Height          =   255
         Left            =   7440
         TabIndex        =   19
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label LabelKonfirmasiAnggota 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Contoh"
         Height          =   255
         Left            =   7440
         TabIndex        =   18
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label LabelGabungAnggota 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Contoh"
         Height          =   255
         Left            =   7440
         TabIndex        =   17
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label LabelNISAnggota 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Contoh"
         Height          =   255
         Left            =   7440
         TabIndex        =   16
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label LabelSekolahAnggota 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Contoh"
         Height          =   255
         Left            =   2280
         TabIndex        =   15
         Top             =   1320
         Width           =   3375
      End
      Begin VB.Label LabelJurusanAnggota 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Contoh"
         Height          =   255
         Left            =   2280
         TabIndex        =   14
         Top             =   960
         Width           =   3375
      End
      Begin VB.Label LabelKelasAnggota 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Contoh"
         Height          =   255
         Left            =   2280
         TabIndex        =   13
         Top             =   600
         Width           =   3375
      End
      Begin VB.Label LabelNamaAnggota 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Contoh"
         Height          =   255
         Left            =   2280
         TabIndex        =   12
         Top             =   240
         Width           =   3375
      End
      Begin VB.Label LabelDenda 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Denda"
         Height          =   255
         Left            =   5880
         TabIndex        =   11
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label LabelKonfirmasiPetugas 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Konfirmasi Petugas"
         Height          =   255
         Left            =   5880
         TabIndex        =   10
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label LabelTanggalBergabung 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Bergabung Sejak"
         Height          =   255
         Left            =   5880
         TabIndex        =   9
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label LabelNIS 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "NIS"
         Height          =   255
         Left            =   5880
         TabIndex        =   8
         Top             =   240
         Width           =   495
      End
      Begin VB.Label LabelSekolah 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Sekolah"
         Height          =   255
         Left            =   1560
         TabIndex        =   7
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label LabelJurusan 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Jurusan"
         Height          =   255
         Left            =   1560
         TabIndex        =   6
         Top             =   960
         Width           =   615
      End
      Begin VB.Image ImageAktif 
         Height          =   975
         Left            =   9240
         Picture         =   "FormClient.frx":7D43
         Stretch         =   -1  'True
         Top             =   360
         Width           =   1335
      End
      Begin VB.Image ImageBelumAktif 
         Height          =   975
         Left            =   9240
         Picture         =   "FormClient.frx":91E5
         Stretch         =   -1  'True
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label LabelKelas 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Kelas"
         Height          =   255
         Left            =   1560
         TabIndex        =   5
         Top             =   600
         Width           =   495
      End
      Begin VB.Label LabelNama 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Nama"
         Height          =   255
         Left            =   1560
         TabIndex        =   4
         Top             =   240
         Width           =   615
      End
      Begin VB.Image ImageWanita 
         Height          =   1335
         Left            =   120
         Picture         =   "FormClient.frx":A4E8
         Stretch         =   -1  'True
         Top             =   240
         Width           =   1215
      End
      Begin VB.Image ImagePria 
         Height          =   1305
         Left            =   120
         Picture         =   "FormClient.frx":D832
         Stretch         =   -1  'True
         Top             =   240
         Width           =   1080
      End
   End
   Begin MSComctlLib.StatusBar StatusBarClient 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   6960
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   882
            MinWidth        =   882
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Enabled         =   0   'False
            Object.Width           =   10636
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
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
End
Attribute VB_Name = "FormClient"
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

Private Sub CmdAnggota_Click()
FrameMessage.Visible = False
FrameRequest.Visible = False
FrameBuku.Visible = False
FrameHistory.Visible = False
FrameAnggota.Visible = True
Call CmdRefreshMessage_Click
End Sub

Private Sub CmdBalas_Click()
TextMessage = FilterInjeksi(TextMessage.Text)
If CmdBalas.Caption = "&Balas" Then
    CmdBalas.Caption = "&Kirim"
    TextMessage.Text = ""
    TextMessage.Locked = False
    TextMessage.SetFocus
ElseIf CmdBalas.Caption = "&Kirim" Then
    If TextMessage.Text = "" Then
        MsgBox "Isi Pesan Dengan Benar", vbExclamation, "Pesan Tidak Di Kirim"
    Else
    If MsgBox("Kirim Pesan Ke " & LabelPengirim.Caption & " ?", vbQuestion + vbYesNo, "Kirim Pesan") = vbYes Then
        If LabelKategori.Caption = "Anggota - Anggota" Then
            Koneksi "SELECT id_anggota FROM tb_anggota WHERE id_anggota='" & StatusBarClient.Panels(2).Text & "'"
            Conn.Execute "INSERT INTO tb_request(id_pengirim, nama_pengirim, nama_penerima, id_anggota_tujuan, perihal_request, konten_request, tanggal_request, waktu_request, status_request) VALUES('" & StatusBarClient.Panels(2).Text & "','" & LabelNamaAnggota.Caption & "','" & LabelPengirim.Caption & "','" & DB!ID_Anggota & "','" & LabelHeaderMessage.Caption & "','" & TextMessage & "','" & StatusBarClient.Panels(5).Text & "','" & StatusBarClient.Panels(4).Text & "','Belum Di Baca')"
            MsgBox "Pesan Berhasil Di Kirim Ke " & LabelPengirim.Caption & "", vbInformation, "Berhasil Mengirim Pesan"
            Call CmdRefreshMessage_Click
        ElseIf LabelKategori.Caption = "Petugas - Anggota" Then
            Koneksi "SELECT id_petugas FROM tb_petugas WHERE nama_petugas='" & LabelPengirim.Caption & "'"
            Conn.Execute "INSERT INTO tb_request(id_pengirim, nama_pengirim, nama_penerima, id_petugas_tujuan, perihal_request, konten_request, tanggal_request, waktu_request, status_request) VALUES('" & StatusBarClient.Panels(2).Text & "','" & LabelNamaAnggota.Caption & "','" & LabelPengirim.Caption & "','" & DB!id_petugas & "','" & LabelHeaderMessage.Caption & "','" & TextMessage & "','" & StatusBarClient.Panels(5).Text & "','" & StatusBarClient.Panels(4).Text & "','Belum Di Baca')"
            MsgBox "Pesan Berhasil Di Kirim Ke " & LabelPengirim.Caption & "", vbInformation, "Berhasil Mengirim Pesan"
            Call CmdRefreshMessage_Click
        End If
        CmdBalas.Caption = "&Balas"
    Else
        TextMessage.Text = ""
        TextMessage.Locked = False
    End If
    End If
End If
End Sub

Private Sub CmdBuku_Click()
FrameMessage.Visible = False
FrameRequest.Visible = False
FrameAnggota.Visible = False
FrameHistory.Visible = False
FrameBuku.Visible = True
Call CmdRefreshMessage_Click
End Sub

Private Sub CmdHapus_Click()
If ListMessage.ListItems.Count > 0 Then
    If MsgBox("Hapus Pesan Request " & ListMessage.SelectedItem.ListSubItems(4).Text & " ?", vbQuestion + vbYesNo, "Hapus Pesan Request") = vbYes Then
        Conn.Execute "DELETE FROM tb_request WHERE id_request='" & ListMessage.SelectedItem.Text & "'"
        MsgBox "Pesan Dari " & ListMessage.SelectedItem.ListSubItems(4).Text & "", vbInformation, "Pesan Berhasil Di Hapus"
        Call CmdRefreshMessage_Click
    End If
Else
    MsgBox "Tidak Ada Pesan", vbExclamation, "Tidak Ada Pesan"
End If
End Sub

Private Sub CmdHistory_Click()
FrameMessage.Visible = False
FrameRequest.Visible = False
FrameAnggota.Visible = False
FrameBuku.Visible = False
FrameHistory.Visible = True
End Sub

Private Sub CmdKirim_Click()
TextTujuan = FilterInjeksi(TextTujuan.Text)
TextJudul = FilterInjeksi(TextJudul.Text)
TextKonten = FilterInjeksi(TextKonten.Text)

If ComboTujuan.Text = "" Or TextJudul.Text = "" Or TextTujuan.Text = "" Or TextKonten.Text = "" Then
    MsgBox "Masukan Data Request Dengan Benar", vbExclamation, "Masukan Semua Data Request"
ElseIf ComboTujuan.Text = "Petugas" Then
    Koneksi "SELECT nama_petugas FROM tb_petugas WHERE id_petugas='" & TextTujuan & "'"
        If DB.EOF Then
            MsgBox "ID Petugas Tidak Di Ketahui", vbExclamation, "Masukan ID Petugas Dengan Benar"
        Else
            If MsgBox("Kirim Pesan Ke " & DB!nama_petugas & " ?", vbQuestion + vbYesNo, "Kirim Pesan") = vbYes Then
                Conn.Execute "INSERT INTO tb_request(id_pengirim, nama_pengirim, nama_penerima, id_petugas_tujuan, perihal_request, konten_request, tanggal_request, waktu_request, status_request) VALUES('" & StatusBarClient.Panels(2).Text & "','" & LabelNamaAnggota.Caption & "','" & DB!nama_petugas & "','" & TextTujuan & "','" & TextJudul & "','" & TextKonten & "','" & StatusBarClient.Panels(5).Text & "','" & StatusBarClient.Panels(4).Text & "','Belum Di Baca')"
                MsgBox "Request Telah Di Kirim Ke " & DB!nama_petugas & "", vbInformation, "Request Berhasil Di Kirim"
                TextJudul.Text = ""
                TextKonten.Text = ""
                TextTujuan.Text = ""
                MasukanMessageClient
            End If
        End If
ElseIf ComboTujuan.Text = "Anggota" Then
    Koneksi "SELECT nama_anggota FROM tb_anggota WHERE id_anggota='" & TextTujuan & "'"
        If DB.EOF Then
            MsgBox "ID Anggota Tidak Di Ketahui", vbExclamation, "Masukan ID Anggota Dengan Benar"
        Else
            If MsgBox("Kirim Pesan Ke " & DB!nama_anggota & " ?", vbQuestion + vbYesNo, "Kirim Pesan") = vbYes Then
                Conn.Execute "INSERT INTO tb_request(id_pengirim, nama_pengirim, nama_penerima, id_anggota_tujuan, perihal_request, konten_request, tanggal_request, waktu_request, status_request) VALUES('" & StatusBarClient.Panels(2).Text & "','" & LabelNamaAnggota.Caption & "','" & DB!nama_anggota & "','" & TextTujuan & "','" & TextJudul & "','" & TextKonten & "','" & StatusBarClient.Panels(5).Text & "','" & StatusBarClient.Panels(4).Text & "','Belum Di Baca')"
                MsgBox "Request Telah Di Kirim Ke " & DB!nama_anggota & "", vbInformation, "Request Berhasil Di Kirim"
                TextJudul.Text = ""
                TextKonten.Text = ""
                TextTujuan.Text = ""
                MasukanMessageClient
            End If
        End If
End If
End Sub

Private Sub CmdLihat_Click()
If ListMessage.ListItems.Count > 0 Then
    Koneksi "SELECT id_pengirim, nama_pengirim, nama_penerima, perihal_request, konten_request, tanggal_request, waktu_request FROM tb_request WHERE id_request='" & ListMessage.SelectedItem.Text & "'"
    If DB.RecordCount > 0 Then
        LabelHeaderMessage.Caption = DB!perihal_request
        LabelPengirim.Caption = DB!nama_pengirim
        LabelPenerima.Caption = DB!nama_penerima
        LabelTanggal.Caption = DB!tanggal_request
        LabelWaktu.Caption = DB!waktu_request
        TextMessage.Text = DB!konten_request
        TextMessage.Locked = True
        If DB!id_pengirim <> StatusBarClient.Panels(2).Text Or DB!nama_pengirim <> LabelPengirim.Caption Then
            Conn.Execute "UPDATE tb_request SET status_request='Di Baca' WHERE id_request='" & ListMessage.SelectedItem.Text & "'"
            CmdBalas.Enabled = True
        Else
            CmdBalas.Enabled = False
        End If
        MasukanMessageClient
        Koneksi "SELECT nama_petugas, nama_anggota FROM tb_petugas, tb_anggota WHERE nama_petugas='" & LabelPengirim.Caption & "' OR nama_anggota='" & LabelPengirim.Caption & "'"
        If LabelPengirim.Caption = DB!nama_petugas Then
            LabelKategori.Caption = "Petugas - Anggota"
        ElseIf LabelPengirim.Caption = DB!nama_anggota Then
            LabelKategori.Caption = "Anggota - Anggota"
        End If
    Else
        MsgBox "Tidak Ada Pesan", vbExclamation, "Tidak Ada Pesan"
    End If
Else
    MsgBox "Tidak Ada Pesan", vbExclamation, "Tidak Ada Pesan"
End If
End Sub

Private Sub CmdLogout_Click()
If MsgBox("" & LabelNamaAnggota.Caption & ", Apakah Anda Yakin Akan Logout Dari Panda Pustaka ?", vbQuestion + vbYesNo, "Logout Dari Panda Pustaka?") = vbYes Then
    CmdBalas.Caption = "&Balas"
    Unload FormPopUp
    Unload FormPanduan
    FormClient.Hide
    Set FormClient = Nothing
    FormLogin.Show
End If
End Sub

Private Sub CmdMessage_Click()
FrameBuku.Visible = False
FrameRequest.Visible = False
FrameAnggota.Visible = False
FrameHistory.Visible = False
FrameMessage.Visible = True
End Sub

Private Sub CmdPanduan_Click()
FormPanduan.Show
End Sub

Private Sub CmdRefreshMessage_Click()
CmdBalas.Enabled = False
CmdBalas.Caption = "&Balas"
MasukanMessageClient
LabelHeaderMessage.Caption = ""
LabelPengirim.Caption = "-"
LabelPenerima.Caption = "-"
LabelTanggal.Caption = "-"
LabelWaktu.Caption = "-"
LabelKategori.Caption = "-"
TextMessage.Text = ""
TextMessage.Locked = False
End Sub

Private Sub CmdRequest_Click()
FrameMessage.Visible = False
FrameBuku.Visible = False
FrameAnggota.Visible = False
FrameHistory.Visible = False
FrameRequest.Visible = True
Call CmdRefreshMessage_Click
End Sub

Private Sub ComboKategoriHistory_Click()
Call ComboKategoriHistoryClient
End Sub

Private Sub ComboKategoriHistory_Change()
Call ComboKategoriHistoryClient
End Sub

Private Sub ComboPetugasAnggota_Click()
If ComboPetugasAnggota.Text = "Petugas" Then
    Koneksi "SELECT id_petugas, nama_petugas FROM tb_petugas ORDER BY id_petugas ASC"
    IsiPetugasAnggota ListPetugasAnggota
ElseIf ComboPetugasAnggota.Text = "Anggota" Then
    Koneksi "SELECT id_anggota, nama_anggota FROM tb_anggota ORDER BY id_anggota ASC"
    IsiPetugasAnggota ListPetugasAnggota
End If
End Sub

Private Sub Form_Load()
TimerClient.Enabled = True
TimerClient.Interval = 100
CmdBalas.Enabled = False
TextJudul.ToolTipText = "Perihal / Judul Request"
TextTujuan.ToolTipText = "ID Petugas / Anggota"

ComboPetugasAnggota.AddItem "Petugas"
ComboPetugasAnggota.AddItem "Anggota"

ComboCariPetugasAnggota.AddItem "ID"
ComboCariPetugasAnggota.AddItem "Nama"

With ComboKategoriHistory
    .AddItem "ID"
    .AddItem "Judul"
    .AddItem "Jumlah"
    .AddItem "Tanggal"
    .AddItem "Status"
End With

Koneksi "SELECT * FROM tb_anggota WHERE id_anggota='" & FormLogin.TextID.Text & "'"
If DB!jenis_kelamin = "Pria" Then
    ImageWanita.Visible = False
    ImagePria.Visible = True
ElseIf DB!jenis_kelamin = "Wanita" Then
    ImagePria.Visible = False
    ImageWanita.Visible = True
End If
If DB!status_anggota = "Aktif" Then
    ImageBelumAktif.Visible = False
    ImageAktif.Visible = True
ElseIf DB!status_anggota = "Tidak Aktif" Then
    ImageAktif.Visible = False
    ImageBelumAktif.Visible = True
End If
Do While Not DB.EOF
    LabelNamaAnggota.Caption = DB!nama_anggota
    LabelKelasAnggota.Caption = DB!kelas_anggota
    LabelJurusanAnggota.Caption = DB!jurusan_anggota
    LabelSekolahAnggota.Caption = DB!sekolah_anggota
    LabelNISAnggota.Caption = DB!nis_anggota
    LabelGabungAnggota.Caption = DB!tanggal_daftar
    LabelKonfirmasiAnggota.Caption = DB!petugas_daftar
    LabelDendaAnggota.Caption = DB!total_denda
DB.MoveNext
Loop

With ListBuku
    .ColumnHeaders.Clear
    .ColumnHeaders.Add , , "ID", "700"
    .ColumnHeaders.Add , , "Nama Buku", "3000"
    .ColumnHeaders.Add , , "Jumlah Buku"
    .ColumnHeaders.Add , , "Nama Pengarang"
    .ColumnHeaders.Add , , "Nama Penerbit"
    .ColumnHeaders.Add , , "Tahun Terbit", "1200"
    .ColumnHeaders.Add , , "Tanggal Daftar"
    .ColumnHeaders.Add , , "Petugas"
    .ColumnHeaders.Add , , "Kategori Buku"
    .GridLines = True
    .FullRowSelect = True
    .LabelEdit = lvwManual
End With

With ComboTujuan
    .AddItem "Petugas"
    .AddItem "Anggota"
End With

With ListMessage
    .ColumnHeaders.Clear
    .ColumnHeaders.Add , , "ID", "650"
    .ColumnHeaders.Add , , "Perihal"
    .ColumnHeaders.Add , , "Status"
    .ColumnHeaders.Add , , "ID Pengirim"
    .ColumnHeaders.Add , , "Pengirim"
    .ColumnHeaders.Add , , "Penerima"
    .ColumnHeaders.Add , , "Konten"
    .ColumnHeaders.Add , , "Tanggal"
    .ColumnHeaders.Add , , "Waktu"
    .GridLines = True
    .FullRowSelect = True
    .LabelEdit = lvwManual
End With

With ListPetugasAnggota
    .ColumnHeaders.Clear
    .ColumnHeaders.Add , , "ID"
    .ColumnHeaders.Add , , "Nama", "6200"
    .GridLines = True
    .FullRowSelect = True
    .LabelEdit = lvwManual
End With

With ListHistory
    .ColumnHeaders.Clear
    .ColumnHeaders.Add , , "ID", "650"
    .ColumnHeaders.Add , , "Judul", "3900"
    .ColumnHeaders.Add , , "Jumlah", "750"
    .ColumnHeaders.Add , , "Tanggal", "1000"
    .ColumnHeaders.Add , , "Status", "1000"
    .GridLines = True
    .FullRowSelect = True
    .LabelEdit = lvwManual
End With

With ComboCariBuku
    .Clear
    .AddItem "ID"
    .AddItem "Judul"
    .AddItem "Jumlah"
    .AddItem "Pengarang"
    .AddItem "Penerbit"
    .AddItem "Tahun"
    .AddItem "Kategori"
End With

Koneksi "SELECT * FROM tb_buku ORDER BY tanggal_daftar DESC"
IsiBukuClient ListBuku

MasukanMessageClientFirst
MasukanHistory
HitungHistory

End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Public Sub MasukanMessageClient()
Koneksi "SELECT id_anggota FROM tb_anggota WHERE id_anggota='" & StatusBarClient.Panels(2).Text & "'"
Koneksi "SELECT id_request, perihal_request, status_request, id_pengirim, nama_pengirim, nama_penerima, konten_request, tanggal_request, waktu_request FROM tb_request WHERE id_pengirim='" & DB!ID_Anggota & "' AND id_anggota_tujuan <> 0 OR id_anggota_tujuan='" & DB!ID_Anggota & "' ORDER BY status_request DESC"
IsiMessage ListMessage
End Sub

Public Sub MasukanMessageClientFirst()
Koneksi "SELECT id_anggota FROM tb_anggota WHERE id_anggota='" & FormLogin.TextID.Text & "'"
Koneksi "SELECT id_request, perihal_request, status_request, id_pengirim, nama_pengirim, nama_penerima, konten_request, tanggal_request, waktu_request FROM tb_request WHERE id_pengirim='" & DB!ID_Anggota & "' AND id_anggota_tujuan <> 0 OR id_anggota_tujuan='" & DB!ID_Anggota & "' ORDER BY status_request DESC"
IsiMessage ListMessage
End Sub

Public Sub MasukanHistory()
Koneksi "SELECT id_pinjam FROM tb_pinjam WHERE id_anggota='" & FormLogin.TextID.Text & "'"
    Do While Not DB.EOF
        Koneksi "SELECT id_pinjam_detail, nama_buku, jumlah_buku, tanggal_pinjam, status_pinjam_detail FROM tb_pinjam_detail WHERE id_pinjam='" & DB!ID_Pinjam & "' ORDER BY tanggal_pinjam DESC"
        Call IsiHistory
    Loop
End Sub

Public Sub HitungHistory()
Koneksi "SELECT id_pinjam FROM tb_pinjam WHERE id_anggota='" & FormLogin.TextID.Text & "'"
    If Not DB.EOF Then
        Koneksi "SELECT SUM(jumlah_buku) AS jumlah_buku_di_pinjam FROM tb_pinjam_detail WHERE id_pinjam='" & DB!ID_Pinjam & "' AND status_pinjam_detail='Pinjam';"
        If DB!jumlah_buku_di_pinjam > 0 Then
            TextIsiPemberitahuan.Text = "Anda Sedang Meminjam " & DB!jumlah_buku_di_pinjam & " Buku." & vbNewLine & "" & vbNewLine & "Jangan Sampai Terlambat Mengembalikan Ya!"
        Else
            TextIsiPemberitahuan.Text = "Anda Tidak Memiliki Buku Yang Sedang Di Pinjam." & vbNewLine & "" & vbNewLine & "Bagus, Jangan Sampai Telat Mengembalikan Buku Ya!"
        End If
    Else
        TextIsiPemberitahuan.Text = "Anda Tidak Memiliki Buku Yang Sedang Di Pinjam." & vbNewLine & "" & vbNewLine & "Bagus, Jangan Sampai Telat Mengembalikan Buku Ya!"
    End If
End Sub

Private Sub TextCariBuku_Change()
TextCariBuku = FilterInjeksi(TextCariBuku.Text)

    If ComboCariBuku.Text = "" Then
        MsgBox "Pilih Kriteria Buku", vbExclamation, "Pilih Kriteria"
    ElseIf TextCariBuku.Text = "" Then
        Koneksi "SELECT * FROM tb_buku ORDER BY tanggal_daftar DESC"
        IsiBukuClient ListBuku
    ElseIf ComboCariBuku.Text = "ID" Then
        Koneksi "SELECT * FROM tb_buku WHERE id_buku='" & TextCariBuku & "'"
        IsiBukuClient ListBuku
    ElseIf ComboCariBuku.Text = "Judul" Then
        Koneksi "SELECT * FROM tb_buku WHERE nama_buku='" & TextCariBuku & "'"
        IsiBukuClient ListBuku
    ElseIf ComboCariBuku.Text = "Jumlah" Then
        Koneksi "SELECT * FROM tb_buku WHERE jumlah_buku='" & TextCariBuku & "'"
        IsiBukuClient ListBuku
    ElseIf ComboCariBuku.Text = "Pengarang" Then
        Koneksi "SELECT * FROM tb_buku WHERE nama_pengarang='" & TextCariBuku & "'"
        IsiBukuClient ListBuku
    ElseIf ComboCariBuku.Text = "Penerbit" Then
        Koneksi "SELECT * FROM tb_buku WHERE nama_penerbit='" & TextCariBuku & "'"
        IsiBukuClient ListBuku
    ElseIf ComboCariBuku.Text = "Tahun" Then
        Koneksi "SELECT * FROM tb_buku WHERE tahun_terbit='" & TextCariBuku & "'"
        IsiBukuClient ListBuku
    ElseIf ComboCariBukuu.Text = "Kategori" Then
        Koneksi "SELECT * FROM tb_buku WHERE kategori_buku='" & TextCariBuku & "'"
        IsiBukuClient ListBuku
    End If
End Sub

Private Sub TextCariPetugasAnggota_Change()
TextCariPetugasAnggota = FilterInjeksi(TextCariPetugasAnggota.Text)

If ComboPetugasAnggota.Text = "" Then
    MsgBox "Pilih Kriteria Anggota / Petugas", vbExclamation, "Pilih Kriteria"
ElseIf ComboCariPetugasAnggota.Text = "" Then
    MsgBox "Pilih Kategori Pencarian", vbExclamation, "Pilih Kategori Pencarian"
ElseIf ComboPetugasAnggota.Text = "Petugas" Then
    If ComboCariPetugasAnggota.Text = "ID" Then
        Koneksi "SELECT id_petugas, nama_petugas FROM tb_petugas WHERE id_petugas LIKE '%' '" & TextCariPetugasAnggota & "' '%'"
        IsiPetugasAnggota ListPetugasAnggota
    ElseIf ComboCariPetugasAnggota.Text = "Nama" Then
        Koneksi "SELECT id_petugas, nama_petugas FROM tb_petugas WHERE nama_petugas LIKE '%' '" & TextCariPetugasAnggota & "' '%'"
        IsiPetugasAnggota ListPetugasAnggota
    End If
ElseIf ComboPetugasAnggota.Text = "Anggota" Then
    If ComboCariPetugasAnggota.Text = "ID" Then
        Koneksi "SELECT id_anggota, nama_anggota FROM tb_anggota WHERE id_anggota LIKE '%' '" & TextCariPetugasAnggota & "' '%'"
        IsiPetugasAnggota ListPetugasAnggota
    ElseIf ComboCariPetugasAnggota.Text = "Nama" Then
        Koneksi "SELECT id_anggota, nama_anggota FROM tb_anggota WHERE nama_anggota LIKE '%' '" & TextCariPetugasAnggota & "' '%'"
        IsiPetugasAnggota ListPetugasAnggota
    End If
End If
End Sub
Private Sub TextSearchHistory_Change()
TextSearchHistory = FilterInjeksi(TextSearchHistory.Text)

If ComboKategoriHistory.Text = "" Then
    MsgBox "Silahkan Pilih Kategori Pencarian", vbExclamation, "Pilih Kategori Pencarian"
ElseIf ComboKategoriHistory.Text = "ID" Then
    Koneksi "SELECT id_pinjam FROM tb_pinjam WHERE id_anggota='" & StatusBarClient.Panels(2).Text & "'"
    Do While Not DB.EOF
        Koneksi "SELECT id_pinjam_detail, nama_buku, jumlah_buku, tanggal_pinjam, status_pinjam_detail FROM tb_pinjam_detail WHERE id_pinjam='" & DB!ID_Pinjam & "' AND id_pinjam_detail LIKE '%' '" & TextSearchHistory & "' '%'"
        Call IsiHistory
    Loop
ElseIf ComboKategoriHistory.Text = "Judul" Then
    Koneksi "SELECT id_pinjam FROM tb_pinjam WHERE id_anggota='" & StatusBarClient.Panels(2).Text & "'"
    Do While Not DB.EOF
        Koneksi "SELECT id_pinjam_detail, nama_buku, jumlah_buku, tanggal_pinjam, status_pinjam_detail FROM tb_pinjam_detail WHERE id_pinjam='" & DB!ID_Pinjam & "' AND nama_buku LIKE '%' '" & TextSearchHistory & "' '%'"
        Call IsiHistory
    Loop
ElseIf ComboKategoriHistory.Text = "Jumlah" Then
    Koneksi "SELECT id_pinjam FROM tb_pinjam WHERE id_anggota='" & StatusBarClient.Panels(2).Text & "'"
    Do While Not DB.EOF
        Koneksi "SELECT id_pinjam_detail, nama_buku, jumlah_buku, tanggal_pinjam, status_pinjam_detail FROM tb_pinjam_detail WHERE id_pinjam='" & DB!ID_Pinjam & "' AND jumlah_buku LIKE '%' '" & TextSearchHistory & "' '%'"
        Call IsiHistory
    Loop
ElseIf ComboKategoriHistory.Text = "Tanggal" Then
    Koneksi "SELECT id_pinjam FROM tb_pinjam WHERE id_anggota='" & StatusBarClient.Panels(2).Text & "'"
    Do While Not DB.EOF
        Koneksi "SELECT id_pinjam_detail, nama_buku, jumlah_buku, tanggal_pinjam, status_pinjam_detail FROM tb_pinjam_detail WHERE id_pinjam='" & DB!ID_Pinjam & "' AND tanggal_pinjam LIKE '%' '" & TextSearchHistory & "' '%'"
        Call IsiHistory
    Loop
ElseIf ComboKategoriHistory.Text = "Status" Then
    Koneksi "SELECT id_pinjam FROM tb_pinjam WHERE id_anggota='" & StatusBarClient.Panels(2).Text & "'"
    Do While Not DB.EOF
        Koneksi "SELECT id_pinjam_detail, nama_buku, jumlah_buku, tanggal_pinjam, status_pinjam_detail FROM tb_pinjam_detail WHERE id_pinjam='" & DB!ID_Pinjam & "' AND status_pinjam_detail LIKE '%' '" & TextSearchHistory & "' '%'"
        Call IsiHistory
    Loop
Else
    MsgBox "Kategori Tidak Di Temukan", vbCritical, "Ganti Kategori"
End If
End Sub

Private Sub TextTujuan_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= Asc("0") & Chr(13) _
     And KeyAscii <= Asc("9") & Chr(13) _
     Or KeyAscii = vbKeyBack _
     Or KeyAscii = vbKeyDelete _
     Or KeyAscii = vbKeySpace) Then
    KeyAscii = 0
End If
End Sub

Private Sub TimerClient_Timer()
StatusBarClient.Panels(4) = Format(Time, "HH:MM:SS")
StatusBarClient.Panels(5) = Format(Date, "YYYY-MM-DD")
End Sub
