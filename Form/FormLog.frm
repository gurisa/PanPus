VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FormLog 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Log"
   ClientHeight    =   6645
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10215
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormLog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6645
   ScaleWidth      =   10215
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame FrameMessage 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Log Message "
      Height          =   5895
      Left            =   120
      TabIndex        =   54
      Top             =   120
      Width           =   9975
      Begin VB.TextBox TextKonten 
         Height          =   2535
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   61
         Top             =   2760
         Width           =   9735
      End
      Begin VB.CommandButton CmdBalas 
         Caption         =   "Balas"
         Height          =   375
         Left            =   8640
         TabIndex        =   60
         Top             =   5400
         Width           =   1215
      End
      Begin VB.CommandButton CmdRefresh 
         Caption         =   "Refresh"
         Height          =   375
         Left            =   8640
         TabIndex        =   58
         Top             =   720
         Width           =   1215
      End
      Begin VB.CommandButton CmdLihat 
         Caption         =   "Lihat"
         Height          =   375
         Left            =   8640
         TabIndex        =   57
         Top             =   1200
         Width           =   1215
      End
      Begin VB.CommandButton CmdHapus 
         Caption         =   "Hapus"
         Height          =   375
         Left            =   8640
         TabIndex        =   56
         Top             =   1680
         Width           =   1215
      End
      Begin VB.CommandButton CmdPesanBaru 
         Caption         =   "Pesan Baru"
         Height          =   375
         Left            =   8640
         TabIndex        =   55
         Top             =   240
         Width           =   1215
      End
      Begin MSComctlLib.ListView ListMessage 
         Height          =   1935
         Left            =   120
         TabIndex        =   59
         Top             =   240
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   3413
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
         TabIndex        =   72
         Top             =   5520
         Width           =   1695
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
         TabIndex        =   71
         Top             =   5520
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
         Left            =   5160
         TabIndex        =   70
         Top             =   5280
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
         TabIndex        =   69
         Top             =   5280
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
         Left            =   5040
         TabIndex        =   68
         Top             =   2520
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
         TabIndex        =   67
         Top             =   2520
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
         Left            =   6120
         TabIndex        =   66
         Top             =   5280
         Width           =   1935
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
         Left            =   6120
         TabIndex        =   65
         Top             =   2520
         Width           =   3615
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
         TabIndex        =   64
         Top             =   2520
         Width           =   3975
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
         TabIndex        =   63
         Top             =   5280
         Width           =   2175
      End
      Begin VB.Label LabelHeaderMessage 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "-"
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
         TabIndex        =   62
         Top             =   2280
         Width           =   9735
      End
   End
   Begin VB.Frame FramePesanBaru 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Pesan Baru "
      Height          =   5895
      Left            =   120
      TabIndex        =   45
      Top             =   120
      Width           =   9975
      Begin VB.TextBox TextMessage 
         Height          =   4215
         Left            =   120
         MaxLength       =   2000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   82
         Top             =   840
         Width           =   6615
      End
      Begin VB.Frame FrameCariAnggota 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   2895
         Left            =   6840
         TabIndex        =   77
         Top             =   720
         Width           =   3015
         Begin VB.ComboBox ComboKategoriTujuan 
            Height          =   330
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   80
            Top             =   240
            Width           =   1455
         End
         Begin VB.ComboBox ComboKategoriCariAnggota 
            Height          =   330
            Left            =   1560
            Style           =   2  'Dropdown List
            TabIndex        =   79
            Top             =   240
            Width           =   1335
         End
         Begin VB.TextBox TextCariAnggota 
            Height          =   315
            Left            =   120
            TabIndex        =   78
            Top             =   600
            Width           =   2775
         End
         Begin MSComctlLib.ListView ListCariAnggota 
            Height          =   1815
            Left            =   120
            TabIndex        =   81
            Top             =   960
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   3201
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
      End
      Begin VB.TextBox TextPenampungNamaPetugas 
         Height          =   255
         Left            =   4680
         TabIndex        =   76
         Top             =   5520
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox TextPenampungIDAnggota 
         Height          =   255
         Left            =   3720
         TabIndex        =   75
         Top             =   5520
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox TextPenampungNamaAnggota 
         Height          =   255
         Left            =   3720
         TabIndex        =   74
         Top             =   5280
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox TextPenampungIDPetugas 
         Height          =   255
         Left            =   4680
         TabIndex        =   73
         Top             =   5280
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton CmdBroadcastPetugas 
         Caption         =   "Broadcast Petugas"
         Height          =   375
         Left            =   1920
         TabIndex        =   53
         Top             =   5280
         Width           =   1695
      End
      Begin VB.CommandButton CmdKirim 
         Caption         =   "Kirim"
         Height          =   375
         Left            =   5520
         TabIndex        =   52
         Top             =   5280
         Width           =   1215
      End
      Begin VB.TextBox TextJudul 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   960
         MaxLength       =   250
         TabIndex        =   49
         Top             =   360
         Width           =   4215
      End
      Begin VB.TextBox TextTujuan 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   8400
         MaxLength       =   250
         TabIndex        =   48
         Top             =   360
         Width           =   1455
      End
      Begin VB.ComboBox ComboTujuan 
         Height          =   330
         Left            =   6840
         Style           =   2  'Dropdown List
         TabIndex        =   47
         Top             =   360
         Width           =   1455
      End
      Begin VB.CommandButton CmdBroadcastAnggota 
         Caption         =   "Broadcast Anggota"
         Height          =   375
         Left            =   120
         TabIndex        =   46
         Top             =   5280
         Width           =   1695
      End
      Begin VB.Image ImagePesan 
         Height          =   1335
         Left            =   7800
         Picture         =   "FormLog.frx":0CCA
         Stretch         =   -1  'True
         Top             =   4200
         Width           =   1695
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
         TabIndex        =   51
         Top             =   360
         Width           =   855
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
         Left            =   6000
         TabIndex        =   50
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.CommandButton CmdLihatLogMessage 
      Caption         =   "Log Message"
      Height          =   375
      Left            =   5520
      TabIndex        =   44
      Top             =   6120
      Width           =   1695
   End
   Begin VB.Frame FrameDenda 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Log Denda "
      Height          =   5895
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   9975
      Begin VB.Frame FrameCariDenda 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Cari Denda "
         Height          =   1575
         Left            =   120
         TabIndex        =   31
         Top             =   4200
         Width           =   9735
         Begin VB.TextBox TextCari 
            Height          =   315
            Left            =   3120
            TabIndex        =   33
            Top             =   960
            Width           =   3495
         End
         Begin VB.ComboBox ComboCariKategori 
            Height          =   330
            Left            =   3120
            Style           =   2  'Dropdown List
            TabIndex        =   32
            Top             =   480
            Width           =   3495
         End
         Begin VB.Image ImageCariDenda 
            Height          =   975
            Left            =   600
            Picture         =   "FormLog.frx":4448
            Stretch         =   -1  'True
            Top             =   360
            Width           =   1455
         End
      End
      Begin VB.CommandButton CmdCariDenda 
         Caption         =   "Cari"
         Height          =   375
         Left            =   8280
         TabIndex        =   30
         Top             =   240
         Width           =   1575
      End
      Begin MSComctlLib.ListView ListAnggotaKembali 
         Height          =   1935
         Left            =   6480
         TabIndex        =   28
         Top             =   2280
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   3413
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
      Begin VB.CommandButton CmdRefreshAnggota 
         Caption         =   "Refresh"
         Height          =   375
         Left            =   8280
         TabIndex        =   27
         Top             =   1680
         Width           =   1575
      End
      Begin VB.CommandButton CmdUbahDenda 
         Caption         =   "Ubah"
         Height          =   375
         Left            =   8280
         TabIndex        =   26
         Top             =   720
         Width           =   1575
      End
      Begin VB.Frame FrameUbahDenda 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Operator "
         Height          =   1575
         Left            =   120
         TabIndex        =   18
         Top             =   4200
         Width           =   9735
         Begin VB.ComboBox ComboIDPinjamDetail 
            BackColor       =   &H00FFFFFF&
            Height          =   330
            Left            =   6240
            Style           =   2  'Dropdown List
            TabIndex        =   29
            Top             =   840
            Width           =   1935
         End
         Begin VB.ComboBox ComboIDAnggota 
            Height          =   330
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   25
            Top             =   360
            Width           =   2415
         End
         Begin VB.TextBox TextNamaAnggota 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   24
            Text            =   "Nama Anggota"
            Top             =   840
            Width           =   4575
         End
         Begin VB.CommandButton CmdEksekusiDenda 
            Caption         =   "Eksekusi"
            Height          =   375
            Left            =   8280
            TabIndex        =   23
            Top             =   840
            Width           =   1335
         End
         Begin VB.OptionButton OptionTambahDenda 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Tambah (+)"
            Height          =   255
            Left            =   120
            TabIndex        =   22
            Top             =   840
            Width           =   1215
         End
         Begin VB.OptionButton OptionKurangDenda 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Kurang (-)"
            Height          =   255
            Left            =   120
            TabIndex        =   21
            Top             =   1200
            Width           =   1215
         End
         Begin VB.TextBox TextJumlahDendaEdit 
            Height          =   375
            Left            =   6240
            TabIndex        =   20
            Text            =   "Jumlah Denda"
            Top             =   360
            Width           =   3375
         End
         Begin VB.TextBox TextJumlahDendaAwal 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   2760
            Locked          =   -1  'True
            TabIndex        =   19
            Text            =   "Jumlah Denda Awal"
            Top             =   360
            Width           =   3375
         End
      End
      Begin VB.CommandButton CmdHapusDenda 
         Caption         =   "Hapus"
         Height          =   375
         Left            =   8280
         TabIndex        =   15
         Top             =   1200
         Width           =   1575
      End
      Begin MSComctlLib.ListView ListAnggota 
         Height          =   1935
         Left            =   120
         TabIndex        =   16
         Top             =   2280
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   3413
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
      Begin MSComctlLib.ListView ListDenda 
         Height          =   2055
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   3625
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
   Begin VB.Frame FramePinjam 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Log Pinjam "
      Height          =   5895
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   9975
      Begin VB.Frame FrameCariPinjamDetail 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Cari Pinjam Detail "
         Height          =   1335
         Left            =   1080
         TabIndex        =   35
         Top             =   3960
         Width           =   8775
         Begin VB.TextBox TextCariDetail 
            Height          =   315
            Left            =   2760
            TabIndex        =   39
            Top             =   840
            Width           =   5895
         End
         Begin VB.ComboBox ComboCariKategoriDetail 
            Height          =   330
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   38
            Top             =   840
            Width           =   2535
         End
         Begin VB.TextBox TextCariPinjam 
            Height          =   315
            Left            =   2760
            TabIndex        =   37
            Top             =   360
            Width           =   5895
         End
         Begin VB.ComboBox ComboCariKategoriPinjam 
            Height          =   330
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   36
            Top             =   360
            Width           =   2535
         End
      End
      Begin VB.CommandButton CmdCariPinjamDetail 
         Caption         =   "Cari"
         Height          =   615
         Left            =   120
         TabIndex        =   34
         Top             =   4440
         Width           =   855
      End
      Begin VB.CommandButton CmdPinjamDetail 
         Caption         =   "Detail"
         Height          =   615
         Left            =   120
         TabIndex        =   11
         Top             =   3720
         Width           =   855
      End
      Begin VB.CommandButton CmdHapusPinjam 
         Caption         =   "Hapus"
         Height          =   615
         Left            =   120
         TabIndex        =   10
         Top             =   3000
         Width           =   855
      End
      Begin VB.CommandButton CmdPinjamkan 
         Caption         =   "Status Pinjam"
         Height          =   375
         Left            =   1080
         TabIndex        =   9
         Top             =   5400
         Width           =   1695
      End
      Begin VB.CommandButton CmdPinjamkanDetail 
         Caption         =   "Status Detail"
         Height          =   375
         Left            =   2880
         TabIndex        =   8
         Top             =   5400
         Width           =   1695
      End
      Begin VB.CommandButton CmdHapusDetailPinjam 
         Caption         =   "Hapus Detail"
         Height          =   375
         Left            =   8160
         TabIndex        =   7
         Top             =   5400
         Width           =   1695
      End
      Begin MSComctlLib.ListView ListPinjam 
         Height          =   2655
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   9735
         _ExtentX        =   17171
         _ExtentY        =   4683
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
      Begin MSComctlLib.ListView ListPinjamDetail 
         Height          =   2415
         Left            =   1080
         TabIndex        =   13
         Top             =   2880
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   4260
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   0
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
   Begin VB.Frame FrameKembali 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Log Kembali "
      Height          =   5895
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   9975
      Begin VB.Frame FrameCariKembali 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Cari Data Kembali "
         Height          =   735
         Left            =   120
         TabIndex        =   40
         Top             =   5040
         Width           =   9735
         Begin VB.CommandButton CmdHapusKembali 
            Caption         =   "Hapus"
            Height          =   315
            Left            =   8280
            TabIndex        =   43
            Top             =   240
            Width           =   1335
         End
         Begin VB.TextBox TextCariKembali 
            Height          =   315
            Left            =   2880
            TabIndex        =   42
            Top             =   240
            Width           =   5295
         End
         Begin VB.ComboBox ComboCariKategoriKembali 
            Height          =   330
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   41
            Top             =   240
            Width           =   2655
         End
      End
      Begin MSComctlLib.ListView ListKembali 
         Height          =   4815
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   9735
         _ExtentX        =   17171
         _ExtentY        =   8493
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
   Begin VB.CommandButton CmdLihatLogDenda 
      Caption         =   "Log Denda"
      Height          =   375
      Left            =   3720
      TabIndex        =   3
      Top             =   6120
      Width           =   1695
   End
   Begin VB.CommandButton CmdLihatLogKembali 
      Caption         =   "Log Kembali"
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   6120
      Width           =   1695
   End
   Begin VB.CommandButton CmdLihatLogPinjam 
      Caption         =   "Log Pinjam"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   6120
      Width           =   1695
   End
   Begin VB.CommandButton CmdKeluar 
      Caption         =   "Keluar"
      Height          =   375
      Left            =   8400
      TabIndex        =   0
      Top             =   6120
      Width           =   1695
   End
End
Attribute VB_Name = "FormLog"
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

Private Sub CmdBalas_Click()
TextKonten = FilterInjeksi(TextKonten.Text)

If CmdBalas.Caption = "Balas" Then
    CmdBalas.Caption = "Kirim"
    TextKonten.Text = ""
    TextKonten.Locked = False
    TextKonten.SetFocus
ElseIf CmdBalas.Caption = "Kirim" Then
    If TextKonten.Text = "" Then
        MsgBox "Isi Pesan Dengan Benar", vbExclamation, "Pesan Tidak Di Kirim"
    Else
        If MsgBox("Balas Pesan Dari " & LabelPengirim.Caption & " ? ", vbInformation + vbYesNo, "Balas Pesan") = vbYes Then
            If LabelKategori.Caption = "Petugas - Petugas" Then
                Koneksi "SELECT id_petugas FROM tb_petugas WHERE nama_petugas='" & LabelPengirim.Caption & "'"
                Conn.Execute "INSERT INTO tb_request(id_pengirim, nama_pengirim, nama_penerima, id_petugas_tujuan, perihal_request, konten_request, tanggal_request, waktu_request, status_request) VALUES('" & FormUtama.StatusBarUtama.Panels(5).Text & "','" & FormUtama.StatusBarUtama.Panels(4).Text & "','" & LabelPengirim.Caption & "','" & DB!id_petugas & "','" & LabelHeaderMessage.Caption & "','" & TextKonten.Text & "','" & FormUtama.StatusBarUtama.Panels(1).Text & "','" & FormUtama.StatusBarUtama.Panels(2).Text & "','Belum Di Baca')"
                Call CmdRefresh_Click
            ElseIf LabelKategori.Caption = "Anggota - Petugas" Then
                Koneksi "SELECT id_anggota FROM tb_anggota WHERE nama_anggota='" & LabelPengirim.Caption & "'"
                Conn.Execute "INSERT INTO tb_request(id_pengirim, nama_pengirim, nama_penerima, id_anggota_tujuan, perihal_request, konten_request, tanggal_request, waktu_request, status_request) VALUES('" & FormUtama.StatusBarUtama.Panels(5).Text & "','" & FormUtama.StatusBarUtama.Panels(4).Text & "','" & LabelPengirim.Caption & "','" & DB!ID_Anggota & "','" & LabelHeaderMessage.Caption & "','" & TextKonten.Text & "','" & FormUtama.StatusBarUtama.Panels(1).Text & "','" & FormUtama.StatusBarUtama.Panels(2).Text & "','Belum Di Baca')"
                Call CmdRefresh_Click
            End If
            MsgBox "Pesan Berhasil Di Kirim", vbInformation, "Berhasil Mengirim Pesan"
        Else
            Exit Sub
        End If
    End If
End If
End Sub

Private Sub CmdBroadcastAnggota_Click()
If TextJudul.Text = "" Or TextMessage.Text = "" Then
    MsgBox "Masukan Data Pesan Dengan Benar", vbExclamation, "Masukan Semua Data Pesan"
Else
Koneksi "SELECT id_anggota, nama_anggota FROM tb_anggota WHERE status_anggota='Aktif'"
    If Not DB.EOF Then
        DB.MoveFirst
    While Not DB.EOF
        Conn.Execute "INSERT INTO tb_request(id_pengirim, nama_pengirim, nama_penerima, id_anggota_tujuan, perihal_request, konten_request, tanggal_request, waktu_request, status_request) VALUES('" & FormUtama.StatusBarUtama.Panels(5).Text & "','" & FormUtama.StatusBarUtama.Panels(4).Text & "','" & DB!nama_anggota & "','" & DB!ID_Anggota & "','" & TextJudul.Text & "','" & TextMessage.Text & "','" & Format(Date, "YYYY-MM-DD") & "','" & Format(Time, "HH:MM:SS") & "','Belum Di Baca')"
        DB.MoveNext
    Wend
    End If
    MsgBox "Pesan Telah Di Kirim Ke Semua Anggota", vbInformation, "Pesan Berhasil Di Kirim"
End If
End Sub

Private Sub CmdBroadcastPetugas_Click()
If TextJudul.Text = "" Or TextMessage.Text = "" Then
    MsgBox "Masukan Data Pesan Dengan Benar", vbExclamation, "Masukan Semua Data Pesan"
Else
    Koneksi "SELECT id_petugas, nama_petugas FROM tb_petugas"
    If Not DB.EOF Then
        DB.MoveFirst
            While Not DB.EOF
                Conn.Execute "INSERT INTO tb_request(id_pengirim, nama_pengirim, nama_penerima, id_petugas_tujuan, perihal_request, konten_request, tanggal_request, waktu_request, status_request) VALUES('" & DB!id_petugas & "','" & FormUtama.StatusBarUtama.Panels(4).Text & "','" & DB!nama_petugas & "','" & DB!id_petugas & "','" & TextJudul.Text & "','" & TextMessage.Text & "','" & Format(Date, "YYYY-MM-DD") & "','" & Format(Time, "HH:MM:SS") & "','Belum Di Baca')"
                DB.MoveNext
            Wend
        MsgBox "Pesan Telah Di Kirim Ke Semua Petugas", vbInformation, "Pesan Berhasil Di Kirim"
    End If
End If
End Sub

Private Sub CmdCariDenda_Click()
If CmdCariDenda.Caption = "Cari" Then
    CmdCariDenda.Caption = "Selesai"
    FrameUbahDenda.Visible = False
    FrameCariDenda.Visible = True
        With ListAnggota
            .Height = 1935
            .Left = 120
            .Top = 2280
            .Width = 6375
        End With
        
        With ListDenda
            .Height = 2055
            .Left = 120
            .Top = 240
            .Width = 8055
        End With
        
        With ListAnggotaKembali
            .Height = 1935
            .Left = 6480
            .Top = 2280
            .Width = 3375
        End With
        CmdUbahDenda.Enabled = False
        CmdHapusDenda.Enabled = False
        CmdRefreshAnggota.Enabled = False
ElseIf CmdCariDenda.Caption = "Selesai" Then
    CmdCariDenda.Caption = "Cari"
    FrameCariDenda.Visible = False
        With ListAnggota
            .Height = 2775
            .Left = 120
            .Top = 3000
            .Width = 6375
        End With
        
        With ListDenda
            .Height = 2775
            .Left = 120
            .Top = 240
            .Width = 8055
        End With
        
        With ListAnggotaKembali
            .Height = 2775
            .Left = 6480
            .Top = 3000
            .Width = 3375
        End With
        CmdUbahDenda.Enabled = True
        CmdHapusDenda.Enabled = True
        CmdRefreshAnggota.Enabled = True
        
        Koneksi "SELECT * FROM tb_denda"
        IsiDenda ListDenda
End If
End Sub

Private Sub CmdCariPinjamDetail_Click()
If CmdCariPinjamDetail.Caption = "Cari" Then
    CmdCariPinjamDetail.Caption = "Selesai"
    FrameCariPinjamDetail.Visible = True
    
    With ListPinjam
    .Width = 9735
    .Height = 1815
    .Top = 240
    .Left = 120
    End With
    
    With ListPinjamDetail
    .Width = 8775
    .Height = 1935
    .Top = 2040
    .Left = 1080
    End With
    
    With CmdHapusPinjam
    .Width = 855
    .Height = 615
    .Top = 2160
    .Left = 120
    End With
    
    With CmdPinjamDetail
    .Width = 855
    .Height = 615
    .Top = 2880
    .Left = 120
    End With
    
    With CmdCariPinjamDetail
    .Width = 855
    .Height = 615
    .Top = 3600
    .Left = 120
    End With
ElseIf CmdCariPinjamDetail.Caption = "Selesai" Then
    CmdCariPinjamDetail.Caption = "Cari"
    FrameCariPinjamDetail.Visible = False
    
    With ListPinjam
    .Width = 9735
    .Height = 2655
    .Top = 240
    .Left = 120
    End With
    
    With ListPinjamDetail
    .Width = 8775
    .Height = 2415
    .Top = 2880
    .Left = 1080
    End With
    
    With CmdHapusPinjam
    .Width = 855
    .Height = 615
    .Top = 3000
    .Left = 120
    End With
    
    With CmdPinjamDetail
    .Width = 855
    .Height = 615
    .Top = 3720
    .Left = 120
    End With
    
    With CmdCariPinjamDetail
    .Width = 855
    .Height = 615
    .Top = 4440
    .Left = 120
    End With
End If
End Sub

Private Sub CmdEksekusiDenda_Click()
TextJumlahDendaEdit = FilterInjeksi(TextJumlahDendaEdit.Text)

If ComboIDAnggota.Text = "" Or TextJumlahDendaAwal.Text = "" Or TextNamaAnggota.Text = "" Or TextJumlahDendaEdit.Text = "" Or ComboIDPinjamDetail.Text = "" Then
    MsgBox "Silahkan Isi Data Denda Terlebih Dahulu", vbCritical, "Isi Data Denda Dengan Benar"
ElseIf OptionTambahDenda.Value = True Then
    Conn.Execute "UPDATE tb_anggota SET total_denda=total_denda + '" & Val(TextJumlahDendaEdit.Text) & "' WHERE id_anggota='" & ComboIDAnggota.Text & "'"
    Conn.Execute "INSERT INTO tb_denda(id_anggota, id_pinjam_detail, banyak_denda, tanggal_denda, petugas_denda) VALUES('" & ComboIDAnggota.Text & "','" & ComboIDPinjamDetail.Text & "','" & TextJumlahDendaEdit.Text & "','" & FormUtama.StatusBarUtama.Panels(1) & "','" & FormUtama.StatusBarUtama.Panels(4) & "')"
    MsgBox "Denda Berhasil Di Tambahkan Dari Total Denda Rp. " & TextJumlahDendaAwal.Text & " Di Tambah Rp. " & TextJumlahDendaEdit.Text & " Total Denda Menjadi Rp. " & Val(TextJumlahDendaAwal) + Val(TextJumlahDendaEdit) & "", vbInformation, "Berhasil Menambahkan Denda " & TextNamaAnggota.Text & ""
    Call MasukanLogDenda
    Call DefaultUbahDenda
ElseIf OptionKurangDenda.Value = True Then
    Koneksi "UPDATE tb_anggota SET total_denda=total_denda - '" & Val(TextJumlahDendaEdit.Text) & "' WHERE id_anggota='" & ComboIDAnggota.Text & "'"
    Conn.Execute "INSERT INTO tb_denda(id_anggota, id_pinjam_detail, banyak_denda, tanggal_denda, petugas_denda) VALUES('" & ComboIDAnggota.Text & "','" & ComboIDPinjamDetail.Text & "','" & TextJumlahDendaEdit.Text & "','" & FormUtama.StatusBarUtama.Panels(1) & "','" & FormUtama.StatusBarUtama.Panels(4) & "')"
    MsgBox "Denda Berhasil Di Kurangi Dari Total Denda Rp. " & TextJumlahDendaAwal.Text & " Di Kurang Rp. " & TextJumlahDendaEdit.Text & " Total Denda Menjadi Rp. " & Val(TextJumlahDendaAwal) - Val(TextJumlahDendaEdit) & "", vbInformation, "Berhasil Mengurangi Denda " & TextNamaAnggota.Text & ""
    Call MasukanLogDenda
    Call DefaultUbahDenda
Else
    MsgBox "Silahkan Pilih Operator Eksekusi Denda", vbExclamation, "Pilih Operator Eksekusi Denda"
End If
End Sub

Private Sub CmdHapus_Click()
If ListMessage.ListItems.Count > 0 Then
    If MsgBox("Hapus Pesan Request " & ListMessage.SelectedItem.ListSubItems(4).Text & " ?", vbQuestion + vbYesNo, "Hapus Pesan Request") = vbYes Then
        Conn.Execute "DELETE FROM tb_request WHERE id_request='" & ListMessage.SelectedItem.Text & "'"
        MsgBox "Pesan Dari " & ListMessage.SelectedItem.ListSubItems(4).Text & " Berhasil Di Hapus", vbInformation, "Pesan Berhasil Di Hapus"
        Call CmdRefresh_Click
        CmdBalas.Enabled = False
    End If
Else
    MsgBox "Tidak Ada Pesan", vbExclamation, "Tidak Ada Pesan"
End If
End Sub

Private Sub CmdHapusDenda_Click()
If ListDenda.ListItems.Count > 0 Then
    If MsgBox("Hapus Data Denda Dengan Nomer Denda " & ListDenda.SelectedItem.Text & " ?", vbQuestion + vbYesNo, "Hapus Data Denda") = vbYes Then
        Conn.Execute "DELETE FROM tb_denda WHERE id_denda='" & ListDenda.SelectedItem.Text & "'"
        MasukanLogDenda
        MsgBox "Data Denda Berhasil Di Hapus", vbInformation, "Data Denda Berhasil Di Hapus"
    End If
Else
    MsgBox "Tidak Ada Data Denda Yang Bisa Di Hapus", vbCritical, "Data Denda Tidak Tersedia"
End If
End Sub

Private Sub CmdHapusDetailPinjam_Click()
If ListPinjamDetail.ListItems.Count > 0 Then
    If MsgBox("Hapus Data Detail Pinjam Dengan Nomer Detail " & ListPinjamDetail.SelectedItem.Text & " ?", vbQuestion + vbYesNo, "Hapus Data Detail Pinjam") = vbYes Then
        Conn.Execute "DELETE FROM tb_pinjam_detail WHERE id_pinjam_detail='" & ListPinjamDetail.SelectedItem.Text & "'"
        IsiLogPinjamDetail ListPinjamDetail
        Call Default
        MsgBox "Data Detail Berhasil Di Hapus", vbInformation, "Data Detail Berhasil Di Hapus"
    End If
Else
    MsgBox "Tidak Ada Data Detail Yang Bisa Di Hapus", vbCritical, "Data Detail Tidak Tersedia"
End If
End Sub

Private Sub CmdHapusKembali_Click()
If ListKembali.ListItems.Count > 0 Then
    If MsgBox("Hapus Data Kembali Dengan Nomer Kembali " & ListKembali.SelectedItem.Text & " ?", vbQuestion + vbYesNo, "Hapus Data Kembali") = vbYes Then
        Conn.Execute "DELETE FROM tb_kembali WHERE id_kembali='" & ListKembali.SelectedItem.Text & "'"
        Call MasukanLogKembali
        Call Default
        MsgBox "Data Kembali Berhasil Di Hapus", vbInformation, "Data Kembali Berhasil Di Hapus"
    End If
Else
    MsgBox "Tidak Ada Data Kembali Yang Bisa Di Hapus", vbCritical, "Data Kembali Tidak Tersedia"
End If
End Sub

Private Sub CmdHapusPinjam_Click()
If ListPinjam.ListItems.Count > 0 Then
    If MsgBox("Hapus Data Pinjam Dengan Nomer Pinjam " & ListPinjam.SelectedItem.Text & " ?", vbQuestion + vbYesNo, "Hapus Data Pinjam") = vbYes Then
        Conn.Execute "DELETE FROM tb_pinjam WHERE id_pinjam='" & ListPinjam.SelectedItem.Text & "'"
        Call MasukanLogPinjam
        Call Default
        ListPinjam.ListItems.Clear
        MsgBox "Data Pinjam Berhasil Di Hapus", vbInformation, "Data Pinjam Berhasil Di Hapus"
    End If
Else
    MsgBox "Tidak Ada Data Pinjam Yang Bisa Di Hapus", vbCritical, "Data Pinjam Tidak Tersedia"
End If
End Sub

Private Sub CmdKeluar_Click()
Unload Me
End Sub

Private Sub CmdKirim_Click()
TextJudul = FilterInjeksi(TextJudul.Text)
TextTujuan = FilterInjeksi(TextTujuan.Text)
TextMessage = FilterInjeksi(TextMessage.Text)

If TextJudul.Text = "" Or TextTujuan.Text = "" Or TextMessage.Text = "" Then
    MsgBox "Masukan Data Pesan Dengan Benar", vbExclamation, "Masukan Semua Data Pesan"
ElseIf ComboTujuan.Text = "Petugas" Then
    Koneksi "SELECT nama_petugas AS penampung_nama_petugas FROM tb_petugas WHERE nama_petugas='" & FormUtama.StatusBarUtama.Panels(4).Text & "'"
    Koneksi "SELECT id_anggota AS penampung_id_petugas FROM tb_anggota WHERE id_anggota='" & TextTujuan.Text & "'"
    Koneksi "SELECT id_petugas FROM tb_petugas WHERE id_petugas='" & TextTujuan.Text & "'"
    If DB.EOF Then
        MsgBox "ID Petugas Tidak Di Ketahui", vbExclamation, "Masukan ID Petugas Dengan Benar"
    Else
        Koneksi "SELECT id_petugas AS penampung_id_petugas FROM tb_petugas WHERE id_petugas='" & TextTujuan.Text & "'"
        TextPenampungIDPetugas.Text = DB!penampung_id_petugas
        Koneksi "SELECT nama_petugas AS penampung_nama_petugas FROM tb_petugas WHERE id_petugas='" & TextPenampungIDPetugas.Text & "'"
        TextPenampungNamaPetugas.Text = DB!penampung_nama_petugas

        Conn.Execute "INSERT INTO tb_request(id_pengirim, nama_pengirim, nama_penerima, id_petugas_tujuan, perihal_request, konten_request, tanggal_request, waktu_request, status_request) VALUES('" & TextPenampungIDAnggota.Text & "','" & FormUtama.StatusBarUtama.Panels(4).Text & "','" & TextPenampungNamaPetugas.Text & "','" & TextTujuan.Text & "','" & TextJudul.Text & "','" & TextMessage.Text & "','" & Format(Date, "YYYY-MM-DD") & "','" & Format(Time, "HH:MM:SS") & "','Belum Di Baca')"
        MsgBox "Pesan Telah Di Kirim Ke " & TextPenampungNamaPetugas.Text & "", vbInformation, "Pesan Berhasil Di Kirim"
        TextJudul.Text = ""
        TextMessage.Text = ""
        TextTujuan.Text = ""
        MasukanMessagePetugas
    End If
ElseIf ComboTujuan.Text = "Anggota" Then
    Koneksi "SELECT nama_anggota FROM tb_anggota AS penampung_nama_anggota WHERE id_anggota='" & TextTujuan.Text & "'"
    Koneksi "SELECT id_petugas FROM tb_petugas AS penampung_id_petugas WHERE nama_petugas='" & FormUtama.StatusBarUtama.Panels(4).Text & "'"
    Koneksi "SELECT id_anggota FROM tb_anggota WHERE id_anggota='" & TextTujuan.Text & "'"
    If DB.EOF Then
        MsgBox "ID Anggota Tidak Di Ketahui", vbExclamation, "Masukan ID Anggota Dengan Benar"
    Else
        Koneksi "SELECT nama_anggota AS penampung_nama_anggota FROM tb_anggota WHERE id_anggota='" & TextTujuan.Text & "'"
        TextPenampungNamaAnggota.Text = DB!penampung_nama_anggota
        Koneksi "SELECT id_petugas AS penampung_id_petugas FROM tb_petugas WHERE nama_petugas='" & FormUtama.StatusBarUtama.Panels(4).Text & "'"
        TextPenampungIDPetugas.Text = DB!penampung_id_petugas
        
        Conn.Execute "INSERT INTO tb_request(id_pengirim, nama_pengirim, nama_penerima, id_anggota_tujuan, perihal_request, konten_request, tanggal_request, waktu_request, status_request) VALUES('" & TextPenampungIDPetugas.Text & "','" & FormUtama.StatusBarUtama.Panels(4).Text & "','" & TextPenampungNamaAnggota.Text & "','" & TextTujuan.Text & "','" & TextJudul.Text & "','" & TextMessage.Text & "','" & Format(Date, "YYYY-MM-DD") & "','" & Format(Time, "HH:MM:SS") & "','Belum Di Baca')"
        MsgBox "Pesan Telah Di Kirim Ke " & TextPenampungNamaAnggota.Text & "", vbInformation, "Pesan Berhasil Di Kirim"
        TextJudul.Text = ""
        TextMessage.Text = ""
        TextTujuan.Text = ""
        MasukanMessagePetugas
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
        TextKonten.Text = DB!konten_request
        TextKonten.Locked = True
        If DB!id_pengirim <> FormUtama.StatusBarUtama.Panels(5).Text Or DB!nama_pengirim <> FormUtama.StatusBarUtama.Panels(4).Text Then
            Conn.Execute "UPDATE tb_request SET status_request='Di Baca' WHERE id_request='" & ListMessage.SelectedItem.Text & "'"
            CmdBalas.Enabled = True
        Else
            CmdBalas.Enabled = False
        End If
        MasukanMessagePetugas
        Koneksi "SELECT nama_petugas, nama_anggota FROM tb_petugas, tb_anggota WHERE nama_petugas='" & LabelPengirim.Caption & "' OR nama_anggota='" & LabelPengirim.Caption & "'"
        If LabelPengirim.Caption = DB!nama_petugas Then
            LabelKategori.Caption = "Petugas - Petugas"
        ElseIf LabelPengirim.Caption = DB!nama_anggota Then
            LabelKategori.Caption = "Anggota - Petugas"
        End If
    Else
        MsgBox "Tidak Ada Pesan", vbExclamation, "Tidak Ada Pesan"
    End If
Else
    MsgBox "Tidak Ada Pesan", vbExclamation, "Tidak Ada Pesan"
End If
End Sub

Private Sub CmdLihatLogDenda_Click()
FramePesanBaru.Visible = False
FrameMessage.Visible = False
FramePinjam.Visible = False
FrameKembali.Visible = False
FrameDenda.Visible = True
End Sub

Private Sub CmdLihatLogKembali_Click()
FramePesanBaru.Visible = False
FrameMessage.Visible = False
FrameDenda.Visible = False
FramePinjam.Visible = False
FrameKembali.Visible = True
End Sub

Private Sub CmdLihatLogMessage_Click()
FramePesanBaru.Visible = False
FramePinjam.Visible = False
FrameKembali.Visible = False
FrameDenda.Visible = False
FrameMessage.Visible = True
End Sub

Private Sub CmdLihatLogPinjam_Click()
FramePesanBaru.Visible = False
FrameMessage.Visible = False
FrameKembali.Visible = False
FrameDenda.Visible = False
FramePinjam.Visible = True
End Sub

Private Sub CmdPesanBaru_Click()
FramePinjam.Visible = False
FrameKembali.Visible = False
FrameDenda.Visible = False
FrameMessage.Visible = False
FramePesanBaru.Visible = True
Call CmdRefresh_Click
End Sub

Private Sub CmdPinjamDetail_Click()
If ListPinjam.ListItems.Count > 0 Then
    Koneksi "SELECT id_pinjam_detail, nama_buku, jumlah_buku, tanggal_pinjam, status_pinjam_detail FROM tb_pinjam_detail  WHERE id_pinjam='" & ListPinjam.SelectedItem.Text & "' AND status_pinjam_detail='Kembali'"
    IsiLogPinjamDetail ListPinjamDetail
    Call Default
Else
    MsgBox "Tidak Terdapat Data Pinjam", vbInformation, "Data Pinjam Tidak Tersedia"
End If
End Sub

Private Sub CmdPinjamkan_Click()
Call CmdPinjamDetail_Click
If ListPinjamDetail.ListItems.Count > 0 Then
    MsgBox "Masih Ada Data Detail Yang Belum Di Rubah", vbInformation, "Silahkan Rubah Data Detail Terlebih Dahulu"
Else
    If ListPinjam.ListItems.Count > 0 Then
        If MsgBox("Rubah Status Pinjam Nomer " & ListPinjam.SelectedItem.Text & " Menjadi Pinjam ?", vbQuestion + vbYesNo, "Rubah Status Kembali Menjadi Pinjam") = vbYes Then
        Conn.Execute "UPDATE tb_pinjam SET status_pinjam='Pinjam' WHERE id_pinjam='" & ListPinjam.SelectedItem.Text & "'"
        MasukanLogPinjam
        Call Default
        MsgBox "Status Peminjaman Berhasil Di Rubah Statusnya Menjadi Pinjam", vbInformation, "Berhasil Merubah Status"
        End If
    Else
        MsgBox "Tidak Ada Data Pinjam Yang Bisa Di Hapus", vbCritical, "Data Pinjam Tidak Tersedia"
    End If
End If
End Sub

Private Sub CmdPinjamkanDetail_Click()
Call CmdPinjamDetail_Click
If ListPinjamDetail.ListItems.Count > 0 Then
    If MsgBox("Rubah Status Detail Nomer " & ListPinjamDetail.SelectedItem.Text & " Menjadi Pinjam ?", vbQuestion + vbYesNo, "Rubah Status Kembali Menjadi Pinjam") = vbYes Then
        Koneksi "SELECT id_pinjam_detail FROM tb_kembali"
        Conn.Execute "DELETE FROM tb_kembali WHERE id_pinjam_detail='" & ListPinjamDetail.SelectedItem.Text & "'"
        Conn.Execute "UPDATE tb_pinjam_detail SET status_pinjam_detail='Pinjam' WHERE id_pinjam_detail='" & ListPinjamDetail.SelectedItem.Text & "'"
        Koneksi "SELECT id_buku FROM tb_pinjam_detail WHERE id_pinjam_detail = '" & ListPinjamDetail.SelectedItem.Text & "'"
        Conn.Execute "UPDATE tb_buku SET jumlah_buku=jumlah_buku - '" & ListPinjamDetail.SelectedItem.SubItems(2) & "' WHERE id_buku='" & DB!ID_Buku & "'"
        MasukanLogPinjam
        MasukanLogKembali
        Call Default
        MsgBox "Status Detail Berhasil Di Rubah Statusnya Menjadi Pinjam", vbInformation, "Berhasil Merubah Status"
        Call CmdPinjamDetail_Click
    End If
Else
    MsgBox "Tidak Ada Data Detail Yang Bisa Di Hapus", vbCritical, "Data Detail Tidak Tersedia"
End If
End Sub

Private Sub CmdRefresh_Click()
MasukanMessagePetugas
TextKonten.Text = ""
LabelHeaderMessage.Caption = "-"
LabelPengirim.Caption = "-"
LabelPenerima.Caption = "-"
LabelTanggal.Caption = "-"
LabelWaktu.Caption = "-"
LabelKategori.Caption = "-"
CmdBalas.Enabled = False
CmdBalas.Caption = "Balas"
End Sub

Private Sub CmdRefreshAnggota_Click()
Call MasukanLogAnggota
Call DefaultUbahDenda
End Sub

Private Sub CmdUbahDenda_Click()
If ListAnggota.ListItems.Count > 0 Then
    If CmdUbahDenda.Caption = "Selesai" Then
        CmdUbahDenda.Caption = "Ubah"
        FrameUbahDenda.Visible = False
        With ListAnggota
            .Height = 2775
            .Left = 120
            .Top = 3000
            .Width = 6375
        End With
        
        With ListDenda
            .Height = 2775
            .Left = 120
            .Top = 240
            .Width = 8055
        End With
        
        With ListAnggotaKembali
            .Height = 2775
            .Left = 6480
            .Top = 3000
            .Width = 3375
        End With
        CmdCariDenda.Enabled = True
        CmdHapusDenda.Enabled = True
        CmdRefreshAnggota.Enabled = True
    ElseIf CmdUbahDenda.Caption = "Ubah" Then
        CmdUbahDenda.Caption = "Selesai"
        FrameUbahDenda.Visible = True
        With ListAnggota
            .Height = 1935
            .Left = 120
            .Top = 2280
            .Width = 6375
        End With
        
        With ListDenda
            .Height = 2055
            .Left = 120
            .Top = 240
            .Width = 8055
        End With
        
        With ListAnggotaKembali
            .Height = 1935
            .Left = 6480
            .Top = 2280
            .Width = 3375
        End With
        CmdCariDenda.Enabled = False
        CmdHapusDenda.Enabled = False
        CmdRefreshAnggota.Enabled = False
    End If
Else
    MsgBox "Tidak Ada Data Anggota Tersedia", vbExclamation, "Tambahkan Data Anggota Terlebih Dahulu"
    FormAnggota.Show
    Unload FormLog
End If
End Sub

Private Sub ComboCariKategori_Click()
If ComboCariKategori.Text = "ID" Then
    TextCari.ToolTipText = "Format ID Integer(25)"
ElseIf ComboCariKategori.Text = "ID Anggota" Then
    TextCari.ToolTipText = "Format ID Anggota Integer(250)"
ElseIf ComboCariKategori.Text = "ID Detail" Then
    TextCari.ToolTipText = "Format ID Detail Integer(250)"
ElseIf ComboCariKategori.Text = "Banyak Denda" Then
    TextCari.ToolTipText = "Format Banyak Denda Integer(250)"
ElseIf ComboCariKategori.Text = "Tanggal Denda" Then
    TextCari.ToolTipText = "Format Tanggal Denda YYYY-MM-DD(1997-07-05)"
ElseIf ComboCariKategori.Text = "Petugas" Then
    TextCari.ToolTipText = "Format Petugas VarChar(250)"
End If
End Sub

Private Sub ComboCariKategoriDetail_Click()
If ComboCariKategoriDetail.Text = "ID Detail" Then
    TextCariDetail.ToolTipText = "Format ID Detail Integer(250)"
ElseIf ComboCariKategoriDetail.Text = "Nama Buku" Then
    TextCariDetail.ToolTipText = "Format Nama Buku VarChar(250)"
ElseIf ComboCariKategoriDetail.Text = "Jumlah Buku" Then
    TextCariDetail.ToolTipText = "Format Jumlah Buku Integer(250)"
ElseIf ComboCariKategoriDetail.Text = "Tanggal Pinjam" Then
    TextCariDetail.ToolTipText = "Format Tanggal Pinjam YYYY-MM-DD(1997-07-05)"
ElseIf ComboCariKategoriDetail.Text = "Status" Then
    TextCariDetail.ToolTipText = "Format Status Enum(Pinjam/Kembali)"
End If
End Sub

Private Sub ComboCariKategoriKembali_Click()
If ComboCariKategoriKembali.Text = "ID Kembali" Then
    TextCariKembali.ToolTipText = "Format ID Kembali Integer(250)"
ElseIf ComboCariKategoriKembali.Text = "ID Detail" Then
    TextCariKembali.ToolTipText = "Format ID Detail Integer(250)"
ElseIf ComboCariKategoriKembali.Text = "ID Anggota" Then
    TextCariKembali.ToolTipText = "Format ID Anggota Integer(250)"
ElseIf ComboCariKategoriKembali.Text = "Petugas" Then
    TextCariKembali.ToolTipText = "Format Petugas VarChar(250)"
ElseIf ComboCariKategoriKembali.Text = "ID Detail" Then
    TextCariKembali.ToolTipText = "Format ID Detail Integer(250)"
ElseIf ComboCariKategoriKembali.Text = "Nama Anggota" Then
    TextCariKembali.ToolTipText = "Format Nama Anggota String(250)"
ElseIf ComboCariKategoriKembali.Text = "Nama Buku" Then
    TextCariKembali.ToolTipText = "Format Nama Buku VarChar(250)"
ElseIf ComboCariKategoriKembali.Text = "Jumlah Buku" Then
    TextCariKembali.ToolTipText = "Format Jumlah Buku Integer(250)"
ElseIf ComboCariKategoriKembali.Text = "Tanggal Pinjam" Then
    TextCariKembali.ToolTipText = "Format Tanggal Pinjam YYYY-MM-DD(1997-07-05)"
ElseIf ComboCariKategoriKembali.Text = "Tanggal Kembali" Then
    TextCariKembali.ToolTipText = "Format Tanggal Kembali YYYY-MM-DD(1997-07-05)"
ElseIf ComboCariKategoriKembali.Text = "Jumlah Denda" Then
    TextCariKembali.ToolTipText = "Format Jumlah Denda Integer(250)"
End If
End Sub

Private Sub ComboCariKategoriPinjam_Click()
If ComboCariKategoriPinjam.Text = "ID Pinjam" Then
    TextCariPinjam.ToolTipText = "Format ID Pinjam Integer(250)"
ElseIf ComboCariKategoriPinjam.Text = "ID Anggota" Then
    TextCariPinjam.ToolTipText = "Format ID Anggota Integer(250)"
ElseIf ComboCariKategoriPinjam.Text = "ID Anggota" Then
    TextCariPinjam.ToolTipText = "Format ID Anggota Integer(250)"
ElseIf ComboCariKategoriPinjam.Text = "Nama Anggota" Then
    TextCariPinjam.ToolTipText = "Format Nama Anggota String(250)"
ElseIf ComboCariKategoriPinjam.Text = "Petugas" Then
    TextCariPinjam.ToolTipText = "Format Petugas VarChar(250)"
ElseIf ComboCariKategoriPinjam.Text = "Tanggal Pinjam" Then
    TextCariPinjam.ToolTipText = "Format Tanggal Pinjam YYYY-MM-DD(1997-07-05)"
ElseIf ComboCariKategoriPinjam.Text = "Status" Then
    TextCariPinjam.ToolTipText = "Format Status Enum(Pinjam/Kembali)"
End If
End Sub

Private Sub ComboIDAnggota_Change()
ComboIDPinjamDetail.Clear
Koneksi "SELECT * FROM tb_anggota WHERE id_anggota='" & ComboIDAnggota.Text & "'"
Do While Not DB.EOF
    TextNamaAnggota.Text = DB!nama_anggota
    TextJumlahDendaAwal.Text = DB!total_denda
    DB.MoveNext
Loop

Koneksi "SELECT id_anggota, nama_anggota, total_denda FROM tb_anggota WHERE id_anggota='" & ComboIDAnggota.Text & "'"
IsiLogAnggota ListAnggota

Koneksi "SELECT id_pinjam_detail FROM tb_kembali WHERE id_anggota='" & ComboIDAnggota.Text & "'"
Do While Not DB.EOF
    ComboIDPinjamDetail.AddItem DB!id_pinjam_detail
    ComboIDPinjamDetail.SetFocus
    DB.MoveNext
    Loop
    
Koneksi "SELECT id_kembali, id_pinjam_detail, id_anggota FROM tb_kembali WHERE id_anggota='" & ComboIDAnggota.Text & "'"
IsiLogAnggotaKembali ListAnggotaKembali
End Sub

Private Sub ComboIDAnggota_Click()
ComboIDPinjamDetail.Clear
Koneksi "SELECT * FROM tb_anggota WHERE id_anggota='" & ComboIDAnggota.Text & "'"
Do While Not DB.EOF
    TextNamaAnggota.Text = DB!nama_anggota
    TextJumlahDendaAwal.Text = DB!total_denda
    DB.MoveNext
Loop

Koneksi "SELECT id_anggota, nama_anggota, total_denda FROM tb_anggota WHERE id_anggota='" & ComboIDAnggota.Text & "'"
IsiLogAnggota FormLog.ListAnggota

Koneksi "SELECT id_pinjam_detail FROM tb_kembali WHERE id_anggota='" & ComboIDAnggota.Text & "'"
Do While Not DB.EOF
    ComboIDPinjamDetail.AddItem DB!id_pinjam_detail
    ComboIDPinjamDetail.SetFocus
    DB.MoveNext
    Loop
        
Koneksi "SELECT id_kembali, id_pinjam_detail, id_anggota FROM tb_kembali WHERE id_anggota='" & ComboIDAnggota.Text & "'"
IsiLogAnggotaKembali ListAnggotaKembali
End Sub

Private Sub ComboKategoriTujuan_Change()
If ComboKategoriTujuan.Text = "Petugas" Then
    Koneksi "SELECT id_petugas, nama_petugas FROM tb_petugas ORDER BY id_petugas ASC"
    IsiCariAnggota ListCariAnggota
ElseIf ComboKategoriTujuan.Text = "Anggota" Then
    Koneksi "SELECT id_anggota, nama_anggota FROM tb_anggota ORDER BY id_anggota ASC"
    IsiCariAnggota ListCariAnggota
End If
End Sub

Private Sub ComboKategoriTujuan_Click()
If ComboKategoriTujuan.Text = "Petugas" Then
    Koneksi "SELECT id_petugas, nama_petugas FROM tb_petugas ORDER BY id_petugas ASC"
    IsiCariAnggota ListCariAnggota
ElseIf ComboKategoriTujuan.Text = "Anggota" Then
    Koneksi "SELECT id_anggota, nama_anggota FROM tb_anggota ORDER BY id_anggota ASC"
    IsiCariAnggota ListCariAnggota
End If
End Sub

Private Sub Form_Load()
    With ListPinjam
        .LabelEdit = lvwManual
        .FullRowSelect = True
        .GridLines = True
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "ID", "1200"
        .ColumnHeaders.Add , , "ID Anggota", "1200"
        .ColumnHeaders.Add , , "Peminjam", "3300"
        .ColumnHeaders.Add , , "Petugas", "1500"
        .ColumnHeaders.Add , , "Tanggal Pinjam"
        .ColumnHeaders.Add , , "Status", "1000"
    End With
    
    With ListPinjamDetail
        .LabelEdit = lvwManual
        .FullRowSelect = True
        .GridLines = True
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "ID Detail", "1000"
        .ColumnHeaders.Add , , "Nama Buku", "4000"
        .ColumnHeaders.Add , , "Jumlah Buku", "1200"
        .ColumnHeaders.Add , , "Tanggal Pinjam"
        .ColumnHeaders.Add , , "Status", "1050"
    End With
    
    With ListKembali
        .LabelEdit = lvwManual
        .FullRowSelect = True
        .GridLines = True
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "ID", "1000"
        .ColumnHeaders.Add , , "ID Detail", "1000"
        .ColumnHeaders.Add , , "ID Anggota", "1000"
        .ColumnHeaders.Add , , "Petugas", "1500"
        .ColumnHeaders.Add , , "Peminjam", "2000"
        .ColumnHeaders.Add , , "Nama Buku", "3000"
        .ColumnHeaders.Add , , "Jumlah Buku", "1200"
        .ColumnHeaders.Add , , "Tanggal Pinjam"
        .ColumnHeaders.Add , , "Tanggal Kembali"
        .ColumnHeaders.Add , , "Denda"
    End With
    
    With ListDenda
        .LabelEdit = lvwManual
        .FullRowSelect = True
        .GridLines = True
        .Height = 2775
        .Left = 120
        .Top = 240
        .Width = 8055
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "ID Denda", "1000"
        .ColumnHeaders.Add , , "ID Anggota", "1000"
        .ColumnHeaders.Add , , "ID Detail", "1000"
        .ColumnHeaders.Add , , "Banyak Denda", "2000"
        .ColumnHeaders.Add , , "Tanggal Denda", "1500"
        .ColumnHeaders.Add , , "Petugas", "1450"
    End With
        
    With ListAnggota
        .LabelEdit = lvwManual
        .FullRowSelect = True
        .GridLines = True
        .Height = 2775
        .Left = 120
        .Top = 3000
        .Width = 6375
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "ID Anggota", "1000"
        .ColumnHeaders.Add , , "Nama Anggota", "3300"
        .ColumnHeaders.Add , , "Total Denda", "2000"
    End With
  
    With ListAnggotaKembali
        .LabelEdit = lvwManual
        .FullRowSelect = True
        .GridLines = True
        .Height = 2775
        .Left = 6480
        .Top = 3000
        .Width = 3375
        .ColumnHeaders.Add , , "ID Kembali", "1100"
        .ColumnHeaders.Add , , "ID Detail", "1100"
        .ColumnHeaders.Add , , "ID Anggota", "1100"
    End With
    
    With ComboCariKategori
        .AddItem "ID"
        .AddItem "ID Anggota"
        .AddItem "ID Detail"
        .AddItem "Banyak Denda"
        .AddItem "Tanggal Denda"
        .AddItem "Petugas"
    End With
    
    With ComboCariKategoriPinjam
        .AddItem "ID Pinjam"
        .AddItem "ID Anggota"
        .AddItem "Nama Anggota"
        .AddItem "Petugas"
        .AddItem "Tanggal Pinjam"
        .AddItem "Status"
        .ToolTipText = "Cari Pinjaman Utama"
    End With
    
    With ComboCariKategoriDetail
        .AddItem "ID Detail"
        .AddItem "Nama Buku"
        .AddItem "Jumlah Buku"
        .AddItem "Tanggal Pinjam"
        .AddItem "Status"
        .ToolTipText = "Cari Pinjaman Detail"
    End With
    
    With ComboCariKategoriKembali
        .AddItem "ID Kembali"
        .AddItem "ID Detail"
        .AddItem "ID Anggota"
        .AddItem "Petugas"
        .AddItem "Nama Anggota"
        .AddItem "Nama Buku"
        .AddItem "Jumlah Buku"
        .AddItem "Tanggal Pinjam"
        .AddItem "Tanggal Kembali"
        .AddItem "Jumlah Denda"
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
    
    With ComboTujuan
        .Clear
        .AddItem "Petugas"
        .AddItem "Anggota"
    End With
    
    With ComboKategoriTujuan
        .Clear
        .AddItem "Petugas"
        .AddItem "Anggota"
    End With
    
    With ComboKategoriCariAnggota
        .Clear
        .AddItem "ID"
        .AddItem "Nama"
    End With
    
    With ListCariAnggota
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "ID", "660"
        .ColumnHeaders.Add , , "Nama", "2000"
        .GridLines = True
        .FullRowSelect = True
        .LabelEdit = lvwManual
    End With
    
Call MasukanLogPinjam
Call MasukanLogKembali
Call MasukanLogDenda
Call MasukanLogAnggota
Call MasukanLogIDAnggota
Call MasukanLogAnggotaKembali
Call CmdLihatLogDenda_Click
FrameUbahDenda.Visible = False
FrameCariDenda.Visible = False
FrameCariPinjamDetail.Visible = False

MasukanMessagePetugas
End Sub

Public Sub MasukanMessagePetugas()
Koneksi "SELECT id_petugas FROM tb_petugas WHERE nama_petugas='" & FormUtama.StatusBarUtama.Panels(4).Text & "'"
Koneksi "SELECT id_request, perihal_request, status_request, id_pengirim, nama_pengirim, nama_penerima, konten_request, tanggal_request, waktu_request FROM tb_request WHERE id_pengirim='" & DB!id_petugas & "' OR id_petugas_tujuan='" & DB!id_petugas & "' ORDER BY status_request DESC"
IsiMessagePetugas ListMessage
End Sub

Private Sub TextCari_Change()
TextCari = FilterInjeksi(TextCari.Text)

If ComboCariKategori.Text = "" Then
    MsgBox "Silahkan Pilih Kategori Pencarian Data Denda", vbExclamation, "Pilih Kategori"
ElseIf ComboCariKategori.Text = "ID" Then
    Koneksi "SELECT * FROM tb_denda WHERE id_denda LIKE '%' '" & TextCari & "' '%'"
    IsiDenda ListDenda
ElseIf ComboCariKategori.Text = "ID Anggota" Then
    Koneksi "SELECT * FROM tb_denda WHERE id_anggota LIKE '%' '" & TextCari & "' '%'"
    IsiDenda ListDenda
ElseIf ComboCariKategori.Text = "ID Detail" Then
    Koneksi "SELECT * FROM tb_denda WHERE id_pinjam_detail LIKE '%' '" & TextCari & "' '%'"
    IsiDenda ListDenda
ElseIf ComboCariKategori.Text = "Banyak Denda" Then
    Koneksi "SELECT * FROM tb_denda WHERE banyak_denda LIKE '%' '" & TextCari & "' '%'"
    IsiDenda ListDenda
ElseIf ComboCariKategori.Text = "Tanggal Denda" Then
    Koneksi "SELECT * FROM tb_denda WHERE tanggal_denda LIKE '%' '" & TextCari & "' '%'"
    IsiDenda ListDenda
ElseIf ComboCariKategori.Text = "Petugas" Then
    Koneksi "SELECT * FROM tb_denda WHERE petugas_denda LIKE '%' '" & TextCari & "' '%'"
    IsiDenda ListDenda
End If
End Sub

Private Sub TextCariAnggota_Change()
TextCariAnggota = FilterInjeksi(TextCariAnggota.Text)

If ComboKategoriTujuan.Text = "" Then
    MsgBox "Pilih Kriteria Anggota / Petugas", vbExclamation, "Pilih Kriteria"
ElseIf ComboKategoriCariAnggota.Text = "" Then
    MsgBox "Pilih Kategori Pencarian", vbExclamation, "Pilih Kategori Pencarian"
ElseIf ComboKategoriTujuan.Text = "Petugas" Then
    If ComboKategoriCariAnggota.Text = "ID" Then
        Koneksi "SELECT id_petugas, nama_petugas FROM tb_petugas WHERE id_petugas LIKE '%' '" & TextCariAnggota & "' '%'"
        IsiCariAnggota ListCariAnggota
    ElseIf ComboKategoriCariAnggota.Text = "Nama" Then
        Koneksi "SELECT id_petugas, nama_petugas FROM tb_petugas WHERE nama_petugas LIKE '%' '" & TextCariAnggota & "' '%'"
        IsiCariAnggota ListCariAnggota
    End If
ElseIf ComboKategoriTujuan.Text = "Anggota" Then
    If ComboKategoriCariAnggota.Text = "ID" Then
        Koneksi "SELECT id_anggota, nama_anggota FROM tb_anggota WHERE id_anggota LIKE '%' '" & TextCariAnggota & "' '%'"
        IsiCariAnggota ListCariAnggota
    ElseIf ComboKategoriCariAnggota.Text = "Nama" Then
        Koneksi "SELECT id_anggota, nama_anggota FROM tb_anggota WHERE nama_anggota LIKE '%' '" & TextCariAnggota & "' '%'"
        IsiCariAnggota ListCariAnggota
    End If
End If
End Sub

Private Sub TextCariDetail_Change()
TextCariDetail = FilterInjeksi(TextCariDetail.Text)

If ComboCariKategoriDetail.Text = "" Then
    MsgBox "Silahkan Pilih Kategori Pencarian Data Detail", vbExclamation, "Pilih Kategori"
ElseIf ComboCariKategoriDetail.Text = "ID Detail" Then
    Koneksi "SELECT id_pinjam_detail, nama_buku, jumlah_buku, tanggal_pinjam, status_pinjam_detail FROM tb_pinjam_detail WHERE id_pinjam_detail LIKE '%' '" & TextCariDetail & "' '%' AND status_pinjam_detail='Kembali'"
    IsiLogPinjamDetail ListPinjamDetail
ElseIf ComboCariKategoriDetail.Text = "Nama Buku" Then
    Koneksi "SELECT id_pinjam_detail, nama_buku, jumlah_buku, tanggal_pinjam, status_pinjam_detail FROM tb_pinjam_detail WHERE nama_buku LIKE '%' '" & TextCariDetail & "' '%' AND status_pinjam_detail='Kembali'"
    IsiLogPinjamDetail ListPinjamDetail
ElseIf ComboCariKategoriDetail.Text = "Jumlah Buku" Then
    Koneksi "SELECT id_pinjam_detail, nama_buku, jumlah_buku, tanggal_pinjam, status_pinjam_detail FROM tb_pinjam_detail WHERE jumlah_buku LIKE '%' '" & TextCariDetail & "' '%' AND status_pinjam_detail='Kembali'"
    IsiLogPinjamDetail ListPinjamDetail
ElseIf ComboCariKategoriDetail.Text = "Tanggal Pinjam" Then
    Koneksi "SELECT id_pinjam_detail, nama_buku, jumlah_buku, tanggal_pinjam, status_pinjam_detail FROM tb_pinjam_detail WHERE tanggal_pinjam LIKE '%' '" & TextCariDetail & "' '%' AND status_pinjam_detail='Kembali'"
    IsiLogPinjamDetail ListPinjamDetail
ElseIf ComboCariKategoriDetail.Text = "Status" Then
    Koneksi "SELECT id_pinjam_detail, nama_buku, jumlah_buku, tanggal_pinjam, status_pinjam_detail FROM tb_pinjam_detail WHERE status_pinjam_detail LIKE '%' '" & TextCariDetail & "' '%'"
    IsiLogPinjamDetail ListPinjamDetail
End If
End Sub

Private Sub TextCariKembali_Change()
TextCariKembali = FilterInjeksi(TextCariKembali.Text)

If ComboCariKategoriKembali.Text = "" Then
    MsgBox "Silahkan Pilih Kategori Pencarian Data Kembali", vbExclamation, "Pilih Kategori"
ElseIf ComboCariKategoriKembali.Text = "ID Kembali" Then
    Koneksi "SELECT * FROM tb_kembali WHERE id_kembali LIKE '%' '" & TextCariKembali & "' '%'"
    IsiLogKembali ListKembali
ElseIf ComboCariKategoriKembali.Text = "ID Detail" Then
    Koneksi "SELECT * FROM tb_kembali WHERE id_pinjam_detail LIKE '%' '" & TextCariKembali & "' '%'"
    IsiLogKembali ListKembali
ElseIf ComboCariKategoriKembali.Text = "ID Anggota" Then
    Koneksi "SELECT * FROM tb_kembali WHERE id_anggota LIKE '%' '" & TextCariKembali & "' '%'"
    IsiLogKembali ListKembali
ElseIf ComboCariKategoriKembali.Text = "Petugas" Then
    Koneksi "SELECT * FROM tb_kembali WHERE nama_petugas LIKE '%' '" & TextCariKembali & "' '%'"
    IsiLogKembali ListKembali
ElseIf ComboCariKategoriKembali.Text = "Nama Anggota" Then
    Koneksi "SELECT * FROM tb_kembali WHERE nama_anggota LIKE '%' '" & TextCariKembali & "' '%'"
    IsiLogKembali ListKembali
ElseIf ComboCariKategoriKembali.Text = "Nama Buku" Then
    Koneksi "SELECT * FROM tb_kembali WHERE nama_buku LIKE '%' '" & TextCariKembali & "' '%'"
    IsiLogKembali ListKembali
ElseIf ComboCariKategoriKembali.Text = "Jumlah Buku" Then
    Koneksi "SELECT * FROM tb_kembali WHERE jumlah_buku LIKE '%' '" & TextCariKembali & "' '%'"
    IsiLogKembali ListKembali
ElseIf ComboCariKategoriKembali.Text = "Tanggal Pinjam" Then
    Koneksi "SELECT * FROM tb_kembali WHERE tanggal_pinjam LIKE '%' '" & TextCariKembali & "' '%'"
    IsiLogKembali ListKembali
ElseIf ComboCariKategoriKembali.Text = "Tanggal Kembali" Then
    Koneksi "SELECT * FROM tb_kembali WHERE tanggal_kembali LIKE '%' '" & TextCariKembali & "' '%'"
    IsiLogKembali ListKembali
ElseIf ComboCariKategoriKembali.Text = "Jumlah Denda" Then
    Koneksi "SELECT * FROM tb_kembali WHERE denda_kembali LIKE '%' '" & TextCariKembali & "' '%'"
    IsiLogKembali ListKembali
End If
End Sub

Private Sub TextCariPinjam_Change()
TextCariPinjam = FilterInjeksi(TextCariPinjam.Text)

ListPinjamDetail.ListItems.Clear
If ComboCariKategoriPinjam.Text = "" Then
    MsgBox "Silahkan Pilih Kategori Pencarian Data Pinjam", vbExclamation, "Pilih Kategori"
ElseIf ComboCariKategoriPinjam.Text = "ID Pinjam" Then
    Koneksi "SELECT * FROM tb_pinjam WHERE id_pinjam LIKE '%' '" & TextCariPinjam & "' '%' AND status_pinjam='Kembali'"
    IsiLogPinjam ListPinjam
ElseIf ComboCariKategoriPinjam.Text = "ID Anggota" Then
    Koneksi "SELECT * FROM tb_pinjam WHERE id_anggota LIKE '%' '" & TextCariPinjam & "' '%' AND status_pinjam='Kembali'"
    IsiLogPinjam ListPinjam
ElseIf ComboCariKategoriPinjam.Text = "Nama Anggota" Then
    Koneksi "SELECT * FROM tb_pinjam WHERE nama_anggota LIKE '%' '" & TextCariPinjam & "' '%' AND status_pinjam='Kembali'"
    IsiLogPinjam ListPinjam
ElseIf ComboCariKategoriPinjam.Text = "Petugas" Then
    Koneksi "SELECT * FROM tb_pinjam WHERE nama_petugas LIKE '%' '" & TextCariPinjam & "' '%' AND status_pinjam='Kembali'"
    IsiLogPinjam ListPinjam
ElseIf ComboCariKategoriPinjam.Text = "Tanggal Pinjam" Then
    Koneksi "SELECT * FROM tb_pinjam WHERE tanggal_pinjam LIKE '%' '" & TextCariPinjam & "' '%' AND status_pinjam='Kembali'"
    IsiLogPinjam ListPinjam
ElseIf ComboCariKategoriPinjam.Text = "Status" Then
    Koneksi "SELECT * FROM tb_pinjam WHERE status_pinjam LIKE '%' '" & TextCariPinjam & "' '%'"
    IsiLogPinjam ListPinjam
End If
End Sub

Private Sub TextJumlahDendaEdit_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= Asc("0") & Chr(13) _
     And KeyAscii <= Asc("9") & Chr(13) _
     Or KeyAscii = vbKeyBack _
     Or KeyAscii = vbKeyDelete _
     Or KeyAscii = vbKeySpace) Then
    TextJumlahDendaEdit.Text = "0"
    KeyAscii = 0
End If
End Sub

Private Sub TextJumlahDendaEdit_click()
If TextJumlahDendaEdit.Text = "Jumlah Denda" Then
    TextJumlahDendaEdit.Text = ""
ElseIf TextJumlahDendaEdit.Text = "" Then
    TextJumlahDendaEdit.Text = "Jumlah Denda"
End If
End Sub

Public Sub DefaultUbahDenda()
TextJumlahDendaAwal.Text = "Jumlah Denda Awal"
TextJumlahDendaEdit.Text = "Jumlah Denda"
TextNamaAnggota.Text = "Nama Anggota"
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
