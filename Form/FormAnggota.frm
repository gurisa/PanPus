VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form FormAnggota 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Anggota"
   ClientHeight    =   4350
   ClientLeft      =   -15
   ClientTop       =   330
   ClientWidth     =   10110
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
   Icon            =   "FormAnggota.frx":0000
   LinkTopic       =   "Anggota"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4350
   ScaleWidth      =   10110
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame FrameKategoriAnggota 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Kategori Anggota"
      Height          =   1455
      Left            =   120
      TabIndex        =   22
      Top             =   120
      Width           =   9855
      Begin VB.TextBox TextKategoriLembaga 
         Height          =   315
         Left            =   4800
         TabIndex        =   37
         Top             =   960
         Width           =   3495
      End
      Begin VB.ComboBox ComboKategoriLembaga 
         Height          =   330
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   36
         Top             =   960
         Width           =   4575
      End
      Begin VB.CommandButton CmdTambahKategoriLembaga 
         Caption         =   "+"
         Height          =   255
         Left            =   8400
         TabIndex        =   35
         Top             =   960
         Width           =   615
      End
      Begin VB.CommandButton CmdHapusKategoriLembaga 
         Caption         =   "-"
         Height          =   255
         Left            =   9120
         TabIndex        =   34
         Top             =   960
         Width           =   495
      End
      Begin VB.CommandButton CmdHapusKategoriJurusan 
         Caption         =   "-"
         Height          =   255
         Left            =   9120
         TabIndex        =   30
         Top             =   600
         Width           =   495
      End
      Begin VB.CommandButton CmdTambahKategoriJurusan 
         Caption         =   "+"
         Height          =   255
         Left            =   8400
         TabIndex        =   29
         Top             =   600
         Width           =   615
      End
      Begin VB.ComboBox ComboKategoriJurusan 
         Height          =   330
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   600
         Width           =   4575
      End
      Begin VB.TextBox TextKategoriJurusan 
         Height          =   315
         Left            =   4800
         TabIndex        =   27
         Top             =   600
         Width           =   3495
      End
      Begin VB.CommandButton CmdHapusKategoriTingkat 
         Caption         =   "-"
         Height          =   255
         Left            =   9120
         TabIndex        =   26
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton CmdTambahKategoriTingkat 
         Caption         =   "+"
         Height          =   255
         Left            =   8400
         TabIndex        =   25
         Top             =   240
         Width           =   615
      End
      Begin VB.ComboBox ComboKategoriTingkat 
         Height          =   330
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   240
         Width           =   4575
      End
      Begin VB.TextBox TextKategoriTingkat 
         Height          =   315
         Left            =   4800
         TabIndex        =   23
         Top             =   240
         Width           =   3495
      End
   End
   Begin Crystal.CrystalReport KartuAnggota 
      Left            =   7560
      Top             =   3960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin VB.CommandButton CmdKartuAnggota 
      Caption         =   "Cetak"
      Height          =   375
      Left            =   5880
      TabIndex        =   21
      Top             =   3840
      Width           =   1335
   End
   Begin VB.Frame FrameCariAnggota 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Cari Anggota "
      Height          =   1455
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Width           =   9855
      Begin VB.TextBox TextCari 
         Height          =   315
         Left            =   2400
         TabIndex        =   18
         Top             =   840
         Width           =   4095
      End
      Begin VB.ComboBox ComboCariKategori 
         Height          =   330
         Left            =   2400
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   360
         Width           =   4095
      End
      Begin VB.Image ImageCariAnggota 
         Height          =   1095
         Left            =   600
         Picture         =   "FormAnggota.frx":0CCA
         Stretch         =   -1  'True
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.CommandButton CmdCariAnggota 
      Caption         =   "Cari"
      Height          =   375
      Left            =   1560
      TabIndex        =   15
      Top             =   3840
      Width           =   1335
   End
   Begin VB.CommandButton CmdDaftarAnggota 
      Caption         =   "Daftar"
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   3840
      Width           =   1335
   End
   Begin VB.Frame FrameDaftarAnggota 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Daftar Anggota "
      Height          =   1455
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   9855
      Begin VB.CommandButton CmdDaftarJurusan 
         Caption         =   "+"
         Height          =   255
         Left            =   7920
         TabIndex        =   33
         Top             =   960
         Width           =   255
      End
      Begin VB.CommandButton CmdDaftarTingkat 
         Caption         =   "+"
         Height          =   255
         Left            =   4560
         TabIndex        =   32
         Top             =   960
         Width           =   255
      End
      Begin VB.CommandButton CmdDaftarLembaga 
         Caption         =   "+"
         Height          =   255
         Left            =   3240
         TabIndex        =   31
         Top             =   960
         Width           =   255
      End
      Begin VB.TextBox TextKonfirmasiPassword 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   4440
         PasswordChar    =   "*"
         TabIndex        =   20
         Top             =   600
         Width           =   4215
      End
      Begin VB.TextBox TextPassword 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   4440
         PasswordChar    =   "*"
         TabIndex        =   19
         Top             =   240
         Width           =   4215
      End
      Begin VB.ComboBox ComboSekolah 
         Height          =   330
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   960
         Width           =   3135
      End
      Begin VB.CommandButton CmdTambah 
         Caption         =   "Tambah"
         Height          =   495
         Left            =   8760
         TabIndex        =   12
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox TextNama 
         Height          =   315
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   4215
      End
      Begin VB.ComboBox ComboJurusan 
         Height          =   330
         Left            =   4920
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   960
         Width           =   3015
      End
      Begin VB.ComboBox ComboKelas 
         Height          =   330
         Left            =   3600
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   960
         Width           =   975
      End
      Begin VB.ComboBox ComboStatus 
         Height          =   330
         Left            =   8280
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox TextNIS 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   2520
         TabIndex        =   7
         Top             =   600
         Width           =   1815
      End
      Begin VB.OptionButton OptionFemale 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Wanita"
         Height          =   255
         Left            =   1200
         TabIndex        =   6
         Top             =   600
         Width           =   855
      End
      Begin VB.OptionButton OptionMale 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Pria"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   735
      End
   End
   Begin MSComctlLib.ListView ListAnggota 
      Height          =   2055
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   3625
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.CommandButton CmdHapus 
      Caption         =   "Hapus"
      Height          =   375
      Left            =   4440
      TabIndex        =   2
      Top             =   3840
      Width           =   1335
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "Keluar"
      Height          =   375
      Left            =   8520
      TabIndex        =   1
      Top             =   3840
      Width           =   1455
   End
   Begin VB.CommandButton CmdUbah 
      Caption         =   "Ubah"
      Height          =   375
      Left            =   3000
      TabIndex        =   0
      Top             =   3840
      Width           =   1335
   End
End
Attribute VB_Name = "FormAnggota"
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

Dim ID_Anggota As Integer

Private Sub CmdCariAnggota_Click()
If CmdCariAnggota.Caption = "Cari" Then
    CmdDaftarAnggota.Enabled = False
    CmdUbah.Enabled = False
    CmdHapus.Enabled = False
    CmdKartuAnggota.Enabled = False
    CmdCariAnggota.Caption = "Selesai"
    With ListAnggota
        .Width = 9855
        .Height = 2055
        .Top = 1680
        .Left = 120
    End With
    FrameCariAnggota.Visible = True
    FrameKategoriAnggota.Visible = False
ElseIf CmdCariAnggota.Caption = "Selesai" Then
    CmdDaftarAnggota.Enabled = True
    CmdUbah.Enabled = True
    CmdHapus.Enabled = True
    CmdKartuAnggota.Enabled = True
    CmdCariAnggota.Caption = "Cari"
    FrameCariAnggota.Visible = False
    With ListAnggota
        .Left = 120
        .Top = 120
        .Height = 3615
        .Width = 9855
    End With
    Koneksi "SELECT * FROM tb_anggota"
    IsiAnggota ListAnggota
    TextCari.Text = ""
End If
End Sub

Private Sub CmdDaftarAnggota_Click()
If CmdDaftarAnggota.Caption = "Selesai" Then
    TextNama.Text = ""
    TextNIS.Text = ""
    TextPassword.Text = ""
    TextKonfirmasiPassword.Text = ""
    FrameDaftarAnggota.Visible = False
    CmdDaftarAnggota.Caption = "Daftar"
    ListAnggota.Top = 120
    With ListAnggota
        .Left = 120
        .Top = 120
        .Height = 3615
        .Width = 9855
    End With
    CmdCariAnggota.Enabled = True
    CmdUbah.Enabled = True
    CmdHapus.Enabled = True
    CmdKartuAnggota.Enabled = True
ElseIf CmdDaftarAnggota.Caption = "Daftar" Then
    TextNama.Text = ""
    TextNIS.Text = ""
    TextPassword.Text = ""
    TextKonfirmasiPassword.Text = ""
    FrameDaftarAnggota.Visible = True
    FrameKategoriAnggota.Visible = False
    CmdDaftarAnggota.Caption = "Selesai"
    With ListAnggota
        .Left = 120
        .Top = 1680
        .Height = 2055
        .Width = 9855
    End With
    CmdCariAnggota.Enabled = False
    CmdUbah.Enabled = False
    CmdHapus.Enabled = False
    CmdKartuAnggota.Enabled = False
End If
FrameCariAnggota.Visible = False
End Sub

Private Sub CmdDaftarJurusan_Click()
Call BukaKategoriAnggota
End Sub

Private Sub CmdDaftarLembaga_Click()
Call BukaKategoriAnggota
End Sub

Public Sub BukaKategoriAnggota()
FrameKategoriAnggota.Visible = True
FrameCariAnggota.Visible = False
FrameDaftarAnggota.Visible = False
CmdCariAnggota.Enabled = True
CmdUbah.Enabled = True
CmdHapus.Enabled = True
CmdKartuAnggota.Enabled = True
CmdDaftarAnggota.Caption = "Daftar"
End Sub

Private Sub CmdDaftarTingkat_Click()
Call BukaKategoriAnggota
End Sub

Private Sub CmdExit_Click()
Unload Me
End Sub

Private Sub MasukanDataAnggota()
Koneksi "SELECT * FROM tb_anggota ORDER BY nama_anggota ASC"
IsiAnggota ListAnggota
End Sub

Private Sub CmdHapus_Click()
If ListAnggota.ListItems.Count > 0 Then
    If MsgBox("Hapus Data Anggota Dengan Nomor " & ListAnggota.SelectedItem.Text & " ?", vbQuestion + vbYesNo, "Hapus Data Anggota") = vbYes Then
        Conn.Execute "DELETE FROM tb_anggota WHERE id_anggota='" & ListAnggota.SelectedItem.Text & "'"
        MasukanDataAnggota
        Call Default
        MsgBox "Data Anggota Berhasil Di Hapus", vbInformation, "Data Anggota Berhasil Di Hapus"
    Else
        Exit Sub
    End If
Else
    MsgBox "Tidak Ada Data Anggota Yang Bisa Di Hapus", vbCritical, "Tidak Ada Data Yang Bisa Di Hapus Lagi"
End If
End Sub

Private Sub CmdHapusKategoriJurusan_Click()
If ComboKategoriJurusan.Text = "" Then
    MsgBox "Pilih Kategori Jurusan Yang Akan Di Hapus", vbExclamation, "Pilih Kategori Jurusan"
Else
    If MsgBox("Hapus Kategori Jurusan " & ComboKategoriJurusan.Text & " ?", vbInformation + vbYesNo, "Hapus Kategori Jurusan") = vbYes Then
        Conn.Execute "DELETE FROM tb_kategori_jurusan WHERE nama_kategori_jurusan='" & ComboKategoriJurusan.Text & "'"
        Call KategoriJurusan
        MsgBox "Kategori Jurusan Berhasil Di Hapus", vbInformation, "Berhasil Menghapus"
    Else
        Exit Sub
    End If
End If
End Sub

Private Sub CmdHapusKategoriLembaga_Click()
If ComboKategoriLembaga.Text = "" Then
    MsgBox "Pilih Kategori Lembaga Yang Akan Di Hapus", vbExclamation, "Pilih Kategori Lembaga"
Else
    If MsgBox("Hapus Kategori Lembaga " & ComboKategoriLembaga.Text & " ?", vbInformation + vbYesNo, "Hapus Kategori Lembaga") = vbYes Then
        Conn.Execute "DELETE FROM tb_kategori_lembaga WHERE nama_kategori_lembaga='" & ComboKategoriLembaga.Text & "'"
        Call KategoriLembaga
        MsgBox "Kategori Lembaga Berhasil Di Hapus", vbInformation, "Berhasil Menghapus"
    Else
        Exit Sub
    End If
End If
End Sub

Private Sub CmdHapusKategoriTingkat_Click()
If ComboKategoriTingkat.Text = "" Then
    MsgBox "Pilih Kategori Tingkat Yang Akan Di Hapus", vbExclamation, "Pilih Kategori Tingkat"
Else
    If MsgBox("Hapus Kategori Tingkat " & ComboKategoriTingkat.Text & " ?", vbInformation + vbYesNo, "Hapus Kategori Tingkat") = vbYes Then
        Conn.Execute "DELETE FROM tb_kategori_tingkat WHERE nama_kategori_tingkat='" & ComboKategoriTingkat.Text & "'"
        Call KategoriTingkat
        MsgBox "Kategori Tingkat Berhasil Di Hapus", vbInformation, "Berhasil Menghapus"
    Else
        Exit Sub
    End If
End If
End Sub

Private Sub CmdKartuAnggota_Click()
If ListAnggota.ListItems.Count > 0 Then
    Koneksi "SELECT status_anggota FROM tb_anggota WHERE id_anggota='" & ListAnggota.SelectedItem.Text & "'"
    If DB!status_anggota = "Aktif" Then
            With KartuAnggota
                .WindowTitle = "Panda Pustaka"
                .ReportFileName = App.Path & "\Report\LaporanKartuAnggota.rpt"
                .SelectionFormula = "{tb_anggota.id_anggota} = " & ListAnggota.SelectedItem.Text & ""
                .RetrieveDataFiles
                .WindowShowPrintSetupBtn = True
                .WindowShowPrintBtn = True
                .Action = 1
            End With
    ElseIf DB!status_anggota = "Tidak Aktif" Then
        MsgBox "Status Keanggotaan " & ListAnggota.SelectedItem.SubItems(2) & " Belum Aktif", vbExclamation, "Anggota Belum Resmi Aktif"
    End If
Else
    MsgBox "Data Anggota Tidak Tersedia", vbExclamation, "Tidak Terdapat Data Anggota"
End If
End Sub

Private Sub CmdTambah_Click()
TextNama = FilterInjeksi(TextNama.Text)
TextPassword = FilterInjeksi(TextPassword.Text)
TextNIS = FilterInjeksi(TextNIS.Text)
TextKonfirmasiPassword = FilterInjeksi(TextKonfirmasiPassword.Text)

If AntiInjeksi(TextNama.Text) = False Or AntiInjeksi(TextPassword.Text) = False Or AntiInjeksi(TextNIS.Text) Or AntiInjeksi(TextKonfirmasiPassword.Text) = False Then
    If TextPassword.Text <> TextKonfirmasiPassword.Text Or TextKonfirmasiPassword.Text <> TextPassword.Text Then
        MsgBox "Password Yang Di Masukan Tidak Sama", vbExclamation, "Silahkan Periksa Password Kembali"
    Else
        If TextNIS.Text = "" Or TextNama.Text = "" Or ComboJurusan.Text = "" Or ComboKelas.Text = "" Or ComboStatus.Text = "" Or ComboSekolah.Text = "" Then
            MsgBox "Masukan Data", vbCritical, "Masukan Data"
        ElseIf OptionMale.Value = True Then
            Conn.Execute "INSERT INTO tb_anggota(nis_anggota, nama_anggota, jenis_kelamin, kelas_anggota, jurusan_anggota, status_anggota, sekolah_anggota, tanggal_daftar, petugas_daftar, password_anggota) VALUES('" & TextNIS.Text & "','" & TextNama.Text & "','Pria','" & ComboKelas.Text & "','" & ComboJurusan.Text & "','" & ComboStatus.Text & "','" & ComboSekolah.Text & "','" & FormUtama.StatusBarUtama.Panels(1).Text & "','" & FormUtama.StatusBarUtama.Panels(4).Text & "','" & TextPassword.Text & "')"
            MsgBox "Berhasil Menambahkan Anggota " & TextNama.Text & "", vbInformation, "Data Anggota Berhasil Di Tambahkan"
            TextNama.Text = ""
            TextNIS.Text = ""
            TextPassword.Text = ""
            TextKonfirmasiPassword.Text = ""
            Call Default
        ElseIf OptionFemale.Value = True Then
            Conn.Execute "INSERT INTO tb_anggota(nis_anggota, nama_anggota, jenis_kelamin, kelas_anggota, jurusan_anggota, status_anggota, sekolah_anggota, tanggal_daftar, petugas_daftar, password_anggota) VALUES('" & TextNIS.Text & "','" & TextNama.Text & "','Wanita','" & ComboKelas.Text & "','" & ComboJurusan.Text & "','" & ComboStatus.Text & "','" & ComboSekolah.Text & "','" & FormUtama.StatusBarUtama.Panels(1).Text & "','" & FormUtama.StatusBarUtama.Panels(4).Text & "','" & TextPassword.Text & "')"
            MsgBox "Berhasil Menambahkan Anggota " & TextNama.Text & "", vbInformation, "Data Anggota Berhasil Di Tambahkan"
            TextNama.Text = ""
            TextNIS.Text = ""
            TextPassword.Text = ""
            TextKonfirmasiPassword.Text = ""
            Call Default
        End If
            MasukanDataAnggota
    End If
Else
    Call ErrorInjeksi
End If
End Sub

Private Sub CmdTambahKategoriJurusan_Click()
TextKategoriJurusan = FilterInjeksi(TextKategoriJurusan.Text)

If AntiInjeksi(TextKategoriJurusan.Text) = False Then
    If TextKategoriJurusan.Text = "" Then
        MsgBox "Silahkan Tulis Kategori Jurusan", vbExclamation, "Kategori Jurusan"
        TextKategoriJurusan.SetFocus
    Else
    Koneksi "SELECT nama_kategori_jurusan FROM tb_kategori_jurusan WHERE nama_kategori_jurusan='" & TextKategoriJurusan & "'"
        If DB.RecordCount > 0 Then
            MsgBox "Kategori Jurusan Sudah Ada", vbExclamation, "Kategori Jurusan Sudah Ada"
        Else
            If MsgBox("Tambahkan Kategori Jurusan " & TextKategoriJurusan.Text & " ?", vbInformation + vbYesNo, "Tambah Kategori Jurusan") = vbYes Then
                Conn.Execute "INSERT INTO tb_kategori_jurusan(nama_kategori_jurusan)VALUES('" & TextKategoriJurusan & "')"
                Call KategoriJurusan
                TextKategoriJurusan.Text = ""
                MsgBox "Kategori Jurusan Berhasil Di Tambahkan", vbInformation, "Berhasil Menambahkan Kategori Jurusan"
            Else
                Exit Sub
            End If
        End If
    End If
Else
    Call ErrorInjeksi
End If
End Sub

Private Sub CmdTambahKategoriLembaga_Click()
TextKategoriLembaga = FilterInjeksi(TextKategoriLembaga.Text)

If AntiInjeksi(TextKategoriLembaga.Text) = False Then
    If TextKategoriLembaga.Text = "" Then
        MsgBox "Silahkan Tulis Kategori Lembaga", vbExclamation, "Kategori Lembaga"
        TextKategoriLembaga.SetFocus
    Else
    Koneksi "SELECT nama_kategori_lembaga FROM tb_kategori_lembaga WHERE nama_kategori_lembaga='" & TextKategoriLembaga & "'"
        If DB.RecordCount > 0 Then
            MsgBox "Kategori Lembaga Sudah Ada", vbExclamation, "Kategori Lembaga Sudah Ada"
        Else
            If MsgBox("Tambahkan Kategori Lembaga " & TextKategoriLembaga.Text & " ?", vbInformation + vbYesNo, "Tambah Kategori Lembaga") = vbYes Then
                Conn.Execute "INSERT INTO tb_kategori_lembaga(nama_kategori_lembaga)VALUES('" & TextKategoriLembaga & "')"
                Call KategoriLembaga
                TextKategoriLembaga.Text = ""
                MsgBox "Kategori Lembaga Berhasil Di Tambahkan", vbInformation, "Berhasil Menambahkan Kategori Lembaga"
            Else
                Exit Sub
            End If
        End If
    End If
Else
    Call ErrorInjeksi
End If
End Sub

Private Sub CmdTambahKategoriTingkat_Click()
TextKategoriTingkat = FilterInjeksi(TextKategoriTingkat.Text)

If AntiInjeksi(TextKategoriTingkat.Text) = False Then
    If TextKategoriTingkat.Text = "" Then
        MsgBox "Silahkan Tulis Kategori Tingkat", vbExclamation, "Kategori Tingkat"
        TextKategoriTingkat.SetFocus
    Else
    Koneksi "SELECT nama_kategori_tingkat FROM tb_kategori_tingkat WHERE nama_kategori_tingkat='" & TextKategoriTingkat.Text & "'"
        If DB.RecordCount > 0 Then
            MsgBox "Kategori Tingkat Sudah Ada", vbExclamation, "Kategori Tingkat Sudah Ada"
        Else
            If MsgBox("Tambahkan Kategori Tingkat " & TextKategoriTingkat.Text & " ?", vbInformation + vbYesNo, "Tambah Kategori Tingkat") = vbYes Then
                Conn.Execute "INSERT INTO tb_kategori_tingkat(nama_kategori_tingkat)VALUES('" & TextKategoriTingkat & "')"
                Call KategoriTingkat
                TextKategoriTingkat.Text = ""
                MsgBox "Kategori Tingkat Berhasil Di Tambahkan", vbInformation, "Berhasil Menambahkan Kategori Tingkat"
            Else
                Exit Sub
            End If
        End If
    End If
Else
    Call ErrorInjeksi
End If
End Sub

Private Sub CmdUbah_Click()
If ListAnggota.ListItems.Count = 0 Then Exit Sub
    If CmdUbah.Caption = "Ubah" Then
        FrameDaftarAnggota.Visible = True
        FrameDaftarAnggota.Caption = "&Ubah Anggota"
        FrameCariAnggota.Visible = False
        FrameKategoriAnggota.Visible = False
        CmdDaftarTingkat.Enabled = False
        CmdDaftarLembaga.Enabled = False
        CmdDaftarJurusan.Enabled = False
        With ListAnggota
            .Width = 9855
            .Height = 2055
            .Top = 1680
            .Left = 120
        End With
        CmdCariAnggota.Enabled = False
        CmdTambah.Enabled = False
        CmdDaftarAnggota.Enabled = False
        CmdHapus.Enabled = False
        CmdKartuAnggota.Enabled = False
        CmdUbah.Caption = "Simpan"
        ID_Anggota = ListAnggota.SelectedItem.Text
        TextNIS.Text = ListAnggota.SelectedItem.SubItems(1)
        TextNama.Text = ListAnggota.SelectedItem.SubItems(2)
        If ListAnggota.SelectedItem.ListSubItems(3) = "Pria" Then
            OptionMale.Value = True
        ElseIf ListAnggota.SelectedItem.ListSubItems(3) = "Wanita" Then
            OptionFemale.Value = True
        End If
        ComboKelas.Text = ListAnggota.SelectedItem.SubItems(4)
        ComboJurusan.Text = ListAnggota.SelectedItem.SubItems(5)
        ComboStatus.Text = ListAnggota.SelectedItem.SubItems(6)
        ComboSekolah.Text = ListAnggota.SelectedItem.SubItems(7)
        TextPassword.Text = ListAnggota.SelectedItem.SubItems(11)
        TextKonfirmasiPassword.Text = ListAnggota.SelectedItem.SubItems(11)
    Else
        TextNama = FilterInjeksi(TextNama.Text)
        TextPassword = FilterInjeksi(TextPassword.Text)
        TextNIS = FilterInjeksi(TextNIS.Text)
        TextKonfirmasiPassword = FilterInjeksi(TextKonfirmasiPassword.Text)
        If AntiInjeksi(TextNama.Text) = False Or AntiInjeksi(TextPassword.Text) = False Or AntiInjeksi(TextNIS.Text) Or AntiInjeksi(TextKonfirmasiPassword.Text) = False Then
            If TextNama.Text = "" Or ComboJurusan.Text = "" Or ComboKelas.Text = "" Or ComboStatus.Text = "" Or ComboSekolah.Text = "" Then
                MsgBox "Lengkapi data", vbCritical, "Error"
            ElseIf ComboJurusan.Text = "Jurusan" Or ComboKelas.Text = "Kelas" Or ComboStatus.Text = "Status" Or ComboSekolah.Text = "Sekolah" Then
                MsgBox "Silahkan Pilih Data Yang Sudah Tersedia", vbInformation, "Pilih Data Anggota"
            Else
                If TextPassword.Text <> TextKonfirmasiPassword.Text Or TextKonfirmasiPassword.Text <> TextPassword.Text Then
                    MsgBox "Password Yang Di Masukan Tidak Sama", vbExclamation, "Silahkan Periksa Password Kembali"
                Else
                    If OptionMale.Value = True Then
                        Conn.Execute "UPDATE tb_anggota SET nis_anggota='" & TextNIS.Text & "', nama_Anggota='" & TextNama.Text & "', jenis_kelamin='Pria', kelas_anggota='" & ComboKelas.Text & "',jurusan_anggota='" & ComboJurusan.Text & "',status_anggota='" & ComboStatus.Text & "', sekolah_anggota='" & ComboSekolah.Text & "', password_anggota='" & TextPassword.Text & "' WHERE id_anggota='" & ID_Anggota & "'"
                        MsgBox "Berhasil Merubah Informasi Keanggotaan " & TextNama.Text & "", vbInformation, "Informasi Anggota Berhasil Di Perbaharui"
                    ElseIf OptionFemale.Value = True Then
                        Conn.Execute "UPDATE tb_anggota SET nis_anggota='" & TextNIS.Text & "', nama_Anggota='" & TextNama.Text & "', jenis_kelamin='Wanita', kelas_anggota='" & ComboKelas.Text & "',jurusan_anggota='" & ComboJurusan.Text & "',status_anggota='" & ComboStatus.Text & "', sekolah_anggota='" & ComboSekolah.Text & "', password_anggota='" & TextPassword.Text & "' WHERE id_anggota='" & ID_Anggota & "'"
                        MsgBox "Berhasil Merubah Informasi Keanggotaan " & TextNama.Text & "", vbInformation, "Informasi Anggota Berhasil Di Perbaharui"
                    End If
                    MasukanDataAnggota
                    CmdUbah.Caption = "Ubah"
                    Call Default
                    FrameDaftarAnggota.Visible = False
                    FrameDaftarAnggota.Caption = "&Daftar Anggota"
                    CmdDaftarTingkat.Enabled = True
                    CmdDaftarLembaga.Enabled = True
                    CmdDaftarJurusan.Enabled = True
                    With ListAnggota
                        .Left = 120
                        .Top = 120
                        .Height = 3615
                        .Width = 9855
                    End With
                    CmdTambah.Enabled = True
                    CmdDaftarAnggota.Enabled = True
                    CmdHapus.Enabled = True
                    CmdCariAnggota.Enabled = True
                    CmdKartuAnggota.Enabled = True
                End If
            End If
        Else
            Call ErrorInjeksi
        End If
End If
End Sub

Private Sub ComboCariKategori_Click()
If ComboCariKategori.Text = "ID" Then
    TextCari.ToolTipText = "Format ID Integer(250)"
ElseIf ComboCariKategori.Text = "NIS" Then
    TextCari.ToolTipText = "Format NIS VarChar XXX.XXX(250)"
ElseIf ComboCariKategori.Text = "Nama" Then
    TextCari.ToolTipText = "Format Nama String(250)"
ElseIf ComboCariKategori.Text = "Jenis Kelamin" Then
    TextCari.ToolTipText = "Format Jenis Kelamin Enum(Pria/Wanita)"
ElseIf ComboCariKategori.Text = "Kelas" Then
    TextCari.ToolTipText = "Format Kelas Enum(10/11/12)"
ElseIf ComboCariKategori.Text = "Jurusan" Then
    TextCari.ToolTipText = "Format Jurusan Enum(TKJ/Otomotif/Mesin/Listrik/Administrasi Perkantoran/Niaga/Jasa Boga/Farmasi')"
ElseIf ComboCariKategori.Text = "Status" Then
    TextCari.ToolTipText = "Format Status Enum(Aktif/Tidak Aktif)"
ElseIf ComboCariKategori.Text = "Sekolah" Then
    TextCari.ToolTipText = "Format Sekolah Enum(SMK WIRAKARYA 1 CIPARAY/SMK WIRAKARYA 2 CIPARAY/SMK AS-SHIFA CIPARAY)"
ElseIf ComboCariKategori.Text = "Tanggal Daftar" Then
    TextCari.ToolTipText = "Format Tanggal (YYYY-MM-DD)"
ElseIf ComboCariKategori.Text = "Petugas" Then
    TextCari.ToolTipText = "Format Petugas String(250)"
End If
End Sub

Private Sub Form_Load()
FrameDaftarAnggota.Visible = False
FrameCariAnggota.Visible = False
FrameKategoriAnggota.Visible = False
With ListAnggota
    .Left = 120
    .Top = 120
    .Height = 3615
    .Width = 9855
End With

With ListAnggota.ColumnHeaders
    .Clear
    .Add , , "ID", "800"
    .Add , , "NIS", "1200"
    .Add , , "Nama", "2500"
    .Add , , "Jenis Kelamin", "1500"
    .Add , , "Kelas", "800"
    .Add , , "Jurusan", "2200"
    .Add , , "Status", "1000"
    .Add , , "Sekolah", "2500"
    .Add , , "Tanggal Daftar"
    .Add , , "Petugas"
    .Add , , "Total Denda"
    .Add , , "Password"
End With

ComboStatus.AddItem "Aktif"
ComboStatus.AddItem "Tidak Aktif"

MasukanKelasAnggota
MasukanJurusanAnggota
MasukanSekolahAnggota
MasukanDataAnggota

ListAnggota.GridLines = True
ListAnggota.FullRowSelect = True
ListAnggota.LabelEdit = lvwManual

With ComboCariKategori
    .AddItem "ID"
    .AddItem "NIS"
    .AddItem "Nama"
    .AddItem "Jenis Kelamin"
    .AddItem "Kelas"
    .AddItem "Jurusan"
    .AddItem "Status"
    .AddItem "Sekolah"
    .AddItem "Tanggal Daftar"
    .AddItem "Petugas"
End With

TextNama.ToolTipText = "Nama Anggota"
TextPassword.ToolTipText = "Password"
TextKonfirmasiPassword.ToolTipText = "Konfirmasi Password"
OptionMale.ToolTipText = "Jenis Kelamin"
OptionFemale.ToolTipText = "Jenis Kelamin"
TextNIS.ToolTipText = "Nomor Induk Siswa"
ComboSekolah.ToolTipText = "Sekolah"
ComboKelas.ToolTipText = "Kelas"
ComboJurusan.ToolTipText = "Jurusan"
ComboStatus.ToolTipText = "Status"

Call KategoriTingkat
Call KategoriJurusan
Call KategoriLembaga
End Sub

Private Sub TextCari_Change()
TextCari = FilterInjeksi(TextCari.Text)

If ComboCariKategori.Text = "" Then
    MsgBox "Silahkan Pilih Kategori Pencarian Data Anggota", vbExclamation, "Pilih Kategori"
ElseIf ComboCariKategori.Text = "ID" Then
    Koneksi "SELECT * FROM tb_anggota WHERE id_anggota LIKE '%' '" & TextCari & "' '%'"
    IsiAnggota ListAnggota
ElseIf ComboCariKategori.Text = "NIS" Then
    Koneksi "SELECT * FROM tb_anggota WHERE nis_anggota LIKE '%' '" & TextCari & "' '%'"
    IsiAnggota ListAnggota
ElseIf ComboCariKategori.Text = "Nama" Then
    Koneksi "SELECT * FROM tb_anggota WHERE nama_anggota LIKE '%' '" & TextCari & "' '%'"
    IsiAnggota ListAnggota
ElseIf ComboCariKategori.Text = "Jenis Kelamin" Then
    Koneksi "SELECT * FROM tb_anggota WHERE jenis_kelamin LIKE '%' '" & TextCari & "' '%'"
    IsiAnggota ListAnggota
ElseIf ComboCariKategori.Text = "Kelas" Then
    Koneksi "SELECT * FROM tb_anggota WHERE kelas_anggota LIKE '%' '" & TextCari & "' '%'"
    IsiAnggota ListAnggota
ElseIf ComboCariKategori.Text = "Jurusan" Then
    Koneksi "SELECT * FROM tb_anggota WHERE jurusan_anggota LIKE '%' '" & TextCari & "' '%'"
    IsiAnggota ListAnggota
ElseIf ComboCariKategori.Text = "Status" Then
    Koneksi "SELECT * FROM tb_anggota WHERE status_anggota LIKE '%' '" & TextCari & "' '%'"
    IsiAnggota ListAnggota
ElseIf ComboCariKategori.Text = "Sekolah" Then
    Koneksi "SELECT * FROM tb_anggota WHERE sekolah_anggota LIKE '%' '" & TextCari & "' '%'"
    IsiAnggota ListAnggota
ElseIf ComboCariKategori.Text = "Tanggal Daftar" Then
    Koneksi "SELECT * FROM tb_anggota WHERE tanggal_daftar LIKE '%' '" & TextCari & "' '%'"
    IsiAnggota ListAnggota
ElseIf ComboCariKategori.Text = "Petugas" Then
    Koneksi "SELECT * FROM tb_anggota WHERE petugas_daftar LIKE '%' '" & TextCari & "' '%'"
    IsiAnggota ListAnggota
End If
End Sub

