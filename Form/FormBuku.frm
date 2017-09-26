VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FormBuku 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Buku"
   ClientHeight    =   5760
   ClientLeft      =   -15
   ClientTop       =   330
   ClientWidth     =   8445
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormBuku.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   8445
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame FrameKategoriBuku 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Kategori Buku"
      Height          =   2295
      Left            =   120
      TabIndex        =   17
      Top             =   120
      Width           =   8175
      Begin VB.CommandButton CmdHapusKategoriBuku 
         Caption         =   "Hapus"
         Height          =   375
         Left            =   6840
         TabIndex        =   21
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox TextKategoriBuku 
         Height          =   375
         Left            =   2280
         MaxLength       =   250
         TabIndex        =   20
         Top             =   960
         Width           =   3255
      End
      Begin VB.CommandButton CmdTambahKategoriBuku 
         Caption         =   "Tambah"
         Height          =   375
         Left            =   5640
         TabIndex        =   19
         Top             =   960
         Width           =   1095
      End
      Begin VB.ComboBox ComboDaftarKategoriBuku 
         Height          =   330
         Left            =   2280
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   480
         Width           =   5655
      End
      Begin VB.Image ImageKategoriBuku 
         Height          =   1695
         Left            =   240
         Picture         =   "FormBuku.frx":0CCA
         Stretch         =   -1  'True
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.CommandButton CmdCariBuku 
      Caption         =   "Cari"
      Height          =   375
      Left            =   1440
      TabIndex        =   16
      Top             =   5160
      Width           =   1215
   End
   Begin VB.Frame FrameCariBuku 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Cari Buku "
      Height          =   2295
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   8175
      Begin VB.ComboBox ComboCariKategori 
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   480
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   600
         Width           =   3975
      End
      Begin VB.TextBox TextCari 
         Height          =   375
         Left            =   480
         TabIndex        =   14
         Top             =   1080
         Width           =   3975
      End
      Begin VB.Image ImageCariBuku 
         Height          =   1695
         Left            =   5760
         Picture         =   "FormBuku.frx":1279B
         Stretch         =   -1  'True
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.CommandButton CmdDaftarBuku 
      Caption         =   "Daftar"
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   5160
      Width           =   1215
   End
   Begin VB.Frame FrameDaftarBuku 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Daftar Buku "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   8175
      Begin VB.CommandButton CmdDaftarKategoriBuku 
         Caption         =   "+"
         Height          =   255
         Left            =   4200
         TabIndex        =   22
         Top             =   1320
         Width           =   375
      End
      Begin VB.CommandButton CmdTambah 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Tambah"
         Height          =   375
         Left            =   6840
         TabIndex        =   11
         Top             =   1800
         Width           =   1215
      End
      Begin VB.TextBox TextNamaBuku 
         Height          =   375
         Left            =   120
         MaxLength       =   250
         ScrollBars      =   1  'Horizontal
         TabIndex        =   10
         Top             =   360
         Width           =   7935
      End
      Begin VB.TextBox TextJumlahBuku 
         Height          =   375
         Left            =   120
         MaxLength       =   250
         TabIndex        =   9
         Top             =   840
         Width           =   2175
      End
      Begin VB.TextBox TextNamaPengarang 
         Height          =   405
         Left            =   4680
         MaxLength       =   250
         ScrollBars      =   1  'Horizontal
         TabIndex        =   8
         Top             =   840
         Width           =   3375
      End
      Begin VB.TextBox TextNamaPenerbit 
         Height          =   375
         Left            =   4680
         MaxLength       =   250
         ScrollBars      =   1  'Horizontal
         TabIndex        =   7
         Top             =   1320
         Width           =   3375
      End
      Begin VB.TextBox TextTahunTerbit 
         Height          =   375
         Left            =   2400
         MaxLength       =   5
         TabIndex        =   6
         Top             =   840
         Width           =   2175
      End
      Begin VB.ComboBox ComboKategoriBuku 
         Height          =   330
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1320
         Width           =   3975
      End
   End
   Begin MSComctlLib.ListView ListBuku 
      Height          =   2535
      Left            =   120
      TabIndex        =   3
      Top             =   2520
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   4471
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
   Begin VB.CommandButton CmdExit 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Keluar"
      Height          =   375
      Left            =   7080
      TabIndex        =   2
      Top             =   5160
      Width           =   1215
   End
   Begin VB.CommandButton CmdUbah 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ubah"
      Height          =   375
      Left            =   2760
      TabIndex        =   1
      Top             =   5160
      Width           =   1215
   End
   Begin VB.CommandButton CmdHapus 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Hapus"
      Height          =   375
      Left            =   4080
      TabIndex        =   0
      Top             =   5160
      Width           =   1215
   End
End
Attribute VB_Name = "FormBuku"
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

Dim ID_Buku As Integer

Private Sub CmdCariBuku_Click()
If CmdCariBuku.Caption = "Cari" Then
    CmdCariBuku.Caption = "Selesai"
    With ListBuku
        .Top = 2520
        .Left = 120
        .Width = 8175
        .Height = 2535
    End With
    FrameCariBuku.Visible = True
    FrameKategoriBuku.Visible = False
    FrameDaftarBuku.Visible = False
    CmdUbah.Enabled = False
    CmdDaftarBuku.Enabled = False
    CmdHapus.Enabled = False
ElseIf CmdCariBuku.Caption = "Selesai" Then
    CmdCariBuku.Caption = "Cari"
    With ListBuku
        .Top = 120
        .Left = 120
        .Width = 8175
        .Height = 4935
    End With
    FrameCariBuku.Visible = False
    FrameKategoriBuku.Visible = False
    FrameDaftarBuku.Visible = False
    CmdUbah.Enabled = True
    CmdDaftarBuku.Enabled = True
    CmdHapus.Enabled = True
    
    Koneksi "SELECT * FROM tb_buku"
    IsiBuku ListBuku
End If
End Sub

Private Sub CmdDaftarBuku_Click()
If CmdDaftarBuku.Caption = "Daftar" Then
    CmdDaftarBuku.Caption = "Selesai"
    With ListBuku
        .Top = 2520
        .Left = 120
        .Width = 8175
        .Height = 2535
    End With
    FrameDaftarBuku.Visible = True
    FrameKategoriBuku.Visible = False
    FrameCariBuku.Visible = False
    FrameDaftarBuku.Caption = "&Daftar Buku"
    CmdUbah.Enabled = False
    CmdCariBuku.Enabled = False
    CmdHapus.Enabled = False
ElseIf CmdDaftarBuku.Caption = "Selesai" Then
    CmdDaftarBuku.Caption = "Daftar"
    With ListBuku
        .Top = 120
        .Left = 120
        .Width = 8175
        .Height = 4935
    End With
    FrameDaftarBuku.Visible = False
    FrameKategoriBuku.Visible = False
    FrameCariBuku.Visible = False
    CmdUbah.Enabled = True
    CmdCariBuku.Enabled = True
    CmdHapus.Enabled = True
End If
End Sub

Private Sub CmdDaftarKategoriBuku_Click()
FrameKategoriBuku.Visible = True
FrameDaftarBuku.Visible = False
FrameCariBuku.Visible = False
CmdUbah.Enabled = True
CmdCariBuku.Enabled = True
CmdHapus.Enabled = True
CmdDaftarBuku.Caption = "Daftar"
End Sub

Private Sub CmdExit_Click()
Unload Me
End Sub

Private Sub CmdHapus_Click()
If ListBuku.ListItems.Count > 0 Then
    If MsgBox("Hapus Data Buku Nomor " & ListBuku.SelectedItem.Text & " ?", vbQuestion + vbYesNo, "Hapus Data Buku") = vbYes Then
        Conn.Execute "DELETE FROM tb_buku WHERE id_buku='" & ListBuku.SelectedItem.Text & "'"
        MasukanDataBuku
        Call Default
        MsgBox "Data Buku Berhasil Di Hapus", vbInformation, "Data Buku Berhasil Di Hapus"
    Else
        Exit Sub
    End If
Else
    MsgBox "Tidak Ada Data Yang Bisa Di Hapus", vbInformation, "Data Buku Sudah Kosong"
End If
End Sub

Public Sub LihatDaftarKategoriBuku()
Koneksi "SELECT nama_kategori_buku FROM tb_buku_kategori ORDER BY nama_kategori_buku ASC"
ComboDaftarKategoriBuku.Clear
Do While Not DB.EOF
    ComboDaftarKategoriBuku.AddItem DB!nama_kategori_buku
    DB.MoveNext
Loop
End Sub

Private Sub CmdHapusKategoriBuku_Click()
If ComboDaftarKategoriBuku.Text = "" Then
    MsgBox "Pilih Kategori Buku Yang Akan Di Hapus", vbExclamation, "Pilih Kategori Buku"
Else
    If MsgBox("Hapus Kategori Buku " & ComboDaftarKategoriBuku.Text & " ?", vbInformation + vbYesNo, "Hapus Kategori Buku") = vbYes Then
        Conn.Execute "DELETE FROM tb_buku_kategori WHERE nama_kategori_buku='" & ComboDaftarKategoriBuku.Text & "'"
        Call LihatDaftarKategoriBuku
        Call KategoriBuku
        MsgBox "Kategori Buku Berhasil Di Hapus", vbInformation, "Berhasil Menghapus"
    Else
        Exit Sub
    End If
End If
End Sub

Private Sub CmdTambah_Click()
TextNamaBuku = FilterInjeksi(TextNamaBuku.Text)
TextJumlahBuku = FilterInjeksi(TextJumlahBuku.Text)
TextTahunTerbit = FilterInjeksi(TextTahunTerbit.Text)
TextNamaPengarang = FilterInjeksi(TextNamaPengarang.Text)
TextNamaPenerbit = FilterInjeksi(TextNamaPenerbit.Text)

If AntiInjeksi(TextNamaBuku.Text) = False Or AntiInjeksi(TextNamaPengarang.Text) = False Or AntiInjeksi(TextNamaPenerbit.Text) = False Then
    If TextNamaBuku.Text = "" Or TextJumlahBuku.Text = "" Or TextTahunTerbit.Text = "" Or TextNamaPenerbit.Text = "" Or TextNamaPengarang.Text = "" Or ComboKategoriBuku.Text = "" Then
        MsgBox "Data Buku Belum Lengkap", vbCritical, "Isi Data Buku"
        Exit Sub
    ElseIf TextNamaBuku.Text = "Nama Buku" Or TextJumlahBuku.Text = "Jumlah" Or TextTahunTerbit.Text = "Tahun Terbit" Or TextNamaPenerbit.Text = "Nama Penerbit" Or TextNamaPengarang.Text = "Nama Pengarang" Or ComboKategoriBuku.Text = "Kategori Buku" Then
        MsgBox "Silahkan Isi Data Buku Dengan Benar", vbCritical, "Isi Data Buku"
        Exit Sub
    Else
        Conn.Execute "INSERT INTO tb_buku(nama_buku, jumlah_buku, nama_pengarang, nama_penerbit, tahun_terbit, tanggal_daftar, petugas_daftar, kategori_buku) VALUES('" & TextNamaBuku & "','" & TextJumlahBuku & "','" & TextNamaPengarang & "','" & TextNamaPenerbit & "','" & TextTahunTerbit & "','" & FormUtama.StatusBarUtama.Panels(1) & "','" & FormUtama.StatusBarUtama.Panels(4) & "','" & ComboKategoriBuku.Text & "')"
        MasukanDataBuku
        MsgBox "Berhasil Menambahkan Buku " & TextNamaBuku.Text & "", vbInformation, "Berhasil Menambahkan Data Buku"
        Call Default
    End If
Else
    Call ErrorInjeksi
End If
End Sub

Private Sub MasukanDataBuku()
Koneksi "SELECT * FROM tb_buku ORDER BY tanggal_daftar DESC"
IsiBuku ListBuku
End Sub

Private Sub CmdTambahKategoriBuku_Click()
TextKategoriBuku = FilterInjeksi(TextKategoriBuku.Text)

If AntiInjeksi(TextKategoriBuku.Text) = False Then
    If TextKategoriBuku.Text = "" Then
        MsgBox "Silahkan Tulis Kategori Buku", vbExclamation, "Kategori Buku"
        TextKategoriBuku.SetFocus
    Else
    Koneksi "SELECT nama_kategori_buku FROM tb_buku_kategori WHERE nama_kategori_buku='" & TextKategoriBuku & "'"
        If DB.RecordCount > 0 Then
            MsgBox "Kategori Buku Sudah Ada", vbExclamation, "Kategori Buku Sudah Ada"
        Else
            If MsgBox("Tambahkan Kategori Buku " & TextKategoriBuku.Text & " ?", vbInformation + vbYesNo, "Tambah Kategori Buku") = vbYes Then
                Conn.Execute "INSERT INTO tb_buku_kategori(nama_kategori_buku)VALUES('" & TextKategoriBuku & "')"
                Call LihatDaftarKategoriBuku
                Call KategoriBuku
                TextKategoriBuku.Text = ""
                MsgBox "Kategori Buku Berhasil Di Tambahkan", vbInformation, "Berhasil Menambahkan Kategori Buku"
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
If ListBuku.ListItems.Count = 0 Then
    MsgBox "Tidak Ada Data Yang Bisa Di Rubah", vbCritical, "Tidak Ada Data Yang Bisa Di Rubah"
    Exit Sub
Else
    If CmdUbah.Caption = "Ubah" Then
        With ListBuku
            .Top = 2520
            .Left = 120
            .Width = 8175
            .Height = 2535
        End With
        CmdDaftarBuku.Enabled = False
        CmdTambah.Enabled = False
        CmdHapus.Enabled = False
        CmdCariBuku.Enabled = False
        CmdDaftarKategoriBuku.Enabled = False
        FrameDaftarBuku.Visible = True
        FrameKategoriBuku.Visible = False
        FrameCariBuku.Visible = False
        FrameDaftarBuku.Caption = "&Ubah Buku"
        CmdUbah.Caption = "Simpan"
        ID_Buku = ListBuku.SelectedItem.Text
        TextNamaBuku.Text = ListBuku.SelectedItem.SubItems(1)
        TextJumlahBuku.Text = ListBuku.SelectedItem.SubItems(2)
        TextNamaPengarang.Text = ListBuku.SelectedItem.SubItems(3)
        TextNamaPenerbit.Text = ListBuku.SelectedItem.SubItems(4)
        TextTahunTerbit.Text = ListBuku.SelectedItem.SubItems(5)
        ComboKategoriBuku.Text = ListBuku.SelectedItem.SubItems(8)
    Else
        TextNamaBuku = FilterInjeksi(TextNamaBuku.Text)
        TextJumlahBuku = FilterInjeksi(TextJumlahBuku.Text)
        TextTahunTerbit = FilterInjeksi(TextTahunTerbit.Text)
        TextNamaPengarang = FilterInjeksi(TextNamaPengarang.Text)
        TextNamaPenerbit = FilterInjeksi(TextNamaPenerbit.Text)
        
        If AntiInjeksi(TextNamaBuku.Text) = False Or AntiInjeksi(TextNamaPengarang.Text) = False Or AntiInjeksi(TextNamaPenerbit.Text) = False Then
            With ListBuku
                .Top = 120
                .Left = 120
                .Width = 8175
                .Height = 4935
            End With
            FrameDaftarBuku.Visible = False
            FrameKategoriBuku.Visible = False
            FrameCariBuku.Visible = False
            FrameDaftarBuku.Caption = "&Daftar Buku"
            CmdDaftarBuku.Enabled = True
            CmdTambah.Enabled = True
            CmdHapus.Enabled = True
            CmdCariBuku.Enabled = True
            CmdDaftarKategoriBuku.Enabled = True
            If TextNamaBuku.Text = "" Or TextJumlahBuku.Text = "" Or TextTahunTerbit.Text = "" Or TextNamaPengarang.Text = "" Or TextNamaPenerbit.Text = "" Or ComboKategoriBuku.Text = "" Or ComboKategoriBuku.Text = "Kategori" Then
                MsgBox "Silahkan Isi Data", vbCritical, "Data Belum Di Isi"
                Exit Sub
            End If
            Conn.Execute "UPDATE tb_buku SET nama_buku='" & TextNamaBuku & "', jumlah_buku='" & TextJumlahBuku & "',nama_pengarang='" & TextNamaPengarang & "',nama_penerbit='" & TextNamaPenerbit & "',tahun_terbit='" & TextTahunTerbit & "',kategori_buku='" & ComboKategoriBuku.Text & "' WHERE id_buku='" & ID_Buku & "'"
            MasukanDataBuku
            CmdUbah.Caption = "Ubah"
            MsgBox "Berhasil Merubah Informasi Buku " & TextNamaBuku.Text & " Dengan Nomor Detail " & ListBuku.SelectedItem.Text & "", vbInformation, "Data Buku Berhasil Di Rubah"
            Call Default
        Else
            Call ErrorInjeksi
        End If
    End If
End If
End Sub

Private Sub ComboCariKategori_Click()
If ComboCariKategori.Text = "ID" Then
    TextCari.ToolTipText = "Format ID Integer(250)"
ElseIf ComboCariKategori.Text = "Nama" Then
    TextCari.ToolTipText = "Format Nama String(250)"
ElseIf ComboCariKategori.Text = "Jumlah" Then
    TextCari.ToolTipText = "Format Jumlah Integer(250)"
ElseIf ComboCariKategori.Text = "Pengarang" Then
    TextCari.ToolTipText = "Format Pengarang String(250)"
ElseIf ComboCariKategori.Text = "Penerbit" Then
    TextCari.ToolTipText = "Format Penerbit String(250)"
ElseIf ComboCariKategori.Text = "Tahun Terbit" Then
    TextCari.ToolTipText = "Format Tahun Terbit Integer(4)"
ElseIf ComboCariKategori.Text = "Tanggal Daftar" Then
    TextCari.ToolTipText = "Format Tanggal Daftar Date(YYYY-MM-DD)"
ElseIf ComboCariKategori.Text = "Petugas" Then
    TextCari.ToolTipText = "Format Petugas String(250)"
ElseIf ComboCariKategori.Text = "Kategori" Then
    TextCari.ToolTipText = "Format Kategori Enum(Sekolah/BSE/Novel/Anak - Anak/Remaja/Umum)"
End If
End Sub

Private Sub Form_Load()
FrameDaftarBuku.Visible = False
FrameCariBuku.Visible = False
FrameKategoriBuku.Visible = False
With ListBuku
    .Top = 120
    .Left = 120
    .Width = 8175
    .Height = 4935
End With

With ListBuku.ColumnHeaders
    .Clear
    .Add , , "ID", "700"
    .Add , , "Nama Buku", "3000"
    .Add , , "Jumlah Buku"
    .Add , , "Nama Pengarang"
    .Add , , "Nama Penerbit"
    .Add , , "Tahun Terbit", "1200"
    .Add , , "Tanggal Daftar"
    .Add , , "Petugas"
    .Add , , "Kategori Buku"
End With

ListBuku.GridLines = True
ListBuku.FullRowSelect = True
ListBuku.LabelEdit = lvwManual

TextNamaBuku.ToolTipText = "Nama Buku"
TextJumlahBuku.ToolTipText = "Jumlah"
TextTahunTerbit.ToolTipText = "Tahun Terbit"
TextNamaPengarang.ToolTipText = "Nama Pengarang"
TextNamaPenerbit.ToolTipText = "Nama Penerbit"
ComboKategoriBuku.ToolTipText = "Kategori Buku"
TextJumlahBuku.MaxLength = 9

MasukanDataBuku
KategoriBuku
Call Default

With ComboCariKategori
    .AddItem "ID"
    .AddItem "Nama"
    .AddItem "Jumlah"
    .AddItem "Pengarang"
    .AddItem "Penerbit"
    .AddItem "Tahun Terbit"
    .AddItem "Tanggal Daftar"
    .AddItem "Petugas"
    .AddItem "Kategori"
End With

Call LihatDaftarKategoriBuku
End Sub

Private Sub TextCari_Change()
TextCari = FilterInjeksi(TextCari.Text)

If ComboCariKategori.Text = "" Then
    MsgBox "Silahkan Pilih Kategori Pencarian Data Buku", vbExclamation, "Pilih Kategori"
ElseIf ComboCariKategori.Text = "ID" Then
    Koneksi "SELECT * FROM tb_buku WHERE id_buku LIKE '%' '" & TextCari & "' '%'"
    IsiBuku ListBuku
ElseIf ComboCariKategori.Text = "Nama" Then
    Koneksi "SELECT * FROM tb_buku WHERE nama_buku LIKE '%' '" & TextCari & "' '%'"
    IsiBuku ListBuku
ElseIf ComboCariKategori.Text = "Jumlah" Then
    Koneksi "SELECT * FROM tb_buku WHERE jumlah_buku LIKE '%' '" & TextCari & "' '%'"
    IsiBuku ListBuku
ElseIf ComboCariKategori.Text = "Pengarang" Then
    Koneksi "SELECT * FROM tb_buku WHERE nama_pengarang LIKE '%' '" & TextCari & "' '%'"
    IsiBuku ListBuku
ElseIf ComboCariKategori.Text = "Penerbit" Then
    Koneksi "SELECT * FROM tb_buku WHERE nama_penerbit LIKE '%' '" & TextCari & "' '%'"
    IsiBuku ListBuku
ElseIf ComboCariKategori.Text = "Tahun Terbit" Then
    Koneksi "SELECT * FROM tb_buku WHERE tahun_terbit LIKE '%' '" & TextCari & "' '%'"
    IsiBuku ListBuku
ElseIf ComboCariKategori.Text = "Tanggal Daftar" Then
    Koneksi "SELECT * FROM tb_buku WHERE tanggal_daftar LIKE '%' '" & TextCari & "' '%'"
    IsiBuku ListBuku
ElseIf ComboCariKategori.Text = "Petugas" Then
    Koneksi "SELECT * FROM tb_buku WHERE petugas_daftar LIKE '%' '" & TextCari & "' '%'"
    IsiBuku ListBuku
ElseIf ComboCariKategori.Text = "Kategori" Then
    Koneksi "SELECT * FROM tb_buku WHERE kategori_buku LIKE '%' '" & TextCari & "' '%'"
    IsiBuku ListBuku
End If
End Sub

Private Sub TextJumlahBuku_Click()
If TextJumlahBuku.Text = "Jumlah" Then
    TextJumlahBuku.Text = ""
ElseIf TextJumlahBuku.Text = "" Then
    TextJumlahBuku.Text = "Jumlah"
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

Private Sub TextTahunTerbit_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= Asc("0") & Chr(13) _
     And KeyAscii <= Asc("9") & Chr(13) _
     Or KeyAscii = vbKeyBack _
     Or KeyAscii = vbKeyDelete _
     Or KeyAscii = vbKeySpace) Then
    TextTahunTerbit.Text = "0"
    KeyAscii = 0
End If
End Sub

Private Sub TextNamaBuku_Click()
If TextNamaBuku.Text = "Nama Buku" Then
    TextNamaBuku.Text = ""
ElseIf TextNamaBuku.Text = "" Then
    TextNamaBuku.Text = "Nama Buku"
End If
End Sub

Private Sub TextNamaPenerbit_Click()
If TextNamaPenerbit.Text = "Nama Penerbit" Then
    TextNamaPenerbit.Text = ""
ElseIf TextNamaPenerbit.Text = "" Then
    TextNamaPenerbit.Text = "Nama Penerbit"
End If
End Sub

Private Sub TextNamaPengarang_Click()
If TextNamaPengarang.Text = "Nama Pengarang" Then
    TextNamaPengarang.Text = ""
ElseIf TextNamaPengarang.Text = "" Then
    TextNamaPengarang.Text = "Nama Pengarang"
End If
End Sub

Private Sub TextTahunTerbit_Click()
If TextTahunTerbit.Text = "Tahun Terbit" Then
    TextTahunTerbit.Text = ""
ElseIf TextTahunTerbit.Text = "" Then
    TextTahunTerbit.Text = "Tahun Terbit"
End If
End Sub
