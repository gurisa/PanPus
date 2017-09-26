Attribute VB_Name = "MdDefault"
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

Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, y, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_SHOWWINDOW = &H40
Public Const TOPMOST_FLAGS = SWP_NOMOVE Or SWP_NOSIZE

Public Sub AlwaysOnTopForm(lngHwnd As Long)
SetWindowPos lngHwnd, HWND_TOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS
End Sub

Public Sub Default()
FormBuku.TextNamaBuku.Text = "Nama Buku"
FormBuku.TextJumlahBuku.Text = "Jumlah"
FormBuku.TextTahunTerbit.Text = "Tahun Terbit"
FormBuku.TextNamaPengarang.Text = "Nama Pengarang"
FormBuku.TextNamaPenerbit.Text = "Nama Penerbit"

FormUtama.ComboNamaAnggota.ToolTipText = "Nama Anggota"
FormUtama.TextNamaBuku.ToolTipText = "Nama Buku"
FormUtama.TextJumlahBuku.ToolTipText = "Jumlah Buku"
FormUtama.ComboIDPinjamDetail.ToolTipText = "ID Pinjam"
FormUtama.ComboIDBuku.ToolTipText = "ID Buku"
FormUtama.ComboIDAnggota.ToolTipText = "ID Anggota"
FormUtama.ComboNamaAnggota.Text = ""
FormUtama.TextNamaBuku.Text = ""

With FormUtama.ListKembali.ColumnHeaders
    .Clear
    .Add , , "ID Kembali", "1200"
    .Add , , "ID Detail", "1200"
    .Add , , "ID Anggota", "1200"
    .Add , , "Petugas", "1200"
    .Add , , "Peminjam", "1500"
    .Add , , "Judul Buku", "3000"
    .Add , , "Jumlah Buku", "1500"
    .Add , , "Tanggal Pinjam", "1700"
    .Add , , "Tanggal Kembali", "1700"
    .Add , , "Denda"
End With

With FormUtama.ListData.ColumnHeaders
    .Clear
    .Add , , "", "0"
    .Add , , "ID Pinjam", "1150"
    .Add , , "ID Buku", "1000"
    .Add , , "Judul Buku", "6000"
    .Add , , "Jumlah Buku", "1500"
End With

With FormUtama.ListPinjam.ColumnHeaders
    .Clear
    .Add , , "ID Pinjam", "1500"
    .Add , , "ID Anggota", "1500"
    .Add , , "Peminjam", "1600"
    .Add , , "Petugas", "1600"
    .Add , , "Tanggal Pinjam", "2000"
    .Add , , "Status", "1475"
End With

With FormUtama.ListPinjamDetail.ColumnHeaders
    .Clear
    .Add , , "ID Detail", "1000"
    .Add , , "ID Buku", "1000"
    .Add , , "Judul Buku", "2700"
    .Add , , "Jumlah Buku", "1400"
    .Add , , "Tanggal Pinjam", "1700"
    .Add , , "Status", "1025"
End With

With FormUtama.ListAdministrator.ColumnHeaders
    .Clear
    .Add , , "ID", "700"
    .Add , , "Username", "2000"
    .Add , , "Alamat", "3000"
End With

FormUtama.CmdHapusListData.Enabled = False
End Sub

Public Sub PerubahanComboPinjaman()
If FormDenda.ComboPinjaman.ListCount > 0 Then
    FormDenda.TextDenda.Enabled = True
    FormDenda.CmdDenda.Enabled = True
ElseIf Val(FormDenda.ComboPinjaman.Text > 0) Then
    FormDenda.TextDenda.Enabled = True
    FormDenda.CmdDenda.Enabled = True
ElseIf FormDenda.ComboPinjaman.Text = 0 Then
    FormDenda.TextDenda.Enabled = False
    FormDenda.CmdDenda.Enabled = False
ElseIf Val(FormDenda.ComboPinjaman.Text = 0) Then
    FormDenda.TextDenda.Enabled = False
    FormDenda.CmdDenda.Enabled = False
Else
    FormDenda.TextDenda.Enabled = False
    FormDenda.CmdDenda.Enabled = False
End If
End Sub

Public Sub IDBuku()
Koneksi "SELECT id_buku FROM tb_buku ORDER BY id_buku"
FormUtama.ComboIDBuku.Clear
Do While Not DB.EOF
    FormUtama.ComboIDBuku.AddItem DB!ID_Buku
    DB.MoveNext
Loop
End Sub

Public Sub IDPinjam()
Koneksi "SELECT id_pinjam_detail FROM tb_kembali WHERE tanggal_kembali-3 > tanggal_pinjam"
FormDenda.ComboPinjaman.Clear
Do While Not DB.EOF
    FormDenda.ComboPinjaman.AddItem DB!id_pinjam_detail
    DB.MoveNext
Loop
End Sub

Public Sub IDPinjamDetail()
Koneksi "SELECT id_pinjam FROM tb_pinjam WHERE status_pinjam='Pinjam'"
FormUtama.ComboIDPinjamDetail.Clear
Do While Not DB.EOF
    FormUtama.ComboIDPinjamDetail.AddItem DB!ID_Pinjam
    DB.MoveNext
Loop
End Sub

Public Sub MasukanDataPinjam()
Koneksi "SELECT * FROM tb_pinjam WHERE status_pinjam='Pinjam'"
IsiPinjam FormUtama.ListPinjam
End Sub

Public Sub MasukanDataKembali()
Koneksi "SELECT * FROM tb_kembali ORDER BY id_kembali DESC"
IsiKembali FormUtama.ListKembali
End Sub

Public Sub MasukanIDAnggota()
Koneksi "SELECT id_anggota, nama_anggota, status_anggota FROM tb_anggota WHERE status_anggota='Aktif' ORDER BY id_anggota "
FormUtama.ComboIDAnggota.Clear
Do While Not DB.EOF
    FormUtama.ComboIDAnggota.AddItem DB!ID_Anggota
    DB.MoveNext
Loop
End Sub

Public Sub MasukanLogIDAnggota()
Koneksi "SELECT id_anggota, nama_anggota, status_anggota FROM tb_anggota WHERE status_anggota='Aktif' ORDER BY id_anggota "
FormLog.ComboIDAnggota.Clear
Do While Not DB.EOF
    FormLog.ComboIDAnggota.AddItem DB!ID_Anggota
    DB.MoveNext
Loop
End Sub

Public Sub MasukanLogPinjam()
Koneksi "SELECT * FROM tb_pinjam WHERE status_pinjam='Kembali'"
IsiLogPinjam FormLog.ListPinjam
End Sub

Public Sub MasukanLogKembali()
Koneksi "SELECT * FROM tb_kembali"
IsiLogKembali FormLog.ListKembali
End Sub

Public Sub MasukanLogDenda()
Koneksi "SELECT * FROM tb_denda"
IsiDenda FormLog.ListDenda
End Sub

Public Sub MasukanLogAnggota()
Koneksi "SELECT id_anggota, nama_anggota, total_denda FROM tb_anggota"
IsiLogAnggota FormLog.ListAnggota
End Sub

Public Sub MasukanLogAnggotaKembali()
Koneksi "SELECT id_kembali, id_pinjam_detail, id_anggota FROM tb_kembali"
IsiLogAnggotaKembali FormLog.ListAnggotaKembali
End Sub

Public Sub MasukanJurusanAnggota()
Koneksi "SELECT status_setting_text AS jurusan_anggota FROM tb_setting WHERE nama_setting='Jurusan Anggota'"
FormAnggota.ComboJurusan.Clear
Do While Not DB.EOF
    FormAnggota.ComboJurusan.AddItem DB!jurusan_anggota
    DB.MoveNext
Loop
End Sub

Public Sub MasukanKelasAnggota()
Koneksi "SELECT status_setting_text AS kelas_anggota FROM tb_setting WHERE nama_setting='Kelas Anggota'"
FormAnggota.ComboKelas.Clear
Do While Not DB.EOF
    FormAnggota.ComboKelas.AddItem DB!kelas_anggota
    DB.MoveNext
Loop
End Sub

Public Sub MasukanSekolahAnggota()
Koneksi "SELECT status_setting_text AS sekolah_anggota FROM tb_setting WHERE nama_setting='Sekolah Anggota'"
FormAnggota.ComboSekolah.Clear
Do While Not DB.EOF
    FormAnggota.ComboSekolah.AddItem DB!sekolah_anggota
    DB.MoveNext
Loop
End Sub

Public Sub MasukanKategoriBuku()
Koneksi "SELECT nama_kategori_buku FROM tb_buku_kategori ORDER BY nama_kategori_buku ASC"
FormReport.ComboKategoriBuku.Clear
Do While Not DB.EOF
    FormReport.ComboKategoriBuku.AddItem DB!nama_kategori_buku
    DB.MoveNext
Loop
End Sub

Public Sub KategoriBuku()
Koneksi "SELECT nama_kategori_buku FROM tb_buku_kategori ORDER BY nama_kategori_buku ASC"
FormBuku.ComboKategoriBuku.Clear
Do While Not DB.EOF
    FormBuku.ComboKategoriBuku.AddItem DB!nama_kategori_buku
    DB.MoveNext
Loop
End Sub

Public Sub KategoriTingkat()
Koneksi "SELECT nama_kategori_tingkat FROM tb_kategori_tingkat ORDER BY nama_kategori_tingkat ASC"
FormAnggota.ComboKategoriTingkat.Clear
FormAnggota.ComboKelas.Clear
Do While Not DB.EOF
    FormAnggota.ComboKategoriTingkat.AddItem DB!nama_kategori_tingkat
    FormAnggota.ComboKelas.AddItem DB!nama_kategori_tingkat
    DB.MoveNext
Loop
End Sub

Public Sub KategoriJurusan()
Koneksi "SELECT nama_kategori_jurusan FROM tb_kategori_jurusan ORDER BY nama_kategori_jurusan ASC"
FormAnggota.ComboKategoriJurusan.Clear
FormAnggota.ComboJurusan.Clear
Do While Not DB.EOF
    FormAnggota.ComboKategoriJurusan.AddItem DB!nama_kategori_jurusan
    FormAnggota.ComboJurusan.AddItem DB!nama_kategori_jurusan
    DB.MoveNext
Loop
End Sub

Public Sub KategoriLembaga()
Koneksi "SELECT nama_kategori_lembaga FROM tb_kategori_lembaga ORDER BY nama_kategori_lembaga ASC"
FormAnggota.ComboKategoriLembaga.Clear
FormAnggota.ComboSekolah.Clear
Do While Not DB.EOF
    FormAnggota.ComboKategoriLembaga.AddItem DB!nama_kategori_lembaga
    FormAnggota.ComboSekolah.AddItem DB!nama_kategori_lembaga
    DB.MoveNext
Loop
End Sub
