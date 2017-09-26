VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form FormReport 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Laporan"
   ClientHeight    =   1815
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3510
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
   Icon            =   "FormReport.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1815
   ScaleWidth      =   3510
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameLaporan 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Laporan"
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3255
      Begin Crystal.CrystalReport Laporan 
         Left            =   0
         Top             =   1200
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowState     =   2
         PrintFileLinesPerPage=   60
         WindowShowSearchBtn=   -1  'True
         WindowShowPrintSetupBtn=   -1  'True
         WindowShowRefreshBtn=   -1  'True
      End
      Begin VB.ComboBox ComboKategoriBuku 
         Height          =   330
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   720
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.CommandButton CmdLihat 
         Caption         =   "Lihat"
         Height          =   375
         Left            =   1680
         TabIndex        =   4
         Top             =   1200
         Width           =   1455
      End
      Begin VB.CommandButton CmdCetak 
         Caption         =   "Cetak"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   1200
         Width           =   1455
      End
      Begin VB.ComboBox ComboKategoriLaporan 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   3015
      End
      Begin MSComCtl2.DTPicker DTAkhir 
         Height          =   375
         Left            =   1680
         TabIndex        =   3
         Top             =   720
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         CalendarBackColor=   16777215
         CalendarForeColor=   0
         CalendarTitleBackColor=   16777215
         CalendarTitleForeColor=   0
         CalendarTrailingForeColor=   255
         Format          =   92274689
         CurrentDate     =   41890
      End
      Begin MSComCtl2.DTPicker DTAwal 
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         CalendarBackColor=   16777215
         CalendarForeColor=   0
         CalendarTitleBackColor=   16777215
         CalendarTitleForeColor=   0
         CalendarTrailingForeColor=   255
         Format          =   92274689
         CurrentDate     =   41890
      End
   End
End
Attribute VB_Name = "FormReport"
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

Private Sub CmdCetak_Click()
On Error GoTo ErrorLihat
With Laporan
    .WindowTitle = "Panda Pustaka"
End With
If ComboKategoriLaporan.Text = "" Then
    MsgBox "Pilih Jenis Laporan", vbExclamation, "Pilih Jenis Laporan"
    ComboKategoriLaporan.SetFocus
Else
    Dim TanggalAwal As Date
    Dim TanggalAkhir As Date
        
    TanggalAwal = Format(DTAwal.Value, "yyyy-MM-dd")
    TanggalAkhir = Format(DTAkhir.Value, "yyyy-MM-dd")
        
    If ComboKategoriLaporan.Text = "Laporan Peminjaman" Then
        If MsgBox("Tampilkan Seluruh Laporan Peminjaman ?", vbInformation + vbYesNo, "Tampilkan Seluruh Data") = vbYes Then
            Koneksi "SELECT SUM(jumlah_buku) AS TotalBuku FROM tb_pinjam_detail"
            With Laporan
                .ReportFileName = App.Path & "\Report\LaporanPeminjaman.rpt"
                .DiscardSavedData = True
                .WindowShowPrintSetupBtn = True
                .WindowShowPrintBtn = True
                .Formulas(0) = "Pencetak = '" & FormUtama.StatusBarUtama.Panels(4).Text & "'"
                .Formulas(1) = "TanggalAwal = '" & TanggalAwal & "'"
                .Formulas(2) = "TanggalAkhir = '" & TanggalAkhir & "'"
                .Formulas(3) = "TotalBuku = '" & DB!TotalBuku & "'"
                .RetrieveDataFiles
                .PrintReport
            End With
        Else
            Koneksi "SELECT SUM(jumlah_buku) AS TotalBuku FROM tb_pinjam_detail WHERE tanggal_pinjam >= '" & Format(DTAwal.Value, "yyyy-MM-dd") & "' AND tanggal_pinjam <= '" & Format(DTAkhir.Value, "yyyy-MM-dd") & "'"
            With Laporan
                .ReportFileName = App.Path & "\Report\LaporanPeminjaman.rpt"
                .DiscardSavedData = True
                .WindowShowPrintSetupBtn = True
                .WindowShowPrintBtn = True
                .Formulas(0) = "Pencetak = '" & FormUtama.StatusBarUtama.Panels(4).Text & "'"
                .Formulas(1) = "TanggalAwal = '" & TanggalAwal & "'"
                .Formulas(2) = "TanggalAkhir = '" & TanggalAkhir & "'"
                .Formulas(3) = "TotalBuku = '" & DB!TotalBuku & "'"
                .SelectionFormula = "{tb_pinjam_detail.tanggal_pinjam} >= #" & Format(DTAwal.Value, "yyyy-MM-dd") & "# AND {tb_pinjam_detail.tanggal_pinjam} <= #" & Format(DTAkhir.Value, "yyyy-MM-dd") & "#"
                .RetrieveDataFiles
                .PrintReport
                .SelectionFormula = ""
            End With
        End If
    ElseIf ComboKategoriLaporan.Text = "Laporan Pengembalian" Then
        If MsgBox("Tampilkan Seluruh Laporan Pengembalian ?", vbInformation + vbYesNo, "Tampilkan Seluruh Data") = vbYes Then
            Koneksi "SELECT SUM(jumlah_buku) AS TotalBuku FROM tb_kembali"
               With Laporan
                .ReportFileName = App.Path & "\Report\LaporanPengembalian.rpt"
                .DiscardSavedData = True
                .WindowShowPrintSetupBtn = True
                .WindowShowPrintBtn = True
                .Formulas(0) = "Pencetak = '" & FormUtama.StatusBarUtama.Panels(4).Text & "'"
                .Formulas(1) = "TanggalAwal = '" & TanggalAwal & "'"
                .Formulas(2) = "TanggalAkhir = '" & TanggalAkhir & "'"
                .Formulas(3) = "TotalBuku = '" & DB!TotalBuku & "'"
                .RetrieveDataFiles
                .PrintReport
            End With
        Else
            Koneksi "SELECT SUM(jumlah_buku) AS TotalBuku FROM tb_kembali WHERE tanggal_kembali >= '" & Format(DTAwal.Value, "yyyy-MM-dd") & "' AND tanggal_kembali <= '" & Format(DTAkhir.Value, "yyyy-MM-dd") & "'"
               With Laporan
                .ReportFileName = App.Path & "\Report\LaporanPengembalian.rpt"
                .DiscardSavedData = True
                .WindowShowPrintSetupBtn = True
                .WindowShowPrintBtn = True
                .Formulas(0) = "Pencetak = '" & FormUtama.StatusBarUtama.Panels(4).Text & "'"
                .Formulas(1) = "TanggalAwal = '" & TanggalAwal & "'"
                .Formulas(2) = "TanggalAkhir = '" & TanggalAkhir & "'"
                .Formulas(3) = "TotalBuku = '" & DB!TotalBuku & "'"
                .SelectionFormula = "{tb_kembali.tanggal_kembali} >= #" & Format(DTAwal.Value, "yyyy-MM-dd") & "# AND {tb_kembali.tanggal_kembali} <= #" & Format(DTAkhir.Value, "yyyy-MM-dd") & "#"
                .RetrieveDataFiles
                .PrintReport
                .SelectionFormula = ""
            End With
        End If
    ElseIf ComboKategoriLaporan.Text = "Laporan Data Buku" Then
        If MsgBox("Tampilkan Seluruh Data Buku ?", vbInformation + vbYesNo, "Tampilkan Seluruh Data") = vbYes Then
            Koneksi "SELECT SUM(jumlah_buku) AS TotalBuku FROM tb_buku"
            With Laporan
                .ReportFileName = App.Path & "\Report\LaporanDataBuku.rpt"
                .DiscardSavedData = True
                .WindowShowPrintSetupBtn = True
                .WindowShowPrintBtn = True
                .Formulas(0) = "Pencetak = '" & FormUtama.StatusBarUtama.Panels(4).Text & "'"
                .Formulas(1) = "TanggalAwal = '" & TanggalAwal & "'"
                .Formulas(2) = "TanggalAkhir = '" & TanggalAkhir & "'"
                .Formulas(3) = "TotalBuku = '" & DB!TotalBuku & "'"
                .RetrieveDataFiles
                .PrintReport
            End With
        ElseIf ComboKategoriBuku.Text = "" Then
            MsgBox "Pilih Kategori Buku Terlebih Dahulu", vbExclamation, "Pilih Kategori Buku"
        Else
            Koneksi "SELECT SUM(jumlah_buku) AS TotalBuku FROM tb_buku WHERE kategori_buku='" & ComboKategoriBuku.Text & "'"
            With Laporan
                .ReportFileName = App.Path & "\Report\LaporanDataBuku.rpt"
                .DiscardSavedData = True
                .WindowShowPrintSetupBtn = True
                .WindowShowPrintBtn = True
                .Formulas(0) = "Pencetak = '" & FormUtama.StatusBarUtama.Panels(4).Text & "'"
                .Formulas(1) = "TanggalAwal = '" & TanggalAwal & "'"
                .Formulas(2) = "TanggalAkhir = '" & TanggalAkhir & "'"
                .Formulas(3) = "TotalBuku = '" & DB!TotalBuku & "'"
                .SelectionFormula = "{tb_buku.kategori_buku} = '" & ComboKategoriBuku.Text & "'"
                .RetrieveDataFiles
                .PrintReport
                .SelectionFormula = ""
            End With
        End If
    End If
End If
Exit Sub
ErrorLihat:
    MsgBox "Silahkan Ulangi Kembali", vbCritical + vbOKOnly, "Terjadi Kesalahan"
    Unload FormReport
    Set FormReport = Nothing
    FormReport.Show
End Sub

Private Sub CmdLihat_Click()
On Error GoTo ErrorLihat
With Laporan
    .WindowTitle = "Panda Pustaka"
End With
If ComboKategoriLaporan.Text = "" Then
    MsgBox "Pilih Jenis Laporan", vbExclamation, "Pilih Jenis Laporan"
    ComboKategoriLaporan.SetFocus
Else
    Dim TanggalAwal As Date
    Dim TanggalAkhir As Date
        
    TanggalAwal = Format(DTAwal.Value, "yyyy-MM-dd")
    TanggalAkhir = Format(DTAkhir.Value, "yyyy-MM-dd")
        
    If ComboKategoriLaporan.Text = "Laporan Peminjaman" Then
        If MsgBox("Tampilkan Seluruh Laporan Peminjaman ?", vbInformation + vbYesNo, "Tampilkan Seluruh Data") = vbYes Then
            Koneksi "SELECT SUM(jumlah_buku) AS TotalBuku FROM tb_pinjam_detail"
            With Laporan
                .ReportFileName = App.Path & "\Report\LaporanPeminjaman.rpt"
                .DiscardSavedData = True
                .WindowShowPrintSetupBtn = True
                .WindowShowPrintBtn = True
                .Formulas(0) = "Pencetak = '" & FormUtama.StatusBarUtama.Panels(4).Text & "'"
                .Formulas(1) = "TanggalAwal = '" & TanggalAwal & "'"
                .Formulas(2) = "TanggalAkhir = '" & TanggalAkhir & "'"
                .Formulas(3) = "TotalBuku = '" & DB!TotalBuku & "'"
                .RetrieveDataFiles
                .Action = 1
            End With
        Else
            Koneksi "SELECT SUM(jumlah_buku) AS TotalBuku FROM tb_pinjam_detail WHERE tanggal_pinjam >= '" & Format(DTAwal.Value, "yyyy-MM-dd") & "' AND tanggal_pinjam <= '" & Format(DTAkhir.Value, "yyyy-MM-dd") & "'"
            With Laporan
                .ReportFileName = App.Path & "\Report\LaporanPeminjaman.rpt"
                .DiscardSavedData = True
                .WindowShowPrintSetupBtn = True
                .WindowShowPrintBtn = True
                .Formulas(0) = "Pencetak = '" & FormUtama.StatusBarUtama.Panels(4).Text & "'"
                .Formulas(1) = "TanggalAwal = '" & TanggalAwal & "'"
                .Formulas(2) = "TanggalAkhir = '" & TanggalAkhir & "'"
                .Formulas(3) = "TotalBuku = '" & DB!TotalBuku & "'"
                .SelectionFormula = "{tb_pinjam_detail.tanggal_pinjam} >= #" & Format(DTAwal.Value, "yyyy-MM-dd") & "# AND {tb_pinjam_detail.tanggal_pinjam} <= #" & Format(DTAkhir.Value, "yyyy-MM-dd") & "#"
                .RetrieveDataFiles
                .Action = 1
                .SelectionFormula = ""
            End With
        End If
    ElseIf ComboKategoriLaporan.Text = "Laporan Pengembalian" Then
        If MsgBox("Tampilkan Seluruh Laporan Pengembalian ?", vbInformation + vbYesNo, "Tampilkan Seluruh Data") = vbYes Then
            Koneksi "SELECT SUM(jumlah_buku) AS TotalBuku FROM tb_kembali"
               With Laporan
                .ReportFileName = App.Path & "\Report\LaporanPengembalian.rpt"
                .DiscardSavedData = True
                .WindowShowPrintSetupBtn = True
                .WindowShowPrintBtn = True
                .Formulas(0) = "Pencetak = '" & FormUtama.StatusBarUtama.Panels(4).Text & "'"
                .Formulas(1) = "TanggalAwal = '" & TanggalAwal & "'"
                .Formulas(2) = "TanggalAkhir = '" & TanggalAkhir & "'"
                .Formulas(3) = "TotalBuku = '" & DB!TotalBuku & "'"
                .RetrieveDataFiles
                .Action = 1
            End With
        Else
            Koneksi "SELECT SUM(jumlah_buku) AS TotalBuku FROM tb_kembali WHERE tanggal_kembali >= '" & Format(DTAwal.Value, "yyyy-MM-dd") & "' AND tanggal_kembali <= '" & Format(DTAkhir.Value, "yyyy-MM-dd") & "'"
               With Laporan
                .ReportFileName = App.Path & "\Report\LaporanPengembalian.rpt"
                .DiscardSavedData = True
                .WindowShowPrintSetupBtn = True
                .WindowShowPrintBtn = True
                .Formulas(0) = "Pencetak = '" & FormUtama.StatusBarUtama.Panels(4).Text & "'"
                .Formulas(1) = "TanggalAwal = '" & TanggalAwal & "'"
                .Formulas(2) = "TanggalAkhir = '" & TanggalAkhir & "'"
                .Formulas(3) = "TotalBuku = '" & DB!TotalBuku & "'"
                .SelectionFormula = "{tb_kembali.tanggal_kembali} >= #" & Format(DTAwal.Value, "yyyy-MM-dd") & "# AND {tb_kembali.tanggal_kembali} <= #" & Format(DTAkhir.Value, "yyyy-MM-dd") & "#"
                .RetrieveDataFiles
                .Action = 1
                .SelectionFormula = ""
            End With
        End If
    ElseIf ComboKategoriLaporan.Text = "Laporan Data Buku" Then
        If MsgBox("Tampilkan Seluruh Data Buku ?", vbInformation + vbYesNo, "Tampilkan Seluruh Data") = vbYes Then
            Koneksi "SELECT SUM(jumlah_buku) AS TotalBuku FROM tb_buku"
            With Laporan
                .ReportFileName = App.Path & "\Report\LaporanDataBuku.rpt"
                .DiscardSavedData = True
                .WindowShowPrintSetupBtn = True
                .WindowShowPrintBtn = True
                .Formulas(0) = "Pencetak = '" & FormUtama.StatusBarUtama.Panels(4).Text & "'"
                .Formulas(1) = "TanggalAwal = '" & TanggalAwal & "'"
                .Formulas(2) = "TanggalAkhir = '" & TanggalAkhir & "'"
                .Formulas(3) = "TotalBuku = '" & DB!TotalBuku & "'"
                .RetrieveDataFiles
                .Action = 1
            End With
        ElseIf ComboKategoriBuku.Text = "" Then
            MsgBox "Pilih Kategori Buku Terlebih Dahulu", vbExclamation, "Pilih Kategori Buku"
        Else
            Koneksi "SELECT SUM(jumlah_buku) AS TotalBuku FROM tb_buku WHERE kategori_buku='" & ComboKategoriBuku.Text & "'"
            With Laporan
                .ReportFileName = App.Path & "\Report\LaporanDataBuku.rpt"
                .DiscardSavedData = True
                .WindowShowPrintSetupBtn = True
                .WindowShowPrintBtn = True
                .Formulas(0) = "Pencetak = '" & FormUtama.StatusBarUtama.Panels(4).Text & "'"
                .Formulas(1) = "TanggalAwal = '" & TanggalAwal & "'"
                .Formulas(2) = "TanggalAkhir = '" & TanggalAkhir & "'"
                .Formulas(3) = "TotalBuku = '" & DB!TotalBuku & "'"
                .SelectionFormula = "{tb_buku.kategori_buku} = '" & ComboKategoriBuku.Text & "'"
                .RetrieveDataFiles
                .Action = 1
                .SelectionFormula = ""
            End With
        End If
    End If
End If
Exit Sub
ErrorLihat:
    MsgBox "Silahkan Ulangi Kembali", vbCritical + vbOKOnly, "Terjadi Kesalahan"
    Unload FormReport
    Set FormReport = Nothing
    FormReport.Show
End Sub

Private Sub ComboKategoriLaporan_Change()
If ComboKategoriLaporan.Text = "Laporan Data Buku" Then
    ComboKategoriBuku.Visible = True
    DTAwal.Visible = False
    DTAkhir.Visible = False
Else
    ComboKategoriBuku.Visible = False
    DTAwal.Visible = True
    DTAkhir.Visible = True
End If
End Sub

Private Sub ComboKategoriLaporan_Click()
If ComboKategoriLaporan.Text = "Laporan Data Buku" Then
    ComboKategoriBuku.Visible = True
    DTAwal.Visible = False
    DTAkhir.Visible = False
Else
    ComboKategoriBuku.Visible = False
    DTAwal.Visible = True
    DTAkhir.Visible = True
End If
End Sub

Private Sub Form_Load()
MasukanKategoriBuku
ComboKategoriLaporan.AddItem "Laporan Peminjaman"
ComboKategoriLaporan.AddItem "Laporan Pengembalian"
ComboKategoriLaporan.AddItem "Laporan Data Buku"
DTAwal.Value = Format(Date, "yyyy-MM-dd")
DTAkhir.Value = Format(Date, "yyyy-MM-dd")
End Sub

