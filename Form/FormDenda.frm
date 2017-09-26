VERSION 5.00
Begin VB.Form FormDenda 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Denda"
   ClientHeight    =   2460
   ClientLeft      =   -15
   ClientTop       =   270
   ClientWidth     =   3285
   BeginProperty Font 
      Name            =   "Consolas"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormDenda.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2460
   ScaleWidth      =   3285
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox TextIDAnggota 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   600
      Width           =   3015
   End
   Begin VB.CommandButton CmdDenda 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Denda"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   1800
      Width           =   3015
   End
   Begin VB.TextBox TextDenda 
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Text            =   "Jumlah Denda"
      Top             =   1080
      Width           =   3015
   End
   Begin VB.ComboBox ComboPinjaman 
      Height          =   345
      ItemData        =   "FormDenda.frx":0CCA
      Left            =   120
      List            =   "FormDenda.frx":0CCC
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "FormDenda"
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
Dim ID_Denda As Integer

Private Sub CmdDenda_Click()
TextDenda = FilterInjeksi(TextDenda.Text)

If TextDenda.Text = "" Or ComboPinjaman.Text = "" Or TextIDAnggota.Text = "" Then
    MsgBox "Masukan Nilai Denda", vbCritical, "Masukan Nilai Denda"
ElseIf TextDenda.Text = "Jumlah Denda" Then
    MsgBox "Masukan Jumlah Denda", vbInformation, "Masukan Jumlah Denda"
ElseIf ComboPinjaman.Text = "Pinjaman" Then
    MsgBox "Pilih Nomor Peminjaman Yang Akan Di Berikan Denda", vbInformation, "Pilih Nomor Peminjaman"
Else
    Conn.Execute "INSERT INTO tb_denda(id_anggota, id_pinjam_detail, banyak_denda, tanggal_denda, petugas_denda) VALUES('" & TextIDAnggota.Text & "','" & ComboPinjaman.Text & "','" & TextDenda.Text & "','" & FormUtama.StatusBarUtama.Panels(1) & "','" & FormUtama.StatusBarUtama.Panels(4) & "')"
    Conn.Execute "UPDATE tb_anggota SET total_denda=total_denda + '" & TextDenda.Text & "' WHERE tb_anggota.id_anggota='" & TextIDAnggota.Text & "'"
    MsgBox "Berhasil Memberikan Denda Untuk Pinjaman Dengan Nomor Detail " & ComboPinjaman.Text & " Senilai Rp. " & Val(TextDenda.Text), vbInformation, "Berhasil"
End If
End Sub

Private Sub ComboPinjaman_Change()
Koneksi "SELECT id_pinjam_detail, id_anggota FROM tb_kembali WHERE id_pinjam_detail='" & ComboPinjaman.Text & "'"
Do While Not DB.EOF
    TextIDAnggota.Text = DB!ID_Anggota
    DB.MoveNext
Loop
Call PerubahanComboPinjaman
End Sub

Private Sub ComboPinjaman_Click()
Koneksi "SELECT id_pinjam_detail, id_anggota FROM tb_kembali WHERE id_pinjam_detail='" & ComboPinjaman.Text & "'"
TextIDAnggota.Text = DB!ID_Anggota
Call PerubahanComboPinjaman
End Sub

Private Sub Form_Load()
Call IDPinjam
TextDenda.Enabled = False
CmdDenda.Enabled = False
TextDenda.Text = ""
End Sub

Private Sub TextDenda_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= Asc("0") & Chr(13) _
     And KeyAscii <= Asc("9") & Chr(13) _
     Or KeyAscii = vbKeyBack _
     Or KeyAscii = vbKeyDelete _
     Or KeyAscii = vbKeySpace) Then
    TextDenda.Text = "0"
    KeyAscii = 0
End If
End Sub
