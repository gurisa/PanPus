Attribute VB_Name = "MdKoneksi"
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

Public DB As ADODB.Recordset
Public Conn As ADODB.Connection
Public TestDB As ADODB.Recordset
Public TestConn As ADODB.Connection

Public Sub Koneksi(SQLCommand As String)
Dim Baris As Integer
Dim Lokasi As String
Dim DRIVER As String, SERVER As String, UID As String, PASSWORD As String, DATABASE As String, PORT As Integer
    Lokasi = App.Path & "\Panel.panpus"
    Baris = FreeFile
    Open Lokasi For Input As #Baris
        Input #Baris, JudulDriver
        Input #Baris, DRIVER
        Input #Baris, KosongDriver
        Input #Baris, JudulServer
        Input #Baris, SERVER
        Input #Baris, KosongServer
        Input #Baris, JudulUID
        Input #Baris, UID
        Input #Baris, KosongUID
        Input #Baris, JudulPassword
        Input #Baris, PASSWORD
        Input #Baris, KosongPassword
        Input #Baris, JudulDatabase
        Input #Baris, DATABASE
        Input #Baris, KosongDatabase
        Input #Baris, JudulPort
        Input #Baris, PORT
        Input #Baris, KosongPort
    Close #Baris
    
On Error GoTo KonfigurasiPanel
    Set Conn = New ADODB.Connection
    Set DB = New ADODB.Recordset
    Conn.Open "Provider=MSDASQL.1;Persist Security Info=False;Extended Properties=DRIVER=" & DRIVER & ";Server=" & SERVER & ";Uid=" & UID & ";Password=" & PASSWORD & ";Database=" & DATABASE & ";PORT=" & PORT & ";"""
    DB.CursorLocation = adUseClient
    DB.Open SQLCommand, Conn
Exit Sub
KonfigurasiPanel:
    FormLogin.TimerExit.Enabled = True
    FormLogin.TimerExit.Interval = 100
    MsgBox "Gagal Melakukan Koneksi", vbCritical, "Gagal Koneksi"
    FormPanel.TextCodeGet.Enabled = False
    FormPanel.TextCodeSend.Enabled = False
    FormPanel.CmdGenerateGet.Enabled = False
    FormPanel.CmdAktivasi.Enabled = False
    FormPanel.ImageDonate.Enabled = False
    FormPanel.CheckPopUpMessage.Enabled = False
    FormPanel.CheckPopUpAbout.Enabled = False
    'FormPanel.Show
    If MsgBox("Terjadi Kesalahan " & vbNewLine & "Hal Ini Dapat Di Sebabkan Oleh : " & vbNewLine & "" & vbNewLine & "1. Gagal Konektivitas" & vbNewLine & "2. Kesalahan Pengaturan" & vbNewLine & "" & vbNewLine & "Untuk Mengatasinya Tekan Yes Untuk Mengubah Pengaturan Aplikasi", vbExclamation + vbYesNo, "Buka Pengaturan") = vbYes Then
        Call EksekusiPanelConfig
    Else
        End
    End If
    End
End Sub

Public Sub TestKoneksi(SQLCommand As String)
Dim Baris As Integer
Dim Lokasi As String
Dim DRIVER As String, SERVER As String, UID As String, PASSWORD As String, DATABASE As String, PORT As Integer
    Lokasi = App.Path & "\Panel.panpus"
    Baris = FreeFile
    Open Lokasi For Input As #Baris
        Input #Baris, JudulDriver
        Input #Baris, DRIVER
        Input #Baris, KosongDriver
        Input #Baris, JudulServer
        Input #Baris, SERVER
        Input #Baris, KosongServer
        Input #Baris, JudulUID
        Input #Baris, UID
        Input #Baris, KosongUID
        Input #Baris, JudulPassword
        Input #Baris, PASSWORD
        Input #Baris, KosongPassword
        Input #Baris, JudulDatabase
        Input #Baris, DATABASE
        Input #Baris, KosongDatabase
        Input #Baris, JudulPort
        Input #Baris, PORT
        Input #Baris, KosongPort
    Close #Baris
    
On Error GoTo GagalTest
    Set TestConn = New ADODB.Connection
    Set TestDB = New ADODB.Recordset
    TestConn.Open "Provider=MSDASQL.1;Persist Security Info=False;Extended Properties=DRIVER=" & DRIVER & ";Server=" & SERVER & ";Uid=" & UID & ";Password=" & PASSWORD & ";Database=" & DATABASE & ";PORT=" & PORT & ";"""
    TestDB.CursorLocation = adUseClient
    TestDB.Open SQLCommand, TestConn
    MsgBox "Koneksi Berhasil", vbInformation, "Koneksi Berhasil"
Exit Sub
GagalTest:
    MsgBox "Gagal Melakukan Koneksi", vbCritical, "Gagal Koneksi"
End Sub

Public Function TemukanError(Data As String) As Boolean
    If Len(Data) - Len(Replace(Data, "'", "")) > 0 Then
        MsgBox "Tidak Boleh Terdapat Tanda Kutip", vbCritical, "Di Larang Menggunakan Tanda Kutip"
        FindError = 1
    Else
        FindError = 0
    End If
End Function

