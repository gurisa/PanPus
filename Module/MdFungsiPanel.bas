Attribute VB_Name = "MdFungsiPanel"
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
Dim Baris As Integer
Public TestDB As ADODB.Recordset
Public TestConn As ADODB.Connection
Dim Lokasi As String, DRIVER As String, SERVER As String, UID As String, PASSWORD As String, DATABASE As String, PORT As Integer


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

Public Function ReadFromPanel()
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
    FormPanelConfig.TextDriver.Text = DRIVER
    FormPanelConfig.TextServer.Text = SERVER
    FormPanelConfig.TextUserID.Text = UID
    FormPanelConfig.TextPassword.Text = PASSWORD
    FormPanelConfig.TextPort.Text = PORT
    FormPanelConfig.ComboDatabase.Text = DATABASE
End Function

Public Function WriteToPanelDefault()
Lokasi = App.Path & "\Panel.panpus"
Baris = FreeFile
    Open Lokasi For Output As #Baris
        Print #Baris, "[DRIVER]"
        Print #Baris, "MySQL ODBC 5.2w Driver"
        Print #Baris, ""
        Print #Baris, "[SERVER]"
        Print #Baris, "192.168.100.1"
        Print #Baris, ""
        Print #Baris, "[UID]"
        Print #Baris, "panda"
        Print #Baris, ""
        Print #Baris, "[PASSWORD]"
        Print #Baris, "pustaka"
        Print #Baris, ""
        Print #Baris, "[DATABASE]"
        Print #Baris, "db_perpus"
        Print #Baris, ""
        Print #Baris, "[PORT]"
        Print #Baris, "3306"
        Print #Baris, ""
    Close #Baris
End Function

Public Function WriteToPanel()
Lokasi = App.Path & "\Panel.panpus"
Baris = FreeFile
    Open Lokasi For Output As #Baris
        Print #Baris, "[DRIVER]"
        Print #Baris, FormPanelConfig.TextDriver.Text
        Print #Baris, ""
        Print #Baris, "[SERVER]"
        Print #Baris, FormPanelConfig.TextServer.Text
        Print #Baris, ""
        Print #Baris, "[UID]"
        Print #Baris, FormPanelConfig.TextUserID.Text
        Print #Baris, ""
        Print #Baris, "[PASSWORD]"
        Print #Baris, FormPanelConfig.TextPassword.Text
        Print #Baris, ""
        Print #Baris, "[DATABASE]"
        Print #Baris, FormPanelConfig.ComboDatabase.Text
        Print #Baris, ""
        Print #Baris, "[PORT]"
        Print #Baris, FormPanelConfig.TextPort.Text
        Print #Baris, ""
    Close #Baris
End Function

Public Function EksekusiPandaPustaka()
If Dir$(App.Path & "\Perpustakaan.exe") <> "" Then
    EksekusiPanelConfig = Shell(App.Path & "\Perpustakaan.exe", vbNormalFocus)
Else
    MsgBox "Panda Pustaka Tidak Tersedia" & vbNewLine & "Hubungi Administrator", vbCritical, "Panpus Tidak Tersedia"
    End
End If
End Function

