Attribute VB_Name = "MdFungsi"
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
Dim Lokasi As String, DRIVER As String, SERVER As String, UID As String, PASSWORD As String, DATABASE As String, PORT As Integer

Public Function ComboKategoriHistoryClient()
If FormClient.ComboKategoriHistory.Text = "" Then
    FormClient.TextSearchHistory.ToolTipText = ""
ElseIf FormClient.ComboKategoriHistory.Text = "ID" Then
    FormClient.TextSearchHistory.ToolTipText = "Format ID Integer(250)"
ElseIf FormClient.ComboKategoriHistory.Text = "Judul" Then
    FormClient.TextSearchHistory.ToolTipText = "Foromat Judul VarChar(250)"
ElseIf FormClient.ComboKategoriHistory.Text = "Jumlah" Then
    FormClient.TextSearchHistory.ToolTipText = "Format Jumlah Integer(250)"
ElseIf FormClient.ComboKategoriHistory.Text = "Tanggal" Then
    FormClient.TextSearchHistory.ToolTipText = "Format Tanggal Date(YYYY-MM-DD)"
ElseIf FormClient.ComboKategoriHistory.Text = "Status" Then
    FormClient.TextSearchHistory.ToolTipText = "Format Status Enum(Pinjam/Kembali)"
End If
End Function

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
    FormPanel.TextDriver.Text = DRIVER
    FormPanel.TextServer.Text = SERVER
    FormPanel.TextUserID.Text = UID
    FormPanel.TextPassword.Text = PASSWORD
    FormPanel.TextPort.Text = PORT
    FormPanel.ComboDatabase.Text = DATABASE
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
        Print #Baris, FormPanel.TextDriver.Text
        Print #Baris, ""
        Print #Baris, "[SERVER]"
        Print #Baris, FormPanel.TextServer.Text
        Print #Baris, ""
        Print #Baris, "[UID]"
        Print #Baris, FormPanel.TextUserID.Text
        Print #Baris, ""
        Print #Baris, "[PASSWORD]"
        Print #Baris, FormPanel.TextPassword.Text
        Print #Baris, ""
        Print #Baris, "[DATABASE]"
        Print #Baris, FormPanel.ComboDatabase.Text
        Print #Baris, ""
        Print #Baris, "[PORT]"
        Print #Baris, FormPanel.TextPort.Text
        Print #Baris, ""
    Close #Baris
End Function

Public Function FilterInjeksi(FInjeksi As String) As String
    FInjeksi = Replace(FInjeksi, "'", "")
    FInjeksi = Replace(FInjeksi, "\", "")
    FInjeksi = Replace(FInjeksi, "-", "")
    FInjeksi = Replace(FInjeksi, "_", "")
    FInjeksi = Replace(FInjeksi, ";", "")
    FilterInjeksi = FInjeksi
End Function

Public Function AntiInjeksi(AInjeksi As String) As Boolean
    AInjeksi = Replace(AInjeksi, "'", "")
    AInjeksi = Replace(AInjeksi, "\", "")
    AInjeksi = Replace(AInjeksi, "-", "")
    AInjeksi = Replace(AInjeksi, "_", "")
    AInjeksi = Replace(AInjeksi, ";", "")
End Function

Public Function ErrorInjeksi()
    MsgBox "Karakter Tidak Valid", vbCritical, "Karakter Tidak Valid"
End Function

Public Function EksekusiPanelConfig()
If Dir$(App.Path & "\Panel.exe") <> "" Then
    EksekusiPanelConfig = Shell(App.Path & "\Panel.exe", vbNormalFocus)
Else
    MsgBox "Panel Tidak Tersedia" & vbNewLine & "Hubungi Administrator", vbCritical, "Panel Tidak Tersedia"
    End
End If
End Function

