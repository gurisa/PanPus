Attribute VB_Name = "MdKoneksiAnggota"
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

Public Sub IsiAnggota(Output As ListView)
Dim List As ListItem
With DB
    FormAnggota.ListAnggota.ListItems.Clear
        Do While Not DB.EOF
            Set List = FormAnggota.ListAnggota.ListItems.Add
                List.Text = DB!ID_Anggota
                List.SubItems(1) = DB!nis_anggota
                List.SubItems(2) = DB!nama_anggota
                List.SubItems(3) = DB!jenis_kelamin
                List.SubItems(4) = DB!kelas_anggota
                List.SubItems(5) = DB!jurusan_anggota
                List.SubItems(6) = DB!status_anggota
                List.SubItems(7) = DB!sekolah_anggota
                List.SubItems(8) = DB!tanggal_daftar
                List.SubItems(9) = DB!petugas_daftar
                List.SubItems(10) = DB!total_denda
                List.SubItems(11) = DB!password_anggota
            DB.MoveNext
        Loop
End With
End Sub

Public Sub IsiLogAnggota(Output As ListView)
    Output.ListItems.Clear
    If DB.RecordCount > 0 Then
        DB.MoveFirst
        Do Until DB.EOF
            Output.ListItems.Add , , DB.Fields(0)
            If DB.Fields.Count > 1 Then
                For X = 2 To DB.Fields.Count
                    Output.ListItems(Output.ListItems.Count).SubItems(X - 1) = DB.Fields(X - 1)
                Next X
            End If
            DB.MoveNext
        Loop
    End If
End Sub

Public Sub IsiLogAnggotaKembali(Output As ListView)
    Output.ListItems.Clear
    If DB.RecordCount > 0 Then
        DB.MoveFirst
        Do Until DB.EOF
            Output.ListItems.Add , , DB.Fields(0)
            If DB.Fields.Count > 1 Then
                For X = 2 To DB.Fields.Count
                    Output.ListItems(Output.ListItems.Count).SubItems(X - 1) = DB.Fields(X - 1)
                Next X
            End If
            DB.MoveNext
        Loop
    End If
End Sub

Public Sub IsiPetugasAnggota(Output As ListView)
    Output.ListItems.Clear
    If DB.RecordCount > 0 Then
        DB.MoveFirst
        Do Until DB.EOF
            Output.ListItems.Add , , DB.Fields(0)
            If DB.Fields.Count > 1 Then
                For X = 2 To DB.Fields.Count
                    Output.ListItems(Output.ListItems.Count).SubItems(X - 1) = DB.Fields(X - 1)
                Next X
            End If
            DB.MoveNext
        Loop
    End If
End Sub

Public Sub IsiCariAnggota(Output As ListView)
    Output.ListItems.Clear
    If DB.RecordCount > 0 Then
        DB.MoveFirst
        Do Until DB.EOF
            Output.ListItems.Add , , DB.Fields(0)
            If DB.Fields.Count > 1 Then
                For X = 2 To DB.Fields.Count
                    Output.ListItems(Output.ListItems.Count).SubItems(X - 1) = DB.Fields(X - 1)
                Next X
            End If
            DB.MoveNext
        Loop
    End If
End Sub
