Attribute VB_Name = "MdKoneksiPinjamDetail"
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

Public Sub IsiPinjamDetail()
Dim List As ListItem
With DB
    FormUtama.ListPinjamDetail.ListItems.Clear
        Do While Not DB.EOF
            Set List = FormUtama.ListPinjamDetail.ListItems.Add
                List.Text = DB!id_pinjam_detail
                List.SubItems(1) = DB!ID_Buku
                List.SubItems(2) = DB!nama_buku
                List.SubItems(3) = DB!jumlah_buku
                List.SubItems(4) = Format(DB!tanggal_pinjam, "yyyy/MM/dd")
                List.SubItems(5) = DB!status_pinjam_detail
            DB.MoveNext
        Loop
End With
End Sub

Public Sub IsiLogPinjamDetail(Output As ListView)
    Output.ListItems.Clear
    If DB.RecordCount > 0 Then
        DB.MoveFirst
        Do Until DB.EOF
            Output.ListItems.Add , , DB.Fields(0)
            If DB.Fields.Count > 1 Then
                For x = 2 To DB.Fields.Count
                    Output.ListItems(Output.ListItems.Count).SubItems(x - 1) = DB.Fields(x - 1)
               Next x
            End If
            DB.MoveNext
        Loop
    End If
End Sub

