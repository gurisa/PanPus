Attribute VB_Name = "MdHistory"
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

Public Sub IsiHistory()
Dim List As ListItem
With DB
    FormClient.ListHistory.ListItems.Clear
        Do While Not DB.EOF
            Set List = FormClient.ListHistory.ListItems.Add
                List.Text = DB!id_pinjam_detail
                List.SubItems(1) = DB!nama_Buku
                List.SubItems(2) = DB!jumlah_Buku
                List.SubItems(3) = Format(DB!tanggal_pinjam, "yyyy/MM/dd")
                List.SubItems(4) = DB!status_pinjam_detail
            DB.MoveNext
        Loop
End With
End Sub
