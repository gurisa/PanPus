Attribute VB_Name = "MdKoneksiClient"
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

Public Sub IsiMessage(Output As ListView)
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

Public Sub IsiMessagePetugas(Output As ListView)
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


