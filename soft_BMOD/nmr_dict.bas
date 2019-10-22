Attribute VB_Name = "read_RDB"
' nmr_dict.bas


Sub read_RDB_dict(RDB_dict_path, RDB_dict_filename, RDB_field_count, tag_count, tag_list)

ReDim RDB_dict_value(RDB_field_count) As String
Dim i As Integer

Open RDB_dict_path + RDB_dict_filename For Input As 1

While EOF(1) <> -1
    For i = 1 To RDB_field_count
        Input #1, RDB_dict_value(i)
        'Debug.Print RDB_dict_value(i)
    Next i
    'Debug.Print
    For i = 1 To tag_count
        If RDB_dict_value(3) = tag_list(i, 1) Then
        If RDB_dict_value(2) = tag_list(i, 2) Then
            tag_list(i, 3) = RDB_dict_value(7)          ' Type
            tag_list(i, 4) = RDB_dict_value(8)          ' Mandatory
            tag_list(i, 7) = RDB_dict_value(9)          ' Enumeration_ID
            tag_list(i, 15) = RDB_dict_value(10)        ' Foreign table
            tag_list(i, 16) = RDB_dict_value(11)        ' Foreign column
            tag_list(i, 17) = RDB_dict_value(12)        ' Secondary index
            tag_list(i, 18) = RDB_dict_value(13)        ' BMRB Internal only
            tag_list(i, 19) = RDB_dict_value(14)        ' DB table name
            tag_list(i, 20) = RDB_dict_value(15)        ' DB column name
            tag_list(i, 21) = RDB_dict_value(16)        ' Loop location mandatory
            tag_list(i, 22) = RDB_dict_value(17)        ' Loop position
            tag_list(i, 23) = RDB_dict_value(4)         ' NMR-STAR form position
            
            
            'Debug.Print i
            'For j = 1 To 5
            'Debug.Print tag_list(i, j)
            'Next j
            'Stop
        End If
        End If
    Next i
Wend
Close 1

End Sub
