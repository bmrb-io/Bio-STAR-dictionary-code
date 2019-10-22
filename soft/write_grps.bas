Attribute VB_Name = "write_grps"
Sub write_grp_tables(NMR_STAR_dict_path, super_grp_tbl_in, super_grp_tbl_out, group_tbl_in, group_tbl_out, RDB_field_count, new_table, item_tbl_out)

'Write out the item table for Steve's software

ReDim new_table_dat(1 To 4000, 1 To RDB_field_count) As String
ReDim view_col(20) As Integer
Dim i, new_tbl_ct As Integer

Open NMR_STAR_dict_path + new_table For Input As 1
Open NMR_STAR_dict_path + item_tbl_out For Output As 2

row_ct = 0: view_ct = 0

While EOF(1) <> -1
    row_ct = row_ct + 1
    For i = 1 To RDB_field_count
        Input #1, new_table_dat(row_ct, i)
        If new_table_dat(row_ct, i) = "View" Then
            view_ct = view_ct + 1
            view_col(view_ct) = i
        End If
        'Debug.Print i, new_table_dat(new_tbl_ct, i)
    Next i
    'Debug.Print
Wend
Close 1

new_table_dat(1, view_col(1)) = "ADIT view flags"
For j = 1 To row_ct
    For i = 1 To RDB_field_count
        For k = 2 To view_ct
            If view_col(k) = i And j > 4 Then
                new_table_dat(j, view_col(1)) = new_table_dat(j, view_col(1)) + new_table_dat(j, view_col(k))
            End If
        Next k
    Next i
    For i = 1 To RDB_field_count
        print_flag = 1
        For k = 2 To view_ct
            If i = view_col(k) Then print_flag = 0
        Next k
        
        If new_table_dat(1, i) = "SG Mandatory" Then print_flag = 0
        If new_table_dat(1, i) = "BMRB current" Then print_flag = 0
        If new_table_dat(1, i) = "BMRB next release" Then print_flag = 0
        
        If print_flag = 1 Then
            If i < RDB_field_count Then Print #2, new_table_dat(j, i) + ",";
            If i = RDB_field_count Then Print #2, new_table_dat(j, i) + ","
        End If
    Next i
Next j
Close 2

'Write out the supergroup table for Steve's software

Open NMR_STAR_dict_path + super_grp_tbl_in For Input As 1
Open NMR_STAR_dict_path + super_grp_tbl_out For Output As 2

ReDim view_col(20)
col_ct = 8
ReDim table_row(200, col_ct) As String
row_ct = 0: view_ct = 0
While EOF(1) <> -1
    row_ct = row_ct + 1
    For i = 1 To col_ct
        Input #1, table_row(row_ct, i)
        If table_row(row_ct, i) = "View" Then
            view_ct = view_ct + 1
            view_col(view_ct) = i
        End If
    Next i
Wend
Close 1

table_row(1, view_col(1)) = "ADIT view flags"
For j = 1 To row_ct
    For i = 1 To col_ct
        For k = 2 To view_ct
            If view_col(k) = i And j > 3 Then
                table_row(j, view_col(1)) = table_row(j, view_col(1)) + table_row(j, view_col(k))
            End If
        Next k
    Next i
    For i = 1 To col_ct
        print_flag = 1
        For k = 2 To view_ct
            If i = view_col(k) Then print_flag = 0
        Next k
        If print_flag = 1 Then
            If i < col_ct Then Print #2, table_row(j, i) + ",";
            If i = col_ct Then Print #2, table_row(j, i)
        End If
    Next i
Next j
Close 2

'Write out category group table for Steve's software

Open NMR_STAR_dict_path + group_tbl_in For Input As 1
Open NMR_STAR_dict_path + group_tbl_out For Output As 2

col_ct = 15
ReDim table_row(200, col_ct) As String
row_ct = 0
While EOF(1) <> -1
    row_ct = row_ct + 1
    For i = 1 To col_ct
        Input #1, table_row(row_ct, i)
    Next i
Wend
Close 1

For i = 1 To row_ct
    For j = 1 To 4
        Print #2, table_row(i, j) + ",";
'        If j = 4 Then Print #2, table_row(i, j)
    Next j
    Print #2, table_row(i, 8) + ",";
    Print #2, table_row(i, 9) + ",";
    Print #2, table_row(i, 13) + ",";
    Print #2, table_row(i, 14) + ","
Next i
Close 2
End Sub
