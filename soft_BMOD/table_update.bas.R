Attribute VB_Name = "table_update"

Sub tbl_update(NMR_STAR_dict_path, old_table, new_table, dict_info, dict_data_ct, RDB_field_count, super_grp_tbl_in, group_tbl_in)

ReDim old_table_dat(1 To 4000, 1 To RDB_field_count) As String
ReDim new_table_dat(1 To 4000, 1 To RDB_field_count) As String

Dim old_tbl_ct, new_tbl_ct As Integer

Open NMR_STAR_dict_path + old_table For Input As 1
Open NMR_STAR_dict_path + new_table For Input As 2

old_tbl_ct = 0: new_tbl_ct = 0

While EOF(1) <> -1
    old_tbl_ct = old_tbl_ct + 1
    For i = 1 To RDB_field_count
        Input #1, old_table_dat(old_tbl_ct, i)
        'Debug.Print old_tbl_ct, i, old_table_dat(old_tbl_ct, i)
    Next i
    'Debug.Print
Wend
Close 1

While EOF(2) <> -1
    new_tbl_ct = new_tbl_ct + 1
    For i = 1 To RDB_field_count
        Input #2, new_table_dat(new_tbl_ct, i)
        'Debug.Print i, new_table_dat(new_tbl_ct, i)
    Next i
    'Debug.Print
        
Wend
Close 2

compare_files old_tbl_ct, old_table_dat, new_tbl_ct, new_table_dat, RDB_field_count

add_text_info new_tbl_ct, new_table_dat, dict_info, dict_data_ct, old_table_dat, RDB_field_count

write_file new_tbl_ct, new_table_dat, RDB_field_count, NMR_STAR_dict_path, new_table, old_tbl_ct, old_table_dat, super_grp_tbl_in, group_tbl_in

End Sub

Sub compare_files(old_tbl_ct, old_table_dat, new_tbl_ct, new_table_dat, RDB_field_count)

Dim i, j, j1, j2, j3, k As Integer
ReDim old_tbl_flag(1 To old_tbl_ct)
j3 = 1: primary_key_ct = 0

For i = 1 To new_tbl_ct
    'Debug.Print i
    'If i > 700 Then j3 = i - 700
    For j = j3 To old_tbl_ct
        If old_tbl_flag(j) = 0 Then
            load_fields = 0
            For j1 = 1 To RDB_field_count
                If old_table_dat(1, j1) = "TagName" And new_table_dat(i, j1) = old_table_dat(j, j1) Then load_fields = load_fields + 1
                If old_table_dat(1, j1) = "SFCategory" And new_table_dat(i, j1) = old_table_dat(j, j1) Then load_fields = load_fields + 1
                If load_fields = 2 Then
                    old_tbl_flag(j) = 1
                    For j2 = 1 To RDB_field_count
                        If old_table_dat(1, j2) <> "Dictionary sequence" Then
                        If old_table_dat(1, j2) <> "SFCategory" Then
                        If old_table_dat(1, j2) <> "TagName" Then
                        If old_table_dat(1, j2) <> "ManDBTableName" Then
                        If old_table_dat(1, j2) <> "ManDBColumnName" Then
                        If old_table_dat(1, j2) <> "Loopflag" Then
                        If old_table_dat(1, j2) <> "Seq" Then
                            new_table_dat(i, j2) = old_table_dat(j, j2)
                            'If old_table_dat(1, j2) = "Primary Key" Then
                            '    If old_table_dat(j, j2) = "Ye" Then
                            '        primary_key_ct = primary_key_ct + 1
                            '        primary_keys(primary_key_ct) = new_table_dat(i, 9)
                            '    End If
                            'End If
                        End If
                        End If
                        End If
                        End If
                        End If
                        End If
                        End If
                    Next j2
                    Exit For
                End If
            Next j1
            If load_fields = 2 Then Exit For
        End If
    Next j
Next i

ReDim old_tbl_flag(1)

End Sub
Sub add_text_info(new_tbl_ct, new_table_dat, dict_info, dict_data_ct, old_table_dat, RDB_field_count)
        
For j1 = 1 To RDB_field_count
    If old_table_dat(1, j1) = "TagName" Then j5 = j1
    If old_table_dat(1, j1) = "Item enumerated" Then j6 = j1
Next j1

For i = 1 To new_tbl_ct
    For j = 1 To dict_data_ct
        If new_table_dat(i, j5) = dict_info(j, 1) Then
            If dict_info(j, 7) = "enumerations" Then
                new_table_dat(i, j6) = "Y"   'enumerations
            Else
                new_table_dat(i, j6) = "N"
            End If
            Exit For
        End If
    Next j
Next i
ReDim dict_info(1, 1)
End Sub
Sub write_file(new_tbl_ct, new_table_dat, RDB_field_count, NMR_STAR_dict_path, new_table, old_tbl_ct, old_table_dat, super_grp_tbl_in, group_tbl_in)

ReDim spg_table_row(0 To 15, 1 To 10) As String
ReDim grp_table_row(0 To 200, 1 To 15) As String
ReDim loop_table_row(1 To 300, 1 To RDB_field_count) As String

Dim i, j, loop_ct As Integer
Dim spg_col_ct, grp_col_ct As Integer
Dim sf_cat(1 To 200, 1 To 2)

Open NMR_STAR_dict_path + super_grp_tbl_in For Input As 1

spg_col_ct = 8: spg_row_ct = 0
While EOF(1) <> -1
    spg_row_ct = spg_row_ct + 1
    For i = 1 To spg_col_ct
        Input #1, spg_table_row(spg_row_ct, i)
    Next i
    If spg_table_row(spg_row_ct, 1) = "TBL_BEGIN" Then spg_row_ct = 0
    'If spg_table_row(spg_row_ct, 1) = "TBL_END" Then spg_row_ct = spg_row_ct - 1
Wend
Close 1

Open NMR_STAR_dict_path + group_tbl_in For Input As 1

grp_col_ct = 15: grp_row_ct = 0
While EOF(1) <> -1
    grp_row_ct = grp_row_ct + 1
    For i = 1 To grp_col_ct
        Input #1, grp_table_row(grp_row_ct, i)
        If grp_table_row(grp_row_ct, 1) = "TBL_BEGIN" Then grp_row_ct = 0
    Next i
    If grp_table_row(grp_row_ct, 1) = "TBL_END" Then grp_row_ct = grp_row_ct - 1
Wend
Close 1

Open NMR_STAR_dict_path + new_table For Output As 5
Open NMR_STAR_dict_path + new_table + "2" For Output As 1

For i = 1 To 4
    For j = 1 To RDB_field_count - 1
        Print #1, old_table_dat(i, j); ",";
        Print #5, old_table_dat(i, j); ",";

        Debug.Print j, old_table_dat(i, j)
        Debug.Print
    Next j
    Print #1, old_table_dat(i, RDB_field_count)
    Print #5, old_table_dat(i, RDB_field_count)
    Debug.Print old_table_dat(i, RDB_field_count)
    'Debug.Print
    'Print #1,
Next i
        
For i = 1 To new_tbl_ct
  For j = 1 To RDB_field_count
    If old_table_dat(1, j) = "Dictionary sequence" Then new_table_dat(i, j) = Trim(Str$(i * 10))
    If old_table_dat(1, j) = "SFCategory" Then
        t4 = new_table_dat(i, j)
        sf_category = new_table_dat(i, j)
        If i = 1 Then sf_cat_last = sf_category
    End If
    If old_table_dat(1, j) = "ManDBTableName" Then
        table_name = new_table_dat(i, j)
    End If
    If old_table_dat(1, j) = "TagName" Then
        t1 = "": t2 = "": t3 = "Y": t5 = "": t6 = "": t7 = "": t8 = "": t9 = ""
        tag_name = new_table_dat(i, j)
        If Right$(new_table_dat(i, j), 3) = "_ID" Then
            t1 = "INTEGER": t2 = "NOT NULL": t3 = "H": t4 = "": t5 = "": t6 = ""
        End If
        If new_table_dat(i, j) = "_Saveframe_category" Then
            t1 = "VARCHAR(127)": t2 = "NOT NULL": t3 = "H": t4 = "": t5 = "": t6 = ""
            
        End If
        If new_table_dat(i, j) = "_Saveframe_framecode" Then
            t1 = "VARCHAR(127)": t2 = "NOT NULL": t3 = "H": t4 = "": t5 = "": t6 = ""
        End If
        'If new_table_dat(i, j) = "_Entry_ID" Then
        '    't4 = "entry_information": t5 = "BMRB_accession_number"
        'End If
        If Left$(new_table_dat(i, j), 5) = "_Loop" Then
            If InStr(1, new_table_dat(i, j), "entry_ID") > 0 Then
                t3 = "H": 't8 = "entry_information": t9 = "BMRB_accession_number"
                t6 = "1"
            End If
            If InStr(1, new_table_dat(i, j), "saveframe_ID") > 0 Then
                t3 = "H": t5 = "_Saveframe_ID": t6 = "2": t8 = sf_category: t9 = "Saveframe_ID"
            End If
        End If
        'If new_table_dat(i, j) = "_Saveframe_ID" Then
        '    t7 = "Ys"
        'End If

    End If
    If old_table_dat(1, j) = "Data type" Then
        i1 = j
        If new_table_dat(i, j) = "" Then
            For j1 = 1 To RDB_field_count
                If old_table_dat(1, j1) = "TagName" Then
    
                If new_table_dat(i, j1) = "_Saveframe_framecode" Then new_table_dat(i, i1) = "VARCHAR(127)"
                If new_table_dat(i, j1) = "_Details" Then new_table_dat(i, i1) = "CLOB"
                If new_table_dat(i, j1) = "_Keyword" Then new_table_dat(i, i1) = "VARCHAR(127)"
                If new_table_dat(i, j1) = "_Synonym" Then new_table_dat(i, i1) = "VARCHAR(127)"
                If new_table_dat(i, j1) = "_Sample_label" Then new_table_dat(i, i1) = "VARCHAR(127)"
                If new_table_dat(i, j1) = "_Sample_conditions_label" Then new_table_dat(i, i1) = "VARCHAR(127)"
                If new_table_dat(i, j1) = "_Software_label" Then new_table_dat(i, i1) = "VARCHAR(127)"
                If new_table_dat(i, j1) = "_Experiment_label" Then new_table_dat(i, i1) = "VARCHAR(127)"
                If new_table_dat(i, j1) = "_Text_data_format" Then new_table_dat(i, i1) = "VARCHAR(127)"
                If new_table_dat(i, j1) = "_Text_data" Then new_table_dat(i, i1) = "TEXT"
                If InStr(1, new_table_dat(i, j1), "ol_system_component_name") > 0 Then new_table_dat(i, i1) = "VARCHAR(127)"
                If Right$(new_table_dat(i, j1), 6) = "_label" Then
                    If Right$(new_table_dat(i, j1), 13) <> "residue_label" Then
                        If new_table_dat(i, j1) <> "_Residue_label" Then
                            new_table_dat(i, i1) = "VARCHAR(127)"
                        End If
                    End If
                End If
                'If Right$(new_table_dat(i, j1), 3) = "_ID" Then new_table_dat(i, i1) = "INTEGER"
                If new_table_dat(i, j1) = "_Residue_label" Then new_table_dat(i, i1) = "VARCHAR(31)"
                If InStr(1, new_table_dat(i, j1), "esidue_label") > 0 Then new_table_dat(i, i1) = "VARCHAR(31)"
                If new_table_dat(i, j1) = "_Residue_seq_code" Then new_table_dat(i, i1) = "INTEGER"
                If InStr(1, new_table_dat(i, j1), "_error") > 0 Then new_table_dat(i, i1) = "FLOAT"
                If InStr(1, new_table_dat(i, j1), "_value") > 0 Then new_table_dat(i, i1) = "FLOAT"
                If InStr(1, new_table_dat(i, j1), "tom_name") > 0 Then new_table_dat(i, i1) = "VARCHAR(31)"
                If new_table_dat(i, j1) = "_Atom_type" Then new_table_dat(i, i1) = "VARCHAR(31)"
                If Right$(new_table_dat(i, j1), 6) = "_count" Then new_table_dat(i, i1) = "INTEGER"
                If Right$(new_table_dat(i, j1), 8) = "_rms_dev" Then new_table_dat(i, i1) = "FLOAT"
                If InStr(1, new_table_dat(i, j1), "PDB_seq_code") Then new_table_dat(i, i1) = "VARCHAR(31)"
                If InStr(1, new_table_dat(i, j1), "esidue_PDB_code") Then new_table_dat(i, i1) = "VARCHAR(31)"
                If InStr(1, new_table_dat(i, j1), "tom_PDB_name") Then new_table_dat(i, i1) = "VARCHAR(31)"
                If InStr(1, new_table_dat(i, j1), "author_seq_code") Then new_table_dat(i, i1) = "VARCHAR(31)"
                If InStr(1, new_table_dat(i, j1), "esidue_author_code") Then new_table_dat(i, i1) = "VARCHAR(31)"
                If InStr(1, new_table_dat(i, j1), "tom_author_name") Then new_table_dat(i, i1) = "VARCHAR(31)"
                If InStr(1, new_table_dat(i, j1), "_residue_seq_code") > 0 Then new_table_dat(i, i1) = "INTEGER"
                If InStr(1, new_table_dat(i, j1), "tom_type") > 0 Then new_table_dat(i, i1) = "VARCHAR(31)"
                If InStr(1, new_table_dat(i, j1), "cartn_x") > 0 Then new_table_dat(i, i1) = "FLOAT"
                If InStr(1, new_table_dat(i, j1), "cartn_y") > 0 Then new_table_dat(i, i1) = "FLOAT"
                If InStr(1, new_table_dat(i, j1), "cartn_z") > 0 Then new_table_dat(i, i1) = "FLOAT"
                If Right$(new_table_dat(i, j1), 9) = "_file_name" Then new_table_dat(i, i1) = "VarChar(127)"
                If new_table_dat(i, j1) = "_Bond_order" Then new_table_dat(i, i1) = "VARCHAR(31)"
                If new_table_dat(i, j1) = "_Bond_type" Then new_table_dat(i, i1) = "VARCHAR(127)"
        
                End If
                If new_table_dat(i, i1) = "" Then new_table_dat(i, i1) = "VARCHAR(127)"
            Next j1
        End If
    End If

  Next j
    For j1 = 1 To RDB_field_count
        comma_flag = 0
        If old_table_dat(1, j1) = "Data Type" Then
            If new_table_dat(i, j1) = "" Then new_table_dat(i, j1) = t1
            If new_table_dat(i, j1) = "" Then new_table_dat(i, j1) = "VARCHAR(127)"
        End If
        If old_table_dat(1, j1) = "Nullable" Then
            If new_table_dat(i, j1) = "" Then new_table_dat(i, j1) = t2
        End If
        If old_table_dat(1, j1) = "User full view" Then
            If new_table_dat(i, j1) = "" Then new_table_dat(i, j1) = t3
            If new_table_dat(i, j1) = "" Then new_table_dat(i, j1) = "Y"
            If new_table_dat(i, 9) = "_Saveframe_framecode" Then new_table_dat(i, j1) = "N"
        End If
        'If old_table_dat(3, j1) = "View" Then
        '    If new_table_dat(i, j1) <> t3 Then new_table_dat(i, j1) = t3
        'End If
            'If old_table_dat(1, j1) = "Primary Key" Then
            '    If new_table_dat(i, j1) = "" Then
            '        new_table_dat(i, j1) = t7
            '    End If
            '    If new_table_dat(i, 9) = "_Saveframe_ID" Then new_table_dat(i, j1) = "Ys"
            'End If

            'If old_table_dat(1, j1) = "Foreign Key Group" Then
            '    If new_table_dat(i, j1) = "" Then
            '        p = InStr(1, new_table_dat(i, 9), "_entry_ID")
            '        If p > 0 Then new_table_dat(i, j1) = "1"
            '        p = InStr(1, new_table_dat(i, 9), "_saveframe_ID")
            '        If p > 0 Then new_table_dat(i, j1) = "2"
            '        p = InStr(1, new_table_dat(i, 9), "_Entry_ID")
            '        If p > 0 Then new_table_dat(i, j1) = "1"
            '    End If
            'End If
            'If old_table_dat(1, j1) = "Foreign Table" Then
            '    If new_table_dat(i, j1) = "" Then
            '        p = InStr(1, new_table_dat(i, 9), "_saveframe_ID")
            '        If p > 0 Then new_table_dat(i, j1) = primary_table
            '        p = InStr(1, new_table_dat(i, 9), "_entry_ID")
            '        If p > 0 Then new_table_dat(i, j1) = "EntryInformation"
            '        p = InStr(1, new_table_dat(i, 9), "_Entry_ID")
            '        If p > 0 Then new_table_dat(i, j1) = "EntryInformation"
            '    End If
            'End If
            'If old_table_dat(1, j1) = "Foreign Column" Then
            '    If new_table_dat(i, j1) = "" Then
            '        p = InStr(1, new_table_dat(i, 9), "_saveframe_ID")
            '        If p > 0 Then new_table_dat(i, j1) = primary_column
            '        p = InStr(1, new_table_dat(i, 9), "_entry_ID")
            '        If p > 0 Then new_table_dat(i, j1) = "EntryID"
            '        p = InStr(1, new_table_dat(i, 9), "_Entry_ID")
            '        If p > 0 Then new_table_dat(i, j1) = "EntryID"
            '    End If
            'End If

            If old_table_dat(1, j1) = "Example" Then
                If new_table_dat(i, j1) = "" Or new_table_dat(i, j1) = Chr$(34) + "?" + Chr$(34) Then new_table_dat(i, j1) = "example_text"
                If new_table_dat(i, j1) = Chr$(34) + Chr$(34) Then new_table_dat(i, j1) = "example_text"
                If new_table_dat(i, j1) = Chr$(34) + "@" + Chr$(34) Then new_table_dat(i, j1) = "example_text"
                comma_flag = 1
            End If
            If old_table_dat(1, j1) = "Help" Then
                If new_table_dat(i, j1) = "" Or new_table_dat(i, j1) = Chr$(34) + "?" + Chr$(34) Then new_table_dat(i, j1) = "help_text"
                If new_table_dat(i, j1) = Chr$(34) + Chr$(34) Then new_table_dat(i, j1) = "help_text"
                If new_table_dat(i, j1) = Chr$(34) + "na" + Chr$(34) Then new_table_dat(i, j1) = "help_text"
                If new_table_dat(i, j1) = Chr$(34) + "." + Chr$(34) Then new_table_dat(i, j1) = "help_text"
                If new_table_dat(i, j1) = "@" Then new_table_dat(i, j1) = "help_text"
                comma_flag = 1
            End If
            If old_table_dat(1, j1) = "Description" Then
                If new_table_dat(i, j1) = "" Or new_table_dat(i, j1) = "?" Then new_table_dat(i, j1) = "description_text"
                If new_table_dat(i, j1) = Chr$(34) + "    ? " + Chr$(34) Then new_table_dat(i, j1) = "description_text"
                If new_table_dat(i, j1) = "@" Then new_table_dat(i, j1) = "description_text"
                If new_table_dat(i, j1) = Chr$(34) + "@" + Chr$(34) Then new_table_dat(i, j1) = "description_text"
                comma_flag = 1
            End If
            If old_table_dat(1, j1) = "User structure view" Then
                If new_table_dat(i, j1) = "" Then new_table_dat(i, j1) = "H"
            End If
            If old_table_dat(1, j1) = "User non-structure view" Then
                If new_table_dat(i, j1) = "" Then new_table_dat(i, j1) = "H"
            End If
            If old_table_dat(1, j1) = "User NMR param. View" Then
                If new_table_dat(i, j1) = "" Then new_table_dat(i, j1) = "H"
            End If
            If old_table_dat(1, j1) = "Annotator full view" Then
                If new_table_dat(i, j1) = "" Then new_table_dat(i, j1) = "H"
            End If
            If old_table_dat(1, j1) = "Item enumeration closed" Then
                new_table_dat(i, j1) = "N"
            End If
            If old_table_dat(1, j1) = "Item enumerated" Then
                If new_table_dat(i, j1) = "" Then new_table_dat(i, j1) = "N"
            End If
           
            If old_table_dat(1, j1) = "Dbspace" Then
                If new_table_dat(i, j1) = "" Then new_table_dat(i, j1) = "1"
                'If i < 902 Or i > 920 Then
                '    new_table_dat(i, j1) = "1"                  'BMRB  DB space
                'Else
                '    new_table_dat(i, j1) = "2"
                'End If
            End If
            If old_table_dat(1, j1) = "ADIT category group ID" Then
                For j2 = 1 To grp_row_ct
                    If new_table_dat(i, 2) = grp_table_row(j2, 4) Then
                        new_table_dat(i, 7) = grp_table_row(j2, 3)
                        new_table_dat(i, 8) = grp_table_row(j2, 14)
                        new_table_dat(i, 5) = grp_table_row(j2, 1)
                        new_table_dat(i, 6) = grp_table_row(j2, 2)
                        Exit For
                    End If
                Next j2
            End If
            'If old_table_dat(1, j1) = "ADIT super category ID" Then
            '    For j2 = 1 To spg_row_ct
            '        If new_table_dat(i, 5) = spg_table_row(j2, 2) Then
            '            new_table_dat(i, 6) = new_table_dat(i, 5)
            '            new_table_dat(i, 5) = spg_table_row(j2, 1)
            '            Exit For
            '        End If
            '    Next j2
            'End If

            
            'If old_table_dat(1, j1) = "ADIT category mandatory" Then
            '    new_table_dat(i, j1) = "Y"
            'End If
            'If old_table_dat(1, j1) = "ADIT category view type" Then
            '    new_table_dat(i, j1) = "R"
            '    If old_table_dat(i, j1) > "" Then
            '        new_table_dat(i, j1) = old_table_dat(i, j1)
            '    End If
            'End If
            'If old_table_dat(1, j1) = "ADIT category group description" Then
            '    new_table_dat(i, j1) = "ADIT category group description"
            'End If
            'If old_table_dat(1, j1) = "ADIT category view name" Then
            '    new_table_dat(i, j1) = sf_category  'category view name
            'End If
        If old_table_dat(1, j1) = "ADIT item view name" Then
            If tag_name > "" Then
                If new_table_dat(i, j1) = "" Then
                    new_table_dat(i, j1) = Right$(tag_name, (Len(tag_name) - 1))
                    new_tag = ""
                    For j3 = 1 To Len(new_table_dat(i, j1))
                        tj5 = Mid$(new_table_dat(i, j1), j3, 1)
                        If tj5 = "_" Then tj5 = " "
                        new_tag = new_tag + tj5
                    Next j3
                    new_table_dat(i, j1) = new_tag
                End If
            End If
        End If
        If old_table_dat(1, j1) = "ManDBColumnName" Then
        'Debug.Print
            p3 = InStr(1, new_table_dat(i, j1), "_entry_ID")
            p4 = InStr(1, new_table_dat(i, j1), "_saveframe_ID")
            If p3 > 1 Or p4 > 1 Then
                If p3 > 1 Then new_table_dat(i, j1) = "EntryID"
                If p4 > 1 Then new_table_dat(i, j1) = "SaveframeID"
            End If
            If p3 < 1 And p4 < 1 Then
                colname = new_table_dat(i, j1)
                remove_hyphen colname
                new_table_dat(i, j1) = colname
                If Len(new_table_dat(i, j1)) > 31 Then
                    colname = new_table_dat(i, j1)
                    shorten_name colname
                    new_table_dat(i, j1) = colname
                    If Len(new_table_dat(i, j1)) > 31 Then
                        Debug.Print new_table_dat(i, j1)
                        Debug.Print
                    End If
                End If
            End If
            If Right$(new_table_dat(i, 9), 5) = "sf_ID" Then
                For j3 = 1 To grp_row_ct
                    p = InStr(1, UCase(new_table_dat(i, 9)), UCase(grp_table_row(j3, 4) + "_SF_ID"))
                    p2 = InStr(1, UCase(new_table_dat(i, 9)), UCase(new_table_dat(i, 2)))
                    If p = 2 Then
                        If p2 = 2 Then
                            sf_cat(j3, 1) = new_table_dat(i, 35)
                            sf_cat(j3, 2) = new_table_dat(i, 36)
                            'Debug.Print
                        End If
                    End If
                Next j3
                'Debug.Print
            End If

            'Debug.Print
        End If
        If old_table_dat(1, j1) = "ManDBTableName" Then
            p = InStr(1, new_table_dat(i, j1), "_entry_ID")
            If p > 0 Then new_table_dat(i, j1) = Left$(new_table_dat(i, j1), p - 1)
            colname = new_table_dat(i, j1)
            'Debug.Print
            remove_hyphen colname
            new_table_dat(i, j1) = colname

            If Len(new_table_dat(i, j1)) > 31 Then
                colname = new_table_dat(i, j1)
                shorten_name colname
                new_table_dat(i, j1) = colname
                If Len(new_table_dat(i, j1)) > 31 Then
                    Debug.Print new_table_dat(i, j1)
                    Debug.Print
                End If
                
            End If
            If new_table_dat(i, 9) = "_Saveframe_category" Then
                    primary_table = new_table_dat(i, j1)
                    primary_column = "SaveframeID"
            End If


        End If

        If old_table_dat(1, j1) = "Loopflag" Then
            If new_table_dat(i, j1) = "Y" Then new_table_dat(i, 4) = "T6"
            If new_table_dat(i, j1) = "N" Then new_table_dat(i, 4) = "R"
        End If
        If old_table_dat(1, j1) = "ADIT category mandatory" Then
            new_table_dat(i, j1) = "Y"
        End If
        If old_table_dat(1, j1) = "SFCategory" Then
            get_out_of_loops = 0
            For j3 = 1 To grp_row_ct
                If new_table_dat(i, j1) = grp_table_row(j3, 4) Then
                    For j4 = 1 To spg_row_ct
                        If grp_table_row(j3, 2) = spg_table_row(j4, 2) Then
                            new_table_dat(i, 5) = spg_table_row(j4, 2)
                            For j5 = 1 To RDB_field_count
                                For j6 = 1 To spg_col_ct
                                    If old_table_dat(1, j5) = spg_table_row(1, j6) Then
                                    If old_table_dat(3, j5) = "View" Then
                                        If new_table_dat(i, j5) = "" Then new_table_dat(i, j5) = spg_table_row(j4, j6)
                                        get_out_of_loops = 1: Exit For
                                    End If
                                    End If
                                Next j6
                            Next j5
                        End If
                        If get_out_of_loops = 1 Then Exit For
                    Next j4
                End If
                If get_out_of_loops = 1 Then Exit For
            Next j3
        End If
        
            If comma_flag = 1 Then
                p = 1: p1 = 1
                While p <> 0
                    p = InStr(p1, new_table_dat(i, j1), ",")
                    'p3 = InStr(p1, new_table_dat(i, j1), "\,")
                    If p > 0 Then
                        'If p - p3 <> 1 Then
                            ln = Len(new_table_dat(i, j1))
                            test1 = Left$(new_table_dat(i, j1), p - 1)
                            test2 = Right$(new_table_dat(i, j1), (ln - p))
                            new_table_dat(i, j1) = test1 + "$" + test2
                        'End If
                        
                        p1 = p + 2
                    End If
                Wend
                'Debug.Print new_table_dat(i, j1)
                'Debug.Print
                comma_flag = 0
            End If
    
    Next j1
 For j2 = 1 To RDB_field_count
    If j2 < RDB_field_count Then Print #1, new_table_dat(i, j2); ",";
    If j2 = RDB_field_count Then Print #1, new_table_dat(i, j2)
    'Debug.Print i, j2, new_table_dat(i, j2)
 Next j2

' The new_table_dat index needs to be updated in these loops when the
' number of columns in the xlschema file is changed.
' The following routine calculates the values for the seq column

 new_table_dat(i, 1) = Str(i * 10)
 If sf_category = sf_cat_last Then
    If new_table_dat(i, 37) = "Y" Then
        loop_ct = loop_ct + 1
        For j2 = 1 To RDB_field_count
            loop_table_row(loop_ct, j2) = new_table_dat(i, j2)
        Next j2
        'Debug.Print
    End If
    If new_table_dat(i, 37) = "N" Then
        For j2 = 1 To RDB_field_count
            If j2 < RDB_field_count Then Print #5, new_table_dat(i, j2); ",";
            If j2 = RDB_field_count Then Print #5, new_table_dat(i, j2)
        Next j2
        'Debug.Print
    End If
 End If
 If sf_category <> sf_cat_last Then
    For j3 = 1 To loop_ct
        For j2 = 1 To RDB_field_count
            If j2 < RDB_field_count Then Print #5, loop_table_row(j3, j2); ",";
            If j2 = RDB_field_count Then Print #5, loop_table_row(j3, j2)
        Next j2
        'Debug.Print
    Next j3
    loop_ct = 0
    ReDim loop_table_row(1 To 300, 1 To RDB_field_count) As String

    If new_table_dat(i, 37) = "Y" Then
        loop_ct = loop_ct + 1
        For j2 = 1 To RDB_field_count
            loop_table_row(loop_ct, j2) = new_table_dat(i, j2)
        Next j2
        'Debug.Print
    End If
    If new_table_dat(i, 37) = "N" Then
        For j2 = 1 To RDB_field_count
            If j2 < RDB_field_count Then Print #5, new_table_dat(i, j2); ",";
            If j2 = RDB_field_count Then Print #5, new_table_dat(i, j2)
        Next j2
        'Debug.Print
    End If
    sf_cat_last = sf_category
 End If
    
Next i

If loop_ct > 0 Then
    For j3 = 1 To loop_ct
        For j2 = 1 To RDB_field_count
            If j2 < RDB_field_count Then Print #5, loop_table_row(j3, j2); ",";
            If j2 = RDB_field_count Then Print #5, loop_table_row(j3, j2)
        Next j2
        'Debug.Print
    Next j3
End If

Print #1, "TBL_END";
For i = 2 To RDB_field_count
    Print #1, ",";
Next i
Print #1, "?"

Print #5, "TBL_END";
For i = 2 To RDB_field_count
    Print #5, ",";
Next i
Print #5, "?"

Close 1
Close 5


' Code to mark RDB primary and foreign keys and to insert table and column
' definitions

tables = 1
If tables = 1 Then
' Reload the variable that holds the Excel file data

Open "c:\bmrb\htdocs\dictionary\htmldocs\nmr_star\adit_files\parent_child.csv" For Output As 1
Open NMR_STAR_dict_path + new_table For Input As 5
new_tbl_ct = 0
While EOF(5) <> -1
    new_tbl_ct = new_tbl_ct + 1
    For i = 1 To RDB_field_count
        Input #5, new_table_dat(new_tbl_ct, i)
    Next i
Wend
Close 5


For i = 1 To new_tbl_ct
    p2 = 0
    ' set up to identify and insert key information for Entry ID
    If new_table_dat(i, 9) = "_Entry_ID" Then
        If new_table_dat(i, 2) = "entry_information" Then
            foreign_table = new_table_dat(i, 35)
            foreign_col = new_table_dat(i, 36)
        End If
    End If
    ' set up to identify and insert key information for mol_system_atom_ID
    If new_table_dat(i, 9) = "_Mol_system_atom_ID" Then
    If new_table_dat(i, 2) = "covalent_definitions" Then
        foreign_table4 = new_table_dat(i, 35)
        foreign_col4 = new_table_dat(i, 36)
    End If
    End If
    
    If new_table_dat(i, 9) = "_Mol_system_component_ID" Then
    If new_table_dat(i, 2) = "mol_system" Then
        foreign_table5 = new_table_dat(i, 35)
        foreign_col5 = new_table_dat(i, 36)
    End If
    End If
    If new_table_dat(i, 9) = "_Chemical_moiety_index_ID" Then
    If new_table_dat(i, 2) = "molecule" Then
        foreign_table6 = new_table_dat(i, 35)
        foreign_col6 = new_table_dat(i, 36)
    End If
    End If
    If new_table_dat(i, 9) = "_chemical_compound_atom_ID" Then
    If new_table_dat(i, 2) = "chemical_compound" Then
        foreign_table7 = new_table_dat(i, 35)
        foreign_col7 = new_table_dat(i, 36)
    End If
    End If
    
    If new_table_dat(i, 9) = "_Saveframe_category" Then
        category = UCase(new_table_dat(i, 2))
    End If
   
    p5 = InStr(1, UCase(new_table_dat(i, 9)), category + "_SF_ID")
    If p5 = 2 Then
        p2 = 1
        new_table_dat(i, 30) = "Y"
        foreign_table2 = new_table_dat(i, 35)
        foreign_col2 = new_table_dat(i, 36)
    End If
    p = InStr(1, UCase(new_table_dat(i, 9)), "SF_ID")
    If p > 0 And p5 <> 2 Then
        If UCase(new_table_dat(i, 2)) = category Then
            p2 = 1
            new_table_dat(i, 30) = "Y"
            new_table_dat(i, 31) = "2"
            new_table_dat(i, 32) = foreign_table2
            new_table_dat(i, 33) = foreign_col2
        End If
    End If
    
    If new_table_dat(i, 9) = "_Keyword" Then new_table_dat(i, 30) = "Y"
    If new_table_dat(i, 9) = "_Synonym" Then new_table_dat(i, 30) = "Y"
    If new_table_dat(i, 9) = "_Biological_function" Then new_table_dat(i, 30) = "Y"
    
    If Right$(new_table_dat(i, 9), 2) = "ID" Then
        If p2 <> 1 Then
            For j3 = 1 To grp_row_ct
                p3 = InStr(1, UCase(new_table_dat(i, 9)), (UCase(grp_table_row(j3, 4) + "_ID")))
                If p3 > 0 Then
                    new_table_dat(i, 31) = Str(j3 + 100)
                    new_table_dat(i, 32) = sf_cat(j3, 1)
                    new_table_dat(i, 33) = sf_cat(j3, 2)
                    'Debug.Print
                End If
            Next j3
        End If
    End If
Next i

For j = 1 To new_tbl_ct
            If new_table_dat(j, 9) = "_Saveframe_category" Then
                category = UCase(new_table_dat(j, 2))
            End If
            If new_table_dat(j, 9) = "_Entry_ID" Then
                foreign_table3 = new_table_dat(j, 35)
                foreign_col3 = new_table_dat(j, 36)
            End If
            p1 = InStr(1, UCase(new_table_dat(j, 9)), "_ENTRY_ID")
            If p1 > 0 Then new_table_dat(j, 30) = "Y"
            If p1 = 1 And category <> "ENTRY_INFORMATION" Then
                new_table_dat(j, 31) = "1"
                new_table_dat(j, 32) = foreign_table
                new_table_dat(j, 33) = foreign_col
                'Debug.Print
            End If
            If p1 > 2 Then
                new_table_dat(j, 30) = "Y"
                new_table_dat(j, 31) = "2"
                new_table_dat(j, 32) = foreign_table3
                new_table_dat(j, 33) = foreign_col3
                'Debug.Print
            End If

            p1 = InStr(1, UCase(new_table_dat(j, 9)), "_CHEMICAL_MOIETY_INDEX_ID")
            If p1 = 1 And category <> "MOLECULE" Then
                new_table_dat(j, 31) = "6"
                new_table_dat(j, 32) = foreign_table6
                new_table_dat(j, 33) = foreign_col6
                'Debug.Print
            End If
            If p1 > 1 Then
                new_table_dat(j, 31) = "6"
                new_table_dat(j, 32) = foreign_table6
                new_table_dat(j, 33) = foreign_col6
                'Debug.Print
            End If
            p1 = InStr(1, UCase(new_table_dat(j, 9)), "_MOL_SYSTEM_COMPONENT_ID")
            If p1 = 1 And category <> "MOL_SYSTEM" Then
                new_table_dat(j, 31) = "5"
                new_table_dat(j, 32) = foreign_table5
                new_table_dat(j, 33) = foreign_col5
                'Debug.Print
            End If
            If p1 > 1 Then
                new_table_dat(j, 31) = "5"
                new_table_dat(j, 32) = foreign_table5
                new_table_dat(j, 33) = foreign_col5
                'Debug.Print
            End If
            p1 = InStr(1, UCase(new_table_dat(j, 9)), "_MOL_SYSTEM_ATOM_ID")
            If p1 = 1 And category <> "COVALENT_DEFINITIONS" Then
                new_table_dat(j, 31) = "4"
                new_table_dat(j, 32) = foreign_table4
                new_table_dat(j, 33) = foreign_col4
                'Debug.Print
            End If
            If p1 > 1 Then
                new_table_dat(j, 31) = "4"
                new_table_dat(j, 32) = foreign_table4
                new_table_dat(j, 33) = foreign_col4
                'Debug.Print
            End If
            p1 = InStr(1, UCase(new_table_dat(j, 9)), "_CHEMICAL_COMPOUND_ATOM_ID")
            If p1 = 1 And category <> "CHEMICAL_COMPOUND" Then
                new_table_dat(j, 31) = "7"
                new_table_dat(j, 32) = foreign_table7
                new_table_dat(j, 33) = foreign_col7
                'Debug.Print
            End If
            If p1 > 1 Then
                new_table_dat(j, 31) = "7"
                new_table_dat(j, 32) = foreign_table7
                new_table_dat(j, 33) = foreign_col7
                'Debug.Print
            End If
            
  'Check for duplicate table names
    'For i = 1 To new_tbl_ct
    '    If UCase(new_table_dat(i, 35)) = UCase(new_table_dat(j, 35)) Then
    '        If i <> j Then
            'If new_table_dat(i, 2) <> new_table_dat(j, 2) Then
    '            Debug.Print "dup table", i, j, new_table_dat(j, 35)
            'End If
    '        End If
    '    End If
    'Next i
            
Next j
    
Open NMR_STAR_dict_path + new_table For Output As 5
For i = 1 To new_tbl_ct
    For j = 1 To RDB_field_count
        If j < RDB_field_count Then Print #5, new_table_dat(i, j) + ",";
        If j = RDB_field_count Then Print #5, new_table_dat(i, j)
    Next j
Next i
Close 1, 5
End If
End Sub

Sub shorten_name(colname)

check_flag = 1
While check_flag = 1
    check_flag = 0
    
    test_string = "Abbreviation": sub_string = "Abbrv"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Absorption": sub_string = "Absorp"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Accession": sub_string = "Access"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Acquisition": sub_string = "Acq"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Action": sub_string = "Act"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Align": sub_string = "Algn"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Alternate": sub_string = "Alt"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Ambiguous": sub_string = "Amb"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Angle": sub_string = "Ang"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Anisotropy": sub_string = "Aniso"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Annotation": sub_string = "Anno"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Applied": sub_string = "Appl"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Assigned": sub_string = "Asgn"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Assign": sub_string = "Asgn"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Atom": sub_string = "Atm"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag

    test_string = "Author": sub_string = "Auth"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
   
    test_string = "Average": sub_string = "Avg"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
   
    test_string = "Bond": sub_string = "Bnd"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "13C-13C": sub_string = "C-C"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Calculation": sub_string = "Calc"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Carbon_Nitrogen": sub_string = "C-N"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Category": sub_string = "Cat"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Characteristic": sub_string = "Char"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Chemical": sub_string = "Chm"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Chem": sub_string = "Chm"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Chirality": sub_string = "Chir"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Citation": sub_string = "Cit"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Combination": sub_string = "Comb"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Commission": sub_string = "Comm"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Common": sub_string = "Com"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Component": sub_string = "Comp"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Compound": sub_string = "Compd"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Conditions": sub_string = "Cond"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Conformation": sub_string = "Conf"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Conformer": sub_string = "Conf"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Constant": sub_string = "Const"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Constraint": sub_string = "Cnstr"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Constraints": sub_string = "Cnstr"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Coordinate": sub_string = "Coord"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Correlation": sub_string = "Corr"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Correlations": sub_string = "Corr"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Coupling": sub_string = "Coup"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Count": sub_string = "Ct"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Criteria": sub_string = "Crit"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Database": sub_string = "Db"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Defining": sub_string = "Def"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Derivation": sub_string = "Der"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Derivative": sub_string = "Deriv"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Derived": sub_string = "Der"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Description": sub_string = "Desc"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Details": sub_string = "Det"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Development": sub_string = "Dev"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Difference": sub_string = "Diff"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Dipolar": sub_string = "Dipo"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Dipole": sub_string = "Dipol"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Distance": sub_string = "Dis"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Energetic": sub_string = "Ener"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Ensemble": sub_string = "Ens"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Equivalent": sub_string = "Equiv"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Error": sub_string = "Err"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Exchange": sub_string = "Exch"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Excipient": sub_string = "Excip"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Expectation": sub_string = "Expect"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Experiment": sub_string = "Exp"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Experimental": sub_string = "Exptl"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
   
    test_string = "Factors": sub_string = "Fact"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Family": sub_string = "Fam"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Features": sub_string = "Feat"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Floating": sub_string = "Float"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Fractionation": sub_string = "Frac"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Function": sub_string = "Func"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Generalized": sub_string = "Gen"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "1H-1H": sub_string = "H-H"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Heterogeneous": sub_string = "Hetero"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Heteronuclear": sub_string = "Hetnuc"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Homology": sub_string = "Homol"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
   
    test_string = "Homonuclear": sub_string = "Homonuc"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Hydrogen": sub_string = "H"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Index": sub_string = "Ind"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Information": sub_string = "Info"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Input": sub_string = "Inp"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Intensity": sub_string = "Intens"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Interaction": sub_string = "Inter"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Intermolecular": sub_string = "IntMol"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Inter-molecular": sub_string = "IntMol"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Intramolecular": sub_string = "IntraMol"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Intraresidue": sub_string = "IntraRes"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    'test_string = "Inter": sub_string = "Inter"
    'p = InStr(1, colname, test_string)
    'If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Int-mol": sub_string = "IntMol"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Intra-mol": sub_string = "IntraMol"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Isotope": sub_string = "Iso"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Letter": sub_string = "Let"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Library": sub_string = "Lib"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Link": sub_string = "Lnk"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Linkage": sub_string = "Lnk"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "List": sub_string = "Lst"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Literature": sub_string = "Lit"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Location": sub_string = "Loc"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Magnetically": sub_string = "Magn"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Magnetization": sub_string = "Magnz"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Manually": sub_string = "Man"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Matched": sub_string = "Mat"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Measurement": sub_string = "Meas"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Member": sub_string = "Memb"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Method": sub_string = "Meth"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Molecular": sub_string = "Mol"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Natural": sub_string = "Nat"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Nomenclature": sub_string = "Nom"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Number": sub_string = "Num"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag

    test_string = "Offset": sub_string = "Off"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Order": sub_string = "Ord"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag

    test_string = "Original": sub_string = "Orig"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Pairwise": sub_string = "Prws"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
     
    test_string = "Paramagnetic": sub_string = "Paramag"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Parameter": sub_string = "Par"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag

    test_string = "Peak": sub_string = "Pk"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Peptide": sub_string = "Pept"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Percentage": sub_string = "Pct"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Person": sub_string = "Prs"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
   
    test_string = "Planarity": sub_string = "Planar"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Polymer": sub_string = "Poly"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Preparation": sub_string = "Prep"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Primary": sub_string = "Pri"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Processing": sub_string = "Proc"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Protection": sub_string = "Pro"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Protocol": sub_string = "Prot"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Proton_Nitrogen": sub_string = "H-N"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Reference": sub_string = "Ref"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Relation": sub_string = "Relat"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Relaxation": sub_string = "Relax"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Reported": sub_string = "Report"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag

    test_string = "Representation": sub_string = "Rep"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Representative": sub_string = "Rep"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Residual": sub_string = "Res"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Residue": sub_string = "Res"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag

    test_string = "Revised": sub_string = "Rev"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
  
    test_string = "Sample": sub_string = "Samp"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Saveframe": sub_string = "Sf"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
   
    test_string = "Schematic": sub_string = "Schem"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Scheme": sub_string = "Sch"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Secondary": sub_string = "Secd"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
   
    test_string = "Second": sub_string = "Sec"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
     
    test_string = "Segment": sub_string = "Seg"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag

    test_string = "Selection": sub_string = "Sel"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Sequence": sub_string = "seq"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag

    test_string = "Sequential": sub_string = "Seq"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Shift": sub_string = "Shif"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Software": sub_string = "Soft"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Solute": sub_string = "Sol"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Solvent": sub_string = "Solv"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag

    test_string = "Source": sub_string = "Src"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
   
    test_string = "Spectral": sub_string = "Spc"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Standard": sub_string = "Std"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Statistics": sub_string = "Stats"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Stereoassignments": sub_string = "Sterassig"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Stereochemistry": sub_string = "Stereo"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag

    test_string = "Structure": sub_string = "Struc"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Submitted": sub_string = "Submit"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
     
    test_string = "Structure": sub_string = "Struc"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Subsystem": sub_string = "Subsys"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Symmetric": sub_string = "Sym"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Systematic": sub_string = "Syst"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "System": sub_string = "Sys"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
   
    test_string = "Tensor": sub_string = "Tens"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Tertiary": sub_string = "Tert"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Theoretical": sub_string = "Theor"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag

    test_string = "Three": sub_string = "Thre"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Torsion": sub_string = "Tor"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Total": sub_string = "Tot"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Transition": sub_string = "Trns"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Unambiguous": sub_string = "Unamb"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Unique": sub_string = "Uniq"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Value": sub_string = "Val"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Violated": sub_string = "Viol"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    
    test_string = "Violation": sub_string = "Viol"
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag
    

Wend
End Sub

Sub shorten_name2(p, colname, test_string, sub_string, check_flag)

ln = Len(colname)
ln2 = Len(test_string)
If p > 1 And p + ln2 <= ln Then colname = Left$(colname, p - 1) + sub_string + Right$(colname, ln - (p + ln2 - 1))
If p + ln2 > ln Then colname = Left$(colname, p - 1) + sub_string
If p = 1 Then colname = sub_string + Right$(colname, ln - ln2)
check_flag = 1

End Sub

Sub remove_hyphen(colname)

ln = Len(colname)
If Asc(Left$(colname, 1)) > 96 Then
    colname = UCase(Left$(colname, 1)) + Right$(colname, ln - 1)
End If

check_flag = 1
While check_flag = 1
check_flag = 0
For i = 65 To 90
    test_string = "_" + Chr$(i): sub_string = Chr$(i)
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag

    test_string = "_" + Chr$(i + 32): sub_string = Chr$(i)
    p = InStr(1, colname, test_string)
    If p > 0 Then shorten_name2 p, colname, test_string, sub_string, check_flag

Next i
Wend
'Debug.Print
End Sub
