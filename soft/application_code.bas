Attribute VB_Name = "application_code"
Sub read_excel(e_tag_ct, excel_tag_dat, excel_header_dat, pathin, input_file, e_col_num)
ReDim excel_tag_dat(10000, e_col_num) As String
ReDim excel_header_dat(5, 200) As String
Dim i, j As Integer

Open pathin + input_file For Input As 1
For i = 1 To 4                              ' load header rows from excel file
    For j = 1 To e_col_num
        Input #1, excel_header_dat(i, j)
    Next j
    'Debug.Print excel_header_dat(i, 106)
Next i
e_tag_ct = 0
While EOF(1) <> -1                          ' load tag rows
    e_tag_ct = e_tag_ct + 1
    For j = 1 To e_col_num
        Input #1, excel_tag_dat(e_tag_ct, j)
    Next j
    'Debug.Print e_tag_ct, excel_tag_dat(e_tag_ct, 1),
    'Debug.Print excel_tag_dat(e_tag_ct, 53),
    'Debug.Print excel_tag_dat(e_tag_ct, 100)
    'If excel_tag_dat(e_tag_ct, 95) <> "?" Then
        'Debug.Print excel_tag_dat(e_tag_ct, 104)
        'Debug.Print
    'End If
    'Debug.Print
Wend
If excel_tag_dat(e_tag_ct, 1) = "TBL_END" Then e_tag_ct = e_tag_ct - 1
Close 1

End Sub

Sub check_excel(e_tag_ct, excel_tag_dat, excel_header_dat, e_col_num, spg_col_ct, spg_row_ct, spg_table, grp_col_ct, grp_row_ct, grp_table, pathin, adit_file_source)
Dim i, p, p1, p2, p3, p4, table_ct, null_ct, sf_tag_flg, entry_tag_flg, sf_table_tag_flg, enum_ct As Integer
Dim tag_col, sf_col, dbtab, dbcol, cat_col, prompt As Integer
Dim usr_view, sf_id, sf_id_set, dat_type, src_key, row_id, adit_spg, adit_cg, loop_col As Integer
Dim enum_test_ct As Integer
Dim enum_test(4) As String

Dim sf_last As String

ReDim table_list(500) As String
ReDim null_list(50)
ReDim enum_list(500) As String

If adit_file_source < 2 Or adit_file_source = 4 Then
Open pathin + "enumerations.txt" For Input As 21
'Debug.Print pathin
'Debug.Print
While EOF(21) <> -1
    t = Input$(1, 21)
    t1 = Asc(t)
    If t1 <> 10 And t1 <> 13 Then tt = tt + t
    If t1 = 10 Then
        p = InStr(1, tt, "save__")
        If p = 1 Then
            ln = Len(tt)
            enum_ct = enum_ct + 1
            enum_list(enum_ct) = Trim(Right$(tt, ln - 5))
        End If
        tt = ""
    End If
Wend
Close 21
Open pathin + "adit_enum_dtl.csv" For Input As 21
While EOF(21) <> -1
    enum_test_ct = enum_test_ct + 1
    For i = 1 To 4
        Input #21, enum_test(i)
    Next i
    If enum_test(3) < " " Then
        syntax_check.Text6 = ""
        syntax_check.Text6.Refresh
        syntax_check.Text7 = ""
        syntax_check.Text7.Refresh
        syntax_check.Text8 = ""
        syntax_check.Text8.Refresh
        syntax_check.Text9 = enum_test_ct
        syntax_check.Text9.Refresh
        syntax_check.Text18 = "Enumeration file corrupted"
        syntax_check.Text18.Refresh

        program_control.Show 1
    End If
Wend
Close 21
End If

For i = 1 To e_col_num
    If excel_header_dat(1, i) = "Tag" Then tag_col = i: null_ct = null_ct + 1: null_list(null_ct) = i
    If excel_header_dat(1, i) = "SFCategory" Then sf_col = i: null_ct = null_ct + 1: null_list(null_ct) = i
    'If excel_header_dat(1, i) = "ManDBColumnName" Then dbcol = i: null_ct = null_ct + 1: null_list(null_ct) = i
    'If excel_header_dat(1, i) = "ManDBTableName" Then dbtab = i: null_ct = null_ct + 1: null_list(null_ct) = i
    If excel_header_dat(1, i) = "ADIT category view name" Then cat_col = i: null_ct = null_ct + 1: null_list(null_ct) = i
    If excel_header_dat(1, i) = "Source Key" Then src_key = i
    
    If excel_header_dat(1, i) = "ADIT category mandatory" Then null_ct = null_ct + 1: null_list(null_ct) = i
    If excel_header_dat(1, i) = "ADIT category view type" Then null_ct = null_ct + 1: null_list(null_ct) = i
    If excel_header_dat(1, i) = "ADIT super category ID" Then adit_spg = i: null_ct = null_ct + 1: null_list(null_ct) = i
    If excel_header_dat(1, i) = "ADIT super category" Then null_ct = null_ct + 1: null_list(null_ct) = i
    If excel_header_dat(1, i) = "ADIT category group ID" Then null_ct = null_ct + 1: null_list(null_ct) = i
    If excel_header_dat(1, i) = "ADIT exists" Then null_ct = null_ct + 1: null_list(null_ct) = i
    If excel_header_dat(1, i) = "User full view" Then usr_view = i: null_ct = null_ct + 1: null_list(null_ct) = i
    If excel_header_dat(1, i) = "User structure view" Then null_ct = null_ct + 1: null_list(null_ct) = i
    If excel_header_dat(1, i) = "User non-structure view" Then null_ct = null_ct + 1: null_list(null_ct) = i
    If excel_header_dat(1, i) = "User NMR param. View" Then null_ct = null_ct + 1: null_list(null_ct) = i
    If excel_header_dat(1, i) = "Annotator full view" Then null_ct = null_ct + 1: null_list(null_ct) = i
    If excel_header_dat(1, i) = "Item enumerated" Then enum_col = i: null_ct = null_ct + 1: null_list(null_ct) = i
    If excel_header_dat(1, i) = "ADIT item view name" Then null_ct = null_ct + 1: null_list(null_ct) = i
    If excel_header_dat(1, i) = "Data Type" Then dat_type = i: null_ct = null_ct + 1: null_list(null_ct) = i
    If excel_header_dat(1, i) = "Row Index Key" Then row_id = i: null_ct = null_ct + 1: null_list(null_ct) = i
    If excel_header_dat(1, i) = "Saveframe ID tag" Then sf_id = i: null_ct = null_ct + 1: null_list(null_ct) = i
    If excel_header_dat(1, i) = "Loopflag" Then loop_col = i: null_ct = null_ct + 1: null_list(null_ct) = i
    If excel_header_dat(1, i) = "Seq" Then null_ct = null_ct + 1: null_list(null_ct) = i
    If excel_header_dat(1, i) = "Dbspace" Then null_ct = null_ct + 1: null_list(null_ct) = i
    'If excel_header_dat(1, i) = "Example" Then null_ct = null_ct + 1: null_list(null_ct) = i
    If excel_header_dat(1, i) = "Prompt" Then prompt = i: null_ct = null_ct + 1: null_list(null_ct) = i
    If excel_header_dat(1, i) = "Dictionary description" Then Desc = i: null_ct = null_ct + 1: null_list(null_ct) = i
    If excel_header_dat(1, i) = "Framecode value flag" Then frame_val = i
    If excel_header_dat(1, i) = "Foreign Table" Then foreignTab = i
    If excel_header_dat(1, i) = "Foreign Column" Then foreignCol = i
    If excel_header_dat(1, i) = "Table Primary Key" Then primarykeyCol = i
    If excel_header_dat(1, i) = "BMRB data type" Then bmrb_data_type = i
    If excel_header_dat(1, i) = "Nullable" Then null_col = i
    If excel_header_dat(1, i) = "Tag category" Then ref_cat = i
    If excel_header_dat(1, i) = "Tag field" Then ref_col = i


Next i

'initialization values
sf_tag_flg = 1
sf_table_tag_flg = 1
entry_tag_flg = 1

If adit_file_source < 2 Or adit_file_source = 4 Then
'search for orphan enumeration tags - enumerated tags that do not exist in the current dictionary
For j = 1 To enum_ct
    found_flg = 0
    
    For i = 1 To e_tag_ct
        If excel_tag_dat(i, tag_col) = enum_list(j) Then
          If excel_tag_dat(i, enum_col) = "Y" Then
            found_flg = 1
            'Debug.Print
            Exit For
          End If
        End If
    Next i
    If found_flg = 0 Then
        syntax_check.Text6 = ""
        syntax_check.Text6.Refresh
        syntax_check.Text7 = ""
        syntax_check.Text7.Refresh
        syntax_check.Text8 = ""
        syntax_check.Text8.Refresh
        syntax_check.Text9 = enum_list(j)
        syntax_check.Text9.Refresh
        syntax_check.Text18 = "Enumerated tag missing from dictionary"
        syntax_check.Text18.Refresh

        program_control.Show 1
    End If
Next j
End If

For i = 1 To e_tag_ct
    syntax_check.Text5 = Str(i)
    syntax_check.Text5.Refresh
    
    'check for nullable primary keys
    If excel_tag_dat(i, null_col) <> "NOT NULL" And excel_tag_dat(i, primarykeyCol) = "Y" Then
        Debug.Print "ERROR", i + 4, excel_tag_dat(i, 9), "Nullable primary key"
    End If
    
    'check enumerations
    If adit_file_source < 2 Or adit_file_source = 4 Then
    If excel_tag_dat(i, enum_col) = "Y" Then
        found_flg = 0
        For j = 1 To enum_ct
            If excel_tag_dat(i, tag_col) = enum_list(j) Then
                found_flg = 1
                Exit For
            End If
        Next j
        If found_flg = 0 Then
            syntax_check.Text6 = Str(i + 4)
            syntax_check.Text6.Refresh
            syntax_check.Text7 = excel_tag_dat(i, enum_col)
            syntax_check.Text7.Refresh
            syntax_check.Text8 = excel_tag_dat(i, cat_col)
            syntax_check.Text8.Refresh
            syntax_check.Text9 = excel_tag_dat(i, tag_col)
            syntax_check.Text9.Refresh
            syntax_check.Text18 = "Tag missing from enumeration list"
            syntax_check.Text18.Refresh

            program_control.Show 1
        End If
        'Debug.Print
    End If
    End If

    'check saveframes for Sf_ID and entry_ID tags
    If excel_tag_dat(i, sf_col) <> sf_last Then
        If sf_tag_flg = 0 Then
            syntax_check.Text6 = Str(i + 4)
            syntax_check.Text6.Refresh
            syntax_check.Text7 = excel_tag_dat(i, sf_col)
            syntax_check.Text7.Refresh
            syntax_check.Text8 = excel_tag_dat(i, cat_col)
            syntax_check.Text8.Refresh
            syntax_check.Text9 = excel_tag_dat(i, tag_col)
            syntax_check.Text9.Refresh
            syntax_check.Text18 = "Saveframe ID tag missing"
            syntax_check.Text18.Refresh

            program_control.Show 1
        
        End If
        If entry_tag_flg = 0 Then
            syntax_check.Text6 = Str(i + 4)
            syntax_check.Text6.Refresh
            syntax_check.Text7 = excel_tag_dat(i, sf_col)
            syntax_check.Text7.Refresh
            syntax_check.Text8 = excel_tag_dat(i, cat_col)
            syntax_check.Text8.Refresh
            syntax_check.Text9 = excel_tag_dat(i, tag_col)
            syntax_check.Text9.Refresh
            syntax_check.Text18 = "Entry ID tag missing"
            syntax_check.Text18.Refresh

            program_control.Show 1
        
        End If
        sf_last = excel_tag_dat(i, sf_col)
        sf_tag_flg = 0
        entry_tag_flg = 0
    End If
    If sf_tag_flg = 0 Then
        If excel_tag_dat(i, loop_col) = "N" Then
            If Right$(excel_tag_dat(i, tag_col), 6) = ".Sf_ID" Then sf_tag_flg = 1
            'Debug.Print
        End If
    End If
    If entry_tag_flg = 0 Then
        If excel_tag_dat(i, loop_col) = "N" Then
            If Right$(excel_tag_dat(i, tag_col), 9) = ".Entry_ID" Then entry_tag_flg = 1
            If excel_tag_dat(i, tag_col) = "_Entry.ID" Then entry_tag_flg = 1
        End If
    End If
    
    'create a list of tables
    If excel_tag_dat(i, cat_col) <> cat_last Then
        table_ct = table_ct + 1
        table_list(table_ct) = excel_tag_dat(i, cat_col)
        cat_last = excel_tag_dat(i, cat_col)
        syntax_check.Text4 = Str(table_ct)
        syntax_check.Text4.Refresh

        'check for missing saveframe ID in the loop
        If sf_id_set = 0 And i > 1 Then
            syntax_check.Text6 = Str(i + 4 - 1)
            syntax_check.Text6.Refresh
            syntax_check.Text7 = excel_tag_dat(i - 1, sf_col)
            syntax_check.Text7.Refresh
            syntax_check.Text8 = excel_tag_dat(i - 1, cat_col)
            syntax_check.Text8.Refresh
            syntax_check.Text9 = excel_tag_dat(i - 1, tag_col)
            syntax_check.Text9.Refresh
            syntax_check.Text18 = "Saveframe ID not defined"
            syntax_check.Text18.Refresh
            
            program_control.Show 1
        
        End If
        sf_id_set = 0
                
        'check for missing Entry ID in the loop
        If entry_id_set = 0 And i > 1 Then
            syntax_check.Text6 = Str(i + 4 - 1)
            syntax_check.Text6.Refresh
            syntax_check.Text7 = excel_tag_dat(i - 1, sf_col)
            syntax_check.Text7.Refresh
            syntax_check.Text8 = excel_tag_dat(i - 1, cat_col)
            syntax_check.Text8.Refresh
            syntax_check.Text9 = excel_tag_dat(i - 1, tag_col)
            syntax_check.Text9.Refresh
            syntax_check.Text18 = "Entry ID not defined"
            syntax_check.Text18.Refresh
            
            program_control.Show 1
        
        End If
        If primary_key = 0 And i > 1 Then
            Debug.Print "ERROR", excel_tag_dat(i - 1, cat_col), "No primary key defined"
        End If
        If foreign_key = 0 And i > 1 Then
            Debug.Print "ERROR", excel_tag_dat(i - 1, cat_col), "No foreign key defined"
        End If
        entry_id_set = 0
        primary_key = 0
        foreign_key = 0
    End If
    If excel_tag_dat(i, cat_col) = cat_last Then
        If excel_tag_dat(i, sf_id) = "Y" Then sf_id_set = 1
    End If
    If excel_tag_dat(i, cat_col) = cat_last Then
        If Right$(excel_tag_dat(i, tag_col), 9) = ".Entry_ID" Then entry_id_set = 1
        If excel_tag_dat(i, tag_col) = "_Entry.ID" Then entry_id_set = 1
        If excel_tag_dat(i, primarykeyCol) = "Y" Then primary_key = 1
        If excel_tag_dat(i, foreignTab) > " " Then foreign_key = 1
    End If
    
    'check saveframe ID data type, should be INTEGER)
    char_type = 0
    If excel_tag_dat(i, sf_id) = "Y" Then char_type = 1
    If char_type = 1 And excel_tag_dat(i, dat_type) <> "INTEGER" Then
        syntax_check.Text6 = Str(i + 4)
        syntax_check.Text6.Refresh
        syntax_check.Text7 = excel_tag_dat(i, sf_col)
        syntax_check.Text7.Refresh
        syntax_check.Text8 = excel_tag_dat(i, cat_col)
        syntax_check.Text8.Refresh
        syntax_check.Text9 = excel_tag_dat(i, tag_col)
        syntax_check.Text9.Refresh
        syntax_check.Text18 = "Data type not set to INTEGER"
        syntax_check.Text18.Refresh

        program_control.Show 1
        
    End If
    
    'check Framecode value flag, should be 'N' if tag does not contain 'label'
    char_type = 0
    If excel_tag_dat(i, frame_val) = "Y" Then char_type = 1
    p = InStr(1, excel_tag_dat(i, tag_col), "label")
    If char_type = 1 And p = 0 Then
        syntax_check.Text6 = Str(i + 4)
        syntax_check.Text6.Refresh
        syntax_check.Text7 = excel_tag_dat(i, sf_col)
        syntax_check.Text7.Refresh
        syntax_check.Text8 = excel_tag_dat(i, cat_col)
        syntax_check.Text8.Refresh
        syntax_check.Text9 = excel_tag_dat(i, tag_col)
        syntax_check.Text9.Refresh
        syntax_check.Text18 = "Non label tag set to 'Y'"
        syntax_check.Text18.Refresh

        program_control.Show 1
        
    End If
    
    'check Framecode value flag, should be 'Y' if tag contains 'label'
    char_type = 0
    If excel_tag_dat(i, frame_val) = "N" Then char_type = 1
    p = InStr(1, excel_tag_dat(i, tag_col), "label")
    p1 = InStr(1, excel_tag_dat(i, tag_col), "labeling")
    If p = p1 Then char_type = 0
    
    If excel_tag_dat(i, tag_col) = "_Sample_component.Isotopic_labeling" Then char_type = 0
    If excel_tag_dat(i, tag_col) = "_Sample_component_atom_isotope.Comp_isotope_label_code" Then char_type = 0
    If excel_tag_dat(i, tag_col) = "_Isotope_effect.Isotope_label_1_ID" Then char_type = 0
    If excel_tag_dat(i, tag_col) = "_Isotope_effect.Isotope_label_1_ID_chem_shift_val" Then char_type = 0
    If excel_tag_dat(i, tag_col) = "_Isotope_effect.Isotope_label_1_ID_chem_shift_val_err" Then char_type = 0
    If excel_tag_dat(i, tag_col) = "_Isotope_effect.Isotope_label_2_ID" Then char_type = 0
    If excel_tag_dat(i, tag_col) = "_Isotope_effect.Isotope_label_2_ID_chem_shift_val" Then char_type = 0
    If excel_tag_dat(i, tag_col) = "_Isotope_effect.Isotope_label_2_ID_chem_shift_val_err" Then char_type = 0
    If excel_tag_dat(i, tag_col) = "_Atom_site.PDBX_label_asym_ID" Then char_type = 0
    If excel_tag_dat(i, tag_col) = "_Atom_site.PDBX_label_seq_ID" Then char_type = 0
    If excel_tag_dat(i, tag_col) = "_Atom_site.PDBX_label_comp_ID" Then char_type = 0
    If excel_tag_dat(i, tag_col) = "_Atom_site.PDBX_label_atom_ID" Then char_type = 0
    If excel_tag_dat(i, tag_col) = "_Atom_site.PDBX_label_entity_ID" Then char_type = 0
    
    p1 = InStr(1, excel_tag_dat(i, tag_col), "Isotope_label_pattern")
    p2 = InStr(1, excel_tag_dat(i, tag_col), "Isotope_label_sample")
    
    If p1 > 0 Or p2 > 0 Then char_type = 0
    If char_type = 1 And p > 0 Then
        syntax_check.Text6 = Str(i + 4)
        syntax_check.Text6.Refresh
        syntax_check.Text7 = excel_tag_dat(i, sf_col)
        syntax_check.Text7.Refresh
        syntax_check.Text8 = excel_tag_dat(i, cat_col)
        syntax_check.Text8.Refresh
        syntax_check.Text9 = excel_tag_dat(i, tag_col)
        syntax_check.Text9.Refresh
        syntax_check.Text18 = "label tag set to 'N'"
        syntax_check.Text18.Refresh

        program_control.Show 1
        
    End If
    
    'check that row index fields have a data type of INTEGER
    char_type = 0
    If excel_tag_dat(i, row_id) = "Y" Then char_type = 1
    If char_type = 1 Then
    If excel_tag_dat(i, dat_type) <> "INTEGER" And excel_tag_dat(i, dat_type) <> "CHAR(12)" Then
        syntax_check.Text6 = Str(i + 4)
        syntax_check.Text6.Refresh
        syntax_check.Text7 = excel_tag_dat(i, sf_col)
        syntax_check.Text7.Refresh
        syntax_check.Text8 = excel_tag_dat(i, cat_col)
        syntax_check.Text8.Refresh
        syntax_check.Text9 = excel_tag_dat(i, tag_col)
        syntax_check.Text9.Refresh
        syntax_check.Text18 = "Data type not set to INTEGER"
        syntax_check.Text18.Refresh

        program_control.Show 1
        
    End If
    End If

    'check parent/child relationships (parent does exist and data types are the same)
    'Debug.Print
    p = InStr(1, excel_tag_dat(i, Desc), "Pointer to '_")       'foreign table or foreign column designation is wrong
    If p > 0 Then
        p = InStr(1, excel_tag_dat(i, Desc), "'")
        p1 = InStr(p + 1, excel_tag_dat(i, Desc), "'")
        ln = Len(excel_tag_dat(i, Desc))
        If p1 > 0 Then
            tag1 = Mid$(excel_tag_dat(i, Desc), p + 1, p1 - (p + 1))
            
            p2 = InStr(1, tag1, ".")
            ln = Len(tag1)
            table1 = Mid$(tag1, 2, p2 - 2)
            column1 = Mid$(tag1, p2 + 1, ln - p2)
            If excel_tag_dat(i, foreignTab) <> table1 Or excel_tag_dat(i, foreignCol) <> column1 Then
            If table1 <> "Entry" Then
                syntax_check.Text6 = Str(i + 4)
                syntax_check.Text6.Refresh
                syntax_check.Text7 = excel_tag_dat(i, sf_col)
                syntax_check.Text7.Refresh
                syntax_check.Text8 = excel_tag_dat(i, cat_col)
                syntax_check.Text8.Refresh
                syntax_check.Text9 = excel_tag_dat(i, foreignTab)
                syntax_check.Text9.Refresh
                syntax_check.Text18 = "Foreign table or foreign column wrong"
                syntax_check.Text18.Refresh

                program_control.Show 1
            End If
            End If
            tag_found = 0: dt_flg = 0: db_dt_flg = 0
            For j = 1 To e_tag_ct
                If tag1 = excel_tag_dat(j, tag_col) Then
                    If excel_tag_dat(j, src_key) = "Y" Then tag_found = 1
                    If excel_tag_dat(i, bmrb_data_type) = excel_tag_dat(j, bmrb_data_type) Then dt_flg = 1
                    If excel_tag_dat(i, dat_type) = excel_tag_dat(j, dat_type) Then db_dt_flg = 1
                    If p = InStr(1, tag1, "Asym_ID") Then
                        Debug.Print
                    End If
                End If
            Next j
            If tag_found = 0 Then
                syntax_check.Text6 = Str(i + 4)
                syntax_check.Text6.Refresh
                syntax_check.Text7 = excel_tag_dat(i, sf_col)
                syntax_check.Text7.Refresh
                syntax_check.Text8 = excel_tag_dat(i, cat_col)
                syntax_check.Text8.Refresh
                syntax_check.Text9 = excel_tag_dat(i, tag_col)
                syntax_check.Text9.Refresh
                syntax_check.Text18 = "Parent tag not found"
                syntax_check.Text18.Refresh

                program_control.Show 1
            End If
            If adit_file_source < 2 Or adit_file_source = 4 Then
            If dt_flg = 0 Then
                syntax_check.Text6 = Str(i + 4)
                syntax_check.Text6.Refresh
                syntax_check.Text7 = excel_tag_dat(i, sf_col)
                syntax_check.Text7.Refresh
                syntax_check.Text8 = excel_tag_dat(i, cat_col)
                syntax_check.Text8.Refresh
                syntax_check.Text9 = excel_tag_dat(i, tag_col)
                syntax_check.Text9.Refresh
                syntax_check.Text18 = "Parent tag BMRB data type mismatch"
                syntax_check.Text18.Refresh

                program_control.Show 1
            End If
            If db_dt_flg = 0 Then
                syntax_check.Text6 = Str(i + 4)
                syntax_check.Text6.Refresh
                syntax_check.Text7 = excel_tag_dat(i, sf_col)
                syntax_check.Text7.Refresh
                syntax_check.Text8 = excel_tag_dat(i, cat_col)
                syntax_check.Text8.Refresh
                syntax_check.Text9 = excel_tag_dat(i, tag_col)
                syntax_check.Text9.Refresh
                syntax_check.Text18 = "Parent tag DB data type mismatch"
                syntax_check.Text18.Refresh
                'Debug.Print
                
                program_control.Show 1
            End If
            End If
        End If
        If p1 < 1 Then
            syntax_check.Text6 = Str(i + 4)
            syntax_check.Text6.Refresh
            syntax_check.Text7 = excel_tag_dat(i, sf_col)
            syntax_check.Text7.Refresh
            syntax_check.Text8 = excel_tag_dat(i, cat_col)
            syntax_check.Text8.Refresh
            syntax_check.Text9 = excel_tag_dat(i, Desc)
            syntax_check.Text9.Refresh
            syntax_check.Text18 = "Missing single quote"
            syntax_check.Text18.Refresh

            program_control.Show 1
        End If
    End If

    'check for parents without children
    If excel_tag_dat(i, src_key) = "Y" Then
        child1 = 0
        For j = 1 To e_tag_ct
            If excel_tag_dat(j, foreignTab) > "" Or excel_tag_dat(j, foreignCol) > "" Then
                tag1 = "_" + excel_tag_dat(j, foreignTab) + "." + excel_tag_dat(j, foreignCol)
                If tag1 = excel_tag_dat(i, tag_col) Then
                    child1 = 1
                    Exit For
                End If
            End If
        Next j
        If child1 = 0 Then
            syntax_check.Text6 = Str(i + 4)
            syntax_check.Text6.Refresh
            syntax_check.Text7 = excel_tag_dat(i, sf_col)
            syntax_check.Text7.Refresh
            syntax_check.Text8 = excel_tag_dat(i, cat_col)
            syntax_check.Text8.Refresh
            syntax_check.Text9 = excel_tag_dat(i, tag_col)
            syntax_check.Text9.Refresh
            syntax_check.Text18 = "Parent without a child"
            syntax_check.Text18.Refresh

            program_control.Show 1
        End If
    End If
    
    
    'check for consistency between db data types and pdbx data types
        
        db_type = 1
        If excel_tag_dat(i, 28) = "INTEGER" And excel_tag_dat(i, 87) <> "int" Then db_type = 0
        If excel_tag_dat(i, 28) = "TEXT" And excel_tag_dat(i, 87) <> "text" Then db_type = 0
        If excel_tag_dat(i, 28) = "FLOAT" And excel_tag_dat(i, 87) <> "float" Then db_type = 0
        If excel_tag_dat(i, 28) = "CHAR(12)" And excel_tag_dat(i, 87) <> "code" Then
            If excel_tag_dat(i, 28) = "CHAR(12)" And excel_tag_dat(i, 87) <> "atcode" Then
                    db_type = 0
            End If
        End If
        If excel_tag_dat(i, 28) = "DATETIME year to day" And excel_tag_dat(i, 87) <> "date" Then
            If excel_tag_dat(i, 28) = "DATETIME year to day" And excel_tag_dat(i, 87) <> "yyyy-mm-dd" Then
                If excel_tag_dat(i, 28) = "DATETIME year to day" And excel_tag_dat(i, 87) <> "yyyy-mm-dd:hh:mm" Then
                    db_type = 0
                End If
            End If
        End If
        
        If excel_tag_dat(i, 28) = "VARCHAR(3)" And excel_tag_dat(i, 87) <> "yes_no" Then
            If excel_tag_dat(i, 28) = "VARCHAR(3)" And excel_tag_dat(i, 87) <> "uchar3" Then
                If excel_tag_dat(i, 28) = "VARCHAR(3)" And excel_tag_dat(i, 87) <> "code" Then
                    If excel_tag_dat(i, 28) = "VARCHAR(3)" And excel_tag_dat(i, 87) <> "ucode" Then
                        db_type = 0
                    End If
                End If
            End If
        End If
        
        If excel_tag_dat(i, 28) = "VARCHAR(15)" And excel_tag_dat(i, 87) <> "code" Then
            If excel_tag_dat(i, 28) = "VARCHAR(15)" And excel_tag_dat(i, 87) <> "line" Then
                If excel_tag_dat(i, 28) = "VARCHAR(15)" And excel_tag_dat(i, 87) <> "ucode" Then
                    If excel_tag_dat(i, 28) = "VARCHAR(15)" And excel_tag_dat(i, 87) <> "uchar3" Then
                        db_type = 0
                    End If
                End If
            End If
        End If
        
        'If excel_tag_dat(i, 28) = "VARCHAR(31)" And excel_tag_dat(i, 87) <> "phrase" Then
            If excel_tag_dat(i, 28) = "VARCHAR(31)" And excel_tag_dat(i, 87) <> "fax" Then
                If excel_tag_dat(i, 28) = "VARCHAR(31)" And excel_tag_dat(i, 87) <> "phone" Then
                    If excel_tag_dat(i, 28) = "VARCHAR(31)" And excel_tag_dat(i, 87) <> "code" Then
                        If excel_tag_dat(i, 28) = "VARCHAR(31)" And excel_tag_dat(i, 87) <> "line" Then
                            db_type = 0
                        End If
                    End If
                End If
            End If
        'End If
        
        If excel_tag_dat(i, 28) = "VARCHAR(127)" And excel_tag_dat(i, 87) <> "line" Then
            If excel_tag_dat(i, 28) = "VARCHAR(127)" And excel_tag_dat(i, 87) <> "email" Then
                If excel_tag_dat(i, 28) = "VARCHAR(127)" And excel_tag_dat(i, 87) <> "framecode" Then
                    If excel_tag_dat(i, 28) = "VARCHAR(127)" And excel_tag_dat(i, 87) <> "uline" Then
                        db_type = 0
                    End If
                End If
            End If
        End If
        
        p = InStr(1, excel_tag_dat(i, 9), "Sf_category")
        If p > 0 Then
            db_type = 1
            If excel_tag_dat(i, 28) = "VARCHAR(127)" And excel_tag_dat(i, 87) <> "code" Then
                db_type = 0
            End If
        End If
        
        p = InStr(1, excel_tag_dat(i, 9), ".Sf_framecode")
        If p > 0 Then
            db_type = 1
            If excel_tag_dat(i, 28) = "VARCHAR(127)" And excel_tag_dat(i, 87) <> "framecode" Then
                db_type = 0
            End If
        End If
        
        'If excel_tag_dat(i, 28) = "VARCHAR(255)" And excel_tag_dat(i, 87) <> "ucode" Then db_type = 0
        'If excel_tag_dat(i, 28) = "VARCHAR(255)" And excel_tag_dat(i, 87) <> "uline" Then db_type = 0
        'If excel_tag_dat(i, 28) = "VARCHAR(255)" And excel_tag_dat(i, 87) <> "name" Then db_type = 0
        'If excel_tag_dat(i, 28) = "VARCHAR(255)" And excel_tag_dat(i, 87) <> "idname" Then db_type = 0
        'If excel_tag_dat(i, 28) = "VARCHAR(255)" And excel_tag_dat(i, 87) <> "any" Then db_type = 0
        If excel_tag_dat(i, 28) = "CHAR(1)" And excel_tag_dat(i, 87) <> "uchar1" Then db_type = 0
        'If excel_tag_dat(i, 28) = "VARCHAR(255)" And excel_tag_dat(i, 87) <> "atcode" Then db_type = 0
        'If excel_tag_dat(i, 28) = "VARCHAR(255)" And excel_tag_dat(i, 87) <> "int-range" Then db_type = 0
        'If excel_tag_dat(i, 28) = "VARCHAR(255)" And excel_tag_dat(i, 87) <> "float-range" Then db_type = 0
        'If excel_tag_dat(i, 28) = "VARCHAR(255)" And excel_tag_dat(i, 87) <> "code30" Then db_type = 0
        'If excel_tag_dat(i, 28) = "VARCHAR(1024)" And excel_tag_dat(i, 87) <> "binary" Then db_type = 0
        'If excel_tag_dat(i, 28) = "VARCHAR(1024)" And excel_tag_dat(i, 87) <> "label" Then db_type = 0
        
        If db_type = 0 Then
            syntax_check.Text6 = Str(i + 4)
            syntax_check.Text6.Refresh
            syntax_check.Text7 = excel_tag_dat(i, sf_col)
            syntax_check.Text7.Refresh
            syntax_check.Text8 = excel_tag_dat(i, cat_col)
            syntax_check.Text8.Refresh
            syntax_check.Text9 = excel_tag_dat(i, tag_col)
            syntax_check.Text9.Refresh
            syntax_check.Text18 = "Data type mismatch db vs pdbx"
            syntax_check.Text18.Refresh

            program_control.Show 1
        End If
      
    'print a list of data types that have not been registered
        db_type = 0
        If excel_tag_dat(i, 28) = "INTEGER" Then db_type = 1
        If excel_tag_dat(i, 28) = "TEXT" Then db_type = 1
        If excel_tag_dat(i, 28) = "FLOAT" Then db_type = 1
        If excel_tag_dat(i, 28) = "DATETIME year to day" Then db_type = 1
        If excel_tag_dat(i, 28) = "CHAR(12)" Then db_type = 1
        If excel_tag_dat(i, 28) = "VARCHAR(3)" Then db_type = 1
        If excel_tag_dat(i, 28) = "VARCHAR(15)" Then db_type = 1
        If excel_tag_dat(i, 28) = "VARCHAR(31)" Then db_type = 1
        If excel_tag_dat(i, 28) = "VARCHAR(127)" Then db_type = 1
        If excel_tag_dat(i, 28) = "VARCHAR(255)" Then db_type = 1
        If excel_tag_dat(i, 28) = "VARCHAR(1024)" Then db_type = 1
        If excel_tag_dat(i, 28) = "yes_no" Then db_type = 1
    
        If db_type = 0 Then Debug.Print excel_tag_dat(i, 28), "Tag number = " + Str(i)
            
    
    'check for incorrect usage of '.' and '_'
    
    p = InStr(1, excel_tag_dat(i, tag_col), ".")
    If p < 1 Then p = 1 Else p = 0
    p1 = InStr(1, excel_tag_dat(i, tag_col), "._")
    p2 = InStr(1, excel_tag_dat(i, tag_col), "_.")
    p3 = InStr(1, excel_tag_dat(i, tag_col), "..")
    p4 = InStr(1, excel_tag_dat(i, tag_col), "__")
    
    If p + p1 + p2 + p3 + p4 > 1 Then
        bad_tag_ct = bad_tag_ct + 1
        syntax_check.Text6 = Str(i + 4)
        syntax_check.Text6.Refresh
        syntax_check.Text7 = excel_tag_dat(i, sf_col)
        syntax_check.Text7.Refresh
        syntax_check.Text9 = excel_tag_dat(i, tag_col)
        syntax_check.Text9.Refresh
        syntax_check.Text10 = Str(bad_tag_ct)
        syntax_check.Text10.Refresh
        If p = 1 Then syntax_check.Text18 = "Missing period (table separator)"
        If p1 + p2 + p3 + p4 > 0 Then syntax_check.Text18 = "Improper period or underscore"
        syntax_check.Text18.Refresh
                
        program_control.Show 1

    End If
            
    'check to see if first character after a '.' is uppercase
    
    p = InStr(1, excel_tag_dat(i, tag_col), ".")
    asc_val = Asc(Mid$(excel_tag_dat(i, tag_col), p + 1, 1))
    If asc_val > 90 Or asc_val < 65 Then
        If asc_val > 57 Or asc_val < 48 Then
            syntax_check.Text6 = Str(i + 4)
            syntax_check.Text6.Refresh
            syntax_check.Text7 = excel_tag_dat(i, sf_col)
            syntax_check.Text7.Refresh
            syntax_check.Text9 = excel_tag_dat(i, tag_col)
            syntax_check.Text9.Refresh
            syntax_check.Text10 = Str(bad_tag_ct)
            syntax_check.Text10.Refresh
            syntax_check.Text18 = "Uppercase error"
            syntax_check.Text18.Refresh
                
            program_control.Show 1
        End If
    End If
  
    'check length of table names, should be less than 31 characters
    If Len(excel_tag_dat(i, dbtab)) > 45 Then
        syntax_check.Text6 = Str(i + 4)
        syntax_check.Text6.Refresh
        syntax_check.Text7 = excel_tag_dat(i, sf_col)
        syntax_check.Text7.Refresh
        syntax_check.Text8 = excel_tag_dat(i, cat_col)
        syntax_check.Text8.Refresh
        syntax_check.Text9 = excel_tag_dat(i, tag_col)
        syntax_check.Text9.Refresh
        syntax_check.Text14 = Str(Len(excel_tag_dat(i, dbtab)))
        syntax_check.Text14.Refresh
        syntax_check.Text15 = Str(Len(excel_tag_dat(i, tag_col)))
        syntax_check.Text15.Refresh
        syntax_check.Text18 = "Category too long"
        syntax_check.Text18.Refresh

        program_control.Show 1
    End If
    
    'check length of column names, should be less than 31 characters
    If Len(excel_tag_dat(i, dbcol)) > 45 Then
        syntax_check.Text6 = Str(i + 4)
        syntax_check.Text6.Refresh
        syntax_check.Text7 = excel_tag_dat(i, sf_col)
        syntax_check.Text7.Refresh
        syntax_check.Text8 = excel_tag_dat(i, cat_col)
        syntax_check.Text8.Refresh
        syntax_check.Text9 = excel_tag_dat(i, tag_col)
        syntax_check.Text9.Refresh
        syntax_check.Text14 = Str(Len(excel_tag_dat(i, dbcol)))
        syntax_check.Text14.Refresh
        syntax_check.Text15 = Str(Len(excel_tag_dat(i, tag_col)))
        syntax_check.Text15.Refresh
        syntax_check.Text18 = "Column name too long"
        syntax_check.Text18.Refresh

        program_control.Show 1
    End If
    
    'Flag missing prompts
    If excel_tag_dat(i, usr_view) = "Y" Or excel_tag_dat(i, usr_view) = "N" Then
        If excel_tag_dat(i, prompt) = "?" Then
            syntax_check.Text6 = Str(i + 4)
            syntax_check.Text6.Refresh
            syntax_check.Text7 = excel_tag_dat(i, sf_col)
            syntax_check.Text7.Refresh
            syntax_check.Text8 = excel_tag_dat(i, cat_col)
            syntax_check.Text8.Refresh
            syntax_check.Text9 = excel_tag_dat(i, tag_col)
            syntax_check.Text9.Refresh
            syntax_check.Text18 = "Prompt missing"
            syntax_check.Text18.Refresh

            program_control.Show 1
        End If
    End If
    
    'flag null fields for all fields that are required for ADIT
    For j = 1 To null_ct
        If excel_tag_dat(i, null_list(j)) = "" Then
            syntax_check.Text6 = Str(i + 4)
            syntax_check.Text6.Refresh
            syntax_check.Text7 = excel_tag_dat(i, sf_col)
            syntax_check.Text7.Refresh
            syntax_check.Text8 = excel_tag_dat(i, cat_col)
            syntax_check.Text8.Refresh
            syntax_check.Text9 = excel_tag_dat(i, tag_col)
            syntax_check.Text9.Refresh
            syntax_check.Text11 = null_list(j)
            syntax_check.Text11.Refresh
            syntax_check.Text18 = "Null field"
            syntax_check.Text18.Refresh

            program_control.Show 1
        End If
    Next j
Next i

'identify duplicate tables
'For i = 1 To table_ct
'    For j = 1 To table_ct
'        If i <> j Then
'            If table_list(i) = table_list(j) Then
'                syntax_check.Text12 = table_list(i)
'                syntax_check.Text12.Refresh
'                syntax_check.Text13 = Str(i)
'                syntax_check.Text13.Refresh
'                syntax_check.Text18 = "Duplicate tables"
'                syntax_check.Text18.Refresh
                                                       
'                program_control.Show 1

'            End If
'        End If
'    Next j
'Next i

'identify duplicate tags
syntax_check.Text8 = ""
syntax_check.Text8.Refresh
For i = 1 To e_tag_ct
    syntax_check.Text5 = Str(i)
    syntax_check.Text5.Refresh
    For j = i + 1 To e_tag_ct
        If UCase(excel_tag_dat(i, tag_col)) = UCase(excel_tag_dat(j, tag_col)) Then
        'If excel_tag_dat(i, 2) = excel_tag_dat(j, 2) Then ' temporary condition to eliminate known duplicates
        If i <> j Then
            syntax_check.Text6 = Str(j)
            syntax_check.Text6.Refresh
            syntax_check.Text7 = excel_tag_dat(j, sf_col)
            syntax_check.Text7.Refresh
            syntax_check.Text8 = excel_tag_dat(j, cat_col)
            syntax_check.Text8.Refresh
            syntax_check.Text9 = excel_tag_dat(i, tag_col)
            syntax_check.Text9.Refresh
            syntax_check.Text18 = "Duplicate tags"
            syntax_check.Text18.Refresh
                                                       
            program_control.Show 1
        End If
        'End If
        End If
    Next j
Next i

'check links defined between tables for the relational data model (reference table and column keys)
'cat = "": tag_list_ct = 0: cat_last = ""
'ReDim tag_list(400)
'For i = 1 To e_tag_ct
'    If i = 1 Then cat_last = excel_tag_dat(i, ref_cat)
'    If excel_tag_dat(i, ref_cat) = cat_last Then    'find all tags belonging to one category
'        tag_list_ct = tag_list_ct + 1
'        tag_list(tag_list_ct) = i
'    End If
'    'Debug.Print
'    If excel_tag_dat(i, ref_cat) <> cat_last Then
'        ReDim mate(60, 6) As String
'        mate_ct = 0
'        For j = 1 To tag_list_ct    'for each tag in the tag category check if it is a foreign key and then if the primary key that it points to exists
'            If excel_tag_dat(tag_list(j), 89) <> "?" And excel_tag_dat(tag_list(j), 89) <> "" Then
'                p = InStr(1, excel_tag_dat(tag_list(j), 89), ";")
'                If p < 1 Then
'                    mate_ct = mate_ct + 1
'                    mate(mate_ct, 1) = excel_tag_dat(tag_list(j), 89)
'                    mate(mate_ct, 2) = excel_tag_dat(tag_list(j), 90)
'                    mate(mate_ct, 3) = excel_tag_dat(tag_list(j), 91)
'                    mate(mate_ct, 6) = Str(tag_list(j))
'                End If
'                If p > 0 Then
'                    mate_ct_start = mate_ct
'                    mate_ct = mate_ct + 1
'                    mate(mate_ct, 1) = Left$(excel_tag_dat(tag_list(j), 89), p - 1)
'                    p1 = 1
'                    While p1 > 0
'                        p1 = InStr(p + 1, excel_tag_dat(tag_list(j), 89), ";")
'                        If p1 > 0 Then
'                            mate_ct = mate_ct + 1
'                            mate(mate_ct, 1) = Mid$(excel_tag_dat(tag_list(j), 89), p + 1, (p1 - p) - 1)
'                            mate(mate_ct, 6) = Str(tag_list(j))
'                            p = p1
'                        End If
'                    Wend
'                    mate_ct = mate_ct + 1
'                    mate(mate_ct, 1) = Right$(excel_tag_dat(tag_list(j), 89), Len(excel_tag_dat(tag_list(j), 89)) - p)
'                    mate(mate_ct, 6) = Str(tag_list(j))
'                End If
'                p = InStr(1, excel_tag_dat(tag_list(j), 90), ";")
'                If p > 0 Then
'                    mate_ct = mate_ct_start
'                    mate_ct = mate_ct + 1
'                    mate(mate_ct, 2) = Left$(excel_tag_dat(tag_list(j), 90), p - 1)
'                    p1 = 1
'                    While p1 > 0
'                        p1 = InStr(p + 1, excel_tag_dat(tag_list(j), 90), ";")
'                        If p1 > 0 Then
'                            mate_ct = mate_ct + 1
'                            mate(mate_ct, 2) = Mid$(excel_tag_dat(tag_list(j), 90), p + 1, (p1 - p) - 1)
'                            p = p1
'                        End If
'                    Wend
'                    mate_ct = mate_ct + 1
'                    mate(mate_ct, 2) = Right$(excel_tag_dat(tag_list(j), 90), Len(excel_tag_dat(tag_list(j), 90)) - p)
'                End If
'                p = InStr(1, excel_tag_dat(tag_list(j), 91), ";")
'                If p > 0 Then
'                    mate_ct = mate_ct_start
'                    mate_ct = mate_ct + 1
'                    mate(mate_ct, 3) = Left$(excel_tag_dat(tag_list(j), 91), p - 1)
'                    p1 = 1
'                    While p1 > 0
'                        p1 = InStr(p + 1, excel_tag_dat(tag_list(j), 91), ";")
'                        If p1 > 0 Then
'                            mate_ct = mate_ct + 1
'                            mate(mate_ct, 3) = Mid$(excel_tag_dat(tag_list(j), 91), p + 1, (p1 - p) - 1)
'                            p = p1
'                        End If
'                    Wend
'                    mate_ct = mate_ct + 1
'                    mate(mate_ct, 3) = Right$(excel_tag_dat(tag_list(j), 91), Len(excel_tag_dat(tag_list(j), 91)) - p)
'                End If
'            End If
'        Next j
'        mates = 1
'        For j = 1 To mate_ct
'            For k = 1 To e_tag_ct
'                If excel_tag_dat(k, ref_cat) = mate(j, 2) Then
'                    If excel_tag_dat(k, ref_col) = mate(j, 3) Then
'                        mate(j, 4) = excel_tag_dat(k, ref_cat)
'                        mate(j, 5) = excel_tag_dat(k, ref_col)
'                        Exit For
'                    End If
'                End If
'            Next k
'        Next j
'
'        For j = 1 To mate_ct
'            For k = j + 1 To mate_ct
'                If mate(j, 1) = mate(k, 1) Then
'                    If mate(j, 2) <> mate(k, 2) Then
'                        syntax_check.Text6 = Str(Val(mate(j, 6)) + 4)
'                        syntax_check.Text6.Refresh
'                        syntax_check.Text7 = excel_tag_dat(Val(mate(j, 6)), sf_col)
'                        syntax_check.Text7.Refresh
'                        syntax_check.Text8 = excel_tag_dat(Val(mate(j, 6)), 90)
'                        syntax_check.Text8.Refresh
'                        syntax_check.Text9 = excel_tag_dat(Val(mate(j, 6)), 91)
'                        syntax_check.Text9.Refresh
'                        syntax_check.Text18 = "primary category mismatch"
'                        syntax_check.Text18.Refresh
'
'                        program_control.Show 1
'                    End If
'                End If
'            Next k
'
'            If mate(j, 4) = "" Or mate(j, 5) = "" Then
'                'If mate(j, 6) = "" Then mate(j, 6) = Str(i + 4)
'                    For k2 = 1 To mate_ct
'                        Debug.Print mate(k2, 1), mate(k2, 2), mate(k2, 3), mate(k2, 4), mate(k2, 5), mate(k2, 6)
'                    Next k2
'                syntax_check.Text6 = Str(Val(mate(j, 6)) + 4)
'                syntax_check.Text6.Refresh
'                syntax_check.Text7 = excel_tag_dat(Val(mate(j, 6)), sf_col)
'                syntax_check.Text7.Refresh
'                syntax_check.Text8 = excel_tag_dat(Val(mate(j, 6)), 90)
'                syntax_check.Text8.Refresh
'                syntax_check.Text9 = excel_tag_dat(Val(mate(j, 6)), 91)
'                syntax_check.Text9.Refresh
'                syntax_check.Text18 = "primary tag not found"
'                syntax_check.Text18.Refresh
'
'                program_control.Show 1
'            End If
'
'        Next j
'        cat_last = excel_tag_dat(i, ref_cat)
'        ReDim tag_list(400)
'        tag_list_ct = 1: tag_list(1) = i
'    End If
'    If i > 4 Then
'    If excel_tag_dat(i, 90) <> "?" And excel_tag_dat(i, 90) <> "" Then
'        If excel_tag_dat(i, 89) = "?" Or excel_tag_dat(i, 89) = "" Then
'            syntax_check.Text6 = Str(i + 4)
'            syntax_check.Text6.Refresh
'            syntax_check.Text7 = excel_tag_dat(i, sf_col)
'            syntax_check.Text7.Refresh
'            syntax_check.Text8 = excel_tag_dat(i, 90)
'            syntax_check.Text8.Refresh
'            syntax_check.Text9 = excel_tag_dat(i, 91)
'            syntax_check.Text9.Refresh
'            syntax_check.Text18 = "missing key group ID"
'            syntax_check.Text18.Refresh
'
'            program_control.Show 1
'        End If
'    End If
'    End If
'
'Next i
                        

'check for consistency between the super group table and category group table
clear_fields
For i = 1 To spg_row_ct
    spg = 0
    For j = 1 To grp_row_ct
        If spg_table(i, 1) = grp_table(j, 1) Then spg = 1
    Next j
    If spg = 0 Then
        syntax_check.Text6 = Str(i)
        syntax_check.Text6.Refresh
        syntax_check.Text19 = spg_table(i, 1)
        syntax_check.Text19.Refresh
        syntax_check.Text18 = "Super group ID from category group table"
        syntax_check.Text18.Refresh
                                                       
        program_control.Show 1
    End If
    spg = 0
    For j = 1 To e_tag_ct
        If spg_table(i, 1) = excel_tag_dat(j, adit_spg) Then spg = 1
    Next j
    If spg = 0 Then
        syntax_check.Text6 = Str(i)
        syntax_check.Text6.Refresh
        syntax_check.Text19 = spg_table(i, 1)
        syntax_check.Text19.Refresh
        syntax_check.Text21 = spg_table(i, 2)
        syntax_check.Text21.Refresh
        syntax_check.Text18 = "Super group ID missing from tag table"
        syntax_check.Text18.Refresh
                                                       
        'program_control.Show 1
    End If
Next i

'check consistency between category group table and supergroup table
'reverse of above check
clear_fields
grp_last = ""
For i = 1 To grp_row_ct
    If grp_table(i, 2) <> grp_last Then
        spg = 0
        For j = 1 To spg_row_ct
            If spg_table(j, 1) = grp_table(i, 1) Then spg = 1
        Next j
        If spg = 0 Then
            syntax_check.Text6 = Str(i)
            syntax_check.Text6.Refresh
            syntax_check.Text19 = spg_table(i, 1)
            syntax_check.Text19.Refresh
            syntax_check.Text21 = spg_table(i, 2)
            syntax_check.Text21.Refresh
            syntax_check.Text18 = "Invalid super group ID in category group"
            syntax_check.Text18.Refresh
                                                       
            program_control.Show 1
        End If
    End If
Next i

'check consistency between category group table and tag table (xlschem_ann)
clear_fields
For i = 1 To grp_row_ct
    grp = 0
    For j = 1 To e_tag_ct
        If grp_table(i, 3) = excel_tag_dat(j, 7) Then grp = 1
    Next j
    If grp = 0 Then
        syntax_check.Text6 = Str(i)
        syntax_check.Text6.Refresh
        syntax_check.Text20 = grp_table(i, 3)
        syntax_check.Text20.Refresh
        syntax_check.Text22 = grp_table(i, 5)
        syntax_check.Text22.Refresh
        syntax_check.Text18 = "Category group ID not found in tag table"
        syntax_check.Text18.Refresh
                                                       
        program_control.Show 1
    End If
Next i

'check consistency between category group table and tag table (xlschem_ann)
clear_fields
For i = 1 To e_tag_ct
    grp = 0
    For j = 1 To grp_row_ct
        If grp_table(j, 3) = excel_tag_dat(i, 7) Then grp = 1
    Next j
    If grp = 0 Then
        syntax_check.Text6 = Str(i)
        syntax_check.Text6.Refresh
        syntax_check.Text20 = excel_tag_dat(i, 7)
        syntax_check.Text20.Refresh
        syntax_check.Text22 = excel_tag_dat(i + 4, 2)
        syntax_check.Text22.Refresh
        syntax_check.Text18 = "Category group ID not found in category group table"
        syntax_check.Text18.Refresh
                                                       
        program_control.Show 1
    End If
Next i


'check consistency between category group table category names and tag table (xlschem_ann)
clear_fields
For i = 1 To grp_row_ct
    grp = 0
    For j = 1 To e_tag_ct
        If grp_table(i, 4) = excel_tag_dat(j, 2) Then grp = 1
    Next j
    If grp = 0 Then
        syntax_check.Text6 = Str(i)
        syntax_check.Text6.Refresh
        syntax_check.Text20 = grp_table(i, 3)
        syntax_check.Text20.Refresh
        syntax_check.Text22 = grp_table(i, 4)
        syntax_check.Text22.Refresh
        syntax_check.Text18 = "Sf category not found in tag table"
        syntax_check.Text18.Refresh
                                                       
        program_control.Show 1
    End If
Next i

End Sub

Sub set_tag_values(tag_ct, tag_char, new_tag_dat, excel_header_dat, e_col_num, spg_col_ct, spg_row_ct, spg_table, grp_col_ct, grp_row_ct, grp_table)
Debug.Print "Set tag values"

Dim i, j, j1 As Integer

For i = 1 To tag_ct
'Debug.Print i, new_tag_dat(i, 2)
'Debug.Print
    If tag_char(i, 0) <> "y" Then
        For j = 1 To e_col_num
            If excel_header_dat(1, j) = "Dictionary sequence" Then
                new_tag_dat(i, j) = Trim(Str(i * 10))
            End If
            If excel_header_dat(1, j) = "SFCategory" Then
                new_tag_dat(i, j) = tag_char(i, 3)
            End If
            If excel_header_dat(1, j) = "ADIT category view type" Then
                If tag_char(i, 4) = "0" Then
                    new_tag_dat(i, j) = "R"
                End If
                If tag_char(i, 4) = "1" Then
                    new_tag_dat(i, 4) = "T6"
                End If
            End If
            'If excel_header_dat(1, j) = "ADIT category view name" Then
            '    ln = Len(tag_char(i, 2))
            '    new_tag_dat(i, j) = ""
            '    For j1 = 1 To ln
            '        t = Mid$(tag_char(i, 2), j1, 1)
            '        If t = "_" Then t = " "
            '        new_tag_dat(i, j) = new_tag_dat(i, j) + t
            '    Next j1
            'End If

            If excel_header_dat(1, j) = "Loopflag" Then
                If tag_char(i, 4) = "0" Then
                    new_tag_dat(i, j) = "N"
                End If
                If tag_char(i, 4) = "1" Then
                    new_tag_dat(i, j) = "Y"
                End If
            End If
            
            If excel_header_dat(1, j) = "BMRB/CCPN status" Then
                If new_tag_dat(i, j) = "" Then new_tag_dat(i, j) = "open"
            End If
            
            If excel_header_dat(1, j) = "Data Type" Then
            '    If new_tag_dat(i, j) = "" Then
            '        If LCase(Right$(tag_char(i, 1), 7)) = "details" Then new_tag_dat(i, j) = "TEXT"
            '        If LCase(Right$(tag_char(i, 1), 7)) = "keyword" Then new_tag_dat(i, j) = "VARCHAR(127)"
            '        If LCase(Right$(tag_char(i, 1), 5)) = "label" Then new_tag_dat(i, j) = "VARCHAR(127)"
            '        If LCase(Right$(tag_char(i, 1), 16)) = "text_data_format" Then new_tag_dat(i, j) = "VARCHAR(31)"
            '        If LCase(Right$(tag_char(i, 1), 9)) = "text_data" Then new_tag_dat(i, j) = "TEXT"
            '        If LCase(Right$(tag_char(i, 1), 2)) = "id" Then new_tag_dat(i, j) = "INTEGER"
            '        If LCase(Right$(tag_char(i, 1), 9)) = "atom_name" Then new_tag_dat(i, j) = "VARCHAR(15)"
            '        If LCase(Right$(tag_char(i, 1), 9)) = "atom_type" Then new_tag_dat(i, j) = "VARCHAR(15)"
                    If LCase(tag_char(i, 1)) = "sf_id" Then new_tag_dat(i, j) = "INTEGER"
            '        If LCase(tag_char(i, 1)) = "sf_framecode" Then new_tag_dat(i, j) = "VARCHAR(127)"
            '        If LCase(tag_char(i, 1)) = "sf_category" Then new_tag_dat(i, j) = "VARCHAR(31)"
            '        If LCase(Right$(tag_char(i, 1), 13)) = "molecule_code" Then new_tag_dat(i, j) = "VARCHAR(15)"
            '        If LCase(Right$(tag_char(i, 1), 14)) = "chem_comp_code" Then new_tag_dat(i, j) = "VARCHAR(15)"
            '        If LCase(Right$(tag_char(i, 1), 4)) = "code" Then new_tag_dat(i, j) = "VARCHAR(15)"
            '        If LCase(Right$(tag_char(i, 1), 7)) = "seq_num" Then new_tag_dat(i, j) = "INTEGER"
            '        If LCase(Right$(tag_char(i, 1), 4)) = "date" Then new_tag_dat(i, j) = "DATETIME year to day"
            '        If LCase(Right$(tag_char(i, 1), 3)) = "val" Then new_tag_dat(i, j) = "FLOAT"
            '        If LCase(Right$(tag_char(i, 1), 3)) = "err" Then new_tag_dat(i, j) = "FLOAT"
            '        If LCase(Right$(tag_char(i, 1), 7)) = "val_err" Then new_tag_dat(i, j) = "FLOAT"
            '        If LCase(Right$(tag_char(i, 1), 9)) = "val_units" Then new_tag_dat(i, j) = "VARCHAR(31)"

            '        If LCase(Right$(tag_char(i, 1), 4)) = "name" Then
            '            If LCase(Right$(tag_char(i, 1), 9)) <> "atom_name" Then new_tag_dat(i, j) = "VARCHAR(127)"
            '        End If
            '    End If
            End If
            If excel_header_dat(1, j) = "Tag" Then
                new_tag_dat(i, j) = tag_char(i, 2) + "." + tag_char(i, 1)
            End If

'            If excel_header_dat(1, j) = "ManDBTableName" Then
'                new_tag_dat(i, j) = Right$(tag_char(i, 2), Len(tag_char(i, 2)) - 1)
'                ln = Len(new_tag_dat(i, j))
'                new_tag = ""
'                For j1 = 1 To ln
'                    t = Mid$(new_tag_dat(i, j), j1, 1)
'                    If t = "_" Then
'                        underscore = 1
'                    End If
'                    If t <> "_" Then
'                        If underscore = 1 Then
'                            t = UCase(t)
'                            underscore = 0
'                        End If
'                        new_tag = new_tag + t
'                    End If
'                Next j1
'                new_tag_dat(i, j) = new_tag
'            End If
            
'            If excel_header_dat(1, j) = "ManDBColumnName" Then
'                new_tag_dat(i, j) = tag_char(i, 1)
'                ln = Len(new_tag_dat(i, j))
'                new_tag = ""
'                For j1 = 1 To ln
'                    t = Mid$(new_tag_dat(i, j), j1, 1)
'                    If t = "_" Then
'                        underscore = 1
'                    End If
'                    If t <> "_" Then
'                        If underscore = 1 Then
'                            t = UCase(t)
'                            underscore = 0
'                        End If
'                        new_tag = new_tag + t
'                    End If
'                Next j1
'                new_tag_dat(i, j) = new_tag
    'special cases for tags that break relational database keyword requirements
    
'                If new_tag_dat(i, j) = "Order" Then new_tag_dat(i, j) = "BondOrder"
'                If new_tag_dat(i, j) = "Offset" Then new_tag_dat(i, j) = "SSOffset"
                'If i > 2900 Then
                'Debug.Print i, tag_char(i, 1), new_tag
                'Debug.Print
                'End If
                'after setting the value for ManDBColumnName then set these additional values
'                For i1 = 1 To e_col_num
'                    If excel_header_dat(1, i1) = "ManDBColumnName" Then manDBcolname = LCase(new_tag_dat(i, i1))
'                    If excel_header_dat(1, i1) = "SFCategory" Then SFCategory = LCase(new_tag_dat(i, i1))
'                    If excel_header_dat(1, i1) = "ManDBTableName" Then
'                        manDBtab = LCase(new_tag_dat(i, i1))
'                        If SFCategory <> SFCategory_last Then
'                            SFmanDBtab = new_tag_dat(i, i1)
'                            SFCategory_last = SFCategory
'                        End If
'                    End If
'                Next i1
'
'            End If
            
            'If excel_header_dat(1, j) = "Data Type" Then
            '    If new_tag_dat(i, j) = "CHAR(12)" Then new_tag_dat(i, 87) = "code"
            '    If new_tag_dat(i, j) = "DATETIME year to day" Then new_tag_dat(i, 87) = "yyyy-mm-dd"
            '    If new_tag_dat(i, j) = "FLOAT" Then new_tag_dat(i, 87) = "float"
            '    If new_tag_dat(i, j) = "INTEGER" Then new_tag_dat(i, 87) = "int"
            '    If new_tag_dat(i, j) = "TEXT" Then new_tag_dat(i, 87) = "text"
            '    If new_tag_dat(i, j) = "VARCHAR(127)" Then new_tag_dat(i, 87) = "line"
            '    If new_tag_dat(i, j) = "VARCHAR(15)" Then new_tag_dat(i, 87) = "code"
            '    If new_tag_dat(i, j) = "VARCHAR(2)" Then new_tag_dat(i, 87) = "uchar3"
            '    If new_tag_dat(i, j) = "VARCHAR(255)" Then new_tag_dat(i, 87) = "text"
            '    If new_tag_dat(i, j) = "VARCHAR(3)" Then new_tag_dat(i, 87) = "uchar3"
            '    If new_tag_dat(i, j) = "VARCHAR(31)" Then new_tag_dat(i, 87) = "name"
            'End If
            'If excel_header_dat(1, j) = "Data Type" Then
            '    If new_tag_dat(i, 87) = "int" Then new_tag_dat(i, j) = "INTEGER"
            'End If
            
            If excel_header_dat(1, j) = "Row Index Key" Then
                If LCase(new_tag_dat(i, 32)) = "ordinal" Then new_tag_dat(i, j) = "Y"
                If new_tag_dat(i, j) = "" Then new_tag_dat(i, j) = "N"
            End If

            'set saveframe ID flag for all tags that end with '.Sf_ID'
            If excel_header_dat(1, j) = "Saveframe ID tag" Then
'                If manDBtab <> SFmanDBtab Then
'                    If LCase(new_tag_dat(i, 32)) = "sfid" Then new_tag_dat(i, j) = "Y"
'                End If
                If new_tag_dat(i, j) = "" Then new_tag_dat(i, j) = "N"
            End If

            If excel_header_dat(1, j) = "ADIT item view name" Then
                ln = Len(tag_char(i, 1))
                new_tag_dat(i, j) = ""
                For j1 = 1 To ln
                    t = Mid$(tag_char(i, 1), j1, 1)
                    If t = "_" Then t = " "
                    new_tag_dat(i, j) = new_tag_dat(i, j) + t
                Next j1
            End If
            
            
            If excel_header_dat(1, j) = "Seq" Then
                If tag_char(i, 4) = "N" Then
                    seq_ct = seq_ct + 10
                    new_tag_dat(i, j) = Trim(Str(seq_ct))
                End If
                If tag_char(i, 4) = "Y" Then
                    If tag_char(i, 2) <> tag_char(i - 1, 2) Then seq_loop_ct = 0
                    seq_loop_ct = seq_loop_ct + 10
                    new_tag_dat(i, j) = Trim(Str(seq_loop_ct))
                End If
            End If
            
            If excel_header_dat(3, j) = "View" Then
                If j < 16 Then
                    If new_tag_dat(i, j) = "" Then new_tag_dat(i, j) = "H"
                End If
            End If
            
            If excel_header_dat(1, j) = "ADIT category group ID" Then
                For i1 = 1 To grp_row_ct
                    If grp_table(i1, 4) = new_tag_dat(i, 2) Then
                        new_tag_dat(i, j) = grp_table(i1, 3)    'group ID
                        'new_tag_dat(i, 3) = grp_table(i1, 6)    'mandatory group flag
                        new_tag_dat(i, 5) = grp_table(i1, 1)    'category super group ID
                        new_tag_dat(i, 6) = grp_table(i1, 2)    'category super group
                        'new_tag_dat(i, 8) = grp_table(i1, 14)   'category group view name
                        Exit For
                    End If
                Next i1
            End If
            
            If excel_header_dat(1, j) = "Table Primary Key" Then
            
                'If manDBcolname = "sfcategory" Then
                '    new_tag_dat(i, j) = "Y"
                'End If
                'If manDBcolname = "entryid" Then
                '    new_tag_dat(i, j) = "Y"
                'End If
                'If manDBcolname = "sfid" Then
                '    new_tag_dat(i, j) = "Y"
                'End If
                'If Right$(manDBcolname, 7) = "ordinal" Then
                '    new_tag_dat(i, j) = "Y"
                'End If
                'If Right$(manDBcolname, 10) = "chemcompid" Then
                '    If manDBtab = "chemcomp" Then new_tag_dat(i, j) = "Y"
                'End If
                
                'If manDBcol = "atom_name" Then
                '    new_tag_dat(i, j) = "Y"
                'End If
                'If Right$(manDBcolname, 6) = "atomid" Then
                '    If manDBtab = "atom" Then
                '        new_tag_dat(i, j) = "Y"
                '    End If
                'End If
                    
            End If
            
            If excel_header_dat(1, j) = "Item enumerated" Then
                If new_tag_dat(i, j) = "" Then new_tag_dat(i, j) = "N"
            End If
            
            If excel_header_dat(1, j) = "Adit form code" Then
                If new_tag_dat(i, j) = "" Or new_tag_dat(i, j) = "0" Then
                    If new_tag_dat(i, 16) = "N" Or new_tag_dat(i, 16) = "Y" Then
                        If new_tag_dat(i, 28) = "TEXT" Then new_tag_dat(i, j) = "2" Else new_tag_dat(i, j) = "1"
                    End If
                    If new_tag_dat(i, 49) = "N" Or new_tag_dat(i, 49) = "Y" Then
                        If new_tag_dat(i, 28) = "TEXT" Then new_tag_dat(i, j) = "2" Else new_tag_dat(i, j) = "1"
                    End If
                End If
            End If
            If excel_header_dat(1, j) = "Tag field" Then
                p = InStr(1, new_tag_dat(i, 9), ".")
                If p > 0 Then
                    new_tag_dat(i, j) = Right$(new_tag_dat(i, 9), Len(new_tag_dat(i, 9)) - p)
                End If
            End If
            If excel_header_dat(1, j) = "Tag category" Then
                p = InStr(1, new_tag_dat(i, 9), ".")
                If p > 0 Then
                    new_tag_dat(i, j) = Mid$(new_tag_dat(i, 9), 2, p - 2)
                End If
            End If
                
            
            'If excel_header_dat(1, j) = "SfLabelFlg" Then
            '    If new_tag_dat(i, 34) = "Y" And new_tag_dat(i, 35) = "Y" Then
            '        new_tag_dat(i, j) = "Y"
            '    Else: new_tag_dat(i, j) = "N"
            '    End If
            'End If
            
            'Fill in the foreign table and foreign column information
            'If excel_header_dat(1, j) = "Source Key" Then
            '    If new_tag_dat(i, j) = "Y" Then
            '        If LCase(new_tag_dat(i, j - 3)) = "id" Then
            '            new_tag_dat(i, j - 3) = new_tag_dat(i, j - 4) + new_tag_dat(i, j - 3)
            '        End If
                    
            '        For j1 = 1 To tag_ct
            '            If j1 <> i Then
            '                p = InStr(1, LCase(new_tag_dat(j1, j - 3)), LCase(new_tag_dat(i, j - 3)))
            '                If p > 0 Then
            '                    If p + Len(new_tag_dat(i, j - 3)) = Len(new_tag_dat(j1, j - 3)) + 1 Or p = 1 Then
            '                        new_tag_dat(j1, j + 3) = new_tag_dat(i, j - 4)
            '                        new_tag_dat(j1, j + 4) = new_tag_dat(i, j - 3)
            '                        new_tag_dat(j1, j - 7) = new_tag_dat(i, j - 7)
            '                        new_tag_dat(j1, j - 6) = new_tag_dat(i, j - 6)
            '                        If new_tag_dat(i, 34) = "Y" Then new_tag_dat(j1, j + 4) = "SfID"
            '                    End If
            '                End If
            '            End If
            '        Next j1
                    
            ' End If
            'End If
            
        Next j
        
    End If
Next i

'locate missing saveframes and set saveframe ids
'For j = 1 To grp_row_ct
'found = 0
'For i = 1 To tag_ct
'    found = 0
'    For j = 1 To grp_row_ct
'        If new_tag_dat(i, 2) = grp_table(j, 4) Then
'            new_tag_dat(i, 7) = grp_table(j, 3)
'            found = 1
'        End If
'    Next i
'    If found = 0 Then
        'Debug.Print
'    End If
'Next j


For i = 1 To tag_ct
    'Debug.Print i, new_tag_dat(i, 2)
    j = 35
    If excel_header_dat(1, j) = "Source Key" Then
        If new_tag_dat(i, j) = "Y" Then
            'test_string = new_tag_dat(i, j - 3)
            j2 = 1: special = ""
            If LCase(new_tag_dat(i, j - 3)) = "id" Then
                new_tag_dat(i, j - 3) = new_tag_dat(i, j - 4) + new_tag_dat(i, j - 3)
            End If
            If LCase(new_tag_dat(i, j - 3)) = "label" Then
                new_tag_dat(i, j - 3) = new_tag_dat(i, j - 4) + new_tag_dat(i, j - 3)
            End If
            'special case parent/child relationships
            'comp_ID
            special_flag = 0
            If LCase(new_tag_dat(i, j - 4)) = "chemcomp" Then
            If LCase(new_tag_dat(i, j - 3)) = "chemcompid" Then
                special_flag = 1
                special = "compid"
            End If
            End If
            
            
            sfid_flag = 0
            If LCase(new_tag_dat(i, j - 3)) = "sfid" Then
                sfid_flag = 1
                j2 = i + 1
            End If
            If LCase(new_tag_dat(i, j - 3)) = "sfcategory" Then
                sfid_flag = 1
                j2 = i + 1
            End If

            For j1 = j2 To tag_ct
                If sfid_flag = 1 And new_tag_dat(j1, 2) <> new_tag_dat(i, 2) Then Exit For
                If j1 <> i Then
                    If LCase(new_tag_dat(j1, j - 3)) = "id" Then
                        new_tag_dat(j1, j - 3) = new_tag_dat(j1, j - 4) + new_tag_dat(j1, j - 3)
                    End If
                    If LCase(new_tag_dat(j1, j - 3)) = "label" Then
                        new_tag_dat(j1, j - 3) = new_tag_dat(j1, j - 4) + new_tag_dat(j1, j - 3)
                    End If
                    If LCase(new_tag_dat(j1, j - 3)) = LCase(new_tag_dat(i, j - 3)) Then
                        If new_tag_dat(i, j) <> "Y" Then
                            'new_tag_dat(j1, j + 3) = new_tag_dat(i, j - 4) 'set foreign table
                            'new_tag_dat(j1, j + 4) = new_tag_dat(i, j - 3) 'set foreign column
                            new_tag_dat(j1, 62) = new_tag_dat(i, 9)
                            'new_tag_dat(j1, j + 2) = "1"                   'set foreign key group
                        End If
                    End If
                    p3 = 0
                    p1 = InStr(1, LCase(new_tag_dat(j1, j - 3)), "entryatomid")
                    If p1 > 0 And LCase(new_tag_dat(i, j - 3)) = "entryatomid" Then p3 = 1
                    If p1 = 0 Then p3 = 1
                    If p3 = 1 Then
                    p = InStr(1, LCase(new_tag_dat(j1, j - 3)), LCase(new_tag_dat(i, j - 3)))
                    If p > 0 Then
                        If p + Len(new_tag_dat(i, j - 3)) = Len(new_tag_dat(j1, j - 3)) + 1 Or p = 1 Then
                            If new_tag_dat(j1, j) <> "Y" Then
                                'new_tag_dat(j1, j + 3) = new_tag_dat(i, j - 4) 'set foreign table
                                'new_tag_dat(j1, j + 4) = new_tag_dat(i, j - 3) 'set foreign column
                                new_tag_dat(j1, 62) = new_tag_dat(i, 9)
                                'new_tag_dat(j1, j + 2) = "1"                   'set foreign key group
                            End If
                        End If
                    End If
                    End If
                    
                    'Special cases
                    If special_flag = 1 Then
                    If LCase(new_tag_dat(j1, j - 3)) = special Then
                        'new_tag_dat(j1, j + 3) = new_tag_dat(i, j - 4)
                        'new_tag_dat(j1, j + 4) = new_tag_dat(i, j - 3)
                        new_tag_dat(j1, 62) = new_tag_dat(i, 9)
                        'new_tag_dat(j1, j + 2) = "1"
                    End If
                    p = InStr(1, LCase(new_tag_dat(j1, j - 3)), special)
                    If p > 0 Then
                        If new_tag_dat(j1, j) <> "Y" Then
                            'new_tag_dat(j1, j + 3) = new_tag_dat(i, j - 4)
                            'new_tag_dat(j1, j + 4) = new_tag_dat(i, j - 3)
                            new_tag_dat(j1, 62) = new_tag_dat(i, 9)
                            'new_tag_dat(j1, j + 2) = "1"
                        End If
                    End If
                    End If
                    
                End If
            Next j1
        End If
    End If
    j = 34
    'If excel_header_dat(1, j) = "Saveframe ID tag" Then
    '    If new_tag_dat(i, j) = "Y" Then
    '        new_tag_dat(i, j - 2) = "SfID"
    '        'If new_tag_dat(i, j + 5) > "" Then
    '        '    new_tag_dat(i, j + 5) = "SfID"
    '        'End If
    '    End If
    'End If
    j = 64
    If excel_header_dat(1, j) = "Parent tag" Then
        p = InStr(1, new_tag_dat(i, 53), "Pointer to '_")       'foreign table or foreign column designation is wrong
        If p > 0 Then
            p = InStr(1, new_tag_dat(i, 53), "'")
            p1 = InStr(p + 1, new_tag_dat(i, 53), "'")
            ln = Len(new_tag_dat(i, 53))
            If p1 > 0 Then
                tag1 = Mid$(new_tag_dat(i, 53), p + 1, p1 - (p + 1))
                new_tag_dat(i, j) = tag1
            End If
        End If
    End If
Next i

End Sub

Sub write_new_excel(pathin, output_file, tag_ct, new_tag_dat, excel_header_dat, e_col_num, adit_file_source)

Debug.Print "Writing new Excel file"

Dim vernum, p As String
Dim i, j, deftl, vernum1 As Integer

Open pathin + output_file For Output As 1
Open pathin + "xlschem_ann_master.csv" For Output As 2

' add header
For i = 1 To 4
    For j = 1 To e_col_num
        If j < e_col_num Then
            If i = 4 And j = 4 Then
                ln = Len(excel_header_dat(i, j))
                vernum = ""
                While p <> "."
                    p = Mid$(excel_header_dat(i, j), ln, 1)
                    If p <> "." Then
                        vernum = p + vernum
                        ln = ln - 1
                    End If
                Wend
                vernum = Trim(Str(Val(vernum) + 1))
                If adit_file_source = 0 Then excel_header_dat(i, j) = Left$(excel_header_dat(i, j), ln) + vernum
                If adit_file_source = 1 Then excel_header_dat(i, j) = Left$(excel_header_dat(i, j), ln) + vernum
                'If adit_file_source = 1 Then excel_header_dat(i, j) = Left$(excel_header_dat(i, j), ln) + vernum
                If adit_file_source = 2 Then excel_header_dat(i, j) = "met" + Left$(excel_header_dat(i, j), ln) + vernum
                If adit_file_source = 3 Then excel_header_dat(i, j) = "sm" + Left$(excel_header_dat(i, j), ln) + vernum
                If adit_file_source = 4 Then excel_header_dat(i, j) = Left$(excel_header_dat(i, j), ln) + vernum
                vernum = excel_header_dat(i, j)
            End If
            
            Print #1, excel_header_dat(i, j) + ",";
            Print #2, excel_header_dat(i, j) + ",";
        End If
        If i = 1 And excel_header_dat(i, j) = "default value" Then deftl = j
        If j = e_col_num Then
            If excel_header_dat(i, j) = "" Then
                Print #1, "?"
                Print #2, "?"
            Else
                Print #1, excel_header_dat(i, j)
                Print #2, excel_header_dat(i, j)
            End If
        End If
    Next j
Next i


For i = 1 To tag_ct
    For j = 1 To e_col_num - 3
        If new_tag_dat(i, 9) = "_Entry.NMR_STAR_version" And j = 1 Then
            new_tag_dat(i, deftl) = vernum
        End If
        'If new_tag_dat(i, 92) = "?" Then new_tag_dat(i, 92) = new_tag_dat(i, 53)

        Print #1, new_tag_dat(i, j) + ",";
        If j = 1 Then Print #2, ",";
        If j > 1 Then Print #2, new_tag_dat(i, j) + ",";
    Next j
    For j = e_col_num - 2 To e_col_num
        If j < e_col_num Then
            If new_tag_dat(i, j) = "" Then new_tag_dat(i, j) = "?"
            If new_tag_dat(i, j) = "help_text" Then new_tag_dat(i, j) = "?"
            If new_tag_dat(i, j) = "example_text" Then new_tag_dat(i, j) = "?"
            Print #1, new_tag_dat(i, j) + ",";
            Print #2, new_tag_dat(i, j) + ",";
        End If
        If j = e_col_num Then
            If new_tag_dat(i, j) = "" Then new_tag_dat(i, j) = "?"
            If new_tag_dat(i, j) = "description_text" Then new_tag_dat(i, j) = "?"
            Print #1, new_tag_dat(i, j)
            Print #2, new_tag_dat(i, j)
        End If
    Next j
Next i

' add last line of Excel table
Print #1, "TBL_END" + ",";
Print #2, "TBL_END" + ",";
For i = 2 To e_col_num
    If i < e_col_num Then
        Print #1, ",";
        Print #2, ",";
    End If
    If i = e_col_num Then
        Print #1, "?";
        Print #2, "?"
    End If
Next i

Close 1, 2

End Sub
Sub write_ccpn_excel(pathin, output_file, tag_ct, new_tag_dat, excel_header_dat, e_col_num)

Debug.Print "Writing new ccpn file"

Dim i, j As Integer
Dim prtcol(50) As Integer

For j = 1 To e_col_num
    If excel_header_dat(1, j) = "Dictionary sequence" Then prtcol(1) = j
    If excel_header_dat(1, j) = "SFCategory" Then prtcol(2) = j
    If excel_header_dat(1, j) = "ADIT super category" Then prtcol(3) = j
    If excel_header_dat(1, j) = "Tag category" Then prtcol(4) = j
    If excel_header_dat(1, j) = "Tag field" Then prtcol(5) = j
    If excel_header_dat(1, j) = "Tag" Then prtcol(6) = j
    If excel_header_dat(1, j) = "Item enumerated" Then prtcol(7) = j
    If excel_header_dat(1, j) = "Item enumeration closed" Then prtcol(8) = j
    If excel_header_dat(1, j) = "Data Type" Then prtcol(9) = j
    If excel_header_dat(1, j) = "Nullable" Then prtcol(10) = j
    If excel_header_dat(1, j) = "Row Index Key" Then prtcol(11) = j
    If excel_header_dat(1, j) = "Saveframe ID tag" Then prtcol(12) = j
    If excel_header_dat(1, j) = "Source Key" Then prtcol(13) = j
    If excel_header_dat(1, j) = "Table Primary Key" Then prtcol(14) = j
    If excel_header_dat(1, j) = "Foreign Table" Then prtcol(15) = j
    If excel_header_dat(1, j) = "Foreign Column" Then prtcol(16) = j
    If excel_header_dat(1, j) = "Loopflag" Then prtcol(17) = j
    If excel_header_dat(1, j) = "Example" Then prtcol(18) = j
    If excel_header_dat(1, j) = "Description" Then prtcol(19) = j
    If excel_header_dat(1, j) = "SfNameFlg" Then prtcol(20) = j
    If excel_header_dat(1, j) = "Sf category flag" Then prtcol(21) = j
    If excel_header_dat(1, j) = "Sf pointer" Then prtcol(22) = j
    If excel_header_dat(1, j) = "Parent tag" Then prtcol(23) = j
    If excel_header_dat(1, j) = "public" Then prtcol(24) = j
    If excel_header_dat(1, j) = "Non-public" Then noprt = j
    'If excel_header_dat(1, j) = "sf category" Then prtcol(26) = j
Next j

prtcol_ct = 24
Open pathin + output_file For Output As 1


' add header
For i = 1 To 4
    For k = 1 To prtcol_ct
    For j = 1 To e_col_num
        If j = prtcol(k) Then
            If k < prtcol_ct Then Print #1, excel_header_dat(i, j) + ",";
            If k = prtcol_ct Then
                If excel_header_dat(i, j) = "" Then
                    Print #1, "?"
                Else
                    Print #1, excel_header_dat(i, j)
                End If
            End If
        End If
    Next j
    Next k
Next i


For i = 1 To tag_ct
    If new_tag_dat(i, noprt) <> "Y" Then
        For k = 1 To prtcol_ct
        For j = 1 To e_col_num
            If j = prtcol(k) Then
                If k < prtcol_ct Then Print #1, new_tag_dat(i, j) + ",";
                If k = prtcol_ct Then Print #1, new_tag_dat(i, j)
            End If
        Next j
        Next k
    End If
Next i

' add last line of Excel table
Print #1, "TBL_END" + ",";
For i = 2 To prtcol_ct
    If i < prtcol_ct Then Print #1, ",";
    If i = prtcol_ct Then Print #1, "?"
Next i

Close 1

End Sub

Sub write_new_adit_files(pathin, output_files, tag_ct, new_tag_dat, e_col_num, excel_header_dat, adit_file_source)

Debug.Print "Writing ADIT files"

Dim i, j, k As Integer
Dim view As String

Open pathin + output_files(3) For Output As 1

'Debug.Print
ReDim col_head(200, 1 To 2) As String
set_column_headings col_head, e_col_num, excel_header_dat

adit_item_col_ct = 56
For i = 1 To adit_item_col_ct
    If i < adit_item_col_ct Then Print #1, col_head(i, 2) + ",";
    If i = adit_item_col_ct Then Print #1, col_head(i, 2)
Next i
Print #1, "TBL_BEGIN" + ",";
For i = 2 To adit_item_col_ct
    If i < adit_item_col_ct Then Print #1, ",";
    If i = adit_item_col_ct Then Print #1, "?"
Next i
For i = 1 To tag_ct
    For k = 1 To adit_item_col_ct
        If col_head(k, 1) = "View" Then
            view = ""
            If adit_file_source = 0 Then view = new_tag_dat(i, 16) + new_tag_dat(i, 17) + new_tag_dat(i, 18) + new_tag_dat(i, 19) + new_tag_dat(i, 20)
            If adit_file_source = 1 Then view = new_tag_dat(i, 16) + new_tag_dat(i, 17) + new_tag_dat(i, 18) + new_tag_dat(i, 19) + new_tag_dat(i, 20)
            'If adit_file_source = 1 Then view = new_tag_dat(i, 16) + new_tag_dat(i, 17) + new_tag_dat(i, 18) + new_tag_dat(i, 19) + new_tag_dat(i, 20)
            If adit_file_source = 2 Then view = new_tag_dat(i, 17) + new_tag_dat(i, 16) + new_tag_dat(i, 18) + new_tag_dat(i, 19) + new_tag_dat(i, 20)
            If adit_file_source = 3 Then view = new_tag_dat(i, 16) + new_tag_dat(i, 17) + new_tag_dat(i, 18) + new_tag_dat(i, 19) + new_tag_dat(i, 20)
            If adit_file_source = 4 Then view = new_tag_dat(i, 16) + new_tag_dat(i, 17) + new_tag_dat(i, 18) + new_tag_dat(i, 19) + new_tag_dat(i, 20)
            
            Print #1, view + ",";
        End If
        If adit_file_source = 2 Then
            If col_head(k, 1) = "ADIT category view name" Then new_tag_dat(i, 8) = new_tag_dat(i, 97)
            If col_head(k, 1) = "Example" Then new_tag_dat(i, 51) = new_tag_dat(i, 98)
            If col_head(k, 1) = "Prompt" Then new_tag_dat(i, 52) = new_tag_dat(i, 99)
            If col_head(k, 1) = "Interface" Then new_tag_dat(i, 53) = new_tag_dat(i, 100)
            If col_head(k, 1) = "default value" Then new_tag_dat(i, 77) = new_tag_dat(i, 105)
        End If
        If adit_file_source = 3 Then
            If col_head(k, 1) = "ADIT category view name" Then new_tag_dat(i, 8) = new_tag_dat(i, 101)
            If col_head(k, 1) = "Example" Then new_tag_dat(i, 51) = new_tag_dat(i, 102)
            If col_head(k, 1) = "Prompt" Then new_tag_dat(i, 52) = new_tag_dat(i, 103)
            If col_head(k, 1) = "Interface" Then new_tag_dat(i, 53) = new_tag_dat(i, 104)
            If col_head(k, 1) = "default value" Then
                new_tag_dat(i, 77) = new_tag_dat(i, 106)
                'Debug.Print
            End If
        End If
        
        If col_head(k, 1) = "Validate" Then
            Validate = ""
            For j = 65 To 70
                Validate = Validate + new_tag_dat(i, j)
            Next j
            Print #1, Validate + ",";
        End If
        If col_head(k, 1) = "Overide" Then
            overide = ""
            For j = 71 To 76
                overide = overide + new_tag_dat(i, j)
            Next j
            Print #1, overide + ",";
        End If
        For j = 1 To e_col_num
            If excel_header_dat(1, j) = col_head(k, 1) Then
                print_val = new_tag_dat(i, j)
                If col_head(k, 2) = "example" Then print_val = "?"
                'If col_head(k, 2) = "prompt" Then print_val = "na"
                If col_head(k, 2) = "description" Then print_val = "?"
                If k < adit_item_col_ct Then Print #1, print_val + ","; 'Debug.Print j, print_val
                If k = adit_item_col_ct Then Print #1, print_val
                Exit For
            End If
        Next j
        'Debug.Print
    Next k
Next i
Print #1, "TBL_END" + ",";
For i = 2 To adit_item_col_ct
    If i < adit_item_col_ct Then Print #1, ",";
    If i = adit_item_col_ct Then Print #1, "?"
Next i
Close 1
End Sub
Sub set_column_headings(col_head, e_col_num, excel_header_dat)
col_head(1, 1) = excel_header_dat(1, 1): col_head(1, 2) = "dictionarySeq"
col_head(2, 1) = excel_header_dat(1, 2): col_head(2, 2) = "originalCategory"
col_head(3, 1) = excel_header_dat(1, 3): col_head(3, 2) = "aditCatManFlg"
col_head(4, 1) = excel_header_dat(1, 4): col_head(4, 2) = "aditCatViewType"
col_head(5, 1) = excel_header_dat(1, 5): col_head(5, 2) = "aditSuperCatID"
col_head(6, 1) = excel_header_dat(1, 6): col_head(6, 2) = "aditSuperCatName"
col_head(7, 1) = excel_header_dat(1, 7): col_head(7, 2) = "aditCatGrpID"
col_head(8, 1) = excel_header_dat(1, 8): col_head(8, 2) = "aditCatViewName"
col_head(9, 1) = excel_header_dat(1, 45): col_head(9, 2) = "aditInitialRows"
col_head(10, 1) = excel_header_dat(1, 9): col_head(10, 2) = "originalTag"
col_head(11, 1) = excel_header_dat(1, 15): col_head(11, 2) = "aditExists"
col_head(12, 1) = excel_header_dat(3, 16): col_head(12, 2) = "aditViewFlgs"
col_head(13, 1) = excel_header_dat(1, 21): col_head(13, 2) = "enumeratedFlg"
col_head(14, 1) = excel_header_dat(1, 22): col_head(14, 2) = "itemEnumClosedFlg"
col_head(15, 1) = excel_header_dat(1, 27): col_head(15, 2) = "aditItemViewName"
col_head(16, 1) = excel_header_dat(1, 78): col_head(16, 2) = "aditFormCode"
col_head(17, 1) = excel_header_dat(1, 28): col_head(17, 2) = "dbType"
col_head(18, 1) = excel_header_dat(1, 87): col_head(18, 2) = "bmrbType"
col_head(19, 1) = excel_header_dat(1, 29): col_head(19, 2) = "dbNullable"
col_head(20, 1) = excel_header_dat(1, 30): col_head(20, 2) = "internalFlg"
col_head(21, 1) = excel_header_dat(1, 33): col_head(21, 2) = "rowIndexFlg"
col_head(22, 1) = excel_header_dat(1, 81): col_head(22, 2) = "lclIDFlg"
col_head(23, 1) = excel_header_dat(1, 96): col_head(23, 2) = "lclSfIdFlg"
col_head(24, 1) = excel_header_dat(1, 34): col_head(24, 2) = "sfIdFlg"
col_head(25, 1) = excel_header_dat(1, 58): col_head(25, 2) = "sfNameFlg"
col_head(26, 1) = excel_header_dat(1, 59): col_head(26, 2) = "sfCategoryFlg"
col_head(27, 1) = excel_header_dat(1, 60): col_head(27, 2) = "sfPointerFlg"
col_head(28, 1) = excel_header_dat(1, 36): col_head(28, 2) = "primaryKey"
col_head(29, 1) = excel_header_dat(1, 37): col_head(29, 2) = "foreignKeyGroup"
col_head(30, 1) = excel_header_dat(1, 38): col_head(30, 2) = "foreignTable"
col_head(31, 1) = excel_header_dat(1, 39): col_head(31, 2) = "foreignColumn"
col_head(32, 1) = excel_header_dat(1, 40): col_head(32, 2) = "indexFlg"
col_head(33, 1) = excel_header_dat(1, 31): col_head(33, 2) = "dbTableName"
col_head(34, 1) = excel_header_dat(1, 32): col_head(34, 2) = "dbColumnName"
col_head(35, 1) = excel_header_dat(1, 79): col_head(35, 2) = "tagCategory"
col_head(36, 1) = excel_header_dat(1, 80): col_head(36, 2) = "tagField"
col_head(37, 1) = excel_header_dat(1, 43): col_head(37, 2) = "loopFlg"
col_head(38, 1) = excel_header_dat(1, 44): col_head(38, 2) = "seq"
col_head(39, 1) = excel_header_dat(1, 57): col_head(39, 2) = "dbFlg"
col_head(40, 1) = excel_header_dat(3, 65): col_head(40, 2) = "validateFlgs"
col_head(41, 1) = excel_header_dat(3, 71): col_head(41, 2) = "valOverideFlgs"
col_head(42, 1) = excel_header_dat(1, 77): col_head(42, 2) = "defaultValue"
col_head(43, 1) = excel_header_dat(1, 54): col_head(43, 2) = "bmrbPdbMatchId"
col_head(44, 1) = excel_header_dat(1, 55): col_head(44, 2) = "bmrbPdbTransFunc"
col_head(45, 1) = excel_header_dat(1, 93): col_head(45, 2) = "variableTypeMatch"
col_head(46, 1) = excel_header_dat(1, 94): col_head(46, 2) = "entryIdFlg"
col_head(47, 1) = excel_header_dat(1, 95): col_head(47, 2) = "OutputMapExistsFlg"
col_head(48, 1) = excel_header_dat(1, 50): col_head(48, 2) = "aditAutoInsert"
col_head(49, 1) = excel_header_dat(1, 82): col_head(49, 2) = "datumCountFlgs"
col_head(50, 1) = excel_header_dat(1, 85): col_head(50, 2) = "metaDataFlgs"
col_head(51, 1) = excel_header_dat(1, 86): col_head(51, 2) = "tagDeleteFlgs"
col_head(52, 1) = excel_header_dat(1, 89): col_head(52, 2) = "RefKeyGroup"
col_head(53, 1) = excel_header_dat(1, 90): col_head(53, 2) = "RefTable"
col_head(54, 1) = excel_header_dat(1, 91): col_head(54, 2) = "RefColumn"
col_head(55, 1) = excel_header_dat(1, 51): col_head(55, 2) = "example"
col_head(56, 1) = excel_header_dat(1, 52): col_head(56, 2) = "prompt"

'col_head(57, 1) = excel_header_dat(1, 53): col_head(57, 2) = "description"

End Sub
Sub set_query_interface_col_headings(col_head, e_col_num, excel_header_dat)
col_head(1, 1) = excel_header_dat(1, 1): col_head(1, 2) = "dictionarySeq"
col_head(2, 1) = excel_header_dat(1, 79): col_head(2, 2) = "tagCategory"
col_head(3, 1) = excel_header_dat(1, 80): col_head(3, 2) = "tagField"
col_head(4, 1) = excel_header_dat(1, 12): col_head(4, 2) = "tagInterfaceFlag"
col_head(5, 1) = excel_header_dat(1, 11): col_head(5, 2) = "tagPrompt"

End Sub

Sub set_val_column_headings(col_head, e_col_num, excel_header_dat)
col_head(1, 1) = excel_header_dat(1, 1): col_head(1, 2) = "dictionarySeq"
col_head(2, 1) = excel_header_dat(1, 2): col_head(2, 2) = "originalCategory"
col_head(3, 1) = excel_header_dat(1, 9): col_head(3, 2) = "originalTag"
col_head(4, 1) = excel_header_dat(1, 22): col_head(4, 2) = "itemEnumClosedFlg"
col_head(5, 1) = excel_header_dat(1, 28): col_head(5, 2) = "dbType"
col_head(6, 1) = excel_header_dat(1, 87): col_head(6, 2) = "bmrbType"
col_head(7, 1) = excel_header_dat(1, 29): col_head(7, 2) = "dbNullable"
col_head(8, 1) = excel_header_dat(1, 36): col_head(8, 2) = "primaryKey"
col_head(9, 1) = excel_header_dat(1, 79): col_head(9, 2) = "tagCategory"
col_head(10, 1) = excel_header_dat(1, 80): col_head(10, 2) = "tagField"
col_head(11, 1) = excel_header_dat(1, 43): col_head(11, 2) = "loopFlg"
col_head(12, 1) = excel_header_dat(1, 77): col_head(12, 2) = "defaultValue"
col_head(13, 1) = excel_header_dat(1, 50): col_head(13, 2) = "aditAutoInsert"
col_head(14, 1) = excel_header_dat(1, 82): col_head(14, 2) = "datumCountFlgs"
col_head(15, 1) = excel_header_dat(1, 89): col_head(15, 2) = "RefKeyGroup"
col_head(16, 1) = excel_header_dat(1, 90): col_head(16, 2) = "RefTable"
col_head(17, 1) = excel_header_dat(1, 91): col_head(17, 2) = "RefColumn"
col_head(18, 1) = excel_header_dat(1, 30): col_head(18, 2) = "NonPublic"
col_head(19, 1) = excel_header_dat(1, 86): col_head(19, 2) = "TagDelete"


End Sub
Sub write_query_interface_file(pathin, output_files, tag_ct, new_tag_dat, e_col_num, excel_header_dat, adit_file_source)
Debug.Print "Writing Query Interface files"

Dim i, j, k As Integer
Dim view As String

Open pathin + output_files(29) For Output As 1
'Debug.Print
ReDim col_head(200, 1 To 2) As String
set_query_interface_col_headings col_head, e_col_num, excel_header_dat

adit_item_col_ct = 5
For i = 1 To adit_item_col_ct
    If i < adit_item_col_ct Then Print #1, col_head(i, 2) + ",";
    If i = adit_item_col_ct Then Print #1, col_head(i, 2)
Next i
Print #1, "TBL_BEGIN" + ",";
For i = 2 To adit_item_col_ct
    If i < adit_item_col_ct Then Print #1, ",";
    If i = adit_item_col_ct Then Print #1, "?"
Next i
For i = 1 To tag_ct
    For k = 1 To adit_item_col_ct
        For j = 1 To e_col_num
            If excel_header_dat(1, j) = col_head(k, 1) Then
                print_val = new_tag_dat(i, j)
                If print_val = "" Then print_val = "?"
                If k < adit_item_col_ct Then Print #1, print_val + ","; 'Debug.Print j, print_val
                If k = adit_item_col_ct Then Print #1, print_val
                Exit For
            End If
        Next j
        'Debug.Print
    Next k
Next i
Print #1, "TBL_END" + ",";
For i = 2 To adit_item_col_ct
    If i < adit_item_col_ct Then Print #1, ",";
    If i = adit_item_col_ct Then Print #1, "?"
Next i
Close 1

End Sub
Sub write_validation_files(pathin, output_files, tag_ct, new_tag_dat, e_col_num, excel_header_dat, adit_file_source)

Debug.Print "Writing Validation files"

Dim i, j, k As Integer
Dim view As String

Open pathin + output_files(25) For Output As 1

'Debug.Print
ReDim col_head(200, 1 To 2) As String
set_val_column_headings col_head, e_col_num, excel_header_dat

adit_item_col_ct = 19
For i = 1 To adit_item_col_ct
    If i < adit_item_col_ct Then Print #1, col_head(i, 2) + ",";
    If i = adit_item_col_ct Then Print #1, col_head(i, 2)
Next i
Print #1, "TBL_BEGIN" + ",";
For i = 2 To adit_item_col_ct
    If i < adit_item_col_ct Then Print #1, ",";
    If i = adit_item_col_ct Then Print #1, "?"
Next i
For i = 1 To tag_ct
    For k = 1 To adit_item_col_ct
        If col_head(k, 1) = "View" Then
            view = ""
            If adit_file_source = 0 Then view = new_tag_dat(i, 16) + new_tag_dat(i, 17) + new_tag_dat(i, 18) + new_tag_dat(i, 19) + new_tag_dat(i, 20)
            If adit_file_source = 1 Then view = new_tag_dat(i, 16) + new_tag_dat(i, 17) + new_tag_dat(i, 18) + new_tag_dat(i, 19) + new_tag_dat(i, 20)
            'If adit_file_source = 1 Then view = new_tag_dat(i, 16) + new_tag_dat(i, 17) + new_tag_dat(i, 18) + new_tag_dat(i, 19) + new_tag_dat(i, 20)
            If adit_file_source = 2 Then view = new_tag_dat(i, 17) + new_tag_dat(i, 16) + new_tag_dat(i, 18) + new_tag_dat(i, 19) + new_tag_dat(i, 20)
            If adit_file_source = 3 Then view = new_tag_dat(i, 16) + new_tag_dat(i, 17) + new_tag_dat(i, 18) + new_tag_dat(i, 19) + new_tag_dat(i, 20)
            If adit_file_source = 4 Then view = new_tag_dat(i, 16) + new_tag_dat(i, 17) + new_tag_dat(i, 18) + new_tag_dat(i, 19) + new_tag_dat(i, 20)
            
            Print #1, view + ",";
        End If
        If adit_file_source = 2 Then
            If col_head(k, 1) = "ADIT category view name" Then new_tag_dat(i, 8) = new_tag_dat(i, 97)
            If col_head(k, 1) = "Example" Then new_tag_dat(i, 51) = new_tag_dat(i, 98)
            If col_head(k, 1) = "Prompt" Then new_tag_dat(i, 52) = new_tag_dat(i, 99)
            If col_head(k, 1) = "Interface" Then new_tag_dat(i, 53) = new_tag_dat(i, 100)
            If col_head(k, 1) = "default value" Then new_tag_dat(i, 77) = new_tag_dat(i, 105)
        End If
        If adit_file_source = 3 Then
            If col_head(k, 1) = "ADIT category view name" Then new_tag_dat(i, 8) = new_tag_dat(i, 101)
            If col_head(k, 1) = "Example" Then new_tag_dat(i, 51) = new_tag_dat(i, 102)
            If col_head(k, 1) = "Prompt" Then new_tag_dat(i, 52) = new_tag_dat(i, 103)
            If col_head(k, 1) = "Interface" Then new_tag_dat(i, 53) = new_tag_dat(i, 104)
            If col_head(k, 1) = "default value" Then
                new_tag_dat(i, 77) = new_tag_dat(i, 106)
                'Debug.Print
            End If
        End If
        
        If col_head(k, 1) = "Validate" Then
            Validate = ""
            For j = 65 To 70
                Validate = Validate + new_tag_dat(i, j)
            Next j
            Print #1, Validate + ",";
        End If
        If col_head(k, 1) = "Overide" Then
            overide = ""
            For j = 71 To 76
                overide = overide + new_tag_dat(i, j)
            Next j
            Print #1, overide + ",";
        End If
        For j = 1 To e_col_num
            If excel_header_dat(1, j) = col_head(k, 1) Then
                print_val = new_tag_dat(i, j)
                If col_head(k, 2) = "example" Then print_val = "?"
                If col_head(k, 2) = "prompt" Then print_val = "na"
                If col_head(k, 2) = "description" Then print_val = "?"
                If k < adit_item_col_ct Then Print #1, print_val + ","; 'Debug.Print j, print_val
                If k = adit_item_col_ct Then Print #1, print_val
                Exit For
            End If
        Next j
        'Debug.Print
    Next k
Next i
Print #1, "TBL_END" + ",";
For i = 2 To adit_item_col_ct
    If i < adit_item_col_ct Then Print #1, ",";
    If i = adit_item_col_ct Then Print #1, "?"
Next i
Close 1
End Sub


Sub write_enum_ties(pathin, output_files, tag_ct, new_tag_dat, e_col_num)
Debug.Print "Writing enumeration ties file"
Dim i, j As Integer

Open pathin + output_files(12) For Output As 12
Print #12, "tag category,Tag,Linked tag category,Linked tag"
Print #12, "TBL_BEGIN,,,?"

For i = 1 To tag_ct
    If new_tag_dat(i, 46) > "" Then
    If new_tag_dat(i, 9) <> "_Experiment.Name" Then
    For j = 1 To tag_ct
        If new_tag_dat(i, 46) <> "19" Then
            If new_tag_dat(i, 46) = new_tag_dat(j, 46) Then
                Print #12, new_tag_dat(i, 79) + ",";
                Print #12, new_tag_dat(i, 9) + ",";
                Print #12, new_tag_dat(j, 79) + ",";
                Print #12, new_tag_dat(j, 9)
            End If
        End If
        If new_tag_dat(j, 46) = "19" Then
            If new_tag_dat(i, 9) = "_Upload_data.Data_file_name" Then
                Print #12, new_tag_dat(i, 79) + ",";
                Print #12, new_tag_dat(i, 9) + ",";
                Print #12, new_tag_dat(j, 79) + ",";
                Print #12, new_tag_dat(j, 9)
            End If
            If new_tag_dat(i, 9) <> "_Upload_data.Data_file_name" Then
                If new_tag_dat(i, 46) = "19" Then
                    If new_tag_dat(j, 9) = "_Upload_data.Data_file_name" Then
                        Print #12, new_tag_dat(i, 79) + ",";
                        Print #12, new_tag_dat(i, 9) + ",";
                        Print #12, new_tag_dat(j, 79) + ",";
                        Print #12, new_tag_dat(j, 9)
                    End If
                End If
            End If
        End If
    Next j
    End If
    End If
Next i

Print #12, "TBL_END,,,?"
End Sub
Sub write_man_over(pathin, input_files, output_files, tag_ct, new_tag_dat, e_col_num, grp_col_ct, grp_row_ct, grp_table)
Debug.Print "Writing mandatory overide file"
Dim i, j, ct As Integer

Open pathin + output_files(13) For Output As 13
Print #13, "Order of operation,Sf category,Tag category,Tag,Override view value,Conditional tag category,Conditional tag,Override value"
Print #13, "TBL_BEGIN,,,,,,,?"
ct = 0
For i = 1 To tag_ct
    If new_tag_dat(i, 47) > "" And new_tag_dat(i, 48) = "" Then
    For j = 1 To tag_ct
        If i <> j Then
            If new_tag_dat(i, 47) = new_tag_dat(j, 47) Then
                If new_tag_dat(j, 48) > "" Then
                    ct = ct + 10
                    Print #13, Trim(Str(ct)) + ",";
                    Print #13, new_tag_dat(j, 2) + ",";
                    Print #13, new_tag_dat(j, 79) + ",";
                    Print #13, new_tag_dat(j, 9) + ",";
                    Print #13, new_tag_dat(j, 49) + ",";
                    Print #13, new_tag_dat(i, 79) + ",";
                    Print #13, new_tag_dat(i, 9) + ",";
                    Print #13, new_tag_dat(j, 48)
                End If
            End If

        End If
    Next j
    End If
Next i
Open pathin + input_files(7) For Input As 14
printflg = 0
ReDim test(8) As String
While EOF(14) <> -1
    For i = 1 To 8
        Input #14, test(i)
    Next i
    If test(1) = "TBL_END" Then printflg = 0
    If printflg = 1 Then
        For i = 1 To 7
            Print #13, test(i) + ",";
        Next i
        Print #13, test(8)
    End If
    If test(1) = "TBL_BEGIN" Then printflg = 1
Wend
Close 14
Print #13, "TBL_END,,,,,,,?"
Close 13

End Sub
Sub write_man_over_prod(pathin, input_files, output_files, tag_ct, new_tag_dat, e_col_num, grp_col_ct, grp_row_ct, grp_table)
Debug.Print "Writing mandatory overide file"
Dim i, j, ct As Integer

Open pathin + output_files(13) For Output As 13
Print #13, "Sf category,Tag category,Tag,Override view value,Conditional tag category,Conditional tag,Override value"
Print #13, "TBL_BEGIN,,,,,,?"
ct = 0
For i = 1 To tag_ct
    If new_tag_dat(i, 47) > "" And new_tag_dat(i, 48) = "" Then
    For j = 1 To tag_ct
        If i <> j Then
            If new_tag_dat(i, 47) = new_tag_dat(j, 47) Then
                If new_tag_dat(j, 48) > "" Then
                    Print #13, new_tag_dat(j, 2) + ",";
                    Print #13, new_tag_dat(j, 79) + ",";
                    Print #13, new_tag_dat(j, 9) + ",";
                    Print #13, new_tag_dat(j, 49) + ",";
                    Print #13, new_tag_dat(i, 79) + ",";
                    Print #13, new_tag_dat(i, 9) + ",";
                    Print #13, new_tag_dat(j, 48)
                End If
            End If

        End If
    Next j
    End If
Next i
Open pathin + input_files(7) For Input As 14
printflg = 0
ReDim test(8) As String
While EOF(14) <> -1
    For i = 1 To 7
        Input #14, test(i)
    Next i
    If test(1) = "TBL_END" Then printflg = 0
    If printflg = 1 Then
        For i = 1 To 6
            Print #13, test(i) + ",";
        Next i
        Print #13, test(7)
    End If
    If test(1) = "TBL_BEGIN" Then printflg = 1
Wend
Close 14
Print #13, "TBL_END,,,,,,?"
Close 13

End Sub
Sub write_tag_validation(pathin, output_files, tag_ct, new_tag_dat, e_col_num)
Debug.Print "Writing tag dependent validation file"

Dim i, j As Integer
Dim intag(9) As String

Open pathin + output_files(18) For Output As 13
Print #13, "Control Sf category,Control tag,Flag Value,Sf category,Tag,validateFlgs"
Print #13, "TBL_BEGIN,,,,,?"

For i = 1 To tag_ct
    If new_tag_dat(i, 47) > "" And new_tag_dat(i, 48) = "" Then
    For j = 1 To tag_ct
        If i <> j Then
            If new_tag_dat(j, 47) = new_tag_dat(i, 47) Then
                If new_tag_dat(j, 48) > " " Then
                    Print #13, new_tag_dat(i, 2) + ",";
                
                    Print #13, new_tag_dat(i, 9) + ",";
                    Print #13, new_tag_dat(j, 48) + ",";
                    Print #13, new_tag_dat(j, 2) + ",";
                
                    Print #13, new_tag_dat(j, 9) + ",";
                    Print #13, new_tag_dat(j, 71) + new_tag_dat(j, 72) + new_tag_dat(j, 73) + new_tag_dat(j, 74) + new_tag_dat(j, 75) + new_tag_dat(j, 76)
                End If
            End If
        End If
    Next j
    End If
Next i

'Open pathin + output_files(13) For Input As 12
'While EOF(12) <> -1
'    For i = 1 To 8
'        Input #12, intag(i)
'    Next i
'    For i = 1 To tag_ct
'        If new_tag_dat(i, 9) = intag(4) Then
'            For j = 1 To tag_ct
'                If new_tag_dat(j, 9) = intag(7) Then
'                    control_tag = j
'                    Exit For
'                End If
'            Next j
'            Print #13, new_tag_dat(control_tag, 2) + ",";
                
'            Print #13, new_tag_dat(control_tag, 9) + ",";
'            Print #13, intag(8) + ",";
'            Print #13, new_tag_dat(i, 2) + ",";
'
'            Print #13, new_tag_dat(i, 9) + ",";
'            Print #13, new_tag_dat(i, 71) + new_tag_dat(i, 72) + new_tag_dat(i, 73) + new_tag_dat(i, 74) + new_tag_dat(i, 75) + new_tag_dat(i, 76)
        
'            Exit For
            
'        End If
'    Next i
'Wend



Print #13, "TBL_END,,,,,?"
Close 13



End Sub
Sub write_anno_star(pathin, output_files, tag_ct, new_tag_dat, e_col_num)

Debug.Print "Writing annotated version of NMR-STAR"

Dim i, ln, ln1 As Integer
tag_prt_ct = 0
ReDim tag_prt(1 To 400, 1 To 3) As String

Open pathin + output_files(7) For Output As 2

Print #2, "data_nmrstarv3.0"
sf_flag = ""
lp_flag_last = "R"
cat_last = ""
sp_ct = 2
last_print = 1

For i = 1 To tag_ct
    print_tag = 1
    If new_tag_dat(i, 8) <> cat_last And cat_last <> "" Then
        If lp_flag_last = "T6" Then
            tag_prt_ct2 = tag_prt_ct
            print_tags tag_prt_ct, tag_prt
            Print #2,
            Print #2, "    ";
            For i1 = 1 To tag_prt_ct2
                Print #2, "@ ";
                If i1 Mod 5 = 0 Then Print #2, "  ";
            Next i1
            Print #2,
            Print #2,
            Print #2, "  stop_"
            'Print #2,
            sp_ct = 2
            last_print = 0
        End If
    End If
    If new_tag_dat(i, 8) <> cat_last And cat_last <> "" Then
        If new_tag_dat(i, 4) = "T6" Then
            print_tags tag_prt_ct, tag_prt
            Print #2,
            Print #2, "  loop_"
            sp_ct = 5
        End If
    End If

    If new_tag_dat(i, 2) <> sf_flag Then
        Print #2,
        If sf_flag <> "" Then
            print_tags tag_prt_ct, tag_prt
            If last_print = 1 Then Print #2,
            Print #2, "save_"
            Print #2,
        End If
        Print #2, "save_<" + new_tag_dat(i, 2) + ">"
        Print #2, "  " + new_tag_dat(i, 9) + Space(5) + new_tag_dat(i, 2)
        sf_flag = new_tag_dat(i, 2)
        print_tag = 0
    End If
    
    lp_flag_last = new_tag_dat(i, 4)
    cat_last = new_tag_dat(i, 8)

    If print_tag = 1 Then
        tag_prt_ct = tag_prt_ct + 1
        tag_prt(tag_prt_ct, 1) = Space(sp_ct) + new_tag_dat(i, 9)
        ln = Len(new_tag_dat(i, 11))
        If new_tag_dat(i, 12) > "" Then tag_prt(tag_prt_ct, 2) = "# " + new_tag_dat(i, 11) + ";" + Space(200 - ln) + new_tag_dat(i, 12)
        If new_tag_dat(i, 12) = "" Then tag_prt(tag_prt_ct, 2) = "# " + new_tag_dat(i, 11)
    End If
    
Next i

If tag_prt_ct > 0 Then print_tags tag_prt_ct, tag_prt
If lp_flag_last = "T6" Then
    Print #2,
    Print #2, "    ";
    For i1 = 1 To tag_prt_ct2
        Print #2, "@ ";
        If i1 Mod 5 = 0 Then Print #2, "  ";
    Next i1
    Print #2,
    Print #2,
    Print #2, "  stop_"
    Print #2,
End If
Print #2, "save_"
Print #2,
Close 2
End Sub
Sub write_template_star(pathin, output_files, tag_ct, new_tag_dat, e_col_num)

Debug.Print "Writing template version of NMR-STAR"

Dim i, ln, ln1 As Integer
tag_prt_ct = 0
ReDim tag_prt(1 To 100, 1 To 2) As String

Open pathin + output_files(14) For Output As 2

Print #2, "data_nmrstarv3.0"
sf_flag = ""
lp_flag_last = "R"
cat_last = ""
sp_ct = 2
last_print = 1

For i = 1 To tag_ct
    print_tag = 1
    If new_tag_dat(i, 8) <> cat_last And cat_last <> "" Then
        If lp_flag_last = "T6" Then
            tag_prt_ct2 = tag_prt_ct
            print_tags tag_prt_ct, tag_prt
            Print #2,
            Print #2, "    ";
            For i1 = 1 To tag_prt_ct2
                Print #2, "@ ";
                If i1 Mod 5 = 0 Then Print #2, "  ";
            Next i1
            Print #2,
            Print #2,
            Print #2, "  stop_"
            'Print #2,
            sp_ct = 2
            last_print = 0
        End If
    End If
    If new_tag_dat(i, 8) <> cat_last And cat_last <> "" Then
        If new_tag_dat(i, 4) = "T6" Then
            print_tags tag_prt_ct, tag_prt
            Print #2,
            Print #2, "  loop_"
            sp_ct = 5
        End If
    End If

    If new_tag_dat(i, 2) <> sf_flag Then
        Print #2,
        If sf_flag <> "" Then
            print_tags tag_prt_ct, tag_prt
            If last_print = 1 Then Print #2,
            Print #2, "save_"
            Print #2,
        End If
        Print #2, "save_<" + new_tag_dat(i, 2) + ">"
        Print #2, "  " + new_tag_dat(i, 9) + Space(5) + new_tag_dat(i, 2)
        sf_flag = new_tag_dat(i, 2)
        print_tag = 0
    End If
    
    lp_flag_last = new_tag_dat(i, 4)
    cat_last = new_tag_dat(i, 8)

    If print_tag = 1 Then
        tag_prt_ct = tag_prt_ct + 1
        tag_prt(tag_prt_ct, 1) = Space(sp_ct) + new_tag_dat(i, 9)
        ln = Len(new_tag_dat(i, 11))
        If new_tag_dat(i, 4) <> "T6" Then
            tag_prt(tag_prt_ct, 2) = "@"
        End If
    End If
    
Next i
            
tag_prt_ct2 = tag_prt_ct
If tag_prt_ct > 0 Then print_tags tag_prt_ct, tag_prt
If lp_flag_last = "T6" Then
    Print #2,
    Print #2, "    ";
    For i1 = 1 To tag_prt_ct2
        Print #2, "@ ";
        If i1 Mod 5 = 0 Then Print #2, "  ";
    Next i1
    Print #2,
    Print #2,
    Print #2, "  stop_"
    Print #2,
End If
Print #2, "save_"
Print #2,
Close 2
End Sub

Sub write_sg_dict(e_tag_ct, excel_tag_dat, tag_ct, tag_list)
Dim i, j, print_flag As Integer

For j = 1 To e_tag_ct
  If Val(excel_tag_dat(j, 7)) = 130 Then print_flag = 1
  If Val(excel_tag_dat(j, 7)) = 140 Then print_flag = 1
  If Val(excel_tag_dat(j, 7)) = 900 Then print_flag = 1
  If Val(excel_tag_dat(j, 7)) = 910 Then print_flag = 1
  
    If excel_tag_dat(j, 12) = "" Then print_flag = 0
    If print_flag = 1 Then
        p = InStr(1, excel_tag_dat(j, 9), ".")
        category1 = Mid$(excel_tag_dat(j, 9), 2, p - 2)
        tag1 = Right$(excel_tag_dat(j, 9), Len(excel_tag_dat(j, 9)) - p)

        If Trim(excel_tag_dat(j, 53)) <> "?" Then
            description1 = excel_tag_dat(j, 53)
        Else
            description1 = "?"
        End If
        If excel_tag_dat(j, 52) <> "?" Then
            prompt = excel_tag_dat(j, 52)
        Else
            prompt = "?"
        End If
        If excel_tag_dat(j, 51) <> "?" Then
            example = excel_tag_dat(j, 51)
        Else
            example = "?"
        End If
        
        Print #13, "save_" + excel_tag_dat(j, 9)
        Print #13,
        Print #13, "   _Tag                  '" + excel_tag_dat(j, 9) + "'"
        Print #13,
        Print #13, "   _Category              " + category1
        Print #13,
        Print #13, "   _Description"
        Print #13, ";"
        Print #13, description1
        Print #13, ";"
        Print #13,
        Print #13, "   _Prompt"
        Print #13, ";"
        Print #13, prompt
        Print #13, ";"
        Print #13,
        Print #13, "   loop_"
        Print #13, "     _Example"
        Print #13,
        Print #13, ";"
        Print #13, example
        Print #13, ";"
        Print #13,
        Print #13, "   stop_"
        Print #13,
        Print #13, "save_"
        Print #13,
    End If
    
Next j

End Sub

Sub write_fake_star(pathin, output_files, tag_ct, new_tag_dat, e_col_num)

Debug.Print "Writing fake populated version of NMR-STAR"

Dim i, ln, ln1, fake_ct As Integer
tag_prt_ct = 0

ReDim tag_prt(1 To 400, 1 To 3) As String
ReDim tag_prt2(1 To 400) As String
Dim tag_fake_val(1 To 30, 1 To 2)

fake_ct = 21
tag_fake_val(1, 1) = "INTEGER": tag_fake_val(1, 2) = "2"
tag_fake_val(2, 1) = "CHAR(3)": tag_fake_val(2, 2) = "yes"
tag_fake_val(3, 1) = "CHAR(12)": tag_fake_val(3, 2) = "1"
tag_fake_val(4, 1) = "VARCHAR(2)": tag_fake_val(4, 2) = Chr$(34) + "Tiny string value" + Chr$(34)
tag_fake_val(5, 1) = "VARCHAR(3)": tag_fake_val(5, 2) = "yes/no"
tag_fake_val(6, 1) = "VARCHAR(15)": tag_fake_val(6, 2) = Chr$(34) + "Short string value" + Chr$(34)
tag_fake_val(7, 1) = "VARCHAR(31)": tag_fake_val(7, 2) = Chr$(34) + "String value" + Chr$(34)
tag_fake_val(8, 1) = "VARCHAR(80)": tag_fake_val(8, 2) = Chr$(34) + "Brief phrase" + Chr$(34)
tag_fake_val(9, 1) = "VARCHAR(127)": tag_fake_val(9, 2) = Chr$(34) + "Long string value" + Chr$(34)
tag_fake_val(10, 1) = "VARCHAR(255)": tag_fake_val(10, 2) = Chr$(34) + "Very long phrase" + Chr$(34)
tag_fake_val(11, 1) = "TEXT": tag_fake_val(11, 2) = Chr$(34) + "Possible multiline text" + Chr$(34)
tag_fake_val(12, 1) = "FLOAT": tag_fake_val(12, 2) = "110.234"
tag_fake_val(13, 1) = "DATETIME year to day": tag_fake_val(13, 2) = "2002-04-15"
tag_fake_val(14, 1) = "TIME": tag_fake_val(14, 2) = "1:02:01"
tag_fake_val(15, 1) = "": tag_fake_val(15, 2) = Chr$(34) + "Missing data type" + Chr$(34)
tag_fake_val(16, 1) = "VARCHAR(1024)": tag_fake_val(16, 2) = Chr$(34) + "Very very long phrase" + Chr$(34)
tag_fake_val(17, 1) = "VARCHAR(511)": tag_fake_val(17, 2) = Chr$(34) + "Very very long phrase" + Chr$(34)
tag_fake_val(18, 1) = "CHAR(1)": tag_fake_val(18, 2) = "A"
tag_fake_val(19, 1) = "CHAR(31)": tag_fake_val(19, 2) = Chr$(34) + "string value" + Chr$(34)
tag_fake_val(20, 1) = "VARCHAR(4096)": tag_fake_val(20, 2) = Chr$(34) + "Very very long phrase" + Chr$(34)
'tag_fake_val(19, 1) = "INTEGER": tag_fake_val(19, 2) = "2"
'tag_fake_val(20, 1) = "INTEGER": tag_fake_val(20, 2) = "2"
tag_fake_val(21, 1) = "BOOLEAN": tag_fake_val(21, 2) = "yes"
tag_fake_val(22, 1) = "VARCHAR(1023)": tag_fake_val(22, 2) = Chr$(34) + "Very very long phrase" + Chr$(34)

Open pathin + output_files(15) For Output As 2

Print #2, "data_nmrstar_v3"
'Print #2,
'Print #2, "#  Copyright © The Board of Regents of the University of Wisconsin System."
sf_flag = ""
lp_flag_last = "R"
cat_last = ""
sp_ct = 2
last_print = 1
'Debug.Print
For i = 1 To tag_ct
    print_tag = 1
    If new_tag_dat(i, 79) <> cat_last And cat_last <> "" Then
        If lp_flag_last = "Y" Then
            tag_prt_ct2 = tag_prt_ct
            For j2 = 1 To tag_prt_ct
                tag_prt2(j2) = tag_prt(j2, 3)
            Next j2
            print_tags tag_prt_ct, tag_prt
            Print #2,
            Print #2, "    ";
            For i1 = 1 To tag_prt_ct2
                printflg = 0
                For j2 = 1 To fake_ct
                    If tag_prt2(i1) = tag_fake_val(j2, 1) Then
                        printflg = 1
                        Print #2, tag_fake_val(j2, 2) + "  ";
                    End If
                Next j2
                If printflg = 0 Then Print #2, Chr$(34) + "new data type" + Chr$(34) + "  ";
                If i1 Mod 5 = 0 Then Print #2, "  ";
            Next i1
            Print #2,
            Print #2,
            Print #2, "  stop_"
            'Print #2,
            sp_ct = 2
            last_print = 0
        End If
    End If
    If new_tag_dat(i, 79) <> cat_last And cat_last <> "" Then
        If new_tag_dat(i, 43) = "Y" Then
            print_tags tag_prt_ct, tag_prt
            Print #2,
            Print #2, "  loop_"
            sp_ct = 5
        End If
    End If

    If new_tag_dat(i, 2) <> sf_flag Then
        Print #2,
        If sf_flag <> "" Then
            print_tags tag_prt_ct, tag_prt
            If last_print = 1 Then Print #2,
            Print #2, "save_"
            Print #2,
        End If
        Print #2, "save_<" + new_tag_dat(i, 2) + ">"
        Print #2, "  " + new_tag_dat(i, 9) + Space(5) + new_tag_dat(i, 2)
        framecode = "<" + new_tag_dat(i, 2) + ">"
        sf_flag = new_tag_dat(i, 2)
        print_tag = 0
    End If
    
    If Right$(new_tag_dat(i, 9), 13) = ".Sf_framecode" Then
        Print #2, "  " + new_tag_dat(i, 9) + Space(5) + framecode
        print_tag = 0
    End If
    
    lp_flag_last = new_tag_dat(i, 43)
    cat_last = new_tag_dat(i, 79)

    If print_tag = 1 Then
        tag_prt_ct = tag_prt_ct + 1
        tag_prt(tag_prt_ct, 1) = Space(sp_ct) + new_tag_dat(i, 9)
        ln = Len(new_tag_dat(i, 11))
        If new_tag_dat(i, 43) = "Y" Then tag_prt(tag_prt_ct, 3) = new_tag_dat(i, 28)
        If new_tag_dat(i, 43) <> "Y" Then
            For j2 = 1 To fake_ct
                If new_tag_dat(i, 28) = tag_fake_val(j2, 1) Then
                    tag_prt(tag_prt_ct, 2) = tag_fake_val(j2, 2)
                End If
            Next j2
        End If
    End If
    
Next i

If tag_prt_ct > 0 Then
    tag_prt_ct2 = tag_prt_ct
    For j2 = 1 To tag_prt_ct
        tag_prt2(j2) = tag_prt(j2, 3)
    Next j2
    print_tags tag_prt_ct, tag_prt
End If
If lp_flag_last = "Y" Then
    Print #2,
    Print #2, "    ";
    For i1 = 1 To tag_prt_ct2
        For j2 = 1 To fake_ct
            If tag_prt2(i1) = tag_fake_val(j2, 1) Then
                Print #2, tag_fake_val(j2, 2) + "  ";
            End If
        Next j2
        If i1 Mod 5 = 0 Then Print #2, "  ";
    Next i1
    Print #2,
    Print #2,
    Print #2, "  stop_"
    Print #2,
End If
Print #2, "save_"
Print #2,
Close 2
End Sub
Sub print_tags(tag_prt_ct, tag_prt)

For i = 1 To tag_prt_ct
    t = Len(tag_prt(i, 1))
    If t > t_big Then t_big = t
Next i
t_big = t_big + 10
For i = 1 To tag_prt_ct
    t = Len(tag_prt(i, 1))
    Print #2, tag_prt(i, 1) + Space(t_big - t) + tag_prt(i, 2)
Next i
tag_prt_ct = 0
ReDim tag_prt(1 To 400, 1 To 3) As String

End Sub
Sub syntax_review(option_flag, data_tag_flag, loop_flag, tag_pair, tag_ct, bad_tag_ct, line_ct, sf_cat_name, table_ct, table_list, sf_ct, sf_list, loop_flag_release, tag_char)
       
Dim p, p1, p2, p3 As Integer
 
If option_flag = 1 Then
        If data_tag_flag = 1 Then
            
            'missing period in tag
            p = InStr(1, tag_pair(1), ".")
            If p < 1 Then
                bad_tag_ct = bad_tag_ct + 1
                syntax_check.Text6 = Str(line_ct)
                syntax_check.Text6.Refresh
                syntax_check.Text7 = sf_cat_name
                syntax_check.Text7.Refresh
                syntax_check.Text9 = tag_pair(1)
                syntax_check.Text9.Refresh
                syntax_check.Text10 = Str(bad_tag_ct)
                syntax_check.Text10.Refresh
                syntax_check.Text18 = "Missing period (table separator)"
                syntax_check.Text18.Refresh
                
                program_control.Show 1

            End If
            
            'check for incorrect usage of . and '_'
            p = InStr(1, tag_pair(1), "._")
            p1 = InStr(1, tag_pair(1), "_.")
            p2 = InStr(1, tag_pair(1), "..")
            p3 = InStr(1, tag_pair(1), "__")
            If p + p1 + p2 + p3 > 0 Then
                syntax_check.Text6 = Str(line_ct)
                syntax_check.Text6.Refresh
                syntax_check.Text7 = sf_cat_name
                syntax_check.Text7.Refresh
                syntax_check.Text9 = tag_pair(1)
                syntax_check.Text9.Refresh
                syntax_check.Text10 = Str(bad_tag_ct)
                syntax_check.Text10.Refresh
                syntax_check.Text18 = "Improper period or underscore"
                syntax_check.Text18.Refresh
                
                program_control.Show 1
            End If
            
            'check to see if first character after a . is uppercase
            p = InStr(1, tag_pair(1), ".")
            asc_val = Asc(Mid$(tag_pair(1), p + 1, 1))
            If asc_val > 90 Or asc_val < 65 Then
            If asc_val > 57 Or asc_val < 48 Then
                syntax_check.Text6 = Str(line_ct)
                syntax_check.Text6.Refresh
                syntax_check.Text7 = sf_cat_name
                syntax_check.Text7.Refresh
                syntax_check.Text9 = tag_pair(1)
                syntax_check.Text9.Refresh
                syntax_check.Text10 = Str(bad_tag_ct)
                syntax_check.Text10.Refresh
                syntax_check.Text18 = "Uppercase error"
                syntax_check.Text18.Refresh
                
                program_control.Show 1
            End If
            End If
           
            'check length of category and item
            syntax_check.Text5 = Str(tag_ct)
            syntax_check.Text5.Refresh
            If p > 0 Then
                category1 = Left$(tag_pair(1), p - 1)
                tag1 = Right$(tag_pair(1), Len(tag_pair(1)) - p)
                syntax_check.Text6 = Str(line_ct)
                syntax_check.Text6.Refresh
                syntax_check.Text7 = sf_cat_name
                syntax_check.Text7.Refresh
                syntax_check.Text8 = category1
                syntax_check.Text8.Refresh
                syntax_check.Text9 = tag1
                syntax_check.Text9.Refresh
                syntax_check.Text14 = Str(Len(category1))
                syntax_check.Text14.Refresh
                syntax_check.Text15 = Str(Len(tag1))
                syntax_check.Text15.Refresh
                If category1 <> category_last2 Then
                    cat_ct = cat_ct + 1
                    syntax_check.Text4 = Str(cat_ct)
                    syntax_check.Text4.Refresh
                    category_last2 = category1
                End If
                If Len(category1) > 30 Then
                    j1 = 0
                    For ji = 1 To Len(category1)
                        If Mid$(category1, ji, 1) = "_" Then j1 = j1 + 1
                    Next ji
                    If Len(category1) - j1 > 30 Then
                        Print #10, sf_cat_name
                        Print #10, category1
                        Print #10,
                        syntax_check.Text18 = "Category too long"
                        syntax_check.Text18.Refresh

                        program_control.Show 1
                    End If
                End If
                If Len(tag1) > 30 Then
                    j1 = 0
                    For ji = 1 To Len(tag1)
                        If Mid$(tag1, ji, 1) = "_" Then j1 = j1 + 1
                    Next ji
                    If Len(tag1) - j1 > 30 Then
                        Print #11, sf_cat_name
                        Print #11, category1
                        Print #11, Len(tag1), tag1
                        Print #11,
                        syntax_check.Text18 = "Tag too long"
                        syntax_check.Text18.Refresh
                                       
                        program_control.Show 1
                    End If
                End If
            End If
            

            'parse and print saveframe category
            p = InStr(1, UCase(tag_pair(1)), ".SF_CATEGORY")
            If p > 0 Then
                sf_ct = sf_ct + 1
                sf_list(sf_ct) = sf_cat_name
                syntax_check.Text3 = Str(sf_ct)
                syntax_check.Text3.Refresh
                p = InStr(1, tag_pair(1), ".")
                If p > 0 Then
                    table_ct = table_ct + 1
                    table_list(table_ct) = Mid$(tag_pair(1), 2, p - 2)
                    syntax_check.Text8 = table_list(table_ct)
                    syntax_check.Text8.Refresh
                    'Debug.Print
                End If
                syntax_check.Text4 = Str(table_ct)
                syntax_check.Text4.Refresh
                syntax_check.Text7 = sf_cat_name
                syntax_check.Text7.Refresh

                syntax_check.Text8 = table_list(table_ct)
                syntax_check.Text8.Refresh
            End If
            'syntax_check.Refresh
        End If
        
        'print out category
        If loop_flag = 0 Then loop_flag_release = 0
        If loop_flag = 1 And data_tag_flag = 1 Then
            If loop_flag = 1 And loop_flag_release = 0 Then
                table_ct = table_ct + 1
                syntax_check.Text4 = Str(table_ct)
                syntax_check.Text4.Refresh
                p = InStr(1, tag_pair(1), ".")
                If p > 0 Then
                    table_list(table_ct) = Mid$(tag_pair(1), 2, p - 2)
                    syntax_check.Text8 = table_list(table_ct)
                    syntax_check.Text8.Refresh
                End If
                loop_flag_release = 1
            End If
        End If
End If
If option_flag = 2 Then
For i = 1 To table_ct
    For j = 1 To table_ct
    
        'check for duplicate table names
        If table_list(i) = table_list(j) Then
        If LCase(table_list(i)) <> "author_assigned_db" Then
        If LCase(table_list(i)) <> "biological_function" Then
        If LCase(table_list(i)) <> "citation" Then
        If LCase(table_list(i)) <> "common_name" Then
        If LCase(table_list(i)) <> "constraint_list" Then
        If LCase(table_list(i)) <> "db" Then
        If LCase(table_list(i)) <> "experiment" Then
        If LCase(table_list(i)) <> "keyword" Then
        If LCase(table_list(i)) <> "observed_conformer" Then
        If LCase(table_list(i)) <> "selection_method" Then
        If LCase(table_list(i)) <> "software" Then
            If i <> j Then
                syntax_check.Text12 = table_list(i)
                syntax_check.Text12.Refresh
                syntax_check.Text13 = Str(i)
                syntax_check.Text13.Refresh
                syntax_check.Text18 = "Duplicate tables"
                syntax_check.Text18.Refresh
                                                       
                program_control.Show 1

            End If
        End If
        End If
        End If
        End If
        End If
        End If
        End If
        End If
        End If
        End If
        End If
        End If
    Next j
Next i
End If
If option_flag = 3 Then
syntax_check.Text8 = ""
syntax_check.Text8.Refresh
For i = 1 To tag_ct
    syntax_check.Text5 = Str(i)
    syntax_check.Text5.Refresh
    
    'check for duplicate tags
    For j = 1 To tag_ct
        If tag_char(i, 2) + "." + tag_char(i, 1) = tag_char(j, 2) + "." + tag_char(j, 1) Then
        If tag_char(i, 3) = tag_char(j, 3) Then
        If i <> j Then
                syntax_check.Text7 = tag_char(i, 3)
                syntax_check.Text7.Refresh
                syntax_check.Text8 = tag_char(i, 2)
                syntax_check.Text8 = Refresh
                syntax_check.Text9 = tag_char(i, 2) + "." + tag_char(i, 1)
                syntax_check.Text9.Refresh
                syntax_check.Text18 = "Duplicate tags"
                syntax_check.Text18.Refresh
                                                       
                program_control.Show 1
        End If
        End If
        End If
    Next j
Next i
End If
End Sub

Sub interface_load(data_tag_flag, tag_pair, tag_ct, tag_list, sf_loc, tag_char, tag_desc)

If data_tag_flag = 1 Then
If tag_pair(1) = "_Tag" Then
    sf_loc = 1
    tag_ct = tag_ct + 1
    tag_list(tag_ct) = tag_pair(2)
    'Debug.Print tag_ct, tag_list(tag_ct)
    'Debug.Print
End If

If tag_pair(1) = "_Description" And sf_loc = 1 Then
    If tag_pair(2) = "@" Then tag_char(tag_ct, 1) = 0
    If tag_pair(2) <> "@" Then tag_char(tag_ct, 1) = 1
End If

If tag_pair(1) = "_Example" Then sf_loc = 0
End If
End Sub

Sub tag_char_load(data_tag_flag, loop_flag, tag_pair, tag_ct, tag_char, sf_cat_name)

If data_tag_flag = 1 Then
    tag_ct = tag_ct + 1
    p = InStr(1, tag_pair(1), ".")
    category1 = Left$(tag_pair(1), p - 1)
    tag1 = Right$(tag_pair(1), Len(tag_pair(1)) - p)
    tag_char(tag_ct, 1) = tag1
    tag_char(tag_ct, 2) = category1
    tag_char(tag_ct, 3) = sf_cat_name
    tag_char(tag_ct, 4) = Trim(Str(loop_flag))
End If

End Sub
Sub update_interface(e_tag_ct, excel_tag_dat, tag_ct, tag_list, adit_file_source)
Dim i, i4, j, ln, print_flag As Integer
Dim t4 As String

Print #13, "data_adit_interface_dictionary"
Print #13,
For j = 1 To e_tag_ct
    print_flag = 1
    If excel_tag_dat(j, 16) = "H" Then print_flag = 1
    If excel_tag_dat(j, 49) = "N" Then print_flag = 1
    If excel_tag_dat(j, 49) = "Y" Then print_flag = 1
    If print_flag = 1 Then
        For i = 1 To tag_ct
            If excel_tag_dat(j, 9) = tag_list(i) Then   'the tag_list array is not populated so all tags are printed out to the file
                print_flag = 0
                Exit For
            End If
        Next i
    End If
    
    If print_flag = 1 Then
        p = InStr(1, excel_tag_dat(j, 9), ".")
        category1 = Chr$(39) + Mid$(excel_tag_dat(j, 9), 2, p - 2) + Chr$(39)
        tag1 = Right$(excel_tag_dat(j, 9), Len(excel_tag_dat(j, 9)) - p)

        If Trim(excel_tag_dat(j, 53)) <> "?" Then
            ln = Len(excel_tag_dat(j, 53))
            description1 = ""
            For i4 = 1 To ln
                t4 = Mid$(excel_tag_dat(j, 53), i4, 1)
                If t4 = "$" Then t4 = ","
                description1 = description1 + t4
            Next i4
        Else
            description1 = "?"
        End If
        If excel_tag_dat(j, 52) <> "n/a" Then
            If adit_file_source = 0 Then tag_num = 52
            If adit_file_source = 1 Then tag_num = 52
            'If adit_file_source = 2 Then tag_num = 99
            'If adit_file_source = 3 Then tag_num = 103
            If adit_file_source = 4 Then tag_num = 52
            ln = Len(excel_tag_dat(j, tag_num))
            prompt = ""
            For i4 = 1 To ln
                t4 = Mid$(excel_tag_dat(j, tag_num), i4, 1)
                If t4 = "$" Then t4 = ","
                prompt = prompt + t4
            Next i4
        Else
            prompt = excel_tag_dat(j, 27)
        End If
        If excel_tag_dat(j, 51) <> "?" Then
            ln = Len(excel_tag_dat(j, 51))
            example = ""
            For i4 = 1 To ln
                t4 = Mid$(excel_tag_dat(j, 51), i4, 1)
                If t4 = "$" Then t4 = ","
                example = example + t4
            Next i4
        Else
            example = "?"
        End If
        
        Print #13, "save_" + excel_tag_dat(j, 9)
        Print #13,
        Print #13, "   _Tag                  '" + excel_tag_dat(j, 9) + "'"
        Print #13,
        Print #13, "   _Category              " + category1
        Print #13,
        Print #13, "   _Description"
        Print #13, ";"
        Print #13, description1
        Print #13, ";"
        Print #13,
        Print #13, "   _Adit_item_view_name"
        Print #13, ";"
        Print #13, prompt
        Print #13, ";"
        Print #13,
        Print #13, "   loop_"
        Print #13, "     _Example"
        Print #13,
        Print #13, ";"
        Print #13, example
        Print #13, ";"
        Print #13,
        Print #13, "   stop_"
        Print #13,
        Print #13, "save_"
        Print #13,
    End If
Next j

End Sub
Sub write_nmrstar_dict(e_tag_ct, excel_tag_dat, tag_ct, tag_list)
Dim i, j, print_flag As Integer

Print #13, "data_NMR-STAR_dictionary_v3.0"
Print #13,

For j = 1 To e_tag_ct
    print_flag = 1
    If print_flag = 1 Then
        For i = 1 To tag_ct
            If excel_tag_dat(j, 9) = tag_list(i) Then
                print_flag = 0
                Exit For
            End If
        Next i
    End If
    
    If print_flag = 1 Then
    
            'set tag information values
        p = InStr(1, excel_tag_dat(j, 9), ".")
        category1 = Mid$(excel_tag_dat(j, 9), 2, p - 2)
        tag1 = Right$(excel_tag_dat(j, 9), Len(excel_tag_dat(j, 9)) - p)

        Select Case excel_tag_dat(j, 28)
            Case "INTEGER"
                data_type = "int"
            Case "TEXT"
                data_type = "text"
            Case "FLOAT"
                data_type = "float"
            Case "DATETIME year to day"
                data_type = "date"
            Case "CHAR(12)"
                data_type = "code"
            Case "VARCHAR(3)"
                data_type = "code"
            Case "VARCHAR(15)"
                data_type = "line"
            Case "VARCHAR(31)"
                data_type = "line"
            Case "VARCHAR(127)"
                data_type = "line"
        End Select
        
        man1 = "no"
        Select Case Trim(excel_tag_dat(j, 10))
            Case "M"
                man1 = "yes"
            Case "MC"
                man1 = "yes"
            Case "S"
                man1 = "yes"
            Case "SC"
                man1 = "yes"
        End Select
        If Trim(excel_tag_dat(j, 62)) <> "?" Then
            parent1 = excel_tag_dat(j, 53)
        Else
            parent1 = "."
        End If
        If Trim(excel_tag_dat(j, 53)) <> "?" Then
            description1 = excel_tag_dat(j, 53)
        Else
            description1 = "?"
        End If
        If excel_tag_dat(j, 52) <> "?" Then
            prompt = excel_tag_dat(j, 52)
        Else
            prompt = "?"
        End If
        If excel_tag_dat(j, 51) <> "?" Then
            example = excel_tag_dat(j, 51)
        Else
            example = "?"
        End If
        
        'print category information
        If excel_tag_dat(j, 8) <> cat_last Then
            cat_last = excel_tag_dat(j, 8)
            
            Print #13, "save_" + UCase(category1)
            Print #13, "   _category.description"
            Print #13, ";"
            Print #13, "text description"
            Print #13, ";"
            Print #13, "   _category.id                   " + category1
            Print #13, "   _category.mandatory_code        ?"
            Print #13,
            Print #13, "   loop_"
            Print #13, "     _category_key.name"
            Print #13, "      ?"
            Print #13,
            Print #13, "   loop_"
            Print #13, "     _category_group.id"
            Print #13, "      ?"
            Print #13,
            Print #13, "   loop_"
            Print #13, "     _category_examples.detail"
            Print #13, "     _category_examples.case"
            Print #13, ":"
            Print #13, "   ?"
            Print #13, ";"
            Print #13, ";"
            Print #13, "   ?"
            Print #13, ";"
            Print #13,
            Print #13, "save_"
            Print #13,
            
        End If
        
        'print tag information
        Print #13, "save_" + LCase(excel_tag_dat(j, 9))
        Print #13,
        Print #13, "   _item_description.description"
        Print #13, ";"
        Print #13, description1
        Print #13, ";"
        Print #13, "   _item.name                  '" + LCase(excel_tag_dat(j, 9)) + "'"
        Print #13, "   _item.category_id              " + LCase(category1)
        Print #13, "   _item.mandatory_code           " + man1
        Print #13, "   _item_type.code                " + data_type
        Print #13, "   _item_linked.child_name        ."
        If parent1 <> "." Then Print #13, "   _item_linked.parent_name       " + Chr$(34) + parent1 + Chr$(34)
        If parent1 = "." Then Print #13, "   _item_linked.parent_name        ."
        Print #13,
        Print #13, "   loop_"
        Print #13, "     _item_examples.case"
        Print #13,
        Print #13, ";"
        Print #13, example
        Print #13, ";"
        Print #13,
        Print #13, "   stop_"
        Print #13,
        Print #13, "   loop_"
        Print #13, "     _item_aliases.alias_name"
        Print #13, "     _item_aliases.dictionary"
        Print #13, "     _item_aliases.version"
        Print #13,
        Print #13, "     ?  ?  ?"
        Print #13,
        Print #13, "   stop_"

'        Print #13, "   _Adit_item_view_name"
'        Print #13, ";"
'        Print #13, prompt
'        Print #13, ";"
'        Print #13,
'        Print #13, "   _Adit_help"
'        Print #13, ";"
'        Print #13, help
'        print #13, ";"
'        print #13,
        Print #13, "save_"
        Print #13,
    End If
Next j

End Sub
Sub write_nmrstar_dict6(e_tag_ct, excel_tag_dat, tag_ct, tag_list, e_col_num, pathin, output_files, lcase_flag, full_dict_flag)

Dim i, j, k, print_flag, link_ct, cat_ct As Integer
Dim enum_values(5000, 5), cat_desc(4000, 3) As String
Dim cif_link(5000, 3) As String
Dim t, ttt, test_str As String

Open pathin + output_files(22) For Input As 5
While cat_desc(cat_ct, 1) <> "Table_end"
    cat_ct = cat_ct + 1
    test_str = ""
    For i = 1 To 3
        Input #5, cat_desc(cat_ct, i)
    Next i
    'Debug.Print cat_ct, cat_desc(cat_ct, 1), cat_desc(cat_ct, 2), cat_desc(cat_ct, 3)
    'Debug.Print
    ln = Len(cat_desc(cat_ct, 3))
    For k = 1 To ln
        t = Mid$(cat_desc(cat_ct, 3), k, 1)
        If t = "$" Then t = ","
        test_str = test_str + t
    Next k
    cat_desc(cat_ct, 3) = test_str
Wend
Close 5

Open pathin + output_files(21) For Input As 5
While EOF(5) <> -1
    enum_ct = enum_ct + 1
    For i = 1 To 5
        Input #5, enum_values(enum_ct, i)
    Next i
    ln = Len(enum_values(enum_ct, 2))
    enum_values(enum_ct, 2) = Mid$(enum_values(enum_ct, 2), 2, ln - 2)
    ln = Len(enum_values(enum_ct, 3))
    enum_values(enum_ct, 3) = Mid$(enum_values(enum_ct, 3), 2, ln - 2)
    ln = Len(enum_values(enum_ct, 3))
    For i23 = 1 To ln
        If Mid$(enum_values(enum_ct, 3), i23, 1) = "$" Then
            enum_values(enum_ct, 3) = Left$(enum_values(enum_ct, 3), i23 - 1) + "," + Right$(enum_values(enum_ct, 3), ln - i23)
            'Debug.Print enum_values(enum_ct, 3)
            'Debug.Print
        End If
    Next i23
    'Debug.Print enum_ct, enum_values(enum_ct, 1), enum_values(enum_ct, 2), enum_values(enum_ct, 3), enum_values(enum_ct, 4), enum_values(enum_ct, 5)
    'Debug.Print


    If enum_values(enum_ct, 3) = "enumeration_text" Then enum_values(enum_ct, 3) = "?"
    ln = Len(enum_values(enum_ct, 2))
    flag1 = 0: flag2 = 0: flag3 = 0
    For i = 1 To ln
        t = Mid$(enum_values(enum_ct, 2), i, 1)
        If t = " " Then flag1 = 1
        If t = Chr$(9) Then flag1 = 1
        If t = Chr$(91) Then flag1 = 1          ' chr [
        If t = Chr$(93) Then flag1 = 1          ' chr ]
        If t = Chr$(39) Then flag2 = 1
        If t = Chr$(96) Then flag2 = 1
        If t = Chr$(34) Then flag3 = 1
        If t = Chr$(36) Then
            enum_values(enum_ct, 2) = Left$(enum_values(enum_ct, 2), i - 1) + "," + Right$(enum_values(enum_ct, 2), ln - i)
        End If
    Next i
    If Left$(enum_values(enum_ct, 2), 1) = Chr$(34) And Right$(enum_values(enum_ct, 2), 1) = Chr$(34) Then flag3 = 4
    If Left$(enum_values(enum_ct, 2), 1) = Chr$(39) And Right$(enum_values(enum_ct, 2), 1) = Chr$(39) Then flag3 = 4
    If Left$(enum_values(enum_ct, 2), 1) = Chr$(96) And Right$(enum_values(enum_ct, 2), 1) = Chr$(96) Then flag3 = 4
    If Left$(enum_values(enum_ct, 2), 1) = Chr$(39) And Right$(enum_values(enum_ct, 2), 1) = Chr$(96) Then flag3 = 4
    If Left$(enum_values(enum_ct, 2), 1) = Chr$(96) And Right$(enum_values(enum_ct, 2), 1) = Chr$(39) Then flag3 = 4
    
    If flag1 = 1 Then
        If flag2 = 0 And flag3 = 0 Then
            enum_values(enum_ct, 2) = Chr$(39) + enum_values(enum_ct, 2) + Chr$(39)
        End If
        If flag2 = 1 And flag3 = 0 Then
            enum_values(enum_ct, 2) = Chr$(34) + enum_values(enum_ct, 2) + Chr$(34)
        End If
        If flag2 = 0 And flag3 = 1 Then
            enum_values(enum_ct, 2) = Chr$(39) + enum_values(enum_ct, 2) + Chr$(39)
        End If
        If flag2 = 1 And flag3 = 1 Then
            enum_values(enum_ct, 2) = Chr$(10) + Chr$(59) + enum_values(enum_ct, 2) + Chr$(10) + Chr$(59) + Chr$(10)
        End If
    End If
    ln = Len(enum_values(enum_ct, 3))
    flag1 = 0: flag2 = 0: flag3 = 0
    For i = 1 To ln
        t = Mid$(enum_values(enum_ct, 3), i, 1)
        If t = " " Then flag1 = 1
        If t = Chr$(9) Then flag1 = 1
        If t = Chr$(39) Then flag2 = 1
        If t = Chr$(96) Then flag2 = 1
        If t = Chr$(34) Then flag3 = 1
    Next i
    If Left$(enum_values(enum_ct, 3), 1) = Chr$(34) And Right$(enum_values(enum_ct, 3), 1) = Chr$(34) Then flag3 = 4
    If Left$(enum_values(enum_ct, 3), 1) = Chr$(39) And Right$(enum_values(enum_ct, 3), 1) = Chr$(39) Then flag3 = 4
    If Left$(enum_values(enum_ct, 3), 1) = Chr$(96) And Right$(enum_values(enum_ct, 3), 1) = Chr$(96) Then flag3 = 4
    If Left$(enum_values(enum_ct, 3), 1) = Chr$(39) And Right$(enum_values(enum_ct, 3), 1) = Chr$(96) Then flag3 = 4
    If Left$(enum_values(enum_ct, 3), 1) = Chr$(96) And Right$(enum_values(enum_ct, 3), 1) = Chr$(39) Then flag3 = 4
    If flag1 = 1 Then
        If flag2 = 0 And flag3 = 0 Then
            enum_values(enum_ct, 3) = Chr$(39) + enum_values(enum_ct, 3) + Chr$(39)
        End If
        If flag2 = 1 And flag3 = 0 Then
            enum_values(enum_ct, 3) = Chr$(34) + enum_values(enum_ct, 3) + Chr$(34)
        End If
        If flag2 = 0 And flag3 = 1 Then
            enum_values(enum_ct, 3) = Chr$(39) + enum_values(enum_ct, 3) + Chr$(39)
        End If
        If flag2 = 1 And flag3 = 1 Then
            enum_values(enum_ct, 3) = Chr$(10) + Chr$(59) + enum_values(enum_ct, 3) + Chr$(10) + Chr$(59) + Chr$(10)
        End If
    End If

    'Debug.Print enum_values(enum_ct, 1), enum_values(enum_ct, 2), enum_values(enum_ct, 3)
    'Debug.Print
Wend
Close 5

Open pathin + "adit_input\nmr_cif_match.csv" For Input As 5
While EOF(5) <> -1
    link_ct = link_ct + 1
    Input #5, cif_link(link_ct, 1), cif_link(link_ct, 2), cif_link(link_ct, 3), t
    'Debug.Print cif_link(link_ct, 1), cif_link(link_ct, 2), cif_link(link_ct, 3)
    'Debug.Print
    If Val(cif_link(link_ct, 1)) < 1 Then link_ct = link_ct - 1
Wend
Close 5

For j = 1 To e_tag_ct
    If excel_tag_dat(j, 9) = "_Entry.NMR_STAR_version" Then
        version_num = excel_tag_dat(j, 77)
        Exit For
    End If
Next j
date_text1 = Date
p = InStr(1, date_text1, "/")
mon = Left$(date_text1, p - 1)
If Len(mon) = 1 Then mon = "0" + mon
p1 = InStr(p + 1, date_text1, "/")
da = Mid$(date_text1, p + 1, (p1 - p) - 1)
If Len(da) = 1 Then da = "0" + da
date_text2 = Right$(date_text1, 4) + "-" + mon + "-" + da

Open pathin + "adit_input\mmcif_NMR-STAR_header.txt" For Input As 5
t = "": ttt = ""
While EOF(5) <> -1
    ttt = Input$(1, 5)
    t1 = Asc(ttt)
    If t1 <> 10 And t1 <> 13 Then t = t + ttt
    If t1 = 10 Then
    p = InStr(1, t, "date_text2")
    If p > 0 Then t = Left$(t, p - 1) + date_text2
    p = InStr(1, t, "_dictionary.version")
    If p > 0 Then
        p = InStr(1, t, "version_num")
        t = Left$(t, p - 1) + version_num
    End If
    p = InStr(1, t, "stop_")
    If p > 0 Then
        Print #13, "     "; version_num; "         ?"
        Print #13, ";"
        Print #13, "     ?"
        Print #13, ";"
        Print #13,
    End If
    Print #13, t
    t = ""
    End If
Wend

Print #13,
Close 5

Print #13, "data_NMR-STAR.dic"
Print #13,
Print #13, "#  version date - "; date_text2
Print #13,
Print #13,
Print #13, "# The majority of the following item type definitions and descriptions for units"
Print #13, "# have been taken from the pdbx and mmCIF dictionaries."
Print #13,
Print #13, "# The NMR-STAR dictionary has been constructed to be congruent with the pdbx"
Print #13, "# where the data being modeled are equivalent. In these cases, where possible"
Print #13, "# identical items have been used and the definitions for these items are"
Print #13, "# intended to be identical."
Print #13,
Print #13,
Print #13, "save_dictionary_header"
Print #13, "  _Dictionary.Sf_category                 dictionary_header"
Print #13, "  _Dictionary.ID                          NMR-STAR"
Print #13, "  _Dictionary.Description"
Print #13, ";"
Print #13, "     This data block contains the NMR-STAR dictionary."
Print #13, ";"
Print #13,
Print #13, "    _Dictionary.Title           NMR-STAR.dic"
Print #13, "    _Dictionary.Datablock_id    NMR-STAR.dic"
Print #13, "    _Dictionary.Version         " + version_num
Print #13,
Print #13, "     loop_"
Print #13, "       _Dictionary_history.Version"
Print #13, "       _Dictionary_history.Update"
Print #13, "       _Dictionary_history.Revision"
Print #13, "       _Dictionary_history.Dictionary_ID"
Print #13,
Print #13, "      " + version_num + "      ?"
Print #13, ";"
Print #13, "       ?"
Print #13, ";"
Print #13,
Print #13, "        NMR-STAR"
Print #13,
Print #13, "     stop_"
Print #13,

'print list of super category groups
Print #13, "     loop_"
Print #13, "       _Super_category_group_list.ID"
Print #13, "       _Super_category_group_list.Parent_id"
Print #13, "       _Super_category_group_list.Description"
Print #13, "       _Super_category_group_list.Dictionary_ID"
Print #13,

For j = 1 To cat_ct
    If cat_desc(j, 1) = "super_group" Then
        Print #13, "   " + cat_desc(j, 2) + "           ."
        Print #13, "   " + cat_desc(j, 3)
        Print #13,
    End If
    Print #13, "   NMR-STAR"
Next j
Print #13, "     stop_"
Print #13,
'Print #13,

'print list of sub_categories
'Print #13, "   loop_"
'Print #13, "     _sub_category.id"
'Print #13, "     _sub_category.description"
'Print #13,
'Print #13, "         ?       ?"
'Print #13,
'Print #13, "   stop_"
'Print #13,
Print #13,



'print list of category groups
Print #13, "     loop_"
Print #13, "       _Category_group_list.ID"
Print #13, "       _Category_group_list.Parent_id"
Print #13, "       _Category_group_list.Description"
Print #13, "       _Category_group_list.Dictionary_ID"
Print #13,
Print #13, "              'inclusive_group'      ."
Print #13, ";"
Print #13, "             Categories that belong to the NMR-STAR dictionary."
Print #13, ";"
Print #13,
Print #13, "               NMR-STAR"
For j = 1 To cat_ct
    If cat_desc(j, 1) = "category_group" Then
        If lcase_flag = 1 Then Print #13, "'" + LCase(cat_desc(j, 2)) + "'"
        If lcase_flag = 0 Then Print #13, "'" + (cat_desc(j, 2)) + "'"
        Print #13, "'" + "inclusive_group" + "'"
        Print #13, ";"
        Print #13, cat_desc(j, 3)
        Print #13, ";"
        Print #13,
        Print #13, "       NMR-STAR"
        Print #13,
    End If
Next j
Print #13,
Print #13, "     stop_"
Print #13,
Print #13,


'print out item ddl types and unit definitions
Open pathin + "adit_input\item_type_units.txt" For Input As 5
print_flg = 1
While EOF(5) <> -1
    t = Input$(1, 5)
    t1 = Asc(t)
    If t1 <> 10 And t1 <> 13 Then tt = tt + t
    If t1 = 10 Then
        Print #13, tt
        tt = ""
    End If
Wend
Print #13, tt
Print #13,

Close 5


For j = 1 To e_tag_ct
    print_flag = 1
    For i = 1 To tag_ct
        If excel_tag_dat(j, 9) = tag_list(i) Then
            print_flag = 0
            Exit For
        End If
    Next i
    If excel_tag_dat(j, 30) = "Y" Then print_flag = 0
    If full_dict_flag = 1 Then print_flag = 1
    
    If print_flag = 1 Then
        For i = 1 To e_col_num
            If excel_tag_dat(j, i) = "" Then excel_tag_dat(j, i) = "?"
        Next i
            'set tag information values
        p = InStr(1, excel_tag_dat(j, 9), ".")
        category1 = Mid$(excel_tag_dat(j, 9), 2, p - 2)
        tag1 = Right$(excel_tag_dat(j, 9), Len(excel_tag_dat(j, 9)) - p)

        
        If Trim(excel_tag_dat(j, 38)) <> "" Then
            parent1 = "_" + excel_tag_dat(j, 38) + "." + excel_tag_dat(j, 39)
        Else
            parent1 = "."
        End If
        'If Trim(excel_tag_dat(j, 53)) <> "?" Then
            description1 = excel_tag_dat(j, 92)
            ln = Len(description1)
            If ln > 80 Then
                tt = "": k2 = 1
                For k1 = 1 To ln
                    t = Mid$(description1, k1, 1)
                    If t = "$" Then t = ","
                    tt = tt + t
                
                    If ln > k2 * 80 Then
                        If k1 > (k2 * 80) - 10 Then
                            If Asc(t) = 32 Then
                                tt = tt + Chr$(13) + Chr$(10)
                                k2 = k2 + 1
                            End If
                        End If
                    End If
                Next k1
                description1 = tt
            End If
        'Else
        '    description1 = "?"
        'End If
        
        If excel_tag_dat(j, 52) <> "?" Then
            prompt = excel_tag_dat(j, 52)
        Else
            prompt = "?"
        End If
        If excel_tag_dat(j, 51) <> "?" Then
            example = excel_tag_dat(j, 51)
            'Debug.Print example, excel_tag_dat(j, 87)
            If excel_tag_dat(j, 87) = "Date" Or excel_tag_dat(j, 87) = "yyyy-mm-dd" Then
            If Right$(example, 1) = "'" Then
                ln = Len(example)
                example = Left$(example, ln - 1)
                'Debug.Print example
                'Debug.Print
            End If
            End If
        Else
            example = "?"
        End If
        
'print category information
        If excel_tag_dat(j, 79) <> cat_last Then
            cat_last = excel_tag_dat(j, 79)
            If lcase_flag = 1 Then Print #13, "save_" + LCase(category1)
            If lcase_flag = 0 Then Print #13, "save_" + category1
            Print #13, "   _Category.Sf_category      ."
            Print #13, "   _Category.ID               ."
            Print #13, "   _Category.Description"
            cat_desc_found = 0
            For k = 1 To cat_ct
                If category1 = cat_desc(k, 2) Then
                If cat_desc(k, 1) = "item_category" Then
                If cat_desc(k, 3) <> "?" Then
                    cat_desc_found = 1
                    Print #13, ";"
                    Print #13, cat_desc(k, 3)
                    Print #13, ";"
                    Exit For
                End If
                End If
                End If
            Next k
            If cat_desc_found = 0 Then
                    Print #13, ";"
                    Print #13, "              category description not available"
                    Print #13, ";"
            End If
            If lcase_flag = 1 Then Print #13, "   _Category.ID                   '" + LCase(category1) + "'"
            If lcase_flag = 0 Then Print #13, "   _Category.ID                   '" + category1 + "'"
'            Print #13, "   _category.name                 '" + category1 + "'"
            If excel_tag_dat(j, 3) = "Y" Then print_text = "yes"
            If excel_tag_dat(j, 3) = "N" Then print_text = "no"
            If excel_tag_dat(j, 3) = "Y" And excel_tag_dat(j, 30) <> "Y" Then Print #13, "   _category.mandatory_code       " + print_text
            If excel_tag_dat(j, 30) = "Y" Then Print #13, "   _Category.Mandatory_code       no"
            Print #13,
            Print #13, "   loop_"
            Print #13, "     _Category_key.Name"
            Print #13, "     _Category_key.Category_ID"
            Print #13,
            key_found = 0
            For k = 1 To e_tag_ct
                If category1 = excel_tag_dat(k, 79) Then
                If excel_tag_dat(k, 36) = "Y" Then
                    If lcase_flag = 1 Then Print #13, "     " + "'" + LCase(excel_tag_dat(k, 9)) + "'"
                    If lcase_flag = 0 Then Print #13, "     " + "'" + excel_tag_dat(k, 9) + "'"
                    key_found = 1
                End If
                End If
            Next k
            If key_found = 0 Then Print #13, "      ?"
            Print #13,
            Print #13, "   stop_"
            Print #13,
            Print #13, "   loop_"
            Print #13, "     _Category_group.ID"
            Print #13, "     _Category_group.Category_ID"
            Print #13,
            Print #13, "     'inclusive_group'"
            key_found = 0
            For k = 1 To e_tag_ct
                If category1 = excel_tag_dat(k, 79) Then
                    If lcase_flag = 1 Then Print #13, "     " + "'" + LCase(excel_tag_dat(k, 2)) + "'"
                    If lcase_flag = 0 Then Print #13, "     " + "'" + excel_tag_dat(k, 2) + "'"
                    key_found = 1
                    Exit For
                End If
            Next k
            'If key_found = 0 Then Print #13, "      ?"
            Print #13,
            Print #13, "   stop_"
            Print #13,
            'Print #13, "   loop_"
            'Print #13, "     _Category_examples.Detail"
            'Print #13, "     _Category_examples.Case"
            'Print #13, "     _Category_examples.Category_ID
            'Print #13,
            'Print #13, ";"
            'Print #13, "   ?"
            'Print #13, ";"
            'Print #13, ";"
            'Print #13, "   ?"
            'Print #13, ";"
            'Print #13, ";"
            'Print #13, "   ?"
            'Print #13, ";"
            'Print #13,
            'Print #13, "   stop_"
            'Print #13,
            Print #13, "save_"
            Print #13,
        End If
'        End If
        
        'print tag information
'    If excel_tag_dat(j, 30) <> "Y" Then
        If lcase_flag = 1 Then Print #13, "save_" + LCase(excel_tag_dat(j, 9))
        If lcase_flag = 0 Then Print #13, "save_" + excel_tag_dat(j, 9)
        Print #13, "   _Item.Sf_category               item_description"
        Print #13, "   _Item.ID                        ."
        Print #13, "   _Item.Description"
        Print #13, ";"
        ln = Len(description1)
        For i = 1 To ln
            If Mid$(description1, i, 1) = "$" Then
                description1 = Left$(description1, i - 1) + "," + Right$(description1, ln - i)
                'Debug.Print description1
                'Debug.Print
            End If
        Next i
        Print #13, description1
        'Debug.Print description1
        'Debug.Print
        Print #13, ";"
        Print #13,
    If excel_tag_dat(j, 35) = "?" Then
        If lcase_flag = 1 Then
            Print #13, "   _Item.Name                                  '" + LCase(excel_tag_dat(j, 9)) + "'"
            Print #13, "   _Item.Category_id                           '" + LCase(category1) + "'"
        End If
        If lcase_flag = 0 Then
            Print #13, "   _Item.Name                                  '" + excel_tag_dat(j, 9) + "'"
            Print #13, "   _Item.Category_id                           '" + category1 + "'"
        End If

        If excel_tag_dat(j, 29) = "NOT NULL" Then Print #13, "   _Item.Mandatory_code                         yes"
        If excel_tag_dat(j, 29) <> "NOT NULL" Then Print #13, "   _Item.Mandatory_code                         no"
        'If excel_tag_dat(j, 21) = "N" Then Print #13, "   _pdbx_item_enumeration_details.flag          no"
        'If excel_tag_dat(j, 21) = "Y" Then Print #13, "   _pdbx_item_enumeration_details.flag          yes"
        If excel_tag_dat(j, 22) = "N" Then Print #13, "   _pdbx_item_enumeration_details.closed_flag   no"
        If excel_tag_dat(j, 22) = "Y" Then Print #13, "   _pdbx_item_enumeration_details.closed_flag   yes"
        
        Print #13, "   _Item.Type_code                             " + "'" + excel_tag_dat(j, 87) + "'"
        If excel_tag_dat(j, 42) <> "?" Then Print #13, "   _Item.Units_code              " + excel_tag_dat(j, 42)
        
        If excel_tag_dat(j, 77) <> "?" Then Print #13, "   _Item.Default_value            " + "'" + excel_tag_dat(j, 77) + "'"
        
        'If excel_tag_dat(j, 77) <> "?" Then Print #13, "   _item_range.minimum            " + "?"
        'If excel_tag_dat(j, 77) <> "?" Then Print #13, "   _item_range.maximum            " + "?"
        Print #13,
        
    End If
    If excel_tag_dat(j, 35) = "Y" Then
            Print #13, "  loop_"
            Print #13, "      _Item.Name"
            Print #13, "      _Item.Category_id"
            Print #13, "      _Item.Mandatory_code"
            Print #13, "      _Item.ID"
            Print #13,
            If lcase_flag = 1 Then Print #13, "  '" + LCase(excel_tag_dat(j, 9)); "'";
            If lcase_flag = 0 Then Print #13, "  '" + excel_tag_dat(j, 9); "'";
            ln = Len(excel_tag_dat(j, 9))
            If lcase_flag = 1 Then
                If ln < 45 Then Print #13, Space(53 - ln) + "'" + LCase(category1) + "'";
                If ln > 44 Then Print #13, Space(72 - ln) + "'" + LCase(category1) + "'";
            End If
            If lcase_flag = 0 Then
                If ln < 45 Then Print #13, Space(53 - ln) + "'" + category1 + "'";
                If ln > 44 Then Print #13, Space(72 - ln) + "'" + category1 + "'";
            End If
            ln = Len(category1)
            If ln < 25 Then nsp = 30 - ln
            If ln > 24 Then nsp = 40 - ln
            If excel_tag_dat(j, 29) = "NOT NULL" Then Print #13, Space(nsp) + "yes"
            If excel_tag_dat(j, 29) <> "NOT NULL" Then Print #13, Space(nsp) + "no"
            'hit_ct = 0
            For k = 1 To e_tag_ct
                If k <> j Then
                If excel_tag_dat(k, 30) <> "Y" Then
                If excel_tag_dat(k, 38) = excel_tag_dat(j, 79) Then
                If excel_tag_dat(k, 39) = excel_tag_dat(j, 80) Then
                    If lcase_flag = 1 Then Print #13, "  '" + LCase(excel_tag_dat(k, 9)) + "'";
                    If lcase_flag = 0 Then Print #13, "  '" + excel_tag_dat(k, 9) + "'";
                    ln = Len(excel_tag_dat(k, 9))
                    If lcase_flag = 1 Then
                        If ln < 45 Then Print #13, Space(53 - ln) + "'" + LCase(excel_tag_dat(k, 79)) + "'";
                        If ln > 44 Then Print #13, Space(85 - ln) + "'" + LCase(excel_tag_dat(k, 79)) + "'";
                    End If
                    If lcase_flag = 0 Then
                        If ln < 45 Then Print #13, Space(53 - ln) + "'" + excel_tag_dat(k, 79) + "'";
                        If ln > 44 Then Print #13, Space(85 - ln) + "'" + excel_tag_dat(k, 79) + "'";
                    End If
                    ln = Len(excel_tag_dat(k, 79))
                    If ln < 25 Then nsp = 30 - ln
                    If ln > 24 Then nsp = 50 - ln
                    If excel_tag_dat(j, 29) = "NOT NULL" Then Print #13, Space(nsp) + "yes"
                    If excel_tag_dat(j, 29) <> "NOT NULL" Then Print #13, Space(nsp) + "no"
                    hit_ct = 1
                End If
                End If
                End If
                End If
                Print #13, "      ."
            Next k
            'If hit_ct = 0 Then Print #13, "   ?      ?      ?      ?"
            
            Print #13,
            Print #13, "  stop_"
            Print #13,

        End If
        
        'Print #13, "   loop_"
        'Print #13, "     _Item_aliases.Alias_name"
        'Print #13, "     _Item_aliases.Dictionary"
        'Print #13, "     _Item_aliases.Version"
        'Print #13, "     _Item_aliases.Item_ID"
        'Print #13,
        'Print #13, "     ?  ?  ?"
        'Print #13,
        'Print #13, "   stop_"
        'Print #13,
        
        If excel_tag_dat(j, 35) = "Y" Then
            Print #13, "  loop_"
            Print #13, "      _Item_linked.Child_name"
            Print #13, "      _Item_linked.Parent_name"
            Print #13, "      _Item_linked.Item_ID"
            Print #13,
            
            hit_ct = 0
            For k = 1 To e_tag_ct
                If k <> j Then
                If excel_tag_dat(k, 30) <> "Y" Then
                If excel_tag_dat(k, 38) = excel_tag_dat(j, 79) Then
                If excel_tag_dat(k, 39) = excel_tag_dat(j, 80) Then
                    If lcase_flag = 1 Then Print #13, "  '" + LCase(excel_tag_dat(k, 9)) + "'";
                    If lcase_flag = 0 Then Print #13, "  '" + excel_tag_dat(k, 9) + "'";
                    ln = Len(excel_tag_dat(k, 9))
                    If lcase_flag = 1 Then
                        If ln < 45 Then Print #13, Space(53 - ln) + "'" + LCase(excel_tag_dat(j, 9)) + "'"
                        If ln > 44 Then Print #13, Space(85 - ln) + "'" + LCase(excel_tag_dat(j, 9)) + "'"
                    End If
                    If lcase_flag = 0 Then
                        If ln < 45 Then Print #13, Space(53 - ln) + "'" + excel_tag_dat(j, 9) + "'"
                        If ln > 44 Then Print #13, Space(85 - ln) + "'" + excel_tag_dat(j, 9) + "'"
                    End If
                    hit_ct = hit_ct + 1
                End If
                End If
                End If
                End If
                Print #13, "     ."
            Next k
            If hit_ct = 0 Then Print #13, "   ?      ?       ?"
            
            Print #13,
            Print #13, "  stop_"
            Print #13,
            Print #13, "  _Item_type.Code                             " + "'" + excel_tag_dat(j, 87) + "'"
            Print #13,

        End If
        
        
        If excel_tag_dat(j, 21) = "Y" Then
'            enum_id = enum_id + 1
            Print #13, "   loop_"
            Print #13, "     _Item_enumeration.Value"
            Print #13, "     _Item_enumeration.Detail"
            Print #13, "     _Item_enumeration.Item_ID"
            Print #13,
            For k = 1 To enum_ct
                If enum_values(k, 1) = excel_tag_dat(j, 9) Then
                    Print #13, "   " + enum_values(k, 2) + "    " + enum_values(k, 3) + "    ."
                End If
            Next k

            If enum_ct < 1 Then Print #13, "    ?     ?      ?"
            Print #13,
            Print #13, "   stop_"
            Print #13,
        End If
        
        If example <> "?" Then
            Print #13, "   loop_"
            Print #13, "     _Item_examples.Case"
            Print #13, "     _Item_examples.Item_ID"
            Print #13,
            Print #13, ";"
            ln = Len(example)
            For i23 = 1 To ln
                If Mid$(example, i23, 1) = "$" Then
                    example = Left$(example, i23 - 1) + "," + Right$(example, ln - i23)
                    'Debug.Print example
                    'Debug.Print
                End If
            Next i23
            Print #13, example
            Print #13, ";"
            Print #13, "     ."
            Print #13,
            Print #13, "   stop_"
        End If
        Print #13,
        

'        Print #13, "   _item_natural.primary_key      " + excel_tag_dat(j, 61)
'        Print #13, "   _itme_natural.foreign_key      '" + excel_tag_dat(j, 62) + "'"
'        Print #13,

'        Print #13, "   _item.star_flag                " + excel_tag_dat(j, 56)
'        Print #13, "   _item.db_flag                  " + excel_tag_dat(j, 57)
'        Print #13, "   _item.save_frame_id_flag       " + excel_tag_dat(j, 34)
'        Print #13, "   _item.non_public_flag          " + excel_tag_dat(j, 30)
'        Print #13, "   _item_db.type_code             " + "'" + excel_tag_dat(j, 28) + "'"
'        Print #13, "   _item_db.null_status           " + "'" + excel_tag_dat(j, 29) + "'"
'        Print #13, "   _item_db.table_name            " + excel_tag_dat(j, 31)
'        Print #13, "   _item_db.column_name           " + excel_tag_dat(j, 32)
'        Print #13, "   _item_db.row_index             " + excel_tag_dat(j, 33)
'        Print #13, "   _item_db.source_key            " + excel_tag_dat(j, 35)
'        Print #13, "   _item_db.table_primary_key     " + excel_tag_dat(j, 36)
'        Print #13, "   _item_db.foreign_key           " + excel_tag_dat(j, 37)
'        Print #13, "   _item_db.foreign_table         " + excel_tag_dat(j, 38)
'        Print #13, "   _item_db.foreign_column        " + excel_tag_dat(j, 39)

'        Print #13, "   _item_sub_category.id          "+"?"
'        Print #13, "   _item_enumeration.flag         " + excel_tag_dat(j, 21)
'        Print #13, "   _item_enumeration.closed_flag  " + excel_tag_dat(j, 22)
        
'        Print #13,
'        Print #13, "#  BMRB NMR-STAR validation information"
'        Print #13,
'        Print #13, "  _item_BMRB_validation.mandatory_master_item_flag        ?"
'        Print #13, "  _item_BMRB_validation.mandatory_master_item_code        " + excel_tag_dat(j, 47)
'        Print #13, "  _item_BMRB_validation.mandatory_master_item             ?"
'        Print #13, "  _item_BMRB_validaiton.mandatory_master_item_value       " + excel_tag_dat(j, 48)
'        Print #13, "  _item_BMRB_validation.mandatory_public                  " + excel_tag_dat(j, 65)
'        Print #13, "  _item_BMRB_validation.mandatory_internal                " + excel_tag_dat(j, 66)
'        Print #13, "  _item_BMRB_validation.mandatory_sg_public               " + excel_tag_dat(j, 67)
'        Print #13, "  _item_BMRB_validation.mandatory_sg_internal             " + excel_tag_dat(j, 68)
'        Print #13, "  _item_BMRB_validation.mandatory_user_public             " + excel_tag_dat(j, 69)
'        Print #13, "  _item_BMRB_validation.mandatory_user_internal           " + excel_tag_dat(j, 70)
'        Print #13, "  _item_BMRB_validation.override_public                   " + excel_tag_dat(j, 71)
'        Print #13, "  _item_BMRB_validation.override_internal                 " + excel_tag_dat(j, 72)
'        Print #13, "  _item_BMRB_validation.override_sg_public                " + excel_tag_dat(j, 73)
'        Print #13, "  _item_BMRB_validation.override_sg_internal              " + excel_tag_dat(j, 74)
'        Print #13, "  _item_BMRB_validation.override_user_public              " + excel_tag_dat(j, 75)
'        Print #13, "  _item_BMRB_validation.override_user_internal            " + excel_tag_dat(j, 76)
'        Print #13, "  _item_BMRB_validation.external_file_validation          ?"
'        Print #13,
'        Print #13, "  _item_BMRB_validation.mandatory_loop_flag               " + excel_tag_dat(j, 43)
'        Print #13, "  _item_BMRB_validation.item_ordinal                      " + excel_tag_dat(j, 1)

'        Print #13,
'        Print #13, "#  ADIT information"
'        Print #13,
'        Print #13, "  _item_ADITNMR_view.adit_item_flag             " + excel_tag_dat(j, 15)
'
'        Print #13,
'        If Val(excel_tag_dat(j, 54)) > 0 Then
'            Print #13, "# NMR-STAR to CIF conversion information"
'            Print #13,
'            Print #13, "  _item_NMRSTAR_CIF_link.id              " + excel_tag_dat(j, 54)
'            Print #13, "  _item_NMRSTAR_CIF_link.transform_code  '" + excel_tag_dat(j, 55) + "'"
'
'            For k = 1 To link_ct
'                If Val(cif_link(k, 1)) = Val(excel_tag_dat(j, 54)) Then
'                    Print #13, "  _item_NMRSTAR_CIF_link.pdb_exch_item   '" + cif_link(k, 3) + "'"
'                    'Debug.Print
'                    Exit For
'                End If
'            Next k
'        End If
        
        'Print #13,
        Print #13, "save_"
        Print #13,
    End If
'    End If

Next j

'mmCIF tags used to describe items
'     _item.name
'     _item.category_id
'     _item.mandatory_code

          '_item.name'                      item                      implicit
          '_category_key.name'              category_key              yes
          '_item_aliases.name'              item_aliases              implicit
          '_item_default.name'              item_default              implicit
          '_item_dependent.name'            item_dependent            implicit
          '_item_dependent.dependent_name'  item_dependent            yes
          '_item_description.name'          item_description          implicit
          '_item_enumeration.name'          item_enumeration          implicit
          '_item_examples.name'             item_examples             implicit
          '_item_linked.child_name'         item_linked               yes
          '_item_linked.parent_name'        item_linked               implicit
          '_item_methods.name'              item_methods              implicit
          '_item_range.name'                item_range                implicit
          '_item_related.name'              item_related              implicit
          '_item_related.related_name'      item_related              yes
          '_item_type.name'                 item_type                 implicit
          '_item_type_conditions.name'      item_type_conditions      implicit
          '_item_structure.name'            item_structure            implicit
          '_item_sub_category.name'         item_sub_category         implicit
          '_item_units.name'                item_units                implicit


End Sub


Sub loop_cont_check(t, loop_item_ct, loop_value_ct)

t = Trim(t)
If Len(t) = 5 Then
    p = InStr(1, t, "stop_")
    If p > 0 Then
        If loop_value_ct Mod loop_item_ct > 0 Then
            syntax_check.Text16 = Str(loop_item_ct)
            syntax_check.Text17 = Str(loop_value_ct)
            syntax_check.Text18 = "Loop count error"
            syntax_check.Refresh
            program_control.Show 1
            syntax_check.Text16 = ""
            syntax_check.Text17 = ""
            syntax_check.Refresh
            
        End If
    End If
End If

End Sub

Sub load_adit_data(pathin, input_files, spg_col_ct, spg_row_ct, spg_table, grp_col_ct, grp_row_ct, grp_table, adit_file_source)

'load supergroup file

Open pathin + input_files(3) For Input As 1

While EOF(1) <> -1
    spg_row_ct = spg_row_ct + 1
    For i = 1 To spg_col_ct
        Input #1, spg_table(spg_row_ct, i)
    Next i
    If spg_table(spg_row_ct, 1) = "TBL_BEGIN" Then spg_row_ct = 0
    If spg_table(spg_row_ct, 1) = "TBL_END" Then spg_row_ct = spg_row_ct - 1
Wend
Close 1


'load category group file

Dim Text1(16) As String

Open pathin + input_files(4) For Input As 1

While EOF(1) <> -1
    grp_row_ct = grp_row_ct + 1
    For i = 1 To grp_col_ct
        Input #1, grp_table(grp_row_ct, i)
        If grp_table(grp_row_ct, 1) = "TBL_BEGIN" Then grp_row_ct = 0
    Next i
    For i = 1 To 16
        Input #1, Text1(i)
    Next i
    If adit_file_source = 2 Then
        grp_table(grp_row_ct, 6) = Text1(1)
        grp_table(grp_row_ct, 7) = Text1(2)
        grp_table(grp_row_ct, 8) = Text1(3)
        grp_table(grp_row_ct, 9) = Text1(4)
        grp_table(grp_row_ct, 10) = Text1(5)
        grp_table(grp_row_ct, 19) = Text1(6)
        grp_table(grp_row_ct, 27) = Text1(7)
        grp_table(grp_row_ct, 28) = Text1(8)
    End If
    
    If adit_file_source = 3 Then
        grp_table(grp_row_ct, 6) = Text1(9)
        grp_table(grp_row_ct, 7) = Text1(10)
        grp_table(grp_row_ct, 8) = Text1(11)
        grp_table(grp_row_ct, 9) = Text1(12)
        grp_table(grp_row_ct, 10) = Text1(13)
        grp_table(grp_row_ct, 19) = Text1(14)
        grp_table(grp_row_ct, 27) = Text1(15)
        grp_table(grp_row_ct, 28) = Text1(16)
    End If

    If grp_table(grp_row_ct, 1) = "TBL_END" Then grp_row_ct = grp_row_ct - 1
Wend
Close 1

End Sub
Sub order_tags(tag_ct, new_tag_dat, e_col_num)

Dim i, j As Integer

ReDim tag_dat(tag_ct, e_col_num) As String
ReDim tag_dat_new(600, e_col_num) As String

tag_ct_tot = 0
new_tag_dat(tag_ct + 1, 2) = ""
SFCategory_last = ""
For i = 1 To tag_ct + 1
    If new_tag_dat(i, 2) <> SFCategory_last Then
        SFCategory_last = new_tag_dat(i, 2)
        For j = 1 To tag_new_ct
            If tag_dat_new(j, 43) = "N" Then
                tag_ct_tot = tag_ct_tot + 1
                For k = 1 To e_col_num
                    tag_dat(tag_ct_tot, k) = tag_dat_new(j, k)
                Next k
            End If
        Next j
        For j = 1 To tag_new_ct
            If tag_dat_new(j, 43) = "Y" Then
                tag_ct_tot = tag_ct_tot + 1
                For k = 1 To e_col_num
                    tag_dat(tag_ct_tot, k) = tag_dat_new(j, k)
                Next k
            End If
        Next j
        tag_new_ct = 0
    End If
    tag_new_ct = tag_new_ct + 1
    For j = 1 To e_col_num
        tag_dat_new(tag_new_ct, j) = new_tag_dat(i, j)
    Next j
    
Next i

For i = 1 To tag_ct
    For j = 1 To e_col_num
        new_tag_dat(i, j) = tag_dat(i, j)
    Next j
Next i

End Sub

Sub write_enumerations(loop_flag, data_tag_flag, loop_tag_list, loop_value_ct, last_loop_value_ct, loop_item_ct, loop_value, saveframe_name, save_names_ct)
'create adit_enum_dtl.csv file

Dim i, ln As Integer
'Debug.Print loop_flag, loop_value_ct, last_loop_value_ct, loop_value(loop_value_ct)
'Debug.Print
If loop_flag = 1 Then
    If data_tag_flag = 0 Then
        If loop_tag_list(1) = "_item_enumeration_value" Then
            If loop_value_ct > last_loop_value_ct Then
                For i = last_loop_value_ct + 1 To loop_value_ct
                    If i = 1 Or i Mod loop_item_ct = 1 Then
                        ln = Len(loop_value(i))
                        If Left$(loop_value(i), 1) = "'" Or Left$(loop_value(i), 1) = Chr$(34) Then
                            loop_value(i) = Right$(loop_value(i), ln - 1)
                            ln = ln - 1
                        End If
                        If Right$(loop_value(i), 1) = "'" Or Right$(loop_value(i), 1) = Chr$(34) Then
                            loop_value(i) = Left$(loop_value(i), ln - 1)
                            ln = ln - 1
                        End If
                        
                        p = 1: p1 = 1
                        While p <> 0
                            p = InStr(p1, loop_value(i), ",")
                            If p > 0 Then
                                ln = Len(loop_value(i))
                                test1 = Left$(loop_value(i), p - 1)
                                test2 = Right$(loop_value(i), (ln - p))
                                loop_value(i) = test1 + test2
                                p1 = p + 2
                            End If
                        Wend
                        p = 1: p1 = 1
                        While p <> 0
                            p = InStr(p1, loop_value(i + 1), ",")
                            If p > 0 Then
                                ln = Len(loop_value(i + 1))
                                test1 = Left$(loop_value(i + 1), p - 1)
                                test2 = Right$(loop_value(i + 1), (ln - p))
                                loop_value(i + 1) = test1 + test2
                                p1 = p + 2
                            End If
                        Wend

                        If loop_value(i + 1) = "." Then loop_value(i + 1) = "enumeration_text"
                        If loop_value(i + 1) = "" Then loop_value(i + 1) = "enumeration_text"
                        If loop_value(i + 1) = "?" Then loop_value(i + 1) = "enumeration_text"
                        
                        Print #2, saveframe_name; ","; loop_value(i); ","; loop_value(i + 1); ","; Str((i + 1) / 2); ","
                        'Debug.Print saveframe_name, loop_value(i), loop_value(i + 1), Str((i + 1) / 2)
                        'Debug.Print
                    End If
                Next i
                'Debug.Print
                last_loop_value_ct = loop_value_ct
            End If
        End If
    End If
End If
    
End Sub
Sub write_enum_for_dictionary(loop_flag, data_tag_flag, loop_tag_list, loop_value_ct, last_loop_value_ct, loop_item_ct, loop_value, saveframe_name, save_names_ct)


Dim i, ln As Integer
'Debug.Print loop_flag, loop_value_ct, last_loop_value_ct, loop_value(loop_value_ct)
'Debug.Print
If loop_flag = 1 Then
    If data_tag_flag = 0 Then
        If loop_tag_list(1) = "_item_enumeration_value" Then
            If loop_value_ct > last_loop_value_ct Then
                For i = last_loop_value_ct + 1 To loop_value_ct
                    If i = 1 Or i Mod loop_item_ct = 1 Then
                        ln = Len(loop_value(i))
                        If Left$(loop_value(i), 1) = "'" Or Left$(loop_value(i), 1) = Chr$(34) Then
                            loop_value(i) = Right$(loop_value(i), ln - 1)
                            ln = ln - 1
                        End If
                        If Right$(loop_value(i), 1) = "'" Or Right$(loop_value(i), 1) = Chr$(34) Then
                            loop_value(i) = Left$(loop_value(i), ln - 1)
                            ln = ln - 1
                        End If
                        
                        p = 1: p1 = 1
                        While p <> 0
                            p = InStr(p1, loop_value(i), ",")
                            If p > 0 Then
                                ln = Len(loop_value(i))
                                test1 = Left$(loop_value(i), p - 1)
                                test2 = Right$(loop_value(i), (ln - p))
                                loop_value(i) = test1 + test2
                                p1 = p + 2
                            End If
                        Wend
                        p = 1: p1 = 1
                        While p <> 0
                            p = InStr(p1, loop_value(i + 1), ",")
                            If p > 0 Then
                                ln = Len(loop_value(i + 1))
                                test1 = Left$(loop_value(i + 1), p - 1)
                                test2 = Right$(loop_value(i + 1), (ln - p))
                                loop_value(i + 1) = test1 + test2
                                p1 = p + 2
                            End If
                        Wend

                        If loop_value(i + 1) = "." Then loop_value(i + 1) = "enumeration_text"
                        If loop_value(i + 1) = "" Then loop_value(i + 1) = "enumeration_text"
                        If loop_value(i + 1) = "?" Then loop_value(i + 1) = "enumeration_text"
                        
                        Print #2, saveframe_name; ","; "'" + loop_value(i) + "'"; ","; "'" + loop_value(i + 1) + "'"; ","; Str((i + 1) / 2); ","
                        'Debug.Print saveframe_name, loop_value(i), loop_value(i + 1), Str((i + 1) / 2)
                        'Debug.Print
                    End If
                Next i
                'Debug.Print
                last_loop_value_ct = loop_value_ct
            End If
        End If
    End If
End If
    
End Sub

Sub write_enum_hdr(pathin, output_files, tag_ct, new_tag_dat, e_col_num, excel_header_dat)

Dim i, j, k, enum_ct, enum_id, enid As Integer
Dim col_num(5) As Integer

ReDim enumer(5000, 5) As String
ReDim enum_list(500, 3) As String
Open pathin + output_files(8) For Input As 1
While EOF(1) <> -1
    enum_ct = enum_ct + 1
    For i = 1 To 5
        Input #1, enumer(enum_ct, i)
    Next i
    'Debug.Print enum_ct, enumer(enum_ct, 1), enumer(enum_ct, 2), enumer(enum_ct, 3), enumer(enum_ct, 4)
    'Debug.Print
Wend
Close 1

Open pathin + output_files(9) For Output As 1       'adit_enum_hdr.csv
Print #1, "Enumeration ID,Sf category,Tag"
Print #1, "TBL_BEGIN,?,?"
k = 0
For j = 1 To e_col_num
    If excel_header_dat(1, j) = "Item enumerated" Then col_num(0) = j
    If excel_header_dat(1, j) = "SFCategory" Then col_num(1) = j
    If excel_header_dat(1, j) = "Tag" Then col_num(2) = j: tag_col = j
    If excel_header_dat(1, j) = "Enum parent SFcategory" Then col_num(3) = j
    If excel_header_dat(1, j) = "Enum parent tag" Then col_num(4) = j
Next j

enum_id = 0
For i = 1 To tag_ct
    If new_tag_dat(i, col_num(0)) = "Y" Then
        'If new_tag_dat(i, col_num(3)) = "" Then
            'Debug.Print
            enum_id = enum_id + 1
            Print #1, Trim(Str(enum_id)) + ",";
            Print #1, new_tag_dat(i, col_num(1)) + ",";
            Print #1, new_tag_dat(i, col_num(2))
            enum_list(enum_id, 1) = Trim(Str(enum_id))
            enum_list(enum_id, 2) = new_tag_dat(i, col_num(1))
            enum_list(enum_id, 3) = new_tag_dat(i, col_num(2))
            For jk = 1 To enum_ct
                If enumer(jk, 1) = new_tag_dat(i, tag_col) Then
                    enumer(jk, 1) = Trim(Str(enum_id))
                    'Debug.Print
                End If
            Next jk
            'Debug.Print
        'End If
    End If
Next i
'For i = 1 To tag_ct
'    If new_tag_dat(i, col_num(0)) = "Y" Then
'        If new_tag_dat(i, col_num(3)) <> "" Then
'            For j = 1 To enum_id
'                If new_tag_dat(i, col_num(3)) = enum_list(j, 2) Then
'                If new_tag_dat(i, col_num(4)) = enum_list(j, 3) Then
'                    Print #1, enum_list(j, 1) + ",";
'                    Print #1, new_tag_dat(i, col_num(1)) + ",";
'                    Print #1, new_tag_dat(i, col_num(2))
'                    'Debug.Print
'                End If
'                End If
'            Next j
'        End If
'    End If
'Next i
Print #1, "TBL_END,?,?"

'output file enum_dtl.csv
Close 1
Open pathin + output_files(8) For Output As 1
Print #1, "Enumeration ID,Enum ordinal,Enum value,Enum value desc"
Print #1, "TBL_BEGIN,?,?,?"
For i = 1 To enum_ct
    enid = Val(Left$(enumer(i, 1), 1))
    If enid > 0 And enid < 10 Then
        ln = Len(enumer(i, 2))
        If Left$(enumer(i, 2), 1) = "'" Then enumer(i, 2) = Chr$(34) + Right$(enumer(i, 2), ln - 1)
        If Right$(enumer(i, 2), 1) = "'" Then enumer(i, 2) = Left$(enumer(i, 2), ln - 1) + Chr$(34)
        
        ln = Len(enumer(i, 3))
        If Left$(enumer(i, 3), 1) = "'" Then enumer(i, 3) = Chr$(34) + Right$(enumer(i, 3), ln - 1)
        If Right$(enumer(i, 3), 1) = "'" Then enumer(i, 3) = Left$(enumer(i, 3), ln - 1) + Chr$(34)
        
        Print #1, enumer(i, 1) + "," + enumer(i, 4) + "," + enumer(i, 2) + "," + enumer(i, 3)
    End If
Next i
Print #1, "TBL_END,?,?,?"
Close 1
End Sub

Sub write_group_tables(pathin, output_files, input_files, adit_file_source)
Debug.Print "Writing category and supercategory group files"

'Write out the supergroup table for Steve's software

Open pathin + input_files(3) For Input As 1
Open pathin + output_files(5) For Output As 2

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

Open pathin + input_files(4) For Input As 1
Open pathin + output_files(4) For Output As 2

col_ct = 36
ReDim table_row(200, col_ct) As String
ReDim test1(8) As String

row_ct = 0
While EOF(1) <> -1
    row_ct = row_ct + 1
    For i = 1 To col_ct
        Input #1, table_row(row_ct, i)
        'Debug.Print row_ct, i, table_row(row_ct, i)
        'Debug.Print
    Next i
    
    For i = 1 To 8
        Input #1, test1(i)
    Next i
    
    If adit_file_source = 2 Then
        table_row(row_ct, 6) = table_row(row_ct, 29)
        table_row(row_ct, 7) = table_row(row_ct, 30)
        table_row(row_ct, 8) = table_row(row_ct, 31)
        table_row(row_ct, 9) = table_row(row_ct, 32)
        table_row(row_ct, 10) = table_row(row_ct, 33)
        table_row(row_ct, 19) = table_row(row_ct, 34)
        table_row(row_ct, 27) = table_row(row_ct, 35)
        table_row(row_ct, 28) = table_row(row_ct, 36)
    End If
    
    If adit_file_source = 3 Then
        table_row(row_ct, 6) = test1(1)
        table_row(row_ct, 7) = test1(2)
        table_row(row_ct, 8) = test1(3)
        table_row(row_ct, 9) = test1(4)
        table_row(row_ct, 10) = test1(5)
        table_row(row_ct, 19) = test1(6)
        table_row(row_ct, 27) = test1(7)
        table_row(row_ct, 28) = test1(8)
    End If
    
    'Debug.Print row_ct, table_row(row_ct, 1)
    'Debug.Print
Wend
Close 1
col_ct = 28

For i = 1 To row_ct
    For j = 1 To 28
        If table_row(3, j) = "validate" Then
            Validate = ""
            For j1 = 13 To 18
                Validate = Validate + table_row(i, j1)
            Next j1
            If i > 3 Then Print #2, Validate + ",";
            If i <= 3 Then Print #2, table_row(i, j) + ",";
        End If
        If table_row(3, j) = "viewFlags" Then
            Validate = ""
            For j1 = 19 To 24
                Validate = Validate + table_row(i, j1)
            Next j1
            If i > 3 Then Print #2, Validate + ",";
            If i <= 3 Then Print #2, table_row(i, j) + ",";
        End If

        If j <= 4 Then Print #2, table_row(i, j) + ",";
        If j > 7 And j < 11 Then Print #2, table_row(i, j) + ",";
        If j = 12 Then Print #2, table_row(i, j) + ",";
        If j = 27 Then Print #2, table_row(i, j) + ",";
        If j = 28 Then Print #2, table_row(i, j)
    Next j

Next i
Close 2

End Sub


Sub compare_tags(tag_ct, tag_char, e_tag_ct, excel_tag_dat, new_tag_dat, e_col_num)
Debug.Print "Compare tags"

Dim i, j, start As Integer

start = 1
For i = 1 To tag_ct
    'Debug.Print i
    If i > 100 Then start = i - 99
    For j = start To e_tag_ct
        'If tag_char(i, 3) = excel_tag_dat(j, 2) Then        'category
        'p = InStr(1, "_" + LCase(tag_char(i, 1)), LCase(excel_tag_dat(j, 9))) 'match tag
        'Debug.Print tag_char(i, 1), excel_tag_dat(j, 9)
        'Debug.Print
        'If Len(tag_char(i, 1)) + 1 <> Len(excel_tag_dat(j, 9)) Then p = 0
        'If p > 0 Then
        If LCase(tag_char(i, 2) + "." + tag_char(i, 1)) = LCase(excel_tag_dat(j, 9)) Then 'compare tags
        If LCase(tag_char(i, 3)) = LCase(excel_tag_dat(j, 2)) Then
            For k = 1 To e_col_num
                new_tag_dat(i, k) = excel_tag_dat(j, k)
            Next k
            tag_char(i, 0) = "y"
            
            'Debug.Print tag_char(i, 1), excel_tag_dat(j, 9)
            'Debug.Print

            Exit For
        End If
        End If
        'End If
    Next j
Next i

End Sub
Sub compare_dep_tags(tag_ct, tag_char, e_tag_ct, excel_tag_dat, new_tag_dat, e_col_num)
Debug.Print "Compare deposition tags"

Dim i, j, tag_found As Integer

For i = 1 To tag_ct
    'Debug.Print i
    tag_found = 0
    For j = 1 To e_tag_ct
        If tag_char(i, 2) + "." + tag_char(i, 1) = excel_tag_dat(j, 9) Then 'compare tags
        If LCase(tag_char(i, 3)) = LCase(excel_tag_dat(j, 2)) Then
            
            tag_found = 1
            Exit For
        End If
        End If
    Next j
    If tag_found = 0 Then
        Debug.Print tag_char(i, 2) + "." + tag_char(i, 1), tag_char(i, 3), "not found"
    End If
'Debug.Print
Next i

End Sub

Sub write_nmrstar_dict3(e_tag_ct, excel_tag_dat, tag_ct, tag_list, e_col_num, pathin, output_files, lcase_flag, full_dict_flag, input_files, excel_header_dat)

Dim i, j, k, print_flag, link_ct, cat_ct, print_tags, print_tags_f As Integer
Dim enum_values(5000, 5), cat_desc(1000, 3) As String
Dim cat_table(120, 50), man_over(2000, 8) As String
Dim cif_link(5000, 3) As String
Dim t, ttt, test_str As String

'Open pathin + output_files(13) For Input As 5       'Load manual overide file
'While EOF(5) <> -1
'    man_over_ct = man_over_ct + 1
'    For i = 1 To 8
'        Input #5, man_over(man_over_ct, i)
'    Next i
'Wend
'Close 5

Open pathin + output_files(22) For Input As 5       'Load category, super category, category group descriptions
While cat_desc(cat_ct, 1) <> "Table_end"
    cat_ct = cat_ct + 1
    test_str = ""
    For i = 1 To 3
        Input #5, cat_desc(cat_ct, i)
    Next i
    'Debug.Print cat_ct, cat_desc(cat_ct, 1), cat_desc(cat_ct, 2), cat_desc(cat_ct, 3)
    'Debug.Print
    ln = Len(cat_desc(cat_ct, 3))
    For k = 1 To ln
        t = Mid$(cat_desc(cat_ct, 3), k, 1)
        If t = "$" Then t = ","
        test_str = test_str + t
    Next k
    cat_desc(cat_ct, 3) = test_str
Wend
Close 5

Open pathin + input_files(4) For Input As 5         'Load category group table
While cat_table(cat_tbl_ct, 1) <> "TBL_END"
    cat_tbl_ct = cat_tbl_ct + 1
    For i = 1 To 44
        Input #5, cat_table(cat_tbl_ct, i)
    Next i
Wend
Close 5

Open pathin + output_files(21) For Input As 5   'Load enumerations
While EOF(5) <> -1
    enum_ct = enum_ct + 1
    For i = 1 To 5
        Input #5, enum_values(enum_ct, i)
    Next i
    ln = Len(enum_values(enum_ct, 2))
    enum_values(enum_ct, 2) = Mid$(enum_values(enum_ct, 2), 2, ln - 2)
    ln = Len(enum_values(enum_ct, 3))
    enum_values(enum_ct, 3) = Mid$(enum_values(enum_ct, 3), 2, ln - 2)
    
    'Debug.Print enum_ct, enum_values(enum_ct, 1), enum_values(enum_ct, 2), enum_values(enum_ct, 3), enum_values(enum_ct, 4), enum_values(enum_ct, 5)
    'Debug.Print
'Wend
'Close 5

'Open "z:\"+"eldonulrich\"+"bmrb\htdocs\dictionary\htmldocs\nmr_star\development_adit_files\adit_input\adit_enum_dtl.csv" For Input As 5
'While EOF(5) <> -1
'    ct = ct + 1
'    Input #5, enum_values(ct, 1), enum_values(ct, 2), enum_values(ct, 3), enum_values(ct, 4)

    If enum_values(enum_ct, 3) = "enumeration_text" Then enum_values(enum_ct, 3) = "?"
    ln = Len(enum_values(enum_ct, 2))
    flag1 = 0: flag2 = 0: flag3 = 0
    For i = 1 To ln
        t = Mid$(enum_values(enum_ct, 2), i, 1)
        If t = " " Then flag1 = 1
        If t = Chr$(9) Then flag1 = 1
        If t = Chr$(91) Then flag1 = 1          ' chr [
        If t = Chr$(93) Then flag1 = 1          ' chr ]
        If t = Chr$(39) Then flag2 = 1
        If t = Chr$(96) Then flag2 = 1
        If t = Chr$(34) Then flag3 = 1
        If t = Chr$(36) Then
            enum_values(enum_ct, 2) = Left$(enum_values(enum_ct, 2), i - 1) + "," + Right$(enum_values(enum_ct, 2), ln - i)
        End If
    Next i
    If Left$(enum_values(enum_ct, 2), 1) = Chr$(34) And Right$(enum_values(enum_ct, 2), 1) = Chr$(34) Then flag3 = 4
    If Left$(enum_values(enum_ct, 2), 1) = Chr$(39) And Right$(enum_values(enum_ct, 2), 1) = Chr$(39) Then flag3 = 4
    If Left$(enum_values(enum_ct, 2), 1) = Chr$(96) And Right$(enum_values(enum_ct, 2), 1) = Chr$(96) Then flag3 = 4
    If Left$(enum_values(enum_ct, 2), 1) = Chr$(39) And Right$(enum_values(enum_ct, 2), 1) = Chr$(96) Then flag3 = 4
    If Left$(enum_values(enum_ct, 2), 1) = Chr$(96) And Right$(enum_values(enum_ct, 2), 1) = Chr$(39) Then flag3 = 4
    
    If flag1 = 1 Then
        If flag2 = 0 And flag3 = 0 Then
            enum_values(enum_ct, 2) = Chr$(39) + enum_values(enum_ct, 2) + Chr$(39)
        End If
        If flag2 = 1 And flag3 = 0 Then
            enum_values(enum_ct, 2) = Chr$(34) + enum_values(enum_ct, 2) + Chr$(34)
        End If
        If flag2 = 0 And flag3 = 1 Then
            enum_values(enum_ct, 2) = Chr$(39) + enum_values(enum_ct, 2) + Chr$(39)
        End If
        If flag2 = 1 And flag3 = 1 Then
            enum_values(enum_ct, 2) = Chr$(10) + Chr$(59) + enum_values(enum_ct, 2) + Chr$(10) + Chr$(59) + Chr$(10)
        End If
    End If
    ln = Len(enum_values(enum_ct, 3))
    flag1 = 0: flag2 = 0: flag3 = 0
    For i = 1 To ln
        t = Mid$(enum_values(enum_ct, 3), i, 1)
        If t = " " Then flag1 = 1
        If t = Chr$(9) Then flag1 = 1
        If t = Chr$(39) Then flag2 = 1
        If t = Chr$(96) Then flag2 = 1
        If t = Chr$(34) Then flag3 = 1
    Next i
    If Left$(enum_values(enum_ct, 3), 1) = Chr$(34) And Right$(enum_values(enum_ct, 3), 1) = Chr$(34) Then flag3 = 4
    If Left$(enum_values(enum_ct, 3), 1) = Chr$(39) And Right$(enum_values(enum_ct, 3), 1) = Chr$(39) Then flag3 = 4
    If Left$(enum_values(enum_ct, 3), 1) = Chr$(96) And Right$(enum_values(enum_ct, 3), 1) = Chr$(96) Then flag3 = 4
    If Left$(enum_values(enum_ct, 3), 1) = Chr$(39) And Right$(enum_values(enum_ct, 3), 1) = Chr$(96) Then flag3 = 4
    If Left$(enum_values(enum_ct, 3), 1) = Chr$(96) And Right$(enum_values(enum_ct, 3), 1) = Chr$(39) Then flag3 = 4
    If flag1 = 1 Then
        If flag2 = 0 And flag3 = 0 Then
            enum_values(enum_ct, 3) = Chr$(39) + enum_values(enum_ct, 3) + Chr$(39)
        End If
        If flag2 = 1 And flag3 = 0 Then
            enum_values(enum_ct, 3) = Chr$(34) + enum_values(enum_ct, 3) + Chr$(34)
        End If
        If flag2 = 0 And flag3 = 1 Then
            enum_values(enum_ct, 3) = Chr$(39) + enum_values(enum_ct, 3) + Chr$(39)
        End If
        If flag2 = 1 And flag3 = 1 Then
            enum_values(enum_ct, 3) = Chr$(10) + Chr$(59) + enum_values(enum_ct, 3) + Chr$(10) + Chr$(59) + Chr$(10)
        End If
    End If

    'Debug.Print enum_values(enum_ct, 1), enum_values(enum_ct, 2), enum_values(enum_ct, 3)
    'Debug.Print
Wend
Close 5

Open pathin + "nmr_cif_match.csv" For Input As 5
While EOF(5) <> -1
    link_ct = link_ct + 1
    Input #5, cif_link(link_ct, 1), cif_link(link_ct, 2), cif_link(link_ct, 3), t
    'Debug.Print cif_link(link_ct, 1), cif_link(link_ct, 2), cif_link(link_ct, 3)
    'Debug.Print
    If Val(cif_link(link_ct, 1)) < 1 Then link_ct = link_ct - 1
Wend
Close 5

For j = 1 To e_tag_ct
    If excel_tag_dat(j, 9) = "_Entry.NMR_STAR_version" Then
        version_num = excel_tag_dat(j, 77)
        Exit For
    End If
Next j
date_text1 = Date
p = InStr(1, date_text1, "/")
mon = Left$(date_text1, p - 1)
If Len(mon) = 1 Then mon = "0" + mon
p1 = InStr(p + 1, date_text1, "/")
da = Mid$(date_text1, p + 1, (p1 - p) - 1)
If Len(da) = 1 Then da = "0" + da
date_text2 = Right$(date_text1, 4) + "-" + mon + "-" + da

Open pathin + "NMR-STAR_STAR_header.txt" For Input As 5
t = "": ttt = ""
While EOF(5) <> -1
    ttt = Input$(1, 5)
    t1 = Asc(ttt)
    If t1 <> 10 And t1 <> 13 Then t = t + ttt
    If t1 = 10 Then
    'p = InStr(1, t, " _")
    'If p > 0 Then
    '    t = "    " + t
    '    p = 0
    'End If
    p = InStr(1, t, "date_text2")
    If p > 0 Then t = Left$(t, p - 1) + date_text2
    p = InStr(1, t, "_dictionary.version")
    If p > 0 Then
        p = InStr(1, t, "version_num")
        t = Left$(t, p - 1) + version_num
    End If
    p = InStr(1, t, "stop_")
    'If p > 0 Then
    '    Print #13, "     "; version_num; "         ?"
    '    Print #13, ";"
    '    Print #13, "     ?"
    '    Print #13, ";"
    '    Print #13,
    'End If
    Print #13, t
    t = ""
    End If
Wend

Print #13,
Close 5


'print list of category groups
Print #13, "     loop_"
Print #13, "       _category_group_list.id"
Print #13, "       _category_group_list.parent_id"
Print #13, "       _category_group_list.description"
Print #13,
Print #13, "'inclusive_group'"
Print #13, "'.'"
Print #13, ";"
Print #13, "Categories that belong to the NMR-STAR dictionary."
Print #13, ";"
Print #13,
For j = 1 To cat_ct
    If cat_desc(j, 1) = "category_group" Then
        If lcase_flag = 1 Then Print #13, "'" + LCase(cat_desc(j, 2)) + "'"
        If lcase_flag = 0 Then Print #13, "'" + (cat_desc(j, 2)) + "'"
        Print #13, "'" + "inclusive_group" + "'"
        Print #13, ";"
        Print #13, cat_desc(j, 3)
        Print #13, ";"
        Print #13,
    End If
Next j
Print #13, "     stop_"
Print #13,
Print #13,

'print out item ddl types and unit definitions
Open pathin + "item_type_units.txt" For Input As 5
print_flg = 1
While EOF(5) <> -1
    t = Input$(1, 5)
    t1 = Asc(t)
    If t1 <> 10 And t1 <> 13 Then tt = tt + t
    If t1 = 10 Then
        Print #13, tt
        tt = ""
    End If
Wend
Print #13, tt
Print #13,

Close 5
For j = 1 To e_tag_ct
    print_flag = 1
    For i = 1 To tag_ct
        If excel_tag_dat(j, 9) = tag_list(i) Then
            print_flag = 0
            Exit For
        End If
    Next i
    If excel_tag_dat(j, 30) = "Y" Then print_flag = 0
    If full_dict_flag = 1 Then print_flag = 1
    
    If print_flag = 1 Then
        For i = 1 To e_col_num
            If excel_tag_dat(j, i) = "" Then excel_tag_dat(j, i) = "?"
        Next i
            'set tag information values
        p = InStr(1, excel_tag_dat(j, 9), ".")
        category1 = Mid$(excel_tag_dat(j, 9), 2, p - 2)
        tag1 = Right$(excel_tag_dat(j, 9), Len(excel_tag_dat(j, 9)) - p)

        
        If Trim(excel_tag_dat(j, 38)) <> "" Then
            parent1 = "_" + excel_tag_dat(j, 38) + "." + excel_tag_dat(j, 39)
        Else
            parent1 = "."
        End If
        
        description1 = excel_tag_dat(j, 92)
        ln = Len(description1)
        If ln > 80 Then
            tt = "": k2 = 1
            For k1 = 1 To ln
                t = Mid$(description1, k1, 1)
                If t = "$" Then t = ","
                tt = tt + t
                
                If ln > k2 * 80 Then
                    If k1 > (k2 * 80) - 10 Then
                        If Asc(t) = 32 Then
                            tt = tt + Chr$(13) + Chr$(10)
                            k2 = k2 + 1
                        End If
                    End If
                End If
            Next k1
            description1 = tt
        End If
        
        If excel_tag_dat(j, 52) <> "?" Then
            prompt = excel_tag_dat(j, 52)
        Else
            prompt = "?"
        End If
        If excel_tag_dat(j, 51) <> "?" Then
            example = excel_tag_dat(j, 51)
            'Debug.Print example, excel_tag_dat(j, 87)
            If excel_tag_dat(j, 87) = "Date" Or excel_tag_dat(j, 87) = "yyyy-mm-dd" Then
            If Right$(example, 1) = "'" Then
                ln = Len(example)
                example = Left$(example, ln - 1)
                'Debug.Print example
                'Debug.Print
            End If
            End If
        Else
            example = "?"
        End If
        
        
        'print category group information
        If excel_tag_dat(j, 2) <> catgroup_last Then
            Print #13, "save_" + UCase(excel_tag_dat(j, 2))
            Print #13, "   _category_group.description"
            catg_desc_found = 0
            For k = 1 To cat_ct
                If excel_tag_dat(j, 2) = cat_desc(k, 2) Then
                If cat_desc(k, 1) = "category_group" Then
                If cat_desc(k, 3) <> "?" Then
                    catg_desc_found = 1
                    Print #13, ";"
                    Print #13, cat_desc(k, 3)
                    Print #13, ";"
                    Exit For
                End If
                End If
                End If
            Next k
            If catg_desc_found = 0 Then
                    Print #13, ";"
                    Print #13, "             Item category group description not available"
                    Print #13, ";"
            End If
            
            For k1 = 1 To cat_tbl_ct
                If excel_tag_dat(j, 2) = cat_table(k1, 5) Then kt = k1
            Next k1

            Print #13,
            Print #13, "   _category_group.id                    " + excel_tag_dat(j, 2)
            'Print #13, "   _category_group.name                  " + excel_tag_dat(j, 2)
           
            Print #13,
            Print #13, "save_"
            Print #13,
            catgroup_last = excel_tag_dat(j, 2)
        End If
        
        'print category information
        If excel_tag_dat(j, 79) <> cat_last Then
            cat_last = excel_tag_dat(j, 79)
            If lcase_flag = 1 Then Print #13, "save_" + LCase(category1)
            If lcase_flag = 0 Then Print #13, "save_" + category1
            Print #13, "   _category.description"
            cat_desc_found = 0
            For k = 1 To cat_ct
                If category1 = cat_desc(k, 2) Then
                If cat_desc(k, 1) = "item_category" Then
                If cat_desc(k, 3) <> "?" Then
                    cat_desc_found = 1
                    Print #13, ";"
                    Print #13, cat_desc(k, 3)
                    Print #13, ";"
                    Exit For
                End If
                End If
                End If
            Next k
            If cat_desc_found = 0 Then
                    Print #13, ";"
                    Print #13, "              category description not available"
                    Print #13, ";"
            End If
            If lcase_flag = 1 Then Print #13, "   _category.id                   '" + LCase(category1) + "'"
            If lcase_flag = 0 Then Print #13, "   _category.id                   '" + category1 + "'"
            
            If excel_tag_dat(j, 3) = "Y" Then print_text = "yes"
            If excel_tag_dat(j, 3) = "N" Then print_text = "no"
            If excel_tag_dat(j, 3) = "Y" And excel_tag_dat(j, 30) <> "Y" Then Print #13, "   _category.mandatory_code       " + print_text
            If excel_tag_dat(j, 30) = "Y" Then Print #13, "   _category.mandatory_code       no"
            Print #13,
            print_tags = 0
                
            If print_tags > 0 Then
                Print #13,
                Print #13, "   stop_"
                Print #13,
            End If
            
            Print #13, "   loop_"
            Print #13, "     _category_key.name"
            Print #13,
            
            key_found = 0
            For k = 1 To e_tag_ct
                If category1 = excel_tag_dat(k, 79) Then
                If excel_tag_dat(k, 36) = "Y" Then
                    If lcase_flag = 1 Then Print #13, "     " + "'" + LCase(excel_tag_dat(k, 9)) + "'"
                    If lcase_flag = 0 Then Print #13, "     " + "'" + excel_tag_dat(k, 9) + "'"
                    key_found = 1
                End If
                End If
            Next k
            
            If key_found = 0 Then Print #13, "      ?"
            
            Print #13,
            Print #13, "   stop_"
            Print #13,
            Print #13, "   loop_"
            Print #13, "     _category_group.id"
            Print #13,
            Print #13, "     'inclusive_group'"
            
            key_found = 0
            For k = 1 To e_tag_ct
                If category1 = excel_tag_dat(k, 79) Then
                    If lcase_flag = 1 Then Print #13, "     " + "'" + LCase(excel_tag_dat(k, 2)) + "'"
                    If lcase_flag = 0 Then Print #13, "     " + "'" + excel_tag_dat(k, 2) + "'"
                    key_found = 1
                    Exit For
                End If
            Next k
            
            Print #13,
            Print #13, "   stop_"
            Print #13,
            Print #13, "   loop_"
            Print #13, "     _category_examples.detail"
            Print #13, "     _category_examples.case"
            Print #13,
            Print #13, ";"
            Print #13, "   ?"
            Print #13, ";"
            Print #13, ";"
            Print #13, "   ?"
            Print #13, ";"
            Print #13,
            Print #13, "   stop_"
            Print #13,
            Print #13, "save_"
            Print #13,
        End If
        
        'print tag information
        
        If lcase_flag = 1 Then Print #13, "save_" + LCase(excel_tag_dat(j, 9))
        If lcase_flag = 0 Then Print #13, "save_" + excel_tag_dat(j, 9)
        
        Print #13,
        Print #13, "   _item_description.name            '" + LCase(excel_tag_dat(j, 9)) + "'"
        Print #13, "   _item_description.description"
        Print #13, ";"
        Print #13, description1
        Print #13, ";"
        Print #13,
        
        If lcase_flag = 1 Then
            Print #13, "   _item.name                         '" + LCase(excel_tag_dat(j, 9)) + "'"
            Print #13, "   _item.category_id                  '" + LCase(category1) + "'"
        End If
        
        If lcase_flag = 0 Then
            Print #13, "   _item.name                         '" + excel_tag_dat(j, 9) + "'"
            Print #13, "   _item.category_id                  '" + category1 + "'"
        End If
        
        If excel_tag_dat(j, 77) = "" Then default_val = "." Else default_val = excel_tag_dat(j, 77)
        
        Print #13, "   _item.default_value                " + "'" + default_val + "'"
        If excel_tag_dat(j, 42) <> "?" Then Print #13, "   _item_units.code                  " + excel_tag_dat(j, 42)
        
        Print #13, "   _pdbx_item_enumeration_details.flag              " + excel_tag_dat(j, 21)
        If excel_tag_dat(j, 22) = "?" Then enum_closed_flag = "N"
        If excel_tag_dat(j, 22) = "Y" Then enum_closed_flag = "Y"
        Print #13, "   _pdbx_item_enumeration_details.closed_flag       " + enum_closed_flag
                
        If excel_tag_dat(j, 43) = "Y" Then loop_flag = "yes"
        If excel_tag_dat(j, 43) = "N" Then loop_flag = "no"
        Print #13, "   _item.loop_flag                     " + loop_flag
        Print #13,
        
        'print database information
        'Print #13, "   _item.star_flag                " + excel_tag_dat(j, 56)
        'Print #13, "   _item.db_flag                  " + excel_tag_dat(j, 57)
        'Print #13, "   _item.save_frame_id_flag       " + excel_tag_dat(j, 34)
        'Print #13, "   _item.non_public_flag          " + excel_tag_dat(j, 30)
        
        Print #13, "   _item_type.name                    '" + LCase(excel_tag_dat(j, 9)) + "'"
        Print #13, "   _item_type.code                    " + "'" + excel_tag_dat(j, 87) + "'"
        Print #13, "   _item_type.db_type_code             " + excel_tag_dat(j, 28)
        
        Print #13,
        Print #13, "   _item_db.name                      '" + LCase(excel_tag_dat(j, 9)) + "'"
        If excel_tag_dat(j, 29) = "NOT NULL" Then Print #13, "   _item_db.table_mandatory_code      " + "'" + excel_tag_dat(j, 29) + "'"
        If excel_tag_dat(j, 29) <> "NOT NULL" Then Print #13, "   _item_db.table_mandatory_code       ?"
        Print #13, "   _item_db.table_name                " + "'" + excel_tag_dat(j, 79) + "'"
        Print #13, "   _item_db.table_column_name         " + "'" + excel_tag_dat(j, 80) + "'"
        Print #13, "   _item_db.table_primary_key          " + excel_tag_dat(j, 36)
        
            If excel_tag_dat(j, 89) <> "?" And excel_tag_dat(j, 89) <> "" Then
                Print #13, "   _item_db.table_foreign_key_flag     Y"
                ReDim mate(30, 3) As String
                mate_ct = 0
                p = InStr(1, excel_tag_dat(j, 89), ";")
                If p < 1 Then
                    mate_ct = 1
                    mate(mate_ct, 1) = excel_tag_dat(j, 89)
                    mate(mate_ct, 2) = excel_tag_dat(j, 90)
                    mate(mate_ct, 3) = excel_tag_dat(j, 91)
                End If
                If p > 0 Then
                    mate_ct = 1
                    mate(mate_ct, 1) = Left$(excel_tag_dat(j, 89), p - 1)
                    p1 = 1
                    While p1 > 0
                        p1 = InStr(p + 1, excel_tag_dat(j, 89), ";")
                        If p1 > 0 Then
                            mate_ct = mate_ct + 1
                            mate(mate_ct, 1) = Mid$(excel_tag_dat(j, 89), p + 1, (p1 - p) - 1)
                            p = p1
                        End If
                    Wend
                    mate_ct = mate_ct + 1
                    mate(mate_ct, 1) = Right$(excel_tag_dat(j, 89), Len(excel_tag_dat(j, 89)) - p)
                End If
                p = InStr(1, excel_tag_dat(j, 90), ";")
                If p > 0 Then
                    mate_ct = 1
                    mate(mate_ct, 2) = Left$(excel_tag_dat(j, 90), p - 1)
                    p1 = 1
                    While p1 > 0
                        p1 = InStr(p + 1, excel_tag_dat(j, 90), ";")
                        If p1 > 0 Then
                            mate_ct = mate_ct + 1
                            mate(mate_ct, 2) = Mid$(excel_tag_dat(j, 90), p + 1, (p1 - p) - 1)
                            p = p1
                        End If
                    Wend
                    mate_ct = mate_ct + 1
                    mate(mate_ct, 2) = Right$(excel_tag_dat(j, 90), Len(excel_tag_dat(j, 90)) - p)
                End If
                p = InStr(1, excel_tag_dat(j, 91), ";")
                If p > 0 Then
                    mate_ct = 1
                    mate(mate_ct, 3) = Left$(excel_tag_dat(j, 91), p - 1)
                    p1 = 1
                    While p1 > 0
                        p1 = InStr(p + 1, excel_tag_dat(j, 91), ";")
                        If p1 > 0 Then
                            mate_ct = mate_ct + 1
                            mate(mate_ct, 3) = Mid$(excel_tag_dat(j, 91), p + 1, (p1 - p) - 1)
                            p = p1
                        End If
                    Wend
                    mate_ct = mate_ct + 1
                    mate(mate_ct, 3) = Right$(excel_tag_dat(j, 91), Len(excel_tag_dat(j, 91)) - p)
                End If
        
        Print #13,
        Print #13, "   loop_"
        Print #13, "      _item_db_foreign_key.group_ID"
        Print #13, "      _item_db_foreign_key.related_table_name"
        Print #13, "      _item_db_foreign_key.related_table_column_name"
        Print #13,
        For i3 = 1 To mate_ct
            Print #13, "   " + mate(i3, 1) + Space(8 - Len(mate(i3, 1))) + mate(i3, 2) + Space(37 - Len(mate(i3, 2))) + mate(i3, 3)
        Next i3
        Print #13,
        Print #13, "   stop_"
            End If
        
        Print #13,
        
        If excel_tag_dat(j, 42) > "" Then
            Print #13, "   _item_units.name          " + "'" + LCase(excel_tag_dat(j, 9)) + "'"
            Print #13, "   _item_units.code          " + "'" + LCase(excel_tag_dat(j, 42)) + "'"
        End If
        
        
        'print #13, "   _itme_range.name              "+lcase(excel_tag_dat(j,9))
        'If excel_tag_dat(j, 77) <> "?" Then Print #13, "   _item_range.minimum            " + "?"
        'If excel_tag_dat(j, 77) <> "?" Then Print #13, "   _item_range.maximum            " + "?"
        

         
        If excel_tag_dat(j, 35) = "Y" Then
            Print #13, "  loop_"
            Print #13, "      _item_linked.child_name"
            Print #13, "      _item_linked.parent_name"
            Print #13,
            
            hit_ct = 0
            For k = 1 To e_tag_ct
                If k <> j Then
                If excel_tag_dat(k, 30) <> "Y" Then
                If excel_tag_dat(k, 38) = excel_tag_dat(j, 79) Then
                If excel_tag_dat(k, 39) = excel_tag_dat(j, 80) Then
                    If lcase_flag = 1 Then Print #13, "  '" + LCase(excel_tag_dat(k, 9)) + "'";
                    If lcase_flag = 0 Then Print #13, "  '" + excel_tag_dat(k, 9) + "'";
                    ln = Len(excel_tag_dat(k, 9))
                    If lcase_flag = 1 Then
                        If ln < 45 Then Print #13, Space(53 - ln) + "'" + LCase(excel_tag_dat(j, 9)) + "'"
                        If ln > 44 Then Print #13, Space(85 - ln) + "'" + LCase(excel_tag_dat(j, 9)) + "'"
                    End If
                    If lcase_flag = 0 Then
                        If ln < 45 Then Print #13, Space(53 - ln) + "'" + excel_tag_dat(j, 9) + "'"
                        If ln > 44 Then Print #13, Space(85 - ln) + "'" + excel_tag_dat(j, 9) + "'"
                    End If
                    hit_ct = hit_ct + 1
                End If
                End If
                End If
                End If
            Next k
            If hit_ct = 0 Then Print #13, "   ?      ?"
            
            Print #13,
            Print #13, "  stop_"
            Print #13,
        End If
        
        
        If excel_tag_dat(j, 21) = "Y" Then
            Print #13, "   loop_"
            Print #13, "     _item_enumeration.value"
            Print #13, "     _item_enumeration.detail"
            Print #13,
            
            For k = 1 To enum_ct
                If enum_values(k, 1) = excel_tag_dat(j, 9) Then
                    Print #13, "   " + enum_values(k, 2) + "    " + enum_values(k, 3)
                End If
            Next k

            If enum_ct < 1 Then Print #13, "    ?     ?"
            Print #13,
            Print #13, "   stop_"
            Print #13,
        End If
        
        Print #13, "   loop_"
        Print #13, "     _item_examples.case"
        Print #13,
        Print #13, ";"
        Print #13, example
        Print #13, ";"
        Print #13,
        Print #13, "   stop_"
        Print #13,

        Print #13, "save_"
        Print #13,
    End If

Next j

End Sub

Sub write_nmrstar_dict4(e_tag_ct, excel_tag_dat, tag_ct, tag_list, e_col_num, pathin, output_files, lcase_flag, full_dict_flag, input_files, excel_header_dat)

Dim i, j, k, print_flag, link_ct, cat_ct, print_tags, print_tags_f As Integer
Dim enum_values(5000, 5), cat_desc(1000, 3) As String
Dim cat_table(120, 50), man_over(2000, 8) As String
Dim cif_link(5000, 3) As String
Dim t, ttt, test_str As String

'Open pathin + output_files(13) For Input As 5       'Load manual overide file
'While EOF(5) <> -1
'    man_over_ct = man_over_ct + 1
'    For i = 1 To 8
'        Input #5, man_over(man_over_ct, i)
'    Next i
'Wend
'Close 5

Open pathin + output_files(22) For Input As 5       'Load category, super category, category group descriptions
While cat_desc(cat_ct, 1) <> "Table_end"
    cat_ct = cat_ct + 1
    test_str = ""
    For i = 1 To 3
        Input #5, cat_desc(cat_ct, i)
    Next i
    'Debug.Print cat_ct, cat_desc(cat_ct, 1), cat_desc(cat_ct, 2), cat_desc(cat_ct, 3)
    'Debug.Print
    ln = Len(cat_desc(cat_ct, 3))
    For k = 1 To ln
        t = Mid$(cat_desc(cat_ct, 3), k, 1)
        If t = "$" Then t = ","
        test_str = test_str + t
    Next k
    cat_desc(cat_ct, 3) = test_str
Wend
Close 5

Open pathin + input_files(4) For Input As 5         'Load category group table
While cat_table(cat_tbl_ct, 1) <> "TBL_END"
    cat_tbl_ct = cat_tbl_ct + 1
    For i = 1 To 44
        Input #5, cat_table(cat_tbl_ct, i)
    Next i
Wend
Close 5

Open pathin + output_files(21) For Input As 5   'Load enumerations
While EOF(5) <> -1
    enum_ct = enum_ct + 1
    For i = 1 To 5
        Input #5, enum_values(enum_ct, i)
    Next i
    ln = Len(enum_values(enum_ct, 2))
    enum_values(enum_ct, 2) = Mid$(enum_values(enum_ct, 2), 2, ln - 2)
    ln = Len(enum_values(enum_ct, 3))
    enum_values(enum_ct, 3) = Mid$(enum_values(enum_ct, 3), 2, ln - 2)
    
    'Debug.Print enum_ct, enum_values(enum_ct, 1), enum_values(enum_ct, 2), enum_values(enum_ct, 3), enum_values(enum_ct, 4), enum_values(enum_ct, 5)
    'Debug.Print
'Wend
'Close 5

'Open "z:\"+"eldonulrich\"+"bmrb\htdocs\dictionary\htmldocs\nmr_star\development_adit_files\adit_input\adit_enum_dtl.csv" For Input As 5
'While EOF(5) <> -1
'    ct = ct + 1
'    Input #5, enum_values(ct, 1), enum_values(ct, 2), enum_values(ct, 3), enum_values(ct, 4)

    If enum_values(enum_ct, 3) = "enumeration_text" Then enum_values(enum_ct, 3) = "?"
    ln = Len(enum_values(enum_ct, 2))
    flag1 = 0: flag2 = 0: flag3 = 0
    For i = 1 To ln
        t = Mid$(enum_values(enum_ct, 2), i, 1)
        If t = " " Then flag1 = 1
        If t = Chr$(9) Then flag1 = 1
        If t = Chr$(91) Then flag1 = 1          ' chr [
        If t = Chr$(93) Then flag1 = 1          ' chr ]
        If t = Chr$(39) Then flag2 = 1
        If t = Chr$(96) Then flag2 = 1
        If t = Chr$(34) Then flag3 = 1
        If t = Chr$(36) Then
            enum_values(enum_ct, 2) = Left$(enum_values(enum_ct, 2), i - 1) + "," + Right$(enum_values(enum_ct, 2), ln - i)
        End If
    Next i
    If Left$(enum_values(enum_ct, 2), 1) = Chr$(34) And Right$(enum_values(enum_ct, 2), 1) = Chr$(34) Then flag3 = 4
    If Left$(enum_values(enum_ct, 2), 1) = Chr$(39) And Right$(enum_values(enum_ct, 2), 1) = Chr$(39) Then flag3 = 4
    If Left$(enum_values(enum_ct, 2), 1) = Chr$(96) And Right$(enum_values(enum_ct, 2), 1) = Chr$(96) Then flag3 = 4
    If Left$(enum_values(enum_ct, 2), 1) = Chr$(39) And Right$(enum_values(enum_ct, 2), 1) = Chr$(96) Then flag3 = 4
    If Left$(enum_values(enum_ct, 2), 1) = Chr$(96) And Right$(enum_values(enum_ct, 2), 1) = Chr$(39) Then flag3 = 4
    
    If flag1 = 1 Then
        If flag2 = 0 And flag3 = 0 Then
            enum_values(enum_ct, 2) = Chr$(39) + enum_values(enum_ct, 2) + Chr$(39)
        End If
        If flag2 = 1 And flag3 = 0 Then
            enum_values(enum_ct, 2) = Chr$(34) + enum_values(enum_ct, 2) + Chr$(34)
        End If
        If flag2 = 0 And flag3 = 1 Then
            enum_values(enum_ct, 2) = Chr$(39) + enum_values(enum_ct, 2) + Chr$(39)
        End If
        If flag2 = 1 And flag3 = 1 Then
            enum_values(enum_ct, 2) = Chr$(10) + Chr$(59) + enum_values(enum_ct, 2) + Chr$(10) + Chr$(59) + Chr$(10)
        End If
    End If
    ln = Len(enum_values(enum_ct, 3))
    flag1 = 0: flag2 = 0: flag3 = 0
    For i = 1 To ln
        t = Mid$(enum_values(enum_ct, 3), i, 1)
        If t = " " Then flag1 = 1
        If t = Chr$(9) Then flag1 = 1
        If t = Chr$(39) Then flag2 = 1
        If t = Chr$(96) Then flag2 = 1
        If t = Chr$(34) Then flag3 = 1
    Next i
    If Left$(enum_values(enum_ct, 3), 1) = Chr$(34) And Right$(enum_values(enum_ct, 3), 1) = Chr$(34) Then flag3 = 4
    If Left$(enum_values(enum_ct, 3), 1) = Chr$(39) And Right$(enum_values(enum_ct, 3), 1) = Chr$(39) Then flag3 = 4
    If Left$(enum_values(enum_ct, 3), 1) = Chr$(96) And Right$(enum_values(enum_ct, 3), 1) = Chr$(96) Then flag3 = 4
    If Left$(enum_values(enum_ct, 3), 1) = Chr$(39) And Right$(enum_values(enum_ct, 3), 1) = Chr$(96) Then flag3 = 4
    If Left$(enum_values(enum_ct, 3), 1) = Chr$(96) And Right$(enum_values(enum_ct, 3), 1) = Chr$(39) Then flag3 = 4
    If flag1 = 1 Then
        If flag2 = 0 And flag3 = 0 Then
            enum_values(enum_ct, 3) = Chr$(39) + enum_values(enum_ct, 3) + Chr$(39)
        End If
        If flag2 = 1 And flag3 = 0 Then
            enum_values(enum_ct, 3) = Chr$(34) + enum_values(enum_ct, 3) + Chr$(34)
        End If
        If flag2 = 0 And flag3 = 1 Then
            enum_values(enum_ct, 3) = Chr$(39) + enum_values(enum_ct, 3) + Chr$(39)
        End If
        If flag2 = 1 And flag3 = 1 Then
            enum_values(enum_ct, 3) = Chr$(10) + Chr$(59) + enum_values(enum_ct, 3) + Chr$(10) + Chr$(59) + Chr$(10)
        End If
    End If

    'Debug.Print enum_values(enum_ct, 1), enum_values(enum_ct, 2), enum_values(enum_ct, 3)
    'Debug.Print
Wend
Close 5

Open pathin + "nmr_cif_match.csv" For Input As 5
While EOF(5) <> -1
    link_ct = link_ct + 1
    Input #5, cif_link(link_ct, 1), cif_link(link_ct, 2), cif_link(link_ct, 3), t
    'Debug.Print cif_link(link_ct, 1), cif_link(link_ct, 2), cif_link(link_ct, 3)
    'Debug.Print
    If Val(cif_link(link_ct, 1)) < 1 Then link_ct = link_ct - 1
Wend
Close 5

For j = 1 To e_tag_ct
    If excel_tag_dat(j, 9) = "_Entry.NMR_STAR_version" Then
        version_num = excel_tag_dat(j, 77)
        Exit For
    End If
Next j
date_text1 = Date
p = InStr(1, date_text1, "/")
mon = Left$(date_text1, p - 1)
If Len(mon) = 1 Then mon = "0" + mon
p1 = InStr(p + 1, date_text1, "/")
da = Mid$(date_text1, p + 1, (p1 - p) - 1)
If Len(da) = 1 Then da = "0" + da
date_text2 = Right$(date_text1, 4) + "-" + mon + "-" + da

Open pathin + "NMR-STAR_STAR_header.txt" For Input As 5
t = "": ttt = ""
While EOF(5) <> -1
    ttt = Input$(1, 5)
    t1 = Asc(ttt)
    If t1 <> 10 And t1 <> 13 Then t = t + ttt
    If t1 = 10 Then
    p = InStr(1, t, "date_text2")
    If p > 0 Then t = Left$(t, p - 1) + date_text2
    p = InStr(1, t, "_dictionary.version")
    If p > 0 Then
        p = InStr(1, t, "version_num")
        t = Left$(t, p - 1) + version_num
    End If
    p = InStr(1, t, "stop_")
    Print #13, t
    t = ""
    End If
Wend

Print #13,
Close 5


'print list of category groups
Print #13, "     loop_"
Print #13, "       _category_group_list.id"
Print #13, "       _category_group_list.parent_id"
Print #13, "       _category_group_list.description"
Print #13,
Print #13, "'inclusive_group'"
Print #13, "'.'"
Print #13, ";"
Print #13, "Categories that belong to the NMR-STAR dictionary."
Print #13, ";"
Print #13,
For j = 1 To cat_ct
    If cat_desc(j, 1) = "category_group" Then
        print_flag = 0
        For i1 = 1 To e_tag_ct
            If LCase(cat_desc(j, 2)) = LCase(excel_tag_dat(i1, 2)) Then
            If excel_tag_dat(i1, 83) = "Y" Then
                print_flag = 1
                Exit For
            End If
            End If
        Next i1
        If print_flag = 1 Then
            If lcase_flag = 1 Then Print #13, "'" + LCase(cat_desc(j, 2)) + "'"
            If lcase_flag = 0 Then Print #13, "'" + (cat_desc(j, 2)) + "'"
            Print #13, "'" + "inclusive_group" + "'"
            Print #13, ";"
            Print #13, cat_desc(j, 3)
            Print #13, ";"
            Print #13,
        End If
    End If
Next j
Print #13, "     stop_"
Print #13,
Print #13,

'print out item ddl types and unit definitions
'Open pathin + "adit_input\item_type_units.txt" For Input As 5
'print_flg = 1
'While EOF(5) <> -1
'    t = Input$(1, 5)
'    t1 = Asc(t)
'    If t1 <> 10 And t1 <> 13 Then tt = tt + t
'    If t1 = 10 Then
'        Print #13, tt
'        tt = ""
'    End If
'Wend
'Print #13, tt
'Print #13,
'Close 5

For j = 1 To e_tag_ct
    print_flag = 0          'All tags begin with the print_flag set to 0
    'For i = 1 To tag_ct
    '    If excel_tag_dat(j, 9) = tag_list(i) Then
    '        print_flag = 0
    '        Exit For
    '    End If
    'Next i
    
    'If excel_tag_dat(j, 30) = "Y" Then print_flag = 0
    
    'If full_dict_flag = 1 Then print_flag = 1
    If excel_tag_dat(j, 83) = "Y" Then print_flag = 1   'Print if Common D&A flag is Y
    
    If print_flag = 1 Then
        For i = 1 To e_col_num
            If excel_tag_dat(j, i) = "" Then excel_tag_dat(j, i) = "?"
        Next i
            'set tag information values
        p = InStr(1, excel_tag_dat(j, 9), ".")
        category1 = Mid$(excel_tag_dat(j, 9), 2, p - 2)
        tag1 = Right$(excel_tag_dat(j, 9), Len(excel_tag_dat(j, 9)) - p)

        
        If Trim(excel_tag_dat(j, 38)) <> "" Then
            parent1 = "_" + excel_tag_dat(j, 38) + "." + excel_tag_dat(j, 39)
        Else
            parent1 = "."
        End If
        
        description1 = excel_tag_dat(j, 92)
        ln = Len(description1)
        If ln > 80 Then
            tt = "": k2 = 1
            For k1 = 1 To ln
                t = Mid$(description1, k1, 1)
                If t = "$" Then t = ","
                tt = tt + t
                
                If ln > k2 * 80 Then
                    If k1 > (k2 * 80) - 10 Then
                        If Asc(t) = 32 Then
                            tt = tt + Chr$(13) + Chr$(10)
                            k2 = k2 + 1
                        End If
                    End If
                End If
            Next k1
            description1 = tt
        End If
        
        If excel_tag_dat(j, 52) <> "?" Then
            prompt = excel_tag_dat(j, 52)
        Else
            prompt = "?"
        End If
        If excel_tag_dat(j, 51) <> "?" Then
            example = excel_tag_dat(j, 51)
            'Debug.Print example, excel_tag_dat(j, 87)
            If excel_tag_dat(j, 87) = "Date" Or excel_tag_dat(j, 87) = "yyyy-mm-dd" Then
            If Right$(example, 1) = "'" Then
                ln = Len(example)
                example = Left$(example, ln - 1)
                'Debug.Print example
                'Debug.Print
            End If
            End If
        Else
            example = "?"
        End If
        
        
        'print category group information
        If excel_tag_dat(j, 2) <> catgroup_last Then
            Print #13, "save_" + UCase(excel_tag_dat(j, 2))
            Print #13, "   _category_group.description"
            catg_desc_found = 0
            For k = 1 To cat_ct
                If excel_tag_dat(j, 2) = cat_desc(k, 2) Then
                If cat_desc(k, 1) = "category_group" Then
                If cat_desc(k, 3) <> "?" Then
                    catg_desc_found = 1
                    Print #13, ";"
                    Print #13, cat_desc(k, 3)
                    Print #13, ";"
                    Exit For
                End If
                End If
                End If
            Next k
            If catg_desc_found = 0 Then
                    Print #13, ";"
                    Print #13, "             Item category group description not available"
                    Print #13, ";"
            End If
            
            For k1 = 1 To cat_tbl_ct
                If excel_tag_dat(j, 2) = cat_table(k1, 5) Then kt = k1
            Next k1

            Print #13,
            Print #13, "   _category_group.id                    " + excel_tag_dat(j, 2)
            'Print #13, "   _category_group.name                  " + excel_tag_dat(j, 2)
           
            Print #13,
            Print #13, "save_"
            Print #13,
            catgroup_last = excel_tag_dat(j, 2)
        End If
        
        'print category information
        If excel_tag_dat(j, 79) <> cat_last Then
            cat_last = excel_tag_dat(j, 79)
            If lcase_flag = 1 Then Print #13, "save_" + LCase(category1)
            If lcase_flag = 0 Then Print #13, "save_" + category1
            Print #13, "   _category.description"
            cat_desc_found = 0
            For k = 1 To cat_ct
                If category1 = cat_desc(k, 2) Then
                If cat_desc(k, 1) = "item_category" Then
                If cat_desc(k, 3) <> "?" Then
                    cat_desc_found = 1
                    Print #13, ";"
                    Print #13, cat_desc(k, 3)
                    Print #13, ";"
                    Exit For
                End If
                End If
                End If
            Next k
            If cat_desc_found = 0 Then
                    Print #13, ";"
                    Print #13, "              category description not available"
                    Print #13, ";"
            End If
            If lcase_flag = 1 Then Print #13, "   _category.id                   '" + LCase(category1) + "'"
            If lcase_flag = 0 Then Print #13, "   _category.id                   '" + category1 + "'"
            
            If excel_tag_dat(j, 3) = "Y" Then print_text = "yes"
            If excel_tag_dat(j, 3) = "N" Then print_text = "no"
            If excel_tag_dat(j, 3) = "Y" And excel_tag_dat(j, 30) <> "Y" Then Print #13, "   _category.mandatory_code       " + print_text
            If excel_tag_dat(j, 30) = "Y" Then Print #13, "   _category.mandatory_code       no"
            Print #13,
            print_tags = 0
                
            If print_tags > 0 Then
                Print #13,
                Print #13, "   stop_"
                Print #13,
            End If
            
            Print #13, "   loop_"
            Print #13, "     _category_key.name"
            Print #13,
            
            key_found = 0
            For k = 1 To e_tag_ct
                If category1 = excel_tag_dat(k, 79) Then
                If excel_tag_dat(k, 36) = "Y" Then
                    If lcase_flag = 1 Then Print #13, "     " + "'" + LCase(excel_tag_dat(k, 9)) + "'"
                    If lcase_flag = 0 Then Print #13, "     " + "'" + excel_tag_dat(k, 9) + "'"
                    key_found = 1
                End If
                End If
            Next k
            
            If key_found = 0 Then Print #13, "      ?"
            
            Print #13,
            Print #13, "   stop_"
            Print #13,
            Print #13, "   loop_"
            Print #13, "     _category_group.id"
            Print #13,
            Print #13, "     'inclusive_group'"
            
            key_found = 0
            For k = 1 To e_tag_ct
                If category1 = excel_tag_dat(k, 79) Then
                    If lcase_flag = 1 Then Print #13, "     " + "'" + LCase(excel_tag_dat(k, 2)) + "'"
                    If lcase_flag = 0 Then Print #13, "     " + "'" + excel_tag_dat(k, 2) + "'"
                    key_found = 1
                    Exit For
                End If
            Next k
            
            Print #13,
            Print #13, "   stop_"
            Print #13,
            Print #13, "   loop_"
            Print #13, "     _category_examples.detail"
            Print #13, "     _category_examples.case"
            Print #13,
            Print #13, ";"
            Print #13, "   ?"
            Print #13, ";"
            Print #13, ";"
            Print #13, "   ?"
            Print #13, ";"
            Print #13,
            Print #13, "   stop_"
            Print #13,
            Print #13, "save_"
            Print #13,
        End If
        
        'print tag information
        
        If lcase_flag = 1 Then Print #13, "save_" + LCase(excel_tag_dat(j, 9))
        If lcase_flag = 0 Then Print #13, "save_" + excel_tag_dat(j, 9)
        
        Print #13,
'        Print #13, "   _item_description.name            '" + LCase(excel_tag_dat(j, 9)) + "'"
        Print #13, "   _item_description.description"
        Print #13, ";"
        Print #13, description1
        Print #13, ";"
        Print #13,
        
        If lcase_flag = 1 Then
            Print #13, "   _item.name                         '" + LCase(excel_tag_dat(j, 9)) + "'"
            Print #13, "   _item.category_id                  '" + LCase(category1) + "'"
        End If
        
        If lcase_flag = 0 Then
            Print #13, "   _item.name                         '" + excel_tag_dat(j, 9) + "'"
            Print #13, "   _item.category_id                  '" + category1 + "'"
        End If
        
        If excel_tag_dat(j, 29) = "NOT NULL" Then Print #13, "   _item.mandatory_code                yes"
        If excel_tag_dat(j, 29) <> "NOT NULL" Then Print #13, "   _item.mandatory_code                no"

        
        If excel_tag_dat(j, 77) = "" Then default_val = "." Else default_val = excel_tag_dat(j, 77)
        
        Print #13, "   _item.default_value                " + "'" + default_val + "'"
        If excel_tag_dat(j, 42) <> "?" Then Print #13, "   _item_units.code                  " + excel_tag_dat(j, 42)
        
        Print #13, "   _pdbx_item_enumeration_details.flag              " + excel_tag_dat(j, 21)
        If excel_tag_dat(j, 22) = "?" Then enum_closed_flag = "N"
        If excel_tag_dat(j, 22) = "Y" Then enum_closed_flag = "Y"
        Print #13, "   _pdbx_item_enumeration_details.closed_flag       " + enum_closed_flag
        
'        If excel_tag_dat(j, 43) = "Y" Then loop_flag = "yes"
'        If excel_tag_dat(j, 43) = "N" Then loop_flag = "no"
'        Print #13, "   _item.loop_flag                     " + loop_flag
'        Print #13,
        
        'print database information
        'Print #13, "   _item.star_flag                " + excel_tag_dat(j, 56)
        'Print #13, "   _item.db_flag                  " + excel_tag_dat(j, 57)
        'Print #13, "   _item.save_frame_id_flag       " + excel_tag_dat(j, 34)
        'Print #13, "   _item.non_public_flag          " + excel_tag_dat(j, 30)
        
'        Print #13, "   _item_type.name                    '" + LCase(excel_tag_dat(j, 9)) + "'"
        Print #13, "   _item_type.code                    " + "'" + excel_tag_dat(j, 87) + "'"
'        Print #13, "   _item_type.db_type_code             " + excel_tag_dat(j, 28)
        
'        Print #13,
'        Print #13, "   _item_db.name                      '" + LCase(excel_tag_dat(j, 9)) + "'"
'        If excel_tag_dat(j, 29) = "NOT NULL" Then Print #13, "   _item_db.table_mandatory_code      " + "'" + excel_tag_dat(j, 29) + "'"
'        If excel_tag_dat(j, 29) <> "NOT NULL" Then Print #13, "   _item_db.table_mandatory_code       ?"
'        Print #13, "   _item_db.table_name                " + "'" + excel_tag_dat(j, 79) + "'"
'        Print #13, "   _item_db.table_column_name         " + "'" + excel_tag_dat(j, 80) + "'"
'        Print #13, "   _item_db.table_primary_key          " + excel_tag_dat(j, 36)
        
'            If excel_tag_dat(j, 89) <> "?" And excel_tag_dat(j, 89) <> "" Then
'                Print #13, "   _item_db.table_foreign_key_flag     Y"
'                ReDim mate(30, 3) As String
'                mate_ct = 0
'                p = InStr(1, excel_tag_dat(j, 89), ";")
'                If p < 1 Then
'                    mate_ct = 1
'                    mate(mate_ct, 1) = excel_tag_dat(j, 89)
'                    mate(mate_ct, 2) = excel_tag_dat(j, 90)
'                    mate(mate_ct, 3) = excel_tag_dat(j, 91)
'                End If
'                If p > 0 Then
'                    mate_ct = 1
'                    mate(mate_ct, 1) = Left$(excel_tag_dat(j, 89), p - 1)
'                    p1 = 1
'                    While p1 > 0
'                        p1 = InStr(p + 1, excel_tag_dat(j, 89), ";")
'                        If p1 > 0 Then
'                            mate_ct = mate_ct + 1
'                            mate(mate_ct, 1) = Mid$(excel_tag_dat(j, 89), p + 1, (p1 - p) - 1)
'                            p = p1
'                        End If
'                    Wend
'                    mate_ct = mate_ct + 1
'                    mate(mate_ct, 1) = Right$(excel_tag_dat(j, 89), Len(excel_tag_dat(j, 89)) - p)
'                End If
'                p = InStr(1, excel_tag_dat(j, 90), ";")
'                If p > 0 Then
'                    mate_ct = 1
'                    mate(mate_ct, 2) = Left$(excel_tag_dat(j, 90), p - 1)
'                    p1 = 1
'                    While p1 > 0
'                        p1 = InStr(p + 1, excel_tag_dat(j, 90), ";")
'                        If p1 > 0 Then
'                            mate_ct = mate_ct + 1
'                            mate(mate_ct, 2) = Mid$(excel_tag_dat(j, 90), p + 1, (p1 - p) - 1)
'                            p = p1
'                        End If
'                    Wend
'                    mate_ct = mate_ct + 1
'                    mate(mate_ct, 2) = Right$(excel_tag_dat(j, 90), Len(excel_tag_dat(j, 90)) - p)
'                End If
'                p = InStr(1, excel_tag_dat(j, 91), ";")
'                If p > 0 Then
'                    mate_ct = 1
'                    mate(mate_ct, 3) = Left$(excel_tag_dat(j, 91), p - 1)
'                    p1 = 1
'                    While p1 > 0
'                        p1 = InStr(p + 1, excel_tag_dat(j, 91), ";")
'                        If p1 > 0 Then
'                            mate_ct = mate_ct + 1
'                            mate(mate_ct, 3) = Mid$(excel_tag_dat(j, 91), p + 1, (p1 - p) - 1)
'                            p = p1
'                        End If
'                    Wend
'                    mate_ct = mate_ct + 1
'                    mate(mate_ct, 3) = Right$(excel_tag_dat(j, 91), Len(excel_tag_dat(j, 91)) - p)
'                End If
'
'        Print #13,
'        Print #13, "   loop_"
'        Print #13, "      _item_db_foreign_key.group_ID"
'        Print #13, "      _item_db_foreign_key.related_table_name"
'        Print #13, "      _item_db_foreign_key.related_table_column_name"
'        Print #13,
'        For i3 = 1 To mate_ct
'            Print #13, "   " + mate(i3, 1) + Space(8 - Len(mate(i3, 1))) + mate(i3, 2) + Space(37 - Len(mate(i3, 2))) + mate(i3, 3)
'        Next i3
'        Print #13,
'        Print #13, "   stop_"
'            End If
        
'        Print #13,
        
        If excel_tag_dat(j, 42) > "" Then
'            Print #13, "   _item_units.name          " + "'" + LCase(excel_tag_dat(j, 9)) + "'"
            Print #13, "   _item_units.code                   " + "'" + LCase(excel_tag_dat(j, 42)) + "'"
        End If
        
        
        'print #13, "   _itme_range.name              "+lcase(excel_tag_dat(j,9))
        'If excel_tag_dat(j, 77) <> "?" Then Print #13, "   _item_range.minimum            " + "?"
        'If excel_tag_dat(j, 77) <> "?" Then Print #13, "   _item_range.maximum            " + "?"
        

         
        If excel_tag_dat(j, 35) = "Y" Then
            Print #13,
            Print #13, "  loop_"
            Print #13, "      _item_linked.child_name"
            Print #13, "      _item_linked.parent_name"
            Print #13,
            
            hit_ct = 0
            For k = 1 To e_tag_ct
                If k <> j Then
                If excel_tag_dat(k, 30) <> "Y" Then
                If excel_tag_dat(k, 38) = excel_tag_dat(j, 79) Then
                If excel_tag_dat(k, 39) = excel_tag_dat(j, 80) Then
                    If lcase_flag = 1 Then Print #13, "  '" + LCase(excel_tag_dat(k, 9)) + "'";
                    If lcase_flag = 0 Then Print #13, "  '" + excel_tag_dat(k, 9) + "'";
                    ln = Len(excel_tag_dat(k, 9))
                    If lcase_flag = 1 Then
                        If ln < 45 Then Print #13, Space(53 - ln) + "'" + LCase(excel_tag_dat(j, 9)) + "'"
                        If ln > 44 Then Print #13, Space(85 - ln) + "'" + LCase(excel_tag_dat(j, 9)) + "'"
                    End If
                    If lcase_flag = 0 Then
                        If ln < 45 Then Print #13, Space(53 - ln) + "'" + excel_tag_dat(j, 9) + "'"
                        If ln > 44 Then Print #13, Space(85 - ln) + "'" + excel_tag_dat(j, 9) + "'"
                    End If
                    hit_ct = hit_ct + 1
                End If
                End If
                End If
                End If
            Next k
            If hit_ct = 0 Then Print #13, "   ?      ?"
            
            Print #13,
            Print #13, "  stop_"
            Print #13,
        End If
        
        
        If excel_tag_dat(j, 21) = "Y" Then
            Print #13,
            Print #13, "   loop_"
            Print #13, "     _item_enumeration.value"
            Print #13, "     _item_enumeration.detail"
            Print #13,
            
            For k = 1 To enum_ct
                If enum_values(k, 1) = excel_tag_dat(j, 9) Then
                    Print #13, "   " + enum_values(k, 2) + "    " + enum_values(k, 3)
                End If
            Next k

            If enum_ct < 1 Then Print #13, "    ?     ?"
            Print #13,
            Print #13, "   stop_"
            Print #13,
        End If
        
        Print #13,
        Print #13, "   loop_"
        Print #13, "     _item_examples.case"
        Print #13,
        Print #13, ";"
        Print #13, example
        Print #13, ";"
        Print #13,
        Print #13, "   stop_"
        Print #13,

        Print #13, "save_"
        Print #13,
    End If

Next j

End Sub
Sub write_nmrstar_dict2(e_tag_ct, excel_tag_dat, tag_ct, tag_list, e_col_num, pathin, output_files, lcase_flag, full_dict_flag, dictionary_name)

Dim i, j, k, print_flag, link_ct, cat_ct As Integer
Dim enum_values(5000, 5), cat_desc(4000, 3) As String
Dim super_grp(20, 2), cat_grp(200, 2)
Dim cif_link(5000, 3) As String
Dim t, ttt, test_str As String
Dim id(10000, 2) As Integer      ' 1 = item_ID; 2 = category_ID for item


'LOAD ALL REQUIRED DATA INTO MEMORY

Open pathin + output_files(22) For Input As 5   'open and load "BMOD_category_desc.csv"
While cat_desc(cat_ct, 1) <> "Table_end"
    cat_ct = cat_ct + 1
    test_str = ""
    For i = 1 To 3
        Input #5, cat_desc(cat_ct, i)
    Next i
    'Debug.Print cat_ct, cat_desc(cat_ct, 1), cat_desc(cat_ct, 2), cat_desc(cat_ct, 3)
    'Debug.Print
    ln = Len(cat_desc(cat_ct, 3))
    For k = 1 To ln
        t = Mid$(cat_desc(cat_ct, 3), k, 1)
        If t = "$" Then t = ","
        test_str = test_str + t
    Next k
    cat_desc(cat_ct, 3) = test_str
Wend
Close 5

Open pathin + output_files(21) For Input As 5   'open  and load "BMOD_enum_dict.csv"
While EOF(5) <> -1
    enum_ct = enum_ct + 1
    For i = 1 To 5
        Input #5, enum_values(enum_ct, i)
    Next i
    ln = Len(enum_values(enum_ct, 2))
    enum_values(enum_ct, 2) = Mid$(enum_values(enum_ct, 2), 2, ln - 2)
    ln = Len(enum_values(enum_ct, 3))
    enum_values(enum_ct, 3) = Mid$(enum_values(enum_ct, 3), 2, ln - 2)
    ln = Len(enum_values(enum_ct, 3))
    For i23 = 1 To ln
        If Mid$(enum_values(enum_ct, 3), i23, 1) = "$" Then
            enum_values(enum_ct, 3) = Left$(enum_values(enum_ct, 3), i23 - 1) + "," + Right$(enum_values(enum_ct, 3), ln - i23)
            'Debug.Print enum_values(enum_ct, 3)
            'Debug.Print
        End If
    Next i23

    If enum_values(enum_ct, 3) = "enumeration_text" Then enum_values(enum_ct, 3) = "?"
    ln = Len(enum_values(enum_ct, 2))
    flag1 = 0: flag2 = 0: flag3 = 0
    For i = 1 To ln
        t = Mid$(enum_values(enum_ct, 2), i, 1)
        If t = " " Then flag1 = 1
        If t = Chr$(9) Then flag1 = 1
        If t = Chr$(91) Then flag1 = 1          ' chr [
        If t = Chr$(93) Then flag1 = 1          ' chr ]
        If t = Chr$(39) Then flag2 = 1
        If t = Chr$(96) Then flag2 = 1
        If t = Chr$(34) Then flag3 = 1
        If t = Chr$(36) Then
            enum_values(enum_ct, 2) = Left$(enum_values(enum_ct, 2), i - 1) + "," + Right$(enum_values(enum_ct, 2), ln - i)
        End If
    Next i
    If Left$(enum_values(enum_ct, 2), 1) = Chr$(34) And Right$(enum_values(enum_ct, 2), 1) = Chr$(34) Then flag3 = 4
    If Left$(enum_values(enum_ct, 2), 1) = Chr$(39) And Right$(enum_values(enum_ct, 2), 1) = Chr$(39) Then flag3 = 4
    If Left$(enum_values(enum_ct, 2), 1) = Chr$(96) And Right$(enum_values(enum_ct, 2), 1) = Chr$(96) Then flag3 = 4
    If Left$(enum_values(enum_ct, 2), 1) = Chr$(39) And Right$(enum_values(enum_ct, 2), 1) = Chr$(96) Then flag3 = 4
    If Left$(enum_values(enum_ct, 2), 1) = Chr$(96) And Right$(enum_values(enum_ct, 2), 1) = Chr$(39) Then flag3 = 4
    
    If flag1 = 1 Then
        If flag2 = 0 And flag3 = 0 Then
            enum_values(enum_ct, 2) = Chr$(39) + enum_values(enum_ct, 2) + Chr$(39)
        End If
        If flag2 = 1 And flag3 = 0 Then
            enum_values(enum_ct, 2) = Chr$(34) + enum_values(enum_ct, 2) + Chr$(34)
        End If
        If flag2 = 0 And flag3 = 1 Then
            enum_values(enum_ct, 2) = Chr$(39) + enum_values(enum_ct, 2) + Chr$(39)
        End If
        If flag2 = 1 And flag3 = 1 Then
            enum_values(enum_ct, 2) = Chr$(10) + Chr$(59) + enum_values(enum_ct, 2) + Chr$(10) + Chr$(59) + Chr$(10)
        End If
    End If
    ln = Len(enum_values(enum_ct, 3))
    flag1 = 0: flag2 = 0: flag3 = 0
    For i = 1 To ln
        t = Mid$(enum_values(enum_ct, 3), i, 1)
        If t = " " Then flag1 = 1
        If t = Chr$(9) Then flag1 = 1
        If t = Chr$(39) Then flag2 = 1
        If t = Chr$(96) Then flag2 = 1
        If t = Chr$(34) Then flag3 = 1
    Next i
    If Left$(enum_values(enum_ct, 3), 1) = Chr$(34) And Right$(enum_values(enum_ct, 3), 1) = Chr$(34) Then flag3 = 4
    If Left$(enum_values(enum_ct, 3), 1) = Chr$(39) And Right$(enum_values(enum_ct, 3), 1) = Chr$(39) Then flag3 = 4
    If Left$(enum_values(enum_ct, 3), 1) = Chr$(96) And Right$(enum_values(enum_ct, 3), 1) = Chr$(96) Then flag3 = 4
    If Left$(enum_values(enum_ct, 3), 1) = Chr$(39) And Right$(enum_values(enum_ct, 3), 1) = Chr$(96) Then flag3 = 4
    If Left$(enum_values(enum_ct, 3), 1) = Chr$(96) And Right$(enum_values(enum_ct, 3), 1) = Chr$(39) Then flag3 = 4
    If flag1 = 1 Then
        If flag2 = 0 And flag3 = 0 Then
            enum_values(enum_ct, 3) = Chr$(39) + enum_values(enum_ct, 3) + Chr$(39)
        End If
        If flag2 = 1 And flag3 = 0 Then
            enum_values(enum_ct, 3) = Chr$(34) + enum_values(enum_ct, 3) + Chr$(34)
        End If
        If flag2 = 0 And flag3 = 1 Then
            enum_values(enum_ct, 3) = Chr$(39) + enum_values(enum_ct, 3) + Chr$(39)
        End If
        If flag2 = 1 And flag3 = 1 Then
            enum_values(enum_ct, 3) = Chr$(10) + Chr$(59) + enum_values(enum_ct, 3) + Chr$(10) + Chr$(59) + Chr$(10)
        End If
    End If

    'Debug.Print enum_values(enum_ct, 1), enum_values(enum_ct, 2), enum_values(enum_ct, 3)
    'Debug.Print
Wend
Close 5

' SET UP ITEM IDS AND ITEM CATEGORY IDS

cat_last = ""
cat_count = 0
For k = 1 To e_tag_ct
    id(k, 1) = k
        If excel_tag_dat(k, 24) <> cat_last Then
            cat_last = excel_tag_dat(k, 24)
            cat_count = cat_count + 1
        End If
    id(k, 2) = cat_count
Next k

'BEGIN PRINTING DICTIONARY FILE

For j = 1 To e_tag_ct
    If excel_tag_dat(j, 7) = "_Entry.BMOD_STAR_version" Then
        version_num = excel_tag_dat(j, 23)
        Exit For
    End If
Next j
date_text1 = Date
p = InStr(1, date_text1, "/")
mon = Left$(date_text1, p - 1)
If Len(mon) = 1 Then mon = "0" + mon
p1 = InStr(p + 1, date_text1, "/")
da = Mid$(date_text1, p + 1, (p1 - p) - 1)
If Len(da) = 1 Then da = "0" + da
date_text2 = Right$(date_text1, 4) + "-" + mon + "-" + da

'load the dictionary header file and insert text
'The dictionary history will be updated in the dictionary header file
'manually

Open pathin + "NMR-STAR_STAR_header.txt" For Input As 5
t = "": ttt = ""
While EOF(5) <> -1
    ttt = Input$(1, 5)
    t1 = Asc(ttt)
    If t1 <> 10 And t1 <> 13 Then t = t + ttt
    If t1 = 10 Then
        p = InStr(1, t, "date_text2")
        If p > 0 Then t = Left$(t, p - 1) + date_text2
        p = InStr(1, t, "_Dictionary.Version")
        If p > 0 Then
            p = InStr(1, t, "version_num")
            t = Left$(t, p - 1) + version_num
        End If
        Print #13, t
        t = ""
    End If
Wend

Close 5

Print #13,
Print #13, "######################################################################"
Print #13,

'PRINT LIST OF SUPER SAVE FRAME GROUPS

Print #13, "     loop_"
Print #13, "       _Super_save_frame_group.ID"
Print #13, "       _Super_save_frame_group.Name"
Print #13, "       _Super_save_frame_group.Description"
Print #13, "       _Super_save_frame_group.Dictionary_ID"
Print #13,
super_cat_grp_ct = 0
For j = 1 To cat_ct
    If cat_desc(j, 1) = "super_group" Then
        super_cat_grp_ct = super_cat_grp_ct + 1
        super_grp(super_cat_grp_ct, 1) = super_cat_grp_ct
        super_grp(super_cat_grp_ct, 2) = cat_desc(j, 2)
        Print #13, "   " + Str(super_cat_grp_ct) + "     ";
        Print #13, cat_desc(j, 2)
        Print #13, ";"
        Print #13, cat_desc(j, 3)
        Print #13, ";"
        Print #13, dictionary_name
        Print #13,
    End If
Next j
Print #13, "     stop_"
Print #13,
Print #13, "######################################################################"
Print #13,

'PRINT LIST OF ALLOWED SAVE FRAME CATEGORIES (DIFFERENT FOR DDL AND DICTIONARY)

cat_grp_ct = 0
Print #13, "     loop_"
Print #13, "       _Save_frame_category.Name"
Print #13, "       _Save_frame_category.Description"
Print #13, "       _Save_frame_category.Super_save_frame_group_ID"
Print #13, "       _Save_frame_category.Super_save_frame_group_name"
Print #13, "       _Save_frame_category.Dictionary_ID"
Print #13,
For j = 1 To cat_ct
    If cat_desc(j, 1) = "category_group" Then
        cat_grp_ct = cat_grp_ct + 1
        cat_grp(cat_grp_ct, 1) = cat_grp_ct
        cat_grp(cat_grp_ct, 2) = cat_desc(j, 2)
        
        'Print #13, "    " + Str(cat_grp_ct) + "     ";
        If lcase_flag = 1 Then Print #13, "'" + LCase(cat_desc(j, 2)) + "'"
        If lcase_flag = 0 Then Print #13, "'" + (cat_desc(j, 2)) + "'"
        Print #13, ";"
        Print #13, cat_desc(j, 3)
        Print #13, ";"
        Print #13,
        For k3 = 1 To e_tag_ct
            If excel_tag_dat(k3, 2) = cat_desc(j, 2) Then
                For k4 = 1 To super_cat_grp_ct
                    If super_grp(k4, 2) = excel_tag_dat(k3, 4) Then
                        Print #13, "    " + Str(super_grp(k4, 1));
                   End If
                Next k4
                Print #13, "     " + "'" + excel_tag_dat(k3, 4) + "'";
                Exit For
            End If
        Next k3
        Print #13, "     " + dictionary_name
        Print #13,
    End If
Next j
Print #13, "     stop_"
Print #13,
Print #13, "######################################################################"
Print #13,

'PRINT CATEGORY GROUP LIST

cat_grp_ct = 0
Print #13, "     loop_"
Print #13, "       _Category_group.ID"
Print #13, "       _Category_group.Name"
'Print #13, "       _Category_group_list.Parent_id"  'PARENT GROUP NOT USED CURRENTLY
Print #13, "       _Category_group.Description"
Print #13, "       _Category_group.Dictionary_ID"
Print #13,
'Print #13, "     1         'inclusive_group'      ?      ?"        'INCLUSIVE GROUP IS NOT USED CURRENTLY AS A CATEGORY GROUP
'Print #13, ";"
'Print #13, "             Categories that belong to the BMOD-STAR dictionary."
'Print #13, ";"
'Print #13, dictionary_name
'Print #13,
For j = 1 To cat_ct
    If cat_desc(j, 1) = "category_group" Then
        cat_grp_ct = cat_grp_ct + 1
        cat_grp(cat_grp_ct, 1) = cat_grp_ct
        cat_grp(cat_grp_ct, 2) = cat_desc(j, 2)
        
        Print #13, "    " + Str(cat_grp_ct) + "     ";
        If lcase_flag = 1 Then Print #13, "'" + LCase(cat_desc(j, 2)) + "'"
        If lcase_flag = 0 Then Print #13, "'" + (cat_desc(j, 2)) + "'"
       ' For k3 = 1 To e_tag_ct               'USED TO INCLUDE SUPER GROUP LINKS IF INFO PLACED IN THIS TABLE - NOT CURRENTLY
       '     If excel_tag_dat(k3, 2) = cat_desc(j, 2) Then
       '         For k4 = 1 To super_cat_grp_ct
       '             If super_grp(k4, 2) = excel_tag_dat(k3, 4) Then
       '                 Print #13, "    " + Str(super_grp(k4, 1));
       '             End If
       '         Next k4
       '         Print #13, "     " + "'" + excel_tag_dat(k3, 4) + "'"
       '         Exit For
       '     End If
       ' Next k3
'        Print #13, "'" + "inclusive_group" + "'"
        Print #13, ";"
        Print #13, cat_desc(j, 3)
        Print #13, ";"
        Print #13, dictionary_name
        Print #13,
    End If
Next j
Print #13, "     stop_"
Print #13,
Print #13, "######################################################################"
Print #13,

'PRINT ALLOWED ITEM DDL TYPES AND UNIT DEFINITIONS TAKEN FROM EXTERNAL TEXT FILE

Open pathin + "item_type_units.txt" For Input As 5
print_flg = 1
While EOF(5) <> -1
    t = Input$(1, 5)
    t1 = Asc(t)
    If t1 <> 10 And t1 <> 13 Then tt = tt + t
    If t1 = 10 Then
        Print #13, tt
        tt = ""
    End If
Wend

Close 5

'PRINT CATEGORY FOLLOWED BY CATEGORY ITEMS - THE ITEM DICTIONARY (SOURCE LARGE EXCEL FILE)

cat_count = 0
cat_last = ""
For j = 1 To e_tag_ct
    print_flag = 1
    For i = 1 To tag_ct
        If excel_tag_dat(j, 9) = tag_list(i) Then
            print_flag = 0
            Exit For
        End If
    Next i
    'If excel_tag_dat(j, 30) = "Y" Then print_flag = 0
    If full_dict_flag = 1 Then print_flag = 1
    
    'PREP AND SETUP INFORMATION EXTRACTED FROM LARGE EXCEL FILE
    
    If print_flag = 1 Then
        For i = 1 To e_col_num
            If excel_tag_dat(j, i) = "" Then excel_tag_dat(j, i) = "?"
        Next i
            'set tag information values
        p = InStr(1, excel_tag_dat(j, 9), ".")
        category1 = Mid$(excel_tag_dat(j, 9), 2, p - 2)
        tag1 = Right$(excel_tag_dat(j, 9), Len(excel_tag_dat(j, 9)) - p)
       
            description1 = excel_tag_dat(j, 92)
            ln = Len(description1)
            If ln > 80 Then
                tt = "": k2 = 1
                For k1 = 1 To ln
                    t = Mid$(description1, k1, 1)
                    If t = "$" Then t = ","
                    tt = tt + t
                
                    If ln > k2 * 80 Then
                        If k1 > (k2 * 80) - 10 Then
                            If Asc(t) = 32 Then
                                tt = tt + Chr$(13) + Chr$(10)
                                k2 = k2 + 1
                            End If
                        End If
                    End If
                Next k1
                description1 = tt
            End If
        
        If excel_tag_dat(j, 19) <> "?" Then
            example = excel_tag_dat(j, 51)
            'Debug.Print example, excel_tag_dat(j, 87)
            If excel_tag_dat(j, 87) = "Date" Or excel_tag_dat(j, 87) = "yyyy-mm-dd" Then
            If Right$(example, 1) = "'" Then
                ln = Len(example)
                example = Left$(example, ln - 1)
                'Debug.Print example
                'Debug.Print
            End If
            End If
        Else
            example = "?"
        End If
                
        'PRINT CATEGORY INFORMATION SAVE FRAME
        
        If excel_tag_dat(j, 24) <> cat_last Then
            cat_last = excel_tag_dat(j, 79)
            If lcase_flag = 1 Then Print #13, "save_" + LCase(category1)
            If lcase_flag = 0 Then Print #13, "save_" + category1
            cat_count = cat_count + 1
            Print #13, "   _Category.Sf_category                category_description"
            Print #13, "   _Category.Sf_framecode               " + category1
            Print #13, "   _Category.ID                        " + Str(cat_count)
            If lcase_flag = 1 Then Print #13, "   _Category.Name                      '" + LCase(category1) + "'"
            If lcase_flag = 0 Then Print #13, "   _Category.Name                      '" + category1 + "'"
'            Print #13, "   _category.name                 '" + category1 + "'"
            
            Print #13, "   _category.Description"
            cat_desc_found = 0
            For k = 1 To cat_ct
                If category1 = cat_desc(k, 2) Then
                If cat_desc(k, 1) = "item_category" Then
                If cat_desc(k, 3) <> "?" Then
                    cat_desc_found = 1
                    Print #13, ";"
                    Print #13, cat_desc(k, 3)
                    Print #13, ";"
                    Exit For
                End If
                End If
                End If
            Next k
            If cat_desc_found = 0 Then
                    Print #13, ";"
                    Print #13, "              category description not available"
                    Print #13, ";"
            End If
            If excel_tag_dat(j, 29) = "NOT NULL" Then Print #13, "   _Category.Mandatory_code             yes"
            If excel_tag_dat(j, 29) = "" Then Print #13, "   _Category.Mandatory_code             no"
            If excel_tag_dat(j, 43) = "Y" Then Print #13, "   _Category.Loop_code                  yes"
            If excel_tag_dat(j, 43) = "N" Then Print #13, "   _Category.Loop_code                  no"
            For k3 = 1 To cat_grp_ct
                If excel_tag_dat(j, 2) = cat_grp(k3, 2) Then
                    Print #13, "   _Category.Category_group_ID         " + Str(cat_grp(k3, 1))
                End If
            Next k3
            
            Print #13, "   _Category.Category_group_name        " + excel_tag_dat(j, 2)
            Print #13, "   _Category.Dictionary_ID              " + dictionary_name
            Print #13,
            
            'PRINT CATEGORY PRIMARY KEY ITEMS
            
            Print #13, "   loop_"
            Print #13, "     _Category_primary_key.Item_ID"
            Print #13, "     _Category_primary_key.Item_label"
            Print #13, "     _Category_primary_key.Item_name"
            Print #13, "     _Category_primary_key.Category_ID"
            Print #13, "     _Category_primary_key.Dictionary_ID"
            Print #13,
            key_found = 0
            For k = 1 To e_tag_ct
                If category1 = excel_tag_dat(k, 79) Then
                If excel_tag_dat(k, 36) = "Y" Then
                    Print #13, "   " + Str(k);
                    Print #13, "   " + "$" + excel_tag_dat(k, 9);
                    If lcase_flag = 1 Then Print #13, "     " + "'" + LCase(excel_tag_dat(k, 9)) + "'";
                    If lcase_flag = 0 Then Print #13, "     " + "'" + excel_tag_dat(k, 9) + "'";
                    key_found = 1
                End If
                End If
            Next k
            If key_found = 0 Then Print #13, "      ?    ?    ?";
            Print #13, "   " + Str(cat_count);
            Print #13, "   " + dictionary_name
            Print #13,
            Print #13, "   stop_"
            
            'THIS IS WHERE THE CATEGORY FOREIGN KEYS WOULD BE PRINTED
            
            'CURRENTLY A CATEGORY CAN BE A MEMBER OF ONLY ONE CATEGORY GROUP AND SO THE FOLLOWING CODE IS NOT USED
            
'            Print #13,
'            Print #13, "   loop_"
'            Print #13, "     _Category_group.ID"
'            Print #13, "     _Category_group.Category_ID"
'            Print #13,
'            key_found = 0
'            For k = 1 To e_tag_ct
'                If category1 = excel_tag_dat(k, 79) Then
'                    If lcase_flag = 1 Then Print #13, "     " + "'" + LCase(excel_tag_dat(k, 2)) + "'";
'                    If lcase_flag = 0 Then Print #13, "     " + "'" + excel_tag_dat(k, 2) + "'";
'                    Print #13, "       "; Str(cat_count)
'                    key_found = 1
'                    Exit For
'                End If
'            Next k
            'If key_found = 0 Then Print #13, "      ?"
'            Print #13,
'            Print #13, "   stop_"
            Print #13,
            Print #13, "save_"
            Print #13,
        End If
'        End If
        
        'PRINT ITEM (TAG) INFORMATION
        
        If lcase_flag = 1 Then Print #13, "save_" + LCase(excel_tag_dat(j, 9))
        If lcase_flag = 0 Then Print #13, "save_" + excel_tag_dat(j, 9)
        item_count = item_count + 1

        Print #13, "   _Item.Sf_category             item_description"
        Print #13, "   _Item.Sf_framecode           " + "'" + excel_tag_dat(j, 9) + "'"
        Print #13, "   _Item.ID                     " + Str(item_count)
        Print #13, "   _Item.Name                   '" + excel_tag_dat(j, 9) + "'"
        Print #13, "   _Item.Description"
        Print #13, ";"
        ln = Len(description1)
        For i = 1 To ln
            If Mid$(description1, i, 1) = "$" Then
                description1 = Left$(description1, i - 1) + "," + Right$(description1, ln - i)
                'Debug.Print description1
                'Debug.Print
            End If
        Next i
        Print #13, description1
        'Debug.Print description1
        'Debug.Print
        Print #13, ";"
        Print #13,

        If lcase_flag = 1 Then
            Print #13, "   _Item.Category_ID                           " + Str(cat_count)
            Print #13, "   _Item.Category_label                        " + "$" + LCase(category1) + "'"
        End If
        If lcase_flag = 0 Then
            Print #13, "   _Item.Category_ID                           " + Str(cat_count)
            Print #13, "   _Item.Category_label                        " + "$" + category1
        End If

        If excel_tag_dat(j, 29) = "NOT NULL" Then Print #13, "   _Item.Mandatory_code                         yes"
        If excel_tag_dat(j, 29) <> "NOT NULL" Then Print #13, "   _Item.Mandatory_code                         no"
        
        Print #13, "   _Item.Type_code                             " + "'" + excel_tag_dat(j, 87) + "'"
        
        'If excel_tag_dat(j, 42) <> "?" Then Print #13, "   _Item.Units_code              " + excel_tag_dat(j, 42)
        
        If excel_tag_dat(j, 77) <> "?" Then Print #13, "   _Item.Default_value                         " + "'" + excel_tag_dat(j, 23) + "'"
        
        If excel_tag_dat(j, 21) = "Y" Then Print #13, "   _Item.Enumerated_code                        yes"
        If excel_tag_dat(j, 22) = "N" Then Print #13, "   _Item.Enumeration_closed_code                no"
        If excel_tag_dat(j, 22) = "Y" Then Print #13, "   _Item.Enumeration_closed_code                yes"
        
        'ITEM RANGE INFORMATION NEEDS TO BE LOOPED
        
        'print #13, "  loop_"
        'If excel_tag_dat(j, 23) <> "?" Then Print #13, "   _item_range.minimum            " + "?"
        'If excel_tag_dat(j, 23) <> "?" Then Print #13, "   _item_range.maximum            " + "?"
        'print #13, "  stop_"
        
        Print #13,
    
    'FOLLOWING CODE USED TO FIND AND PRINT CHILDREN OF THE TAG
    
    If excel_tag_dat(j, 35) = "Y" Then
            Print #13, "  loop_"
            Print #13, "      _Item_linked.Child_item_ID"
            Print #13, "      _Item_linked.Child_item_label"
            Print #13, "      _Item_linked.Child_item_name"
            Print #13, "      _Item_linked.Child_item_category_ID"
            Print #13, "      _Item_linked.Child_item_category_label"
            Print #13, "      _Item_linked.Item_ID"
            Print #13, "      _Item_linked.Item_name"
            Print #13, "      _Item_linked.Category_ID"
            Print #13, "      _Item_linked.Dictionary_ID"
            Print #13,
        
            hit_ct = 0
            For k = 1 To e_tag_ct
                If k <> j Then
                If excel_tag_dat(k, 15) = excel_tag_dat(j, 79) Then   'category of test tag pointer against category of parent tag
                If excel_tag_dat(k, 16) = excel_tag_dat(j, 80) Then   'item of test tag pointer against item of parent tag
                    Print #13, "   " + Str(k);
                    If lcase_flag = 1 Then Print #13, "  $" + LCase(excel_tag_dat(k, 79));
                    If lcase_flag = 0 Then Print #13, "  $" + excel_tag_dat(k, 9);
                    If lcase_flag = 1 Then Print #13, "  '" + LCase(excel_tag_dat(k, 9)) + "'";
                    If lcase_flag = 0 Then Print #13, "  '" + excel_tag_dat(k, 9) + "'";
                    Print #13, "   " + Str(id(k, 2));
                    If lcase_flag = 1 Then Print #13, "  $" + LCase(excel_tag_dat(k, 79));
                    If lcase_flag = 0 Then Print #13, "  $" + excel_tag_dat(k, 79);
                    Print #13, "   " + Str(id(j, 1));
                    If lcase_flag = 1 Then Print #13, "  '" + LCase(excel_tag_dat(j, 9)) + "'";
                    If lcase_flag = 0 Then Print #13, "  '" + excel_tag_dat(j, 9) + "'";
                    Print #13, "   " + Str(cat_count);
                    Print #13, "   " + dictionary_name

                    hit_ct = 1
                End If
                End If
                End If
            Next k
            If hit_ct = 0 Then Print #13, "   ?   ?   ?   ?   ?   ?   ?   ?   ?"

            Print #13,
            Print #13, "  stop_"
            Print #13,

        End If
        
' THIS IS WHERE TAG ALIASES WOULD BE PRINTED OUT IF IMPLEMENTED
' CODE FOR THIS DOES NOT EXIST

        'Print #13, "   loop_"
        'Print #13, "     _item_aliases.Item_name"
        'Print #13, "     _item_aliases.Alias_item_name"
        'Print #13, "     _item_aliases.Alias_dictionary_name"
        'Print #13, "     _item_aliases.Alias_dictionary_version"
        'Print #13, "     _item_aliases.Item_ID"
        'Print #13, "     _item_aliases.Category_ID"
        'Print #13, "     _item_aliases.Dictionary_ID"
        'Print #13,
        'Print #13, "     ?  ?  ?  ?  ?  ?  ?"
        'Print #13,
        'Print #13, "   stop_"
        'Print #13,
        
'  THIS IS WHERE ITEM ENUMERATIONS ARE PRINTED
        
        If excel_tag_dat(j, 21) = "Y" Then
            Print #13, "   loop_"
            Print #13, "     _Item_enumeration.Item_name"
            Print #13, "     _Item_enumeration.Value"
            Print #13, "     _Item_enumeration.Detail"
            Print #13, "     _Item_enumeration.Item_ID"
            Print #13, "     _Item_enumeration.Category_ID"
            Print #13, "     _Item_enumeration.Dictionary_ID"
            Print #13,
            For k = 1 To enum_ct
                If enum_values(k, 1) = excel_tag_dat(j, 9) Then
                    Print #13, "  '" + excel_tag_dat(j, 9) + "'" + "   " + enum_values(k, 2) + "    " + enum_values(k, 3) + "     " + Str(item_count) + "     " + Str(cat_count) + "     " + dictionary_name
                End If
            Next k

            If enum_ct < 1 Then
                Print #13, "    ?     ?     ?     ?     ?     ?"
            End If
            Print #13,
            Print #13, "   stop_"
            Print #13,
        End If
        
        If example <> "?" Then
            Print #13, "   loop_"
            Print #13, "     _Item_examples.Item_name"
            Print #13, "     _Item_examples.Case"
            Print #13, "     _Item_examples.Detail"
            Print #13, "     _Item_examples.Item_ID"
            Print #13, "     _Item_examples.Category_ID"
            Print #13, "     _Item_examples.Dictionary_ID"
            Print #13,
            Print #13, "  '" + excel_tag_dat(j, 9) + "'"

            Print #13, ";"
            ln = Len(example)
            For i23 = 1 To ln
                If Mid$(example, i23, 1) = "$" Then
                    example = Left$(example, i23 - 1) + "," + Right$(example, ln - i23)
                    'Debug.Print example
                    'Debug.Print
                End If
            Next i23
            Print #13, example
            Print #13, ";"
            Print #13, "   ?"                           'detail not extracted
            Print #13, Str(item_count);                 'item_ID
            Print #13, "    " + Str(cat_count);         'category_ID
            Print #13, "    " + dictionary_name         'dictionary_ID
            Print #13,
            Print #13, "   stop_"
        End If
        Print #13,
        
        Print #13, "save_"
        Print #13,
    End If
Next j

End Sub

Sub write_dict_comp(e_tag_ct, excel_tag_dat, excel_header_dat, pathin, input_file, e_col_num, input_files)
Debug.Print "Writing Dictionary Update File"

'input new dictionary

e_col_num = 106     'number of columns in the xlschem_ann.csv file
    
pathmy = "\eldonulrich"
pathout_adit = "z:" + pathmy + "\bmrb\htdocs\dictionary\htmldocs\nmr_star\bmrb_star_v3_files\"
input_file = input_files(2)
pathin = pathout_adit

read_excel e_tag_ct, excel_tag_dat, excel_header_dat, pathin, input_file, e_col_num
'read category group file to get list of save frame categories


'copy new dictionary to new variables

ReDim new_tag_dat(e_tag_ct, e_col_num)
For i = 1 To e_tag_ct
    For j = 1 To e_col_num
        new_tag_dat(i, j) = excel_tag_dat(i, j)
    Next j
Next i
new_tag_ct = e_tag_ct

'input old dictionary
'read old category group file

pathin = "z:" + pathmy + "\bmrb\htdocs\dictionary\htmldocs\nmr_star\last_v3_dictionary\bmrb_star_v3_files\"
input_file = input_files(2)

read_excel e_tag_ct, excel_tag_dat, excel_header_dat, pathin, input_file, e_col_num


'Print #13, "Dictionary update information"
'Print #13,

'compare new and old save frames

'Print #13, "Save frames added"
'Print #13,

'Print #13, "Save frames deleted"
'Print #13,

'compare tag categories
'Print #13, "Tag categories added"
'Print #13,
'Print #13, "Tag categories deleted"
'Print #13,

'compare individual tags
'Print #13, "Specific tags added within previously existing tag categories"
'Print #13,
'Print #13, "Specific tags deleted deleted within previously existing tag categories"
'Print #13,

'compare individual tag item characteristics (data type, definition, etc.)
'Print #13, "Tag item characteristics updated"
'Print #13,

Debug.Print new_tag_ct
Debug.Print e_tag_ct

For i = 1 To new_tag_ct             'new dictionary tag count
    For j = 1 To e_tag_ct           'old dictionary tag count
        If new_tag_dat(i, 9) = excel_tag_dat(j, 9) Then
            Debug.Print i, new_tag_dat(i, 9)
            For i1 = 1 To e_col_num
                If new_tag_dat(i, i1) <> excel_tag_dat(j, i1) Then
                    Debug.Print new_tag_dat(i, i1) + "  ";
                    Debug.Print
                End If
            Next i1
            Debug.Print
        End If
    Next j
Next i
Debug.Print

'tag enumeration comparison
    'load new enumerations
    
'load_enumerations
    
    'load old enumerations

pathin = "/last_v3_dictionary/"
'load_enumerations



End Sub
Sub load_enumerations(enum_ct, enum_list, pathin)

ReDim enum_list(500) As String

Open pathin + "enumerations.txt" For Input As 21
While EOF(21) <> -1
    t = Input$(1, 21)
    t1 = Asc(t)
    If t1 <> 10 And t1 <> 13 Then tt = tt + t
    If t1 = 10 Then
        p = InStr(1, tt, "save__")
        If p = 1 Then
            ln = Len(tt)
            enum_ct = enum_ct + 1
            enum_list(enum_ct) = Trim(Right$(tt, ln - 5))
        End If
        tt = ""
    End If
Wend
Close 21
Open pathin + "adit_input/adit_enum_dtl.csv" For Input As 21
While EOF(21) <> -1
    enum_test_ct = enum_test_ct + 1
    For i = 1 To 4
        Input #21, enum_test(i)
    Next i
Wend

End Sub

Sub write_xml(output_files, e_tag_ct, excel_tag_dat, excel_header_dat, pathin, input_file, e_col_num, spg_col_ct, spg_row_ct, spg_table, grp_col_ct, grp_row_ct, grp_table, adit_file_source)

Debug.Print "Writing XML specification"

Open pathin + "xml/" + output_files(28) For Output As 1
Open pathin + "xml/" + "header_xml_spec.txt" For Input As 2

Dim enum_ct, cat_ct As Integer
ReDim enumer(3000, 4), cat_desc(500, 3) As String

'begin loading data required to print xsd file
'load enumeration information
Open pathin + "adit_enum_dict1.csv" For Input As 21
While EOF(21) <> -1
    enum_ct = enum_ct + 1
    For i = 1 To 4
        Input #21, enumer(enum_ct, i)
    Next i
Wend
Close 21

'load category super_group category_group and item_category descriptions
Open pathin + "adit_category_desc.csv" For Input As 21
While EOF(21) <> -1
    cat_ct = cat_ct + 1
    For i = 1 To 3
        Input #21, cat_desc(cat_ct, i)
    Next i
Wend
Close 21


'begin printing of xsd file
'print header for xsd file
tt = "": t = ""
While EOF(2) <> -1
    t = Input$(1, 2)
    If Asc(t) <> 10 Then
        If Asc(t) <> 13 Then
            tt = tt + t
        End If
    End If
    If Asc(t) = 10 Then
'        Print #1, tt
        tt = ""
    End If
Wend
Close 2

'begin printing the body of the xsd file
last_sfcat = "": sfcat_num = 0
last_cat = "": cat_num = 0

'Print #1, "      <xsd:sequence>"

Print #1, "<!-- definition of simple types (NMR-STAR unique data types) -->"
    'regular expressions
    'unique types
        'VARCHAR(3), VARCHAR(15), VARCHAR(31), VARCHAR(127), VARCHAR(255)
        'CHAR(12), INTEGER, FLOAT, DATE, TEXT, BOOLEAN
        'framecode?, label

Print #1, "<!-- definition of simple elements (NMR-STAR unique tags) -->"
For i = 1 To e_tag_ct  'list of all tags

'print individual tag descriptions
    print_attr = 0
    If excel_tag_dat(i, 89) <> "?" Then print_attr = 1
    If excel_tag_dat(i, 34) = "Y" Then print_attr = 1
    If excel_tag_dat(i, 35) = "Y" Then print_attr = 1
    If excel_tag_dat(i, 96) = "Y" Then print_attr = 1

    If print_attr = 0 Then
        Print #1, "                                 <xsd:element name=" + Chr(34) + excel_tag_dat(i, 80) + Chr(34) + " minOccurs=" + Chr(34) + "1"; Chr(34) + " maxOccurs=" + Chr(34) + Trim(Str(1)) + Chr(34) + " nillable=";
        If excel_tag_dat(i, 29) = "NOT NULL" Then t = "false" Else t = "true"
        Print #1, Chr(34) + t + Chr(34) + " type=" + Chr(34) + "xsd:" + excel_tag_dat(i, 28) + Chr(34) + ">"
        Print #1, "                                    <xsd:annotation>"
        Print #1, "                                       <xsd:documentation xml:lang="; Chr(34); "en"; Chr(34); ">"
        Print #1, excel_tag_dat(i, 92)
        Print #1, "                                       </xsd:documentation>"
        Print #1, "                                    </xsd:annotation>"

        'check to see if tag is enumerated and print enumerations if needed
        print_xml_tag = 1: print_enum = 0
        For i1 = 1 To enum_ct
            If excel_tag_dat(i, 9) = enumer(i1, 1) Then
                print_enum = 1
                If print_enum = 1 Then
                    If print_xml_tag = 1 Then
                
                        Print #1, "                                    <xsd:simpleType>"
                        Print #1, "                                       <xsd:restriction base=" + Chr(34) + "xsd:string" + Chr(34) + ">"
                        print_xml_tag = 0
                    End If
                    Print #1, "                                          <xsd:enumeration value=" + Chr(34) + enumer(i1, 2) + Chr(34) + " />"
                End If
            End If
        Next i1

        If print_enum = 1 Then
            Print #1, "                                       </xsd:restriction>"
            Print #1, "                                    </xsd:simpleType>"
        End If

        Print #1, "                                 </xsd:element>"
    End If

    If print_attr = 1 Then
        If excel_tag_dat(i, 35) = "Y" Then Print #1, "                                 <xsd:attribute name=" + Chr(34) + excel_tag_dat(i, 80) + Chr(34) + " type=" + Chr(34) + "xsd:ID" + Chr(34) + " minOccurs=" + Chr(34) + "1"; Chr(34) + " maxOccurs=" + Chr(34) + Trim(Str(1)) + Chr(34) + " nillable=";
        If excel_tag_dat(i, 35) = "" Then Print #1, "                                 <xsd:attribute name=" + Chr(34) + excel_tag_dat(i, 80) + Chr(34) + " type=" + Chr(34) + "xsd:IDREF" + Chr(34) + " minOccurs=" + Chr(34) + "1"; Chr(34) + " maxOccurs=" + Chr(34) + Trim(Str(1)) + Chr(34) + " nillable=";
        If excel_tag_dat(i, 29) = "NOT NULL" Then t = "false" Else t = "true"
        Print #1, Chr(34) + t + Chr(34) + " type=" + Chr(34) + "xsd:" + excel_tag_dat(i, 28) + Chr(34) + ">"
        Print #1, "                                    <xsd:annotation>"
        Print #1, "                                       <xsd:documentation xml:lang="; Chr(34); "en"; Chr(34); ">"
        Print #1, excel_tag_dat(i, 92)
        Print #1, "                                       </xsd:documentation>"
        Print #1, "                                    </xsd:annotation>"
        Print #1, "                                 </xsd:attribute>"

    End If
Next i


Print #1, "<!-- definition of complex types (NMR-STAR categories) -->"


Print #1, "<!-- definition of objects (save frames) -->"


    sfcat = excel_tag_dat(i, 2)
    cat = excel_tag_dat(i, 79)

If sfcat <> last_sfcat Then 'print closing tags for save frame categories
    If sfcat_num = 1 Then
        Print #1, "                           </xsd:complexType>"
        Print #1, "                        </xsd:element>"
        Print #1, "                     </xsd:sequence>"
        'Print #1, "               </xsd:all>"
        Print #1, "            </xsd:complexType>"
        Print #1, "         </xsd:element>"
    End If
    sfcat_num = 1
    cat_num = 0
    
    'print opening tags for save frame categories
    For i1 = 1 To grp_row_ct     'extract min and max occurance for save frames
        If excel_tag_dat(i, 2) = grp_table(i1, 4) Then
            If grp_table(i1, 8) = "1" Then minocc = grp_table(i1, 8)
            If grp_table(i1, 8) = "" Then minocc = "0"
            If grp_table(i1, 9) = "Y" Then maxocc = "unbounded"
            If grp_table(i1, 9) = "N" Then maxocc = "1"
            Exit For
        End If
    Next i1
    
    Print #1, "         <xsd:element name=" + Chr(34) + excel_tag_dat(i, 2) + Chr(34) + " minOccurs=" + Chr(34) + minocc + Chr(34) + " maxOccurs=" + Chr(34) + maxocc + Chr(34) + ">"
    Print #1, "            <xsd:complexType>"
For i1 = 1 To cat_ct
    If excel_tag_dat(i, 2) = cat_desc(i1, 2) Then
        Print #1, "               <xsd:annotation>"
        Print #1, "                  <xsd:documentation xml:lang="; Chr(34); "en"; Chr(34); ">"
        Print #1, cat_desc(i1, 3)
        Print #1, "                  </xsd:documentation>"
        Print #1, "               </xsd:annotation>"
        Exit For
    End If
Next i1
    'Print #1, "               <xsd:all>"
    last_sfcat = sfcat
    
End If


If cat <> last_cat Then
    If cat_num = 1 Then     'print closing tags for item categories
        'Print #1, "                              </xsd:all>"
        Print #1, "                           </xsd:complexType>"
        Print #1, "                        </xsd:element>"
        'Print #1, "                     </xsd:sequence>"
    End If
    cat_num = 1
End If

If cat <> last_cat Then             'print opening tags for item categories
    Print #1, "                     <xsd:sequence>"
    Print #1, "                        <xsd:element name=" + Chr(34) + excel_tag_dat(i, 79) + Chr(34) + " minOccurs=" + Chr(34) + "1" + Chr(34) + " maxOccurs=" + Chr(34) + "1" + Chr(34) + ">"
    Print #1, "                           <xsd:complexType>"
For i1 = 1 To cat_ct
    If excel_tag_dat(i, 79) = cat_desc(i1, 2) Then
        Print #1, "                              <xsd:annotation>"
        Print #1, "                                 <xsd:documentation xml:lang="; Chr(34); "en"; Chr(34); ">"
        Print #1, cat_desc(i1, 3)
        Print #1, "                                 </xsd:documentation>"
        Print #1, "                               </xsd:annotation>"
        Exit For
    End If
Next i1
    'Print #1, "                              <xsd:all>"
    last_cat = cat
End If

Next i

Print #1, "                           </xsd:complexType>"
Print #1, "                        </xsd:element>"
Print #1, "                     </xsd:sequence>"
'Print #1, "               </xsd:all>"
Print #1, "            </xsd:complexType>"
Print #1, "         </xsd:element>"
Print #1, "      </xsd:sequence>"

Close
End Sub

Sub write_28_column_excel_file(excel_tag_dat, excel_header_dat, e_col_num, e_tag_ct, output_files, pathout_adit)
Debug.Print "Write out 28 column excel dictionary file"

Open pathout_adit + output_files(30) For Output As 2

'FOR NMR-STAR COLUMN = 106 CONVERSION TO 28 COLUMN FORMAT
If e_col_num = 106 Then
    For i = 1 To e_tag_ct
        'excel_tag_dat(i, 1) = excel_tag_dat(i, 1)
        excel_tag_dat(i, 1) = ""
        excel_tag_dat(i, 2) = excel_tag_dat(i, 2)
        excel_tag_dat(i, 3) = excel_tag_dat(i, 5)
        excel_tag_dat(i, 4) = excel_tag_dat(i, 6)
        excel_tag_dat(i, 5) = excel_tag_dat(i, 7)
        excel_tag_dat(i, 6) = excel_tag_dat(i, 8)
        excel_tag_dat(i, 7) = excel_tag_dat(i, 9)
        excel_tag_dat(i, 8) = excel_tag_dat(i, 21)
        excel_tag_dat(i, 9) = excel_tag_dat(i, 22)
        excel_tag_dat(i, 10) = excel_tag_dat(i, 28)
        excel_tag_dat(i, 11) = excel_tag_dat(i, 29)
        excel_tag_dat(i, 12) = excel_tag_dat(i, 33)
        excel_tag_dat(i, 13) = excel_tag_dat(i, 35)
        excel_tag_dat(i, 14) = excel_tag_dat(i, 36)
        excel_tag_dat(i, 15) = excel_tag_dat(i, 38)
        excel_tag_dat(i, 16) = excel_tag_dat(i, 39)
        excel_tag_dat(i, 17) = excel_tag_dat(i, 43)
        excel_tag_dat(i, 18) = excel_tag_dat(i, 44)
        excel_tag_dat(i, 19) = excel_tag_dat(i, 51)
        excel_tag_dat(i, 20) = excel_tag_dat(i, 58)
        excel_tag_dat(i, 21) = excel_tag_dat(i, 59)
        excel_tag_dat(i, 22) = excel_tag_dat(i, 60)
        excel_tag_dat(i, 23) = excel_tag_dat(i, 77)
        excel_tag_dat(i, 24) = excel_tag_dat(i, 79)
        excel_tag_dat(i, 25) = excel_tag_dat(i, 80)
        excel_tag_dat(i, 26) = excel_tag_dat(i, 87)
        excel_tag_dat(i, 27) = excel_tag_dat(i, 92)
        excel_tag_dat(i, 28) = excel_tag_dat(i, 106)
        If i < 5 Then
            excel_header_dat(i, 1) = excel_header_dat(i, 1)
            excel_header_dat(i, 2) = excel_header_dat(i, 2)
            If i = 4 Then excel_header_dat(i, 3) = excel_header_dat(i, 4)
            If i <> 4 Then excel_header_dat(i, 3) = excel_header_dat(i, 5)
            excel_header_dat(i, 4) = excel_header_dat(i, 6)
            excel_header_dat(i, 5) = excel_header_dat(i, 7)
            excel_header_dat(i, 6) = excel_header_dat(i, 8)
            excel_header_dat(i, 7) = excel_header_dat(i, 9)
            excel_header_dat(i, 8) = excel_header_dat(i, 21)
            excel_header_dat(i, 9) = excel_header_dat(i, 22)
            excel_header_dat(i, 10) = excel_header_dat(i, 28)
            excel_header_dat(i, 11) = excel_header_dat(i, 29)
            excel_header_dat(i, 12) = excel_header_dat(i, 33)
            excel_header_dat(i, 13) = excel_header_dat(i, 35)
            excel_header_dat(i, 14) = excel_header_dat(i, 36)
            excel_header_dat(i, 15) = excel_header_dat(i, 38)
            excel_header_dat(i, 16) = excel_header_dat(i, 39)
            excel_header_dat(i, 17) = excel_header_dat(i, 43)
            excel_header_dat(i, 18) = excel_header_dat(i, 44)
            excel_header_dat(i, 19) = excel_header_dat(i, 51)
            excel_header_dat(i, 20) = excel_header_dat(i, 58)
            excel_header_dat(i, 21) = excel_header_dat(i, 59)
            excel_header_dat(i, 22) = excel_header_dat(i, 60)
            excel_header_dat(i, 23) = excel_header_dat(i, 77)
            excel_header_dat(i, 24) = excel_header_dat(i, 79)
            excel_header_dat(i, 25) = excel_header_dat(i, 80)
            excel_header_dat(i, 26) = excel_header_dat(i, 87)
            excel_header_dat(i, 27) = excel_header_dat(i, 92)
            excel_header_dat(i, 28) = "?"
        End If

    Next i
    For i = 1 To 28
        excel_header_dat(3, i) = Trim(Str(i))
    Next i
    e_col_num = 28
End If
For i = 1 To 4
    For j = 1 To e_col_num - 1
        Print #2, excel_header_dat(i, j) + ",";
    Next j
    Print #2, excel_header_dat(i, e_col_num)
Next i
For i = 1 To e_tag_ct
    For j = 1 To e_col_num - 1
        Print #2, excel_tag_dat(i, j) + ",";
    Next j
    Print #2, excel_tag_dat(i, e_col_num)
Next i
Print #2, "TBL_END" + ",";
For i = 2 To e_col_num - 1
    Print #2, ",";
Next i
Print #2, "?"

    
End Sub
