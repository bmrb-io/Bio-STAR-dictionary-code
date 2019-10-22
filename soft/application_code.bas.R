Attribute VB_Name = "application_code"
Sub read_excel(e_tag_ct, excel_tag_dat, excel_header_dat, pathin, input_file, e_col_num)
ReDim excel_tag_dat(4000, e_col_num) As String
ReDim excel_header_dat(5, e_col_num) As String
Dim i, j As Integer

Open pathin + input_file For Input As 1
For i = 1 To 4
    For j = 1 To e_col_num
        Input #1, excel_header_dat(i, j)
    Next j
Next i
e_tag_ct = 0
While EOF(1) <> -1
    e_tag_ct = e_tag_ct + 1
    For j = 1 To e_col_num
        Input #1, excel_tag_dat(e_tag_ct, j)
    Next j
    'Debug.Print excel_tag_dat(e_tag_ct, 1)
    'Debug.Print excel_tag_dat(e_tag_ct, 52)
    'Debug.Print
Wend
If excel_tag_dat(e_tag_ct, 1) = "TBL_END" Then e_tag_ct = e_tag_ct - 1
Close 1


End Sub

Sub check_excel(e_tag_ct, excel_tag_dat, excel_header_dat, e_col_num, spg_col_ct, spg_row_ct, spg_table, grp_col_ct, grp_row_ct, grp_table)
Dim i, p, p1, p2, p3, p4, table_ct, null_ct As Integer
Dim tag_col, sf_col, dbtab, dbcol, cat_col, prompt As Integer
Dim usr_view, sf_id, sf_id_set, dat_type, src_key, row_id, adit_spg, adit_cg As Integer

ReDim table_list(500) As String
ReDim null_list(50)

For i = 1 To e_col_num
    If excel_header_dat(1, i) = "Tag" Then tag_col = i: null_ct = null_ct + 1: null_list(null_ct) = i
    If excel_header_dat(1, i) = "SFCategory" Then sf_col = i: null_ct = null_ct + 1: null_list(null_ct) = i
    If excel_header_dat(1, i) = "ManDBColumnName" Then dbcol = i: null_ct = null_ct + 1: null_list(null_ct) = i
    If excel_header_dat(1, i) = "ManDBTableName" Then dbtab = i: null_ct = null_ct + 1: null_list(null_ct) = i
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
    If excel_header_dat(1, i) = "Item enumerated" Then null_ct = null_ct + 1: null_list(null_ct) = i
    If excel_header_dat(1, i) = "ADIT item view name" Then null_ct = null_ct + 1: null_list(null_ct) = i
    If excel_header_dat(1, i) = "Data Type" Then dat_type = i: null_ct = null_ct + 1: null_list(null_ct) = i
    If excel_header_dat(1, i) = "Row Index Key" Then row_id = i: null_ct = null_ct + 1: null_list(null_ct) = i
    If excel_header_dat(1, i) = "Saveframe ID tag" Then sf_id = i: null_ct = null_ct + 1: null_list(null_ct) = i
    If excel_header_dat(1, i) = "Loopflag" Then null_ct = null_ct + 1: null_list(null_ct) = i
    If excel_header_dat(1, i) = "Seq" Then null_ct = null_ct + 1: null_list(null_ct) = i
    If excel_header_dat(1, i) = "Dbspace" Then null_ct = null_ct + 1: null_list(null_ct) = i
    If excel_header_dat(1, i) = "Example" Then null_ct = null_ct + 1: null_list(null_ct) = i
    If excel_header_dat(1, i) = "Prompt" Then prompt = i: null_ct = null_ct + 1: null_list(null_ct) = i
    If excel_header_dat(1, i) = "Description" Then null_ct = null_ct + 1: null_list(null_ct) = i
Next i

For i = 1 To e_tag_ct
    syntax_check.Text5 = Str(i)
    syntax_check.Text5.Refresh
    
    'create a list of tables
    If excel_tag_dat(i, cat_col) <> cat_last Then
        table_ct = table_ct + 1
        table_list(table_ct) = excel_tag_dat(i, cat_col)
        cat_last = excel_tag_dat(i, cat_col)
        syntax_check.Text4 = Str(table_ct)
        syntax_check.Text4.Refresh

        If sf_id_set = 0 And i > 1 Then
            syntax_check.Text6 = Str(i + 4)
            syntax_check.Text6.Refresh
            syntax_check.Text7 = excel_tag_dat(i, sf_col)
            syntax_check.Text7.Refresh
            syntax_check.Text8 = excel_tag_dat(i, cat_col)
            syntax_check.Text8.Refresh
            syntax_check.Text9 = excel_tag_dat(i, tag_col)
            syntax_check.Text9.Refresh
            syntax_check.Text18 = "Saveframe ID not defined"
            syntax_check.Text18.Refresh

            program_control.Show 1
        
        End If
        sf_id_set = 0
    End If
    If excel_tag_dat(i, cat_col) = cat_last Then
        If excel_tag_dat(i, sf_id) = "Y" Then sf_id_set = 1
    End If
    
    'check saveframe ID data type, should be CHAR(12)
    char_type = 0
    If excel_tag_dat(i, sf_id) = "Y" Then char_type = 1
    If excel_tag_dat(i, src_key) = "Y" Then char_type = 1
    If excel_tag_dat(i, row_id) = "Y" Then char_type = 0
    If char_type = 1 And excel_tag_dat(i, dat_type) <> "CHAR(12)" Then
    If char_type = 1 And excel_tag_dat(i, dat_type) <> "INTEGER" Then
        syntax_check.Text6 = Str(i + 4)
        syntax_check.Text6.Refresh
        syntax_check.Text7 = excel_tag_dat(i, sf_col)
        syntax_check.Text7.Refresh
        syntax_check.Text8 = excel_tag_dat(i, cat_col)
        syntax_check.Text8.Refresh
        syntax_check.Text9 = excel_tag_dat(i, tag_col)
        syntax_check.Text9.Refresh
        syntax_check.Text18 = "Data type not set to CHAR(12)"
        syntax_check.Text18.Refresh

        program_control.Show 1
        
    End If
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

    'check for incorrect usage of . and '_'
    
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
  
    'check length of table names, should be less than 32 characters
    If Len(excel_tag_dat(i, dbtab)) > 31 Then
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
        If excel_tag_dat(i, tag_col) = excel_tag_dat(j, tag_col) Then
        If excel_tag_dat(i, 2) = excel_tag_dat(j, 2) Then ' temporary condition to eliminat known duplicates
        If i <> j Then
            syntax_check.Text6 = Str(j)
            syntax_check.Text6.Refresh
            syntax_check.Text7 = excel_tag_dat(j, sf_col)
            syntax_check.Text7.Refresh
            syntax_check.Text8 = excel_tag_dat(j, cat_col)
            syntax_check.Text8 = Refresh
            syntax_check.Text9 = excel_tag_dat(i, tag_col)
            syntax_check.Text9.Refresh
            syntax_check.Text18 = "Duplicate tags"
            syntax_check.Text18.Refresh
                                                       
            program_control.Show 1
        End If
        End If
        End If
    Next j
Next i

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

End Sub

Sub set_tag_values(tag_ct, tag_char, new_tag_dat, excel_header_dat, e_col_num, spg_col_ct, spg_row_ct, spg_table, grp_col_ct, grp_row_ct, grp_table)
Debug.Print "Set tag values"

Dim i, j, j1 As Integer

For i = 1 To tag_ct
'Debug.Print i, new_tag_dat(i, 2)
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
            If excel_header_dat(1, j) = "ADIT category view name" Then
                ln = Len(tag_char(i, 2))
                new_tag_dat(i, j) = ""
                For j1 = 1 To ln
                    t = Mid$(tag_char(i, 2), j1, 1)
                    If t = "_" Then t = " "
                    new_tag_dat(i, j) = new_tag_dat(i, j) + t
                Next j1
            End If

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
            'If excel_header_dat(1, j) = "Data Type" Then
            '    If new_tag_dat(i, j) = "" Then
            '        If LCase(Right$(tag_char(i, 1), 7)) = "details" Then new_tag_dat(i, j) = "TEXT"
            '        If LCase(Right$(tag_char(i, 1), 7)) = "keyword" Then new_tag_dat(i, j) = "VARCHAR(127)"
            '        If LCase(Right$(tag_char(i, 1), 5)) = "label" Then new_tag_dat(i, j) = "VARCHAR(127)"
            '        If LCase(Right$(tag_char(i, 1), 16)) = "text_data_format" Then new_tag_dat(i, j) = "VARCHAR(31)"
            '        If LCase(Right$(tag_char(i, 1), 9)) = "text_data" Then new_tag_dat(i, j) = "TEXT"
            '        If LCase(Right$(tag_char(i, 1), 2)) = "id" Then new_tag_dat(i, j) = "INTEGER"
            '        If LCase(Right$(tag_char(i, 1), 9)) = "atom_name" Then new_tag_dat(i, j) = "VARCHAR(15)"
            '        If LCase(Right$(tag_char(i, 1), 9)) = "atom_type" Then new_tag_dat(i, j) = "VARCHAR(15)"
            '        If LCase(tag_char(i, 1)) = "sf_id" Then new_tag_dat(i, j) = "INTEGER"
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
            'End If
            If excel_header_dat(1, j) = "Tag" Then
                new_tag_dat(i, j) = tag_char(i, 2) + "." + tag_char(i, 1)
            End If

            If excel_header_dat(1, j) = "ManDBTableName" Then
                new_tag_dat(i, j) = Right$(tag_char(i, 2), Len(tag_char(i, 2)) - 1)
                ln = Len(new_tag_dat(i, j))
                new_tag = ""
                For j1 = 1 To ln
                    t = Mid$(new_tag_dat(i, j), j1, 1)
                    If t = "_" Then
                        underscore = 1
                    End If
                    If t <> "_" Then
                        If underscore = 1 Then
                            t = UCase(t)
                            underscore = 0
                        End If
                        new_tag = new_tag + t
                    End If
                Next j1
                new_tag_dat(i, j) = new_tag
            End If
            If excel_header_dat(1, j) = "ManDBColumnName" Then
                new_tag_dat(i, j) = tag_char(i, 1)
                ln = Len(new_tag_dat(i, j))
                new_tag = ""
                For j1 = 1 To ln
                    t = Mid$(new_tag_dat(i, j), j1, 1)
                    If t = "_" Then
                        underscore = 1
                    End If
                    If t <> "_" Then
                        If underscore = 1 Then
                            t = UCase(t)
                            underscore = 0
                        End If
                        new_tag = new_tag + t
                    End If
                Next j1
                new_tag_dat(i, j) = new_tag
                
                'after setting the value for ManDBColumnName then set these additional values
            For i1 = 1 To e_col_num
                If excel_header_dat(1, i1) = "ManDBColumnName" Then manDBcolname = LCase(new_tag_dat(i, i1))
                If excel_header_dat(1, i1) = "SFCategory" Then SFCategory = LCase(new_tag_dat(i, i1))
                If excel_header_dat(1, i1) = "ManDBTableName" Then
                    manDBtab = LCase(new_tag_dat(i, i1))
                    If SFCategory <> SFCategory_last Then
                        SFmanDBtab = new_tag_dat(i, i1)
                        SFCategory_last = SFCategory
                    End If
                End If
            Next i1
            
            End If
            
            If excel_header_dat(1, j) = "Row Index Key" Then
                If LCase(new_tag_dat(i, 32)) = "ordinal" Then new_tag_dat(i, j) = "Y"
                If new_tag_dat(i, j) = "" Then new_tag_dat(i, j) = "N"
            End If

            'If excel_header_dat(1, j) = "Saveframe ID tag" Then
            '    If manDBtab <> SFmanDBtab Then
            '        If LCase(new_tag_dat(i, 32)) = "sfid" Then new_tag_dat(i, j) = "Y"
            '    End If
            '    If new_tag_dat(i, j) = "" Then new_tag_dat(i, j) = "N"
            'End If

            If excel_header_dat(1, j) = "ADIT item view name" Then
                ln = Len(tag_char(i, 1))
                new_tag_dat(i, j) = ""
                For j1 = 1 To ln
                    t = Mid$(tag_char(i, 1), j1, 1)
                    If t = "_" Then t = " "
                    new_tag_dat(i, j) = new_tag_dat(i, j) + t
                Next j1
            End If
            If excel_header_dat(1, j) = "Dbspace" Then
                new_tag_dat(i, j) = "1"
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
                If new_tag_dat(i, j) = "" Then new_tag_dat(i, j) = "H"
            End If
            If excel_header_dat(1, j) = "ADIT category group ID" Then
                For i1 = 1 To grp_row_ct
                    If grp_table(i1, 4) = new_tag_dat(i, 2) Then
                        new_tag_dat(i, j) = grp_table(i1, 3)    'group ID
                        new_tag_dat(i, 3) = grp_table(i1, 6)    'mandatory group flag
                        new_tag_dat(i, 5) = grp_table(i1, 1)    'category super group ID
                        new_tag_dat(i, 6) = grp_table(i1, 2)    'category super group
                        new_tag_dat(i, 8) = grp_table(i1, 14)   'category group view name
                        Exit For
                    End If
                Next i1
            End If
            
            If excel_header_dat(1, j) = "Table Primary Key" Then
            
                If manDBcolname = "sfcategory" Then
                    new_tag_dat(i, j) = "Y"
                End If
                If manDBcolname = "entryid" Then
                    new_tag_dat(i, j) = "Y"
                End If
                If manDBcolname = "sfid" Then
                    new_tag_dat(i, j) = "Y"
                End If
                If Right$(manDBcolname, 7) = "ordinal" Then
                    new_tag_dat(i, j) = "Y"
                End If
                'If Right$(manDBcolname, 10) = "moleculeid" Then
                '    If manDBtab = "molecule" Then new_tag_dat(i, j) = "Y"
                ' End If
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
            If excel_header_dat(1, j) = "Source Key" Then
                If new_tag_dat(i, j) = "Y" Then
                    If LCase(new_tag_dat(i, j - 3)) = "id" Then
                        new_tag_dat(i, j - 3) = new_tag_dat(i, j - 4) + new_tag_dat(i, j - 3)
                    End If
                    'If new_tag_dat(i, j - 3) = "AssignedChemicalShiftsID" Then
                        'Debug.Print
                    'End If

                    For j1 = 1 To tag_ct
                        If j1 <> i Then
                            p = InStr(1, LCase(new_tag_dat(j1, j - 3)), LCase(new_tag_dat(i, j - 3)))
                            If p > 0 Then
                                If p + Len(new_tag_dat(i, j - 3)) = Len(new_tag_dat(j1, j - 3)) + 1 Or p = 1 Then
                                    new_tag_dat(j1, j + 3) = new_tag_dat(i, j - 4)
                                    new_tag_dat(j1, j + 4) = new_tag_dat(i, j - 3)
                                    new_tag_dat(j1, j - 7) = new_tag_dat(i, j - 7)
                                    new_tag_dat(j1, j - 6) = new_tag_dat(i, j - 6)
                                    'If new_tag_dat(j1, j - 3) = "AssignedChemicalShiftsID" Then
                                    '    Debug.Print
                                    'End If
                                    'Debug.Print
                                    If new_tag_dat(i, 34) = "Y" Then new_tag_dat(j1, j + 4) = "Sf_ID"
                                End If
                            End If
                        End If
                    Next j1
                End If
            End If
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
            If LCase(new_tag_dat(i, j - 3)) = "id" Then
                new_tag_dat(i, j - 3) = new_tag_dat(i, j - 4) + new_tag_dat(i, j - 3)
            End If
            For j1 = 1 To tag_ct
                If j1 <> i Then
                    If LCase(new_tag_dat(j1, j - 3)) = "id" Then
                        new_tag_dat(j1, j - 3) = new_tag_dat(j1, j - 4) + new_tag_dat(j1, j - 3)
                    End If
                    If LCase(new_tag_dat(j1, j - 3)) = LCase(new_tag_dat(i, j - 3)) Then
                        'Debug.Print
                        If p + Len(new_tag_dat(i, j - 3)) = Len(new_tag_dat(j1, j - 3)) + 1 Then
                            new_tag_dat(j1, j + 3) = new_tag_dat(i, j - 4)
                            new_tag_dat(j1, j + 4) = new_tag_dat(i, j - 3)
                            'Debug.Print
                        End If
                    End If
                End If
            Next j1
        End If
    End If
    j = 34
    If excel_header_dat(1, j) = "Saveframe ID tag" Then
        If new_tag_dat(i, j) = "Y" Then
            new_tag_dat(i, j - 2) = "Sf_ID"
            If new_tag_dat(i, j + 5) > "" Then
                new_tag_dat(i, j + 5) = "Sf_ID"
            End If
        End If
    End If
            

Next i
End Sub

Sub write_new_excel(pathin, output_file, tag_ct, new_tag_dat, excel_header_dat, e_col_num)

Debug.Print "Writing new Excel file"

Dim i, j As Integer

Open pathin + output_file For Output As 1

' add header
For i = 1 To 4
    For j = 1 To e_col_num
        If j < e_col_num Then Print #1, excel_header_dat(i, j) + ",";
        If j = e_col_num Then
            If excel_header_dat(i, j) = "" Then
                Print #1, "?"
            Else
                Print #1, excel_header_dat(i, j)
            End If
        End If
    Next j
Next i


For i = 1 To tag_ct
    For j = 1 To e_col_num - 3
        Print #1, new_tag_dat(i, j) + ",";
    Next j
    For j = e_col_num - 2 To e_col_num
        If j < e_col_num Then
            If new_tag_dat(i, j) = "" Then new_tag_dat(i, j) = "?"
            If new_tag_dat(i, j) = "help_text" Then new_tag_dat(i, j) = "?"
            If new_tag_dat(i, j) = "example_text" Then new_tag_dat(i, j) = "?"
            Print #1, new_tag_dat(i, j) + ",";
        End If
        If j = e_col_num Then
            If new_tag_dat(i, j) = "" Then new_tag_dat(i, j) = "?"
            If new_tag_dat(i, j) = "description_text" Then new_tag_dat(i, j) = "?"
            Print #1, new_tag_dat(i, j)
        End If
    Next j
Next i

' add last line of Excel table
Print #1, "TBL_END" + ",";
For i = 2 To e_col_num
    If i < e_col_num Then Print #1, ",";
    If i = e_col_num Then Print #1, "?"
Next i

Close 1

End Sub
Sub write_new_adit_files(pathin, output_files, tag_ct, new_tag_dat, e_col_num, excel_header_dat)

Debug.Print "Writing ADIT files"

Dim i, j, k As Integer
Dim view As String

Open pathin + output_files(3) For Output As 1

ReDim col_head(e_col_num, 1 To 2) As String
set_column_headings col_head, e_col_num, excel_header_dat

adit_item_col_ct = 35
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
            For j = 16 To 20
                view = view + new_tag_dat(i, j)
            Next j
            Print #1, view + ",";
        End If
        For j = 1 To e_col_num
            If excel_header_dat(1, j) = col_head(k, 1) Then
                If k < adit_item_col_ct Then Print #1, new_tag_dat(i, j) + ",";
                If k = adit_item_col_ct Then Print #1, new_tag_dat(i, j)
                Exit For
            End If
        Next j
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
col_head(11, 1) = excel_header_dat(1, 11): col_head(11, 2) = "depMandatory"
col_head(12, 1) = excel_header_dat(1, 15): col_head(12, 2) = "aditExists"
col_head(13, 1) = excel_header_dat(3, 16): col_head(13, 2) = "aditViewFlgs"
col_head(14, 1) = excel_header_dat(1, 21): col_head(14, 2) = "enumeratedFlg"
col_head(15, 1) = excel_header_dat(1, 22): col_head(15, 2) = "itemEnumClosedFlg"
col_head(16, 1) = excel_header_dat(1, 25): col_head(16, 2) = "derivedEnumTbl"
col_head(17, 1) = excel_header_dat(1, 26): col_head(17, 2) = "derivedEnumCol"
col_head(18, 1) = excel_header_dat(1, 27): col_head(18, 2) = "aditItemViewName"
col_head(19, 1) = excel_header_dat(1, 28): col_head(19, 2) = "dbType"
col_head(20, 1) = excel_header_dat(1, 29): col_head(20, 2) = "dbNullable"
col_head(21, 1) = excel_header_dat(1, 30): col_head(21, 2) = "internalFlg"
col_head(22, 1) = excel_header_dat(1, 33): col_head(22, 2) = "rowIndexFlg"
col_head(23, 1) = excel_header_dat(1, 34): col_head(23, 2) = "sfIdFlg"
col_head(24, 1) = excel_header_dat(1, 36): col_head(24, 2) = "primaryKey"
col_head(25, 1) = excel_header_dat(1, 37): col_head(25, 2) = "foreignKeyGroup"
col_head(26, 1) = excel_header_dat(1, 38): col_head(26, 2) = "foreignTable"
col_head(27, 1) = excel_header_dat(1, 39): col_head(27, 2) = "foreignColumn"
col_head(28, 1) = excel_header_dat(1, 40): col_head(28, 2) = "indexFlg"
col_head(29, 1) = excel_header_dat(1, 31): col_head(29, 2) = "dbTableName"
col_head(30, 1) = excel_header_dat(1, 32): col_head(30, 2) = "dbColumnName"
col_head(31, 1) = excel_header_dat(1, 43): col_head(31, 2) = "loopFlg"
col_head(32, 1) = excel_header_dat(1, 44): col_head(32, 2) = "seq"
col_head(33, 1) = excel_header_dat(1, 50): col_head(33, 2) = "dbSpace"
col_head(34, 1) = excel_header_dat(1, 51): col_head(34, 2) = "example"
col_head(35, 1) = excel_header_dat(1, 52): col_head(35, 2) = "prompt"
col_head(36, 1) = excel_header_dat(1, 53): col_head(36, 2) = "description"

End Sub
Sub write_enum_ties(pathin, output_files, tag_ct, new_tag_dat, e_col_num)
Debug.Print "Writing enumeration ties file"
Dim i, j As Integer

Open pathin + output_files(12) For Output As 12
Print #12, "TBL_BEGIN,,,?"

For i = 1 To tag_ct
    If new_tag_dat(i, 46) > "" Then
    For j = 1 To tag_ct
        If new_tag_dat(i, 46) = new_tag_dat(j, 46) Then
            Print #12, new_tag_dat(i, 31) + ",";
            Print #12, "_" + new_tag_dat(i, 31) + "." + new_tag_dat(i, 32) + ",";
            Print #12, new_tag_dat(j, 31) + ",";
            Print #12, "_" + new_tag_dat(j, 31) + "." + new_tag_dat(j, 32)
        End If
    Next j
    End If
Next i

Print #12, "TBL_END,,,?"
End Sub
Sub write_man_over(pathin, output_files, tag_ct, new_tag_dat, e_col_num, grp_col_ct, grp_row_ct, grp_table)
Debug.Print "Writing mandatory overide file"
Dim i, j As Integer

Open pathin + output_files(13) For Output As 13
Print #13, "TBL_BEGIN,,,,,,?"

For i = 1 To tag_ct
    If new_tag_dat(i, 47) > "" And new_tag_dat(i, 48) = "" Then
    For j = 1 To tag_ct
        If i <> j Then
        If new_tag_dat(i, 47) = new_tag_dat(j, 47) Then
            Print #13, new_tag_dat(j, 2) + ",";
            Print #13, new_tag_dat(j, 31) + ",";
            Print #13, "_" + new_tag_dat(j, 31) + "." + new_tag_dat(j, 32) + ",";
            Print #13, new_tag_dat(j, 49) + ",";
            Print #13, new_tag_dat(i, 31) + ",";
            Print #13, "_" + new_tag_dat(i, 31) + "." + new_tag_dat(i, 32) + ",";
            Print #13, new_tag_dat(j, 48)
        End If
        End If
    Next j
    End If
Next i

Print #13, "TBL_END,,,,,,?"

End Sub
Sub write_anno_star(pathin, output_files, tag_ct, new_tag_dat, e_col_num)

Debug.Print "Writing annotated version of NMR-STAR"

Dim i, ln, ln1 As Integer
tag_prt_ct = 0
ReDim tag_prt(1 To 100, 1 To 2) As String

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
        If new_tag_dat(i, 12) > "" Then tag_prt(tag_prt_ct, 2) = "# " + new_tag_dat(i, 11) + ";" + Space(8 - ln) + new_tag_dat(i, 12)
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

Dim i, ln, ln1 As Integer
tag_prt_ct = 0

ReDim tag_prt(1 To 100, 1 To 3) As String
ReDim tag_prt2(1 To 100) As String
Dim tag_fake_val(1 To 20, 1 To 2)

tag_fake_val(1, 1) = "INTEGER": tag_fake_val(1, 2) = "2"
tag_fake_val(2, 1) = "CHAR(3)": tag_fake_val(2, 2) = "Y"
tag_fake_val(3, 1) = "CHAR(12)": tag_fake_val(3, 2) = "1"
tag_fake_val(4, 1) = "VARCHAR(2)": tag_fake_val(4, 2) = Chr$(34) + "Tiny string value" + Chr$(34)
tag_fake_val(5, 1) = "VARCHAR(3)": tag_fake_val(5, 2) = "Y/N"
tag_fake_val(6, 1) = "VARCHAR(15)": tag_fake_val(6, 2) = Chr$(34) + "Short string value" + Chr$(34)
tag_fake_val(7, 1) = "VARCHAR(31)": tag_fake_val(7, 2) = Chr$(34) + "String value" + Chr$(34)
tag_fake_val(8, 1) = "VARCHAR(80)": tag_fake_val(8, 2) = Chr$(34) + "Brief phrase" + Chr$(34)
tag_fake_val(9, 1) = "VARCHAR(127)": tag_fake_val(9, 2) = Chr$(34) + "Long string value" + Chr$(34)
tag_fake_val(10, 1) = "VARCHAR(255)": tag_fake_val(10, 2) = Chr$(34) + "Very long phrase" + Chr$(34)
tag_fake_val(11, 1) = "TEXT": tag_fake_val(11, 2) = Chr$(34) + "Possible multiline text" + Chr$(34)
tag_fake_val(12, 1) = "FLOAT": tag_fake_val(12, 2) = "110.234"
tag_fake_val(13, 1) = "DATETIME year to day": tag_fake_val(13, 2) = "2002-04-15"
tag_fake_val(14, 1) = "": tag_fake_val(14, 2) = Chr$(34) + "Missing data type" + Chr$(34)
'tag_fake_val(15, 1) = "INTEGER": tag_fake_val(15, 2) = "2"
'tag_fake_val(16, 1) = "INTEGER": tag_fake_val(16, 2) = "2"
'tag_fake_val(17, 1) = "INTEGER": tag_fake_val(17, 2) = "2"
'tag_fake_val(18, 1) = "INTEGER": tag_fake_val(18, 2) = "2"
'tag_fake_val(19, 1) = "INTEGER": tag_fake_val(19, 2) = "2"
'tag_fake_val(20, 1) = "INTEGER": tag_fake_val(20, 2) = "2"

Open pathin + output_files(15) For Output As 2

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
            For j2 = 1 To tag_prt_ct
                tag_prt2(j2) = tag_prt(j2, 3)
            Next j2
            print_tags tag_prt_ct, tag_prt
            Print #2,
            Print #2, "    ";
            For i1 = 1 To tag_prt_ct2
            
                For j2 = 1 To 14
                    If tag_prt2(i1) = tag_fake_val(j2, 1) Then
                        Print #2, tag_fake_val(j2, 2) + "  ";
                    End If
                Next j2

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
        If new_tag_dat(i, 4) = "T6" Then tag_prt(tag_prt_ct, 3) = new_tag_dat(i, 28)
        If new_tag_dat(i, 4) <> "T6" Then
            For j2 = 1 To 14
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
If lp_flag_last = "T6" Then
    Print #2,
    Print #2, "    ";
    For i1 = 1 To tag_prt_ct2
        For j2 = 1 To 14
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
ReDim tag_prt(1 To 100, 1 To 3) As String

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
                If Len(category1) > 31 Then
                    j1 = 0
                    For ji = 1 To Len(category1)
                        If Mid$(category1, ji, 1) = "_" Then j1 = j1 + 1
                    Next ji
                    If Len(category1) - j1 > 31 Then
                        Print #10, sf_cat_name
                        Print #10, category1
                        Print #10,
                        syntax_check.Text18 = "Category too long"
                        syntax_check.Text18.Refresh

                        program_control.Show 1
                    End If
                End If
                If Len(tag1) > 31 Then
                    j1 = 0
                    For ji = 1 To Len(tag1)
                        If Mid$(tag1, ji, 1) = "_" Then j1 = j1 + 1
                    Next ji
                    If Len(tag1) - j1 > 31 Then
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


Sub update_interface(e_tag_ct, excel_tag_dat, tag_ct, tag_list)
Dim i, j, print_flag As Integer

For j = 1 To e_tag_ct
    print_flag = 1
    If excel_tag_dat(j, 16) = "H" Then print_flag = 0
    If excel_tag_dat(j, 49) = "N" Then print_flag = 1
    If excel_tag_dat(j, 49) = "Y" Then print_flag = 1
    If print_flag = 1 Then
        For i = 1 To tag_ct
            If excel_tag_dat(j, 9) = tag_list(i) Then
                print_flag = 0
                Exit For
            End If
        Next i
    End If
    
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

Sub load_adit_data(pathin, input_files, spg_col_ct, spg_row_ct, spg_table, grp_col_ct, grp_row_ct, grp_table)

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

Open pathin + input_files(4) For Input As 1

While EOF(1) <> -1
    grp_row_ct = grp_row_ct + 1
    For i = 1 To grp_col_ct
        Input #1, grp_table(grp_row_ct, i)
        If grp_table(grp_row_ct, 1) = "TBL_BEGIN" Then grp_row_ct = 0
    Next i
    If grp_table(grp_row_ct, 1) = "TBL_END" Then grp_row_ct = grp_row_ct - 1
Wend
Close 1

End Sub
Sub order_tags(tag_ct, new_tag_dat, e_col_num)

Dim i, j As Integer

ReDim tag_dat(tag_ct, e_col_num) As String
ReDim tag_dat_new(200, e_col_num) As String

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
Dim i, ln As Integer
'Debug.Print loop_value_ct, last_loop_value_ct, loop_value(loop_value_ct)
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
                        
                        Print #2, saveframe_name; ","; loop_value(i); ","; loop_value(i + 1); ","
                        'Debug.Print saveframe_name, loop_value(i)
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

ReDim enumer(5000, 4) As String
ReDim enum_list(500, 3) As String
Open pathin + output_files(8) For Input As 1
While EOF(1) <> -1
    enum_ct = enum_ct + 1
    For i = 1 To 4
        Input #1, enumer(enum_ct, i)
    Next i
Wend
Close 1

Open pathin + output_files(9) For Output As 1
Print #1, "TBL_BEGIN,?,?"
k = 0
For j = 1 To e_col_num
    If excel_header_dat(1, j) = "Item enumerated" Then col_num(0) = j
    If excel_header_dat(1, j) = "SFCategory" Then col_num(1) = j
    If excel_header_dat(1, j) = "Tag" Then col_num(2) = j: tag_col = j
    If excel_header_dat(1, j) = "Enum parent SFcategory" Then col_num(3) = j
    If excel_header_dat(1, j) = "Enum parent tag" Then col_num(4) = j
Next j

For i = 1 To tag_ct
    If new_tag_dat(i, col_num(0)) = "Y" Then
        If new_tag_dat(i, col_num(3)) = "" Then
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
                    enumer(jk, 1) = enum_id
                End If
            Next jk
            'Debug.Print
        End If
    End If
Next i
For i = 1 To tag_ct
    If new_tag_dat(i, col_num(0)) = "Y" Then
        If new_tag_dat(i, col_num(3)) <> "" Then
            For j = 1 To enum_id
                If new_tag_dat(i, col_num(3)) = enum_list(j, 2) Then
                If new_tag_dat(i, col_num(4)) = enum_list(j, 3) Then
                    Print #1, enum_list(j, 1) + ",";
                    Print #1, new_tag_dat(i, col_num(1)) + ",";
                    Print #1, new_tag_dat(i, col_num(2))
                    'Debug.Print
                End If
                End If
            Next j
        End If
    End If
Next i
Print #1, "TBL_END,?,?"

Close 1
Open pathin + output_files(8) For Output As 1
Print #1, "TBL_BEGIN,?,?,"
For i = 1 To enum_ct
    enid = Val(Left$(enumer(i, 1), 1))
    If enid > 0 And enid < 10 Then
        Print #1, enumer(i, 1) + "," + enumer(i, 2) + "," + enumer(i, 3)
    End If
Next i
Print #1, "TBL_END,?,?,"
Close 1
End Sub

Sub write_group_tables(pathin, output_files, input_files)
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
    Print #2, table_row(i, 14) + ",";
    
    If table_row(i, 1) = "TBL_BEGIN" Or table_row(i, 1) = "TBL_END" Then
        Print #2, table_row(i, 15)
    Else
        Print #2, table_row(i, 15)
    End If

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
