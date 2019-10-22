Attribute VB_Name = "set_frm_star"
Sub set_from_star(tag_ct, tag_char, new_tag_dat, excel_header_dat, e_col_num, spg_col_ct, spg_row_ct, spg_table, grp_col_ct, grp_row_ct, grp_table)
Debug.Print "Set tag values from the STAR file"

Dim i, j, j1 As Integer

For i = 1 To tag_ct
'Debug.Print i, new_tag_dat(i, 2)
    If tag_char(i, 0) <> "y" Then
        Debug.Print i, new_tag_dat(i, 8), new_tag_dat(i, 9)
        Debug.Print
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
                new_tag_dat(i, j) = Trim(new_tag_dat(i, j))
            End If

            If excel_header_dat(1, j) = "Loopflag" Then
                If tag_char(i, 4) = "0" Then
                    new_tag_dat(i, j) = "N"
                End If
                If tag_char(i, 4) = "1" Then
                    new_tag_dat(i, j) = "Y"
                End If
            End If
            'If excel_header_dat(1, j) = "BMRB/CCPN status" Then
            '    If new_tag_dat(i, j) = "" Then new_tag_dat(i, j) = "open"
            'End If
            If excel_header_dat(1, j) = "Data Type" Then
                If new_tag_dat(i, j) = "" Then
                    If LCase(Right$(tag_char(i, 1), 7)) = "details" Then new_tag_dat(i, j) = "TEXT"
                    If LCase(Right$(tag_char(i, 1), 7)) = "keyword" Then new_tag_dat(i, j) = "VARCHAR(127)"
                    If LCase(Right$(tag_char(i, 1), 5)) = "label" Then new_tag_dat(i, j) = "VARCHAR(127)"
                    If LCase(Right$(tag_char(i, 1), 16)) = "text_data_format" Then new_tag_dat(i, j) = "VARCHAR(31)"
                    If LCase(Right$(tag_char(i, 1), 9)) = "text_data" Then new_tag_dat(i, j) = "TEXT"
                    If LCase(Right$(tag_char(i, 1), 2)) = "id" Then new_tag_dat(i, j) = "CHAR(12)"
                    If LCase(Right$(tag_char(i, 1), 9)) = "atom_name" Then new_tag_dat(i, j) = "VARCHAR(15)"
                    If LCase(Right$(tag_char(i, 1), 9)) = "atom_type" Then new_tag_dat(i, j) = "VARCHAR(15)"
                    If LCase(tag_char(i, 1)) = "sf_id" Then new_tag_dat(i, j) = "INTEGER"
                    If LCase(tag_char(i, 1)) = "sf_framecode" Then new_tag_dat(i, j) = "VARCHAR(127)"
                    If LCase(tag_char(i, 1)) = "sf_category" Then new_tag_dat(i, j) = "VARCHAR(31)"
                    If LCase(Right$(tag_char(i, 1), 13)) = "molecule_code" Then new_tag_dat(i, j) = "VARCHAR(15)"
                    If LCase(Right$(tag_char(i, 1), 14)) = "chem_comp_code" Then new_tag_dat(i, j) = "VARCHAR(15)"
                    If LCase(Right$(tag_char(i, 1), 4)) = "code" Then new_tag_dat(i, j) = "VARCHAR(15)"
                    If LCase(Right$(tag_char(i, 1), 7)) = "seq_num" Then new_tag_dat(i, j) = "INTEGER"
                    If LCase(Right$(tag_char(i, 1), 4)) = "date" Then new_tag_dat(i, j) = "DATETIME year to day"
                    If LCase(Right$(tag_char(i, 1), 3)) = "val" Then new_tag_dat(i, j) = "FLOAT"
                    If LCase(Right$(tag_char(i, 1), 3)) = "err" Then new_tag_dat(i, j) = "FLOAT"
                    If LCase(Right$(tag_char(i, 1), 7)) = "val_err" Then new_tag_dat(i, j) = "FLOAT"
                    If LCase(Right$(tag_char(i, 1), 9)) = "val_units" Then new_tag_dat(i, j) = "VARCHAR(31)"

                    If LCase(Right$(tag_char(i, 1), 4)) = "name" Then
                        If LCase(Right$(tag_char(i, 1), 9)) <> "atom_name" Then new_tag_dat(i, j) = "VARCHAR(127)"
                    End If
                End If
            End If
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
                            'If new_tag_dat(i, j - 3) = new_tag_dat(j1, j - 3) Then
                            '    new_tag_dat(j1, j + 3) = new_tag_dat(i, j - 4)
                            '    new_tag_dat(j1, j + 4) = new_tag_dat(i, j - 3)
                            '    Debug.Print
                            'End If
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
                    'If new_tag_dat(i, j - 3) = new_tag_dat(j1, j - 3) Then
                    '    new_tag_dat(j1, j + 3) = new_tag_dat(i, j - 4)
                    '    new_tag_dat(j1, j + 4) = new_tag_dat(i, j - 3)
                    '    Debug.Print
                    'End If
                    'p = InStr(1, LCase(new_tag_dat(j1, j - 3)), LCase(test_string))
                    'If p > 0 Then
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
    'Next j
Next i
End Sub
