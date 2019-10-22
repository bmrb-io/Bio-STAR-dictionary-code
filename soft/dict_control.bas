Attribute VB_Name = "control_prog"

Public run_prog, run_option As Integer

Sub main()

'DECLARE SUB dictionarychk (t$, token$(), linecount, tottoken%)
'NMR-STAR dictionary processing               (dict_control.bas  3/30/03)

'---------------------------------------


run_prog = 0
run_option = 0
Load launch_form
launch_form.Check1.Value = False
launch_form.Check2.Value = False
launch_form.Check3.Value = False
launch_form.Check4.Value = False
launch_form.Check5.Value = False
launch_form.Check6.Value = False
launch_form.Check7.Value = False
launch_form.Check8.Value = False
launch_form.Check9.Value = False

launch_form.Command1.Enabled = True
launch_form.Command2.Enabled = True
'launch_form.Command3.Enabled = True
launch_form.Command1.Value = False
launch_form.Command2.Value = False
'launch_form.Command3.Value = False
launch_form.Show

Load program_control
program_control.Command1.Enabled = True
program_control.Command2.Enabled = True
program_control.Command1.Value = False
program_control.Command2.Value = False

End Sub

Sub go_yahoo()

If run_option = 0 Then main

Dim pathin, pathout As String
Dim output_files(30), input_files(20) As String
Dim tag_ct, table_ct, sf_ct As Integer
ReDim tag_char(1, 3) As String

ReDim tag_list(10000), sf_list(200), table_list(500) As String

'control variables

'set up selection for each project

adit_file_source = 4    '0 = NMR-STAR v3.1 development (bmrb_star_v3_files)
                        '1 = NMR-STAR v4 development (bmrb_star_v4_files)
                        '2 = metabolomics (metabolomics_adit_files)
                        '3 = small molecule structure depositions (small_molecule_adit_files)
                        '4 = new BMRB only for ADIT-NMR (01/29/2014)
                        
If adit_file_source = 0 Then
    e_col_num = 106     'number of columns in the xlschem_ann.csv file
    
    pathmy = "\eldonulrich"
    'pathin = "z:\"+"eldonulrich\"+"bmrb\htdocs\dictionary\htmldocs\nmr_star\bmrb_star_v3_files\"
    pathout_adit = "z:" + pathmy + "\bmrb\htdocs\dictionary\htmldocs\nmr_star\bmrb_star_v3_files\"
    
    pathout2 = "z:\share\elu\dictionary2\bmrb_star_v3_files\adit_input\"
    pathout3 = "z:\share\elu\dictionary2\bmrb_star_v3_files\"
    pathout4 = "z:\" + "eldonulrich\" + "bmrb\htdocs\dictionary\htmldocs\nmr_star\bmrb_star_v3_files\"

End If
If adit_file_source = 1 Then
    e_col_num = 106     'number of columns in the xlschem_ann.csv file

    pathmy = "\eldonulrich"
    'pathin = "z:\"+"eldonulrich\"+"bmrb\htdocs\dictionary\htmldocs\nmr_star\bmrb_star_v4_files\"
    pathout_adit = "z:" + pathmy + "\bmrb\htdocs\dictionary\htmldocs\nmr_star\bmrb_star_v4_files\"
    pathout2 = "z:\share\elu\dictionary2\bmrb_star_v4_files\adit_input\"
    pathout3 = "z:\share\elu\dictionary2\bmrb_star_v4_files\"
    pathout4 = "z:" + pathmy + "\bmrb\htdocs\dictionary\htmldocs\nmr_star\bmrb_star_v4_files\"

End If
If adit_file_source = 2 Then
    e_col_num = 106     'number of columns in the xlschem_ann.csv file

    pathin = "z:\" + "eldonulrich\" + "bmrb\htdocs\dictionary\htmldocs\nmr_star\development_adit_files\"
    pathout_adit = "z:\" + "eldonulrich\" + "bmrb\htdocs\dictionary\htmldocs\nmr_star\metabolomics_adit_files\"
    pathout2 = "g:\elu\dictionary2\metabolomics_adit_files\adit_input\"
    pathout3 = "g:\elu\dictionary2\metabolomics_adit_files\"
    pathout4 = "z:\" + "eldonulrich\" + "bmrb\htdocs\dictionary\htmldocs\nmr_star\metabolomics_adit_files\"

End If
If adit_file_source = 3 Then
    e_col_num = 106     'number of columns in the xlschem_ann.csv file

    pathin = "z:\" + "eldonulrich\" + "bmrb\htdocs\dictionary\htmldocs\nmr_star\development_adit_files\"
    pathout_adit = "z:\" + "eldonulrich\" + "bmrb\htdocs\dictionary\htmldocs\nmr_star\small_molecule_adit_files\"
    pathout2 = "g:\elu\dictionary2\small_molecule_adit_files\adit_input\"
    pathout3 = "g:\elu\dictionary2\small_molecule_adit_files\"
    pathout4 = "z:\" + "eldonulrich\" + "bmrb\htdocs\dictionary\htmldocs\nmr_star\small_molecule_adit_files\"
End If
If adit_file_source = 4 Then
    e_col_num = 106     'number of columns in the xlschem_ann.csv file

    pathmy = "\eldonulrich"
    pathout_adit = "z:" + pathmy + "\bmrb\htdocs\dictionary\htmldocs\nmr_star\NMR-STAR\internal_106_source\"
    pathout3 = "z:" + pathmy + "\bmrb\htdocs\dictionary\htmldocs\nmr_star\NMR-STAR\internal_106_distribution\"
    pathout4 = "z:" + pathmy + "\bmrb\htdocs\dictionary\htmldocs\nmr_star\NMR-STAR\source\"

End If

If run_prog = 1 Then
    launch_form.Hide

'Debug.Print "stop for now"
'Debug.Print

dictionary_name = "NMR-STAR"

input_files(1) = "nmrstar3_edit.txt"
input_files(2) = "xlschem_ann.csv"                      '"xlschem_extended.csv"
input_files(3) = "adit_super_grp_i.csv"                      'adit_super_grp_i.csv"
input_files(4) = "adit_cat_grp_i.csv"                        'adit_cat_grp_i.csv"
input_files(5) = "enumerations.txt"
input_files(7) = "adit_man_over.csv"                            'adit_man_over.csv"
'input_files(7) = "adit_man_over_add.csv"
If adit_file_source = 0 Then input_files(8) = "nmrstar_v3_full_exmpl.str"
If adit_file_source = 1 Then input_files(8) = "nmrstar_v4_full_exmpl.str"
If adit_file_source = 4 Then input_files(8) = "nmrstar_v3_full_exmpl.str"
input_files(9) = "nmr_cif_match.csv"
If adit_file_source = 0 Then input_files(10) = "old_xlschem_ann.csv"
If adit_file_source = 4 Then input_files(10) = "old_xlschem_ann.csv"

If adit_file_source = 2 Then                        'metabolomics
    'input_files(5) = "enumerations_met.txt"
    input_files(7) = "adit_input\adit_man_over_add_met.csv"

        FileCopy pathin + input_files(2), pathout4 + input_files(2)
        FileCopy pathin + input_files(3), pathout4 + input_files(3)
        FileCopy pathin + input_files(4), pathout4 + input_files(4)
        FileCopy pathin + input_files(5), pathout4 + input_files(5)
        FileCopy pathin + input_files(7), pathout4 + input_files(7)
        FileCopy pathin + input_files(8), pathout4 + input_files(8)
        'FileCopy pathin + "adit_input/" + input_files(9), pathout4 + "adit_input/"+input_files(9)
        'FileCopy pathin + "adit_category_desc.csv", pathout4 + "adit_category_desc.csv"
        'FileCopy pathin + "adit_enum_dict.csv", pathout4 + "adit_enum_dict.csv"

End If
If adit_file_source = 3 Then                        'small molecule structure
    input_files(5) = "enumerations.txt"
    input_files(7) = "adit_man_over_add_sm.csv"

        FileCopy pathin + input_files(2), pathout4 + input_files(2)
        FileCopy pathin + input_files(3), pathout4 + input_files(3)
        FileCopy pathin + input_files(4), pathout4 + input_files(4)
        FileCopy pathin + input_files(5), pathout4 + input_files(5)
        FileCopy pathin + input_files(7), pathout4 + input_files(7)
        FileCopy pathin + input_files(8), pathout4 + input_files(8)
        FileCopy pathin + "adit_input/" + input_files(9), pathout4 + "adit_input/" + input_files(9)
        FileCopy pathin + "adit_category_desc.csv", pathout4 + "adit_category_desc.csv"
        FileCopy pathin + "adit_enum_dict.csv", pathout4 + "adit_enum_dict.csv"

End If

output_files(1) = "adit_interface_dict.txt"                     'adit_interface_dict.txt"
output_files(2) = "new_xlschem.csv"
output_files(3) = "adit_item_tbl_o.csv"                          'adit_item_tbl_o.csv"
output_files(4) = "adit_cat_grp_o.csv"                           'adit_cat_grp_o.csv"
output_files(5) = "adit_super_grp_o.csv"                         'adit_super_grp_o.csv"
output_files(6) = "NMR-STAR.dic"
If adit_file_source = 0 Then output_files(7) = "nmrstar3_anno.txt"
If adit_file_source = 1 Then output_files(7) = "nmrstar4_anno.txt"
If adit_file_source = 4 Then output_files(7) = "nmrstar3_anno.txt"
output_files(8) = "adit_enum_dtl.csv"                            'adit_enum_dtl.csv"
output_files(9) = "adit_enum_hdr.csv"                            'adit_enum_hdr.csv"
output_files(10) = "table_err.txt"
output_files(11) = "tag_err.txt"
output_files(12) = "adit_enum_ties.csv"                          'adit_enum_ties.csv"
output_files(13) = "adit_man_over.csv"                           'adit_man_over.csv"
If adit_file_source = 0 Then output_files(14) = "nmrstar3.txt"
If adit_file_source = 1 Then output_files(14) = "nmrstar4.txt"
If adit_file_source = 4 Then output_files(14) = "nmrstar3.txt"
If adit_file_source = 0 Then output_files(15) = "nmrstar3_fake.txt"
If adit_file_source = 1 Then output_files(15) = "nmrstar4_fake.txt"
If adit_file_source = 4 Then output_files(15) = "nmrstar3_fake.txt"
output_files(16) = "nmrstar_sg_dict.txt"
If adit_file_source = 0 Then output_files(17) = "nmrstar3.dic"
If adit_file_source = 1 Then output_files(17) = "nmrstar4.dic"
If adit_file_source = 4 Then output_files(17) = "nmrstar3.dic"
output_files(18) = "adit_tag_validation.csv"                            'adit_tag_validation.csv"
'output_files(19) = "ccpn_excel.csv"
'output_files(20) = "default-entry.cif"
output_files(21) = "adit_enum_dict.csv"                                  'adit_enum_dict.csv"
output_files(22) = "adit_category_desc.csv"                              'adit_category_desc.csv"
output_files(23) = "NMR-STAR_internal.dic"
'output_files(24) = "val_overide_add.csv"
output_files(25) = "val_item_tbl.csv"
output_files(26) = "commonDA_tag.dic"
'If adit_file_source = 0 Then output_files(27) = "dict_diff_report.txt"
'If adit_file_source = 4 Then output_files(27) = "dict_diff_report.txt"
'output_files(28) = "BMRB-ML.xsd"
output_files(29) = "query_interface.csv"
output_files(30) = "NMR-STAR_28_xlschem_master.csv"

'set up variables and arrays

spg_col_ct = 8: spg_row_ct = 0      'number of columns in the ADIT super group excel file
grp_col_ct = 28: grp_row_ct = 0     'number of columns in the ADIT group excel file


    If launch_form.Check1.Value = 1 Then
        Debug.Print "Excel file integrity check"
        clear_fields
        'run_option = 1
        ReDim grp_table(200, grp_col_ct) As String
        ReDim spg_table(20, spg_col_ct) As String

        Open pathout_adit + output_files(10) For Output As 10
        Open pathout_adit + output_files(11) For Output As 11
        
        pathin = pathout_adit
        input_file = input_files(2)
        
        syntax_check.Text1.Text = input_file
        syntax_check.Text2.Text = pathin
        syntax_check.Show
        syntax_check.Refresh
        
    'write_dict_comp e_tag_ct, excel_tag_dat, excel_header_dat, pathin, input_file, e_col_num, input_files
        
        load_adit_data pathin, input_files, spg_col_ct, spg_row_ct, spg_table, grp_col_ct, grp_row_ct, grp_table, adit_file_source

        read_excel e_tag_ct, excel_tag_dat, excel_header_dat, pathin, input_file, e_col_num
        
        check_excel e_tag_ct, excel_tag_dat, excel_header_dat, e_col_num, spg_col_ct, spg_row_ct, spg_table, grp_col_ct, grp_row_ct, grp_table, pathin, adit_file_source
        
        'read_star tag_list, tag_ct, tag_char, table_list, table_ct, sf_list, sf_ct, pathin, input_file, run_option, tag_desc

        'Debug.Print
        
        'Dictionary comparison (New vs. Old)
        
        'write_dict_comp
        
        FileCopy pathin + input_files(2), pathout3 + input_files(2)                     '"xlschem_ann.csv"
        
    End If

    If launch_form.Check2.Value = 1 Then
        Debug.Print "Option #2:  Write NMR-STAR dictionary"
        run_option = 3
        pathin = pathout_adit
        launch_form.Hide
        
        ReDim grp_table(200, grp_col_ct) As String
        ReDim spg_table(20, spg_col_ct) As String
                
        load_adit_data pathin, input_files, spg_col_ct, spg_row_ct, spg_table, grp_col_ct, grp_row_ct, grp_table, adit_file_source

        
        input_file = input_files(2)
        read_excel e_tag_ct, excel_tag_dat, excel_header_dat, pathin, input_file, e_col_num
        
        input_file = input_files(5)
        Open pathin + output_files(21) For Output As 2

        run_option = 15
        read_star tag_list, tag_ct, tag_char, table_list, table_ct, sf_list, sf_ct, pathin, input_file, run_option, tag_desc
        Close 2
        
        'input_file = output_files(1)
        'read_star tag_list, tag_ct, tag_char, table_list, table_ct, sf_list, sf_ct, pathin, input_file, run_option, tag_desc

        'Open pathout_adit + output_files(17) For Output As 13        'for limited dictionary
        'write_nmrstar_dict e_tag_ct, excel_tag_dat, tag_ct, tag_list
        
        lcase_flag = 1: full_dict_flag = 0 'lower case tags and do not print non-public tags
        Open pathout_adit + "adit_input/" + output_files(6) For Output As 13
        write_nmrstar_dict2 e_tag_ct, excel_tag_dat, tag_ct, tag_list, e_col_num, pathin, output_files, lcase_flag, full_dict_flag, dictionary_name

        Close 13
        
        Open pathout_adit + "adit_input/" + "nmr_star_val_dict.str" For Output As 13
        write_nmrstar_dict3 e_tag_ct, excel_tag_dat, tag_ct, tag_list, e_col_num, pathin, output_files, lcase_flag, full_dict_flag, input_files, excel_header_dat

        Close 13
        
       
        
        lcase_flag = 0: full_dict_flag = 1 'Do not lower case tags and print all tags including the non-public tags
        Open pathout_adit + "adit_input/" + output_files(23) For Output As 13
        write_nmrstar_dict2 e_tag_ct, excel_tag_dat, tag_ct, tag_list, e_col_num, pathin, output_files, lcase_flag, full_dict_flag, dictionary_name
        

        Close 13
        
    End If
    
    If launch_form.Check3.Value = 1 Then
        Debug.Print "Write interface items"
        run_option = 3
        pathin = pathout_adit
        launch_form.Hide
        'Tag_def.Show
        
        input_file = input_files(2)
        read_excel e_tag_ct, excel_tag_dat, excel_header_dat, pathin, input_file, e_col_num

        input_file = output_files(1)

        read_star tag_list, tag_ct, tag_char, table_list, table_ct, sf_list, sf_ct, pathin, input_file, run_option, tag_desc

   '     run_option = 13
        Open pathout_adit + output_files(1) For Append As 13
        update_interface e_tag_ct, excel_tag_dat, tag_ct, tag_list, adit_file_source

   '     pathin = pathout_adit
   '     input_file = input_files(1)
   '     read_star tag_list, tag_ct, tag_char, table_list, table_ct, sf_list, sf_ct, pathin, input_file, run_option,tag_desc
        Close 13
        
    End If
    
    If launch_form.Check4.Value = 1 Then
        Debug.Print "Option #4:  Update from the Excel file"
        pathin = pathout_adit
        
        ReDim tag_char(10000, 4) As String
        ReDim new_tag_dat(10000, e_col_num)
        ReDim grp_table(200, grp_col_ct) As String
        ReDim spg_table(20, spg_col_ct) As String
        
        write_group_tables pathin, output_files, input_files, adit_file_source
        
        
        tag_ct = 0
        'read_star tag_list, tag_ct, tag_char, table_list, table_ct, sf_list, sf_ct, pathin, input_file, run_option, tag_desc
    
        input_file = input_files(2)
        read_excel e_tag_ct, excel_tag_dat, excel_header_dat, pathin, input_file, e_col_num
        
        For i = 1 To e_tag_ct
            For j = 1 To e_col_num
                new_tag_dat(i, j) = excel_tag_dat(i, j)
                'Debug.Print new_tag_dat(i, j)
            Next j
            p = InStr(1, new_tag_dat(i, 9), ".")
            category1 = Left$(new_tag_dat(i, 9), p - 1)
            tag1 = Right$(new_tag_dat(i, 9), Len(new_tag_dat(i, 9)) - p)

            tag_char(i, 1) = tag1
            tag_char(i, 2) = category1
            tag_char(i, 3) = new_tag_dat(i, 2)
            tag_char(i, 4) = new_tag_dat(i, 43)

        Next i
        tag_ct = e_tag_ct
        
   '     compare_tags tag_ct, tag_char, e_tag_ct, excel_tag_dat, new_tag_dat, e_col_num
        
        load_adit_data pathin, input_files, spg_col_ct, spg_row_ct, spg_table, grp_col_ct, grp_row_ct, grp_table, adit_file_source
        
' temporary placement for the write_xml subroutine while debugging
        'write_xml output_files, e_tag_ct, excel_tag_dat, excel_header_dat, pathin, input_file, e_col_num, spg_col_ct, spg_row_ct, spg_table, grp_col_ct, grp_row_ct, grp_table, adit_file_source
        
        set_tag_values tag_ct, tag_char, new_tag_dat, excel_header_dat, e_col_num, spg_col_ct, spg_row_ct, spg_table, grp_col_ct, grp_row_ct, grp_table
        
        order_tags tag_ct, new_tag_dat, e_col_num
        
        output_file = output_files(2)
        write_new_excel pathin, output_file, tag_ct, new_tag_dat, excel_header_dat, e_col_num, adit_file_source
        
        output_file = output_files(19)
        'write_ccpn_excel pathin, output_file, tag_ct, new_tag_dat, excel_header_dat, e_col_num
        
        write_new_adit_files pathin, output_files, tag_ct, new_tag_dat, e_col_num, excel_header_dat, adit_file_source
        
        write_validation_files pathin, output_files, tag_ct, new_tag_dat, e_col_num, excel_header_dat, adit_file_source
        
        write_query_interface_file pathin, output_files, tag_ct, new_tag_dat, e_col_num, excel_header_dat, adit_file_source
        
        write_tag_validation pathin, output_files, tag_ct, new_tag_dat, e_col_num
        
        write_anno_star pathin, output_files, tag_ct, new_tag_dat, e_col_num
        
        write_template_star pathin, output_files, tag_ct, new_tag_dat, e_col_num

        write_fake_star pathin, output_files, tag_ct, new_tag_dat, e_col_num
        
        Debug.Print "Writing ADIT enumerations file"
        input_file = input_files(5)
        Open pathin + output_files(8) For Output As 2
        
        run_option = 14
        read_star tag_list, tag_ct, tag_char, table_list, table_ct, sf_list, sf_ct, pathin, input_file, run_option, tag_desc
        
        Close 2
        
        write_enum_hdr pathin, output_files, tag_ct, new_tag_dat, e_col_num, excel_header_dat
        write_enum_ties pathin, output_files, tag_ct, new_tag_dat, e_col_num
        
   '     If adit_file_source = 0 Then write_man_over pathin, input_files, output_files, tag_ct, new_tag_dat, e_col_num, grp_col_ct, grp_row_ct, grp_table
   '     If adit_file_source = 1 Then write_man_over pathin, input_files, output_files, tag_ct, new_tag_dat, e_col_num, grp_col_ct, grp_row_ct, grp_table
'        If adit_file_source = 1 Then write_man_over_prod pathin, input_files, output_files, tag_ct, new_tag_dat, e_col_num, grp_col_ct, grp_row_ct, grp_table
   '     If adit_file_source = 2 Then write_man_over pathin, input_files, output_files, tag_ct, new_tag_dat, e_col_num, grp_col_ct, grp_row_ct, grp_table
   '     If adit_file_source = 3 Then write_man_over pathin, input_files, output_files, tag_ct, new_tag_dat, e_col_num, grp_col_ct, grp_row_ct, grp_table
   '     If adit_file_source = 4 Then write_man_over pathin, input_files, output_files, tag_ct, new_tag_dat, e_col_num, grp_col_ct, grp_row_ct, grp_table
        
        Debug.Print "Writing interface items"
        'retrieve new tag data from Excel file xlschem_ann.csv
        input_file = input_files(2)
'        read_excel e_tag_ct, excel_tag_dat, excel_header_dat, pathin, input_file, e_col_num

        'retrieve tags from existing dictionary file
        'input_file = output_files(1)
        'read_star tag_list, tag_ct, tag_char, table_list, table_ct, sf_list, sf_ct, pathin, input_file, run_option, tag_desc
        Close

        'write out new dictionary either appending information or creating a completely new file
        
        'just added
        input_file = input_files(5)
        Open pathin + output_files(21) For Output As 2

        run_option = 15
        read_star tag_list, tag_ct, tag_char, table_list, table_ct, sf_list, sf_ct, pathin, input_file, run_option, tag_desc
        Close 2

        'Open pathout_adit + output_files(1) For Append As 13
        Open pathout_adit + output_files(1) For Output As 13
        update_interface e_tag_ct, excel_tag_dat, tag_ct, tag_list, adit_file_source
        
        Close
        lcase_flag = 1: full_dict_flag = 0 'lower case tags and do not print non-public tags
        Open pathout_adit + output_files(6) For Output As 13
        write_nmrstar_dict2 e_tag_ct, excel_tag_dat, tag_ct, tag_list, e_col_num, pathin, output_files, lcase_flag, full_dict_flag, dictionary_name

        Close 13
        
        Open pathout_adit + "nmr_star_val_dict.str" For Output As 13
        write_nmrstar_dict3 e_tag_ct, excel_tag_dat, tag_ct, tag_list, e_col_num, pathin, output_files, lcase_flag, full_dict_flag, input_files, excel_header_dat

        Close 13
        
        'write out tags for incorporation into master pdbx dictionary
        Open pathout_adit + output_files(26) For Output As 13
        write_nmrstar_dict4 e_tag_ct, excel_tag_dat, tag_ct, tag_list, e_col_num, pathin, output_files, lcase_flag, full_dict_flag, input_files, excel_header_dat

        Close 13

        Close
        
        lcase_flag = 0: full_dict_flag = 1 'Do not lower case tags and print all tags including the non-public tags
        Open pathout_adit + output_files(23) For Output As 13
        write_nmrstar_dict2 e_tag_ct, excel_tag_dat, tag_ct, tag_list, e_col_num, pathin, output_files, lcase_flag, full_dict_flag, dictionary_name

        Close
        
        write_28_column_excel_file excel_tag_dat, excel_header_dat, e_col_num, e_tag_ct, output_files, pathout_adit
        
        'Write dictionary update file
        'Open pathout_adit + "adit_input/" + output_files(27) For Output As 13
        
        'write_dict_comp e_tag_ct, excel_tag_dat, excel_header_dat, pathin, input_file, e_col_num, input_files
  
        Close
        
'        write_xml pathin, output_files, e_tag_ct, excel_tag_dat, excel_header_dat, pathin, input_file, e_col_num
        
'Copy files from "internal_106_source" to "source"

        FileCopy pathin + "adit_nmr_upload_tags.csv", pathout4 + "NMR-STAR_nmr_upload_tags.csv" 'This file manually edited only not by software

        'FileCopy pathin + input_files(2), pathout4 + "adit_input\" + input_files(2)        '"xlschem_ann.csv"
        FileCopy pathin + input_files(3), pathout4 + "NMR-STAR_super_grp_i.csv"             '"adit_super_grp_i.csv"
        FileCopy pathin + input_files(4), pathout4 + "NMR-STAR_cat_grp_i.csv"               '"adit_cat_grp_i.csv"
        FileCopy pathin + input_files(5), pathout4 + "NMR-STAR_enumerations.txt"            '"enumerations.txt"
        'FileCopy pathin + input_files(6), pathout4 + "adit_input\" + input_files(6)
        'FileCopy pathin + input_files(7), pathout4 + "adit_input\" + input_files(7)        '"adit_man_over_add_pdb_chem_shift_man.csv"
        'FileCopy pathin + input_files(8), pathout4 + "adit_input\" + input_files(8)        '"nmrstar_v3_full_exmpl.str"
        'FileCopy pathin + input_files(9), pathout4 + input_files(9)                        '"nmr_cif_match.csv"
        'FileCopy pathin + input_files(10), pathout4 + "adit_input\" + input_files(10)      '"old_xlschem_ann.csv"
        
        FileCopy pathin + output_files(1), pathout4 + "NMR-STAR_interface_dict.txt"         '"adit_interface_dict.txt"
        'FileCopy pathin + output_files(2), pathout4 + "adit_input\" + output_files(2)      '"new_xlschem.csv"
        FileCopy pathin + output_files(4), pathout4 + "NMR-STAR_cat_grp_o.csv"              '"adit_cat_grp_o.csv"
        FileCopy pathin + output_files(5), pathout4 + "NMR-STAR_super_grp_o.csv"            '"adit_super_grp_o.csv"
        'FileCopy pathin + output_files(6), pathout4 + "adit_input\" + output_files(6)      '"NMR-STAR.dic"
        
        'FileCopy pathin + output_files(7), pathout4 + "adit_input\" + output_files(7)      '"nmrstar3_anno.txt"
        'FileCopy pathin + output_files(10), pathout4 + "adit_input\" + output_files(10)    '"table_err.txt"
        'FileCopy pathin + output_files(11), pathout4 + "adit_input\" + output_files(11)    '"tag_err.txt"
        'FileCopy pathin + output_files(14), pathout4 + "adit_input\" + output_files(14)    '"nmrstar3.txt"
        FileCopy pathin + output_files(15), pathout4 + "NMR-STAR_fake_file.txt"             '"nmrstar3_fake.txt"
        
        'FileCopy pathin + output_files(20), pathout4 + "adit_input\" + output_files(20)    '"default-entry.cif"
        FileCopy pathin + output_files(21), pathout4 + "NMR-STAR_enum_dict.csv"             '"adit_enum_dict.csv"
        FileCopy pathin + output_files(22), pathout4 + "NMR-STAR_category_desc.csv"         '"adit_category_desc.csv"
        FileCopy pathin + output_files(23), pathout4 + output_files(23)                     '"NMR-STAR_internal.dic"
        'FileCopy pathin + output_files(24), pathout4 + "adit_input\" + output_files(24)    '"val_overide_add.csv"
        'FileCopy pathin + output_files(25), pathout4 + "adit_input\" + output_files(25)    '"val_item_tbl.csv"
        'FileCopy pathin + output_files(26), pathout4 + "adit_input\" + output_files(26)    '"commonDA_tag.dic"
        'FileCopy pathin + output_files(27), pathout4 + "adit_input\" + output_files(27)    '"dict_diff_report.txt"
        'FileCopy pathin + output_files(28), pathout4 + "adit_input\" + output_files(28)    '"BMRB-ML.xsd"
        FileCopy pathin + output_files(29), pathout4 + "NMR-STAR_query_interface.csv"       '"query_interface.csv"
        FileCopy pathin + output_files(30), pathout4 + "NMR-STAR_xlschem_master.csv"        '"NMR-STAR_28_xlschem_master.csv"
        'If adit_file_source = 4 Then FileCopy pathin + input_files(8), pathout4 + "adit_input\" + input_files(8)   '"adit_enum_dtl.csv"
        
'Copy files from "internal_106_source" to "internal_106_distribution"

        FileCopy pathin + "adit_nmr_upload_tags.csv", pathout3 + "adit_nmr_upload_tags.csv" 'This file manually edited only not by software

        FileCopy pathin + input_files(2), pathout3 + input_files(2)                     '"xlschem_ann.csv"
        FileCopy pathin + input_files(3), pathout3 + "adit_super_grp_i.csv"             '"adit_super_grp_i.csv"
        FileCopy pathin + input_files(4), pathout3 + "adit_cat_grp_i.csv"               '"adit_cat_grp_i.csv"
        FileCopy pathin + input_files(5), pathout3 + "enumerations.txt"                 '"enumerations.txt"
        ''FileCopy pathin + input_files(6), pathout3 + "adit_input\" + input_files(6)
        FileCopy pathin + input_files(7), pathout3 + input_files(7)                     '"adit_man_over.csv"
        'FileCopy pathin + input_files(8), pathout3 + "adit_input\" + input_files(8)    '"nmrstar_v3_full_exmpl.str"
        FileCopy pathin + input_files(9), pathout3 + input_files(9)                     '"nmr_cif_match.csv"
        'FileCopy pathin + input_files(10), pathout3 + "adit_input\" + input_files(10)  '"old_xlschem_ann.csv"
        
        FileCopy pathin + output_files(1), pathout3 + "adit_interface_dict.txt"         '"adit_interface_dict.txt"
        'FileCopy pathin + output_files(2), pathout3 + output_files(2)                   '"new_xlschem.csv"
        FileCopy pathin + output_files(3), pathout3 + "adit_item_tbl_o.csv"             '"adit_item_tbl_o.csv"
        FileCopy pathin + output_files(4), pathout3 + "adit_cat_grp_o.csv"              '"adit_cat_grp_o.csv"
        FileCopy pathin + output_files(5), pathout3 + "adit_super_grp_o.csv"            '"adit_super_grp_o.csv"
        FileCopy pathin + output_files(6), pathout3 + output_files(6)                   '"NMR-STAR.dic"
        
        FileCopy pathin + output_files(7), pathout3 + output_files(7)                       '"nmrstar3_anno.txt"
        FileCopy pathin + output_files(8), pathout3 + output_files(8)                       '"adit_enum_dtl.csv"
        FileCopy pathin + output_files(9), pathout3 + output_files(9)                       '"adit_enum_hdr.csv"
        FileCopy pathin + output_files(10), pathout3 + output_files(10)                     '"table_err.txt"
        FileCopy pathin + output_files(11), pathout3 + output_files(11)                     '"tag_err.txt"
        FileCopy pathin + output_files(12), pathout3 + output_files(12)                     '"adit_enum_ties.csv"
        FileCopy pathin + output_files(14), pathout3 + output_files(14)                     '"nmrstar3.txt"
        FileCopy pathin + output_files(15), pathout3 + "nmrstar3_fake_file.txt"             '"nmrstar3_fake.txt"
        
        'FileCopy pathin + output_files(20), pathout3 + "adit_input\" + output_files(20)    '"default-entry.cif"
        FileCopy pathin + output_files(21), pathout3 + "adit_enum_dict.csv"                 '"adit_enum_dict.csv"
        FileCopy pathin + output_files(22), pathout3 + "adit_category_desc.csv"             '"adit_category_desc.csv"
        FileCopy pathin + output_files(23), pathout3 + output_files(23)                     '"NMR-STAR_internal.dic"
        'FileCopy pathin + output_files(24), pathout3 + "adit_input\" + output_files(24)    '"val_overide_add.csv"
        'FileCopy pathin + output_files(25), pathout3 + "adit_input\" + output_files(25)    '"val_item_tbl.csv"
        'FileCopy pathin + output_files(26), pathout3 + "adit_input\" + output_files(26)    '"commonDA_tag.dic"
        'FileCopy pathin + output_files(27), pathout3 + "adit_input\" + output_files(27)    '"dict_diff_report.txt"
        'FileCopy pathin + output_files(28), pathout3 + "adit_input\" + output_files(28)    '"BMRB-ML.xsd"
        FileCopy pathin + output_files(29), pathout3 + "query_interface.csv"                '"query_interface.csv"
        'FileCopy pathin + output_files(30), pathout3 + "NMR-STAR_xlschem_master.csv"        '"NMR-STAR_28_xlschem_master.csv"
                
        Debug.Print "tag count ="; tag_ct
        
    End If
    
    If launch_form.Check5.Value = 1 Then
        Debug.Print "Update keys in Excel file"
        pathin = pathout_adit
        input_file = output_files(3)
    
    
    End If
    
    If launch_form.Check6.Value = 1 Then
        Debug.Print "Generate ADIT files"
        pathin = pathout_adit
        input_file = output_files(3)
    
    
    End If
    
    If launch_form.Check7.Value = 1 Then
        Debug.Print "Update from the NMR-STAR file"
        pathin = pathout_adit
        input_file = input_files(1)
        ReDim tag_char(10000, 4) As String
        ReDim new_tag_dat(10000, e_col_num)
        ReDim grp_table(200, grp_col_ct) As String
        ReDim spg_table(20, spg_col_ct) As String
        
        tag_ct = 0
        read_star tag_list, tag_ct, tag_char, table_list, table_ct, sf_list, sf_ct, pathin, input_file, run_option, tag_desc
    
        input_file = input_files(2)
        read_excel e_tag_ct, excel_tag_dat, excel_header_dat, pathin, input_file, e_col_num
        
        compare_tags tag_ct, tag_char, e_tag_ct, excel_tag_dat, new_tag_dat, e_col_num
        
        load_adit_data pathin, input_files, spg_col_ct, spg_row_ct, spg_table, grp_col_ct, grp_row_ct, grp_table, adit_file_source
        
        set_from_star tag_ct, tag_char, new_tag_dat, excel_header_dat, e_col_num, spg_col_ct, spg_row_ct, spg_table, grp_col_ct, grp_row_ct, grp_table
        
        order_tags tag_ct, new_tag_dat, e_col_num
        
        output_file = output_files(2)
        write_new_excel pathin, output_file, tag_ct, new_tag_dat, excel_header_dat, e_col_num, adit_file_source
        
        Debug.Print "tag count ="; tag_ct
    
    End If
    
    If launch_form.Check8.Value = 1 Then
        Debug.Print "Write a SG stub dictionary"
        run_option = 8
        pathin = pathout_adit
        launch_form.Hide
        
        input_file = input_files(6)
        read_excel e_tag_ct, excel_tag_dat, excel_header_dat, pathin, input_file, e_col_num

        Open pathout_adit + output_files(16) For Output As 13
        write_sg_dict e_tag_ct, excel_tag_dat, tag_ct, tag_list

        Close 13
    End If
    
    If launch_form.Check9.Value = 1 Then
        Debug.Print "Checking a deposition file for tag consistency"
        run_option = 9
        pathin = pathout_adit
        launch_form.Hide
        
        input_file = input_files(2)
        read_excel e_tag_ct, excel_tag_dat, excel_header_dat, pathin, input_file, e_col_num
        
        pathin = "z:\" + "eldonulrich\" + "bmrb\htdocs\dictionary\htmldocs\nmr_star\mike_baran_test\PfR48_example\bmr6173\"
        input_file = "bmr6173.str_3.0_elu_12.txt"
        read_star tag_list, tag_ct, tag_char, table_list, table_ct, sf_list, sf_ct, pathin, input_file, run_option, tag_desc
        compare_dep_tags tag_ct, tag_char, e_tag_ct, excel_tag_dat, new_tag_dat, e_col_num
        Close
    End If
End If

Debug.Print "Finished"

End Sub

Sub clear_fields()

syntax_check.Text1.Text = ""
syntax_check.Text2.Text = ""
syntax_check.Text3.Text = ""
syntax_check.Text4.Text = ""
syntax_check.Text5.Text = ""
syntax_check.Text6.Text = ""
syntax_check.Text7.Text = ""
syntax_check.Text8.Text = ""
syntax_check.Text9.Text = ""
syntax_check.Text10.Text = ""
syntax_check.Text11.Text = ""
syntax_check.Text12.Text = ""
syntax_check.Text13.Text = ""
syntax_check.Text14.Text = ""
syntax_check.Text15.Text = ""
syntax_check.Text16.Text = ""
syntax_check.Text17.Text = ""
syntax_check.Text18.Text = ""
syntax_check.Text19.Text = ""
syntax_check.Text20.Text = ""
syntax_check.Text21.Text = ""
syntax_check.Text22.Text = ""

syntax_check.Refresh


End Sub


