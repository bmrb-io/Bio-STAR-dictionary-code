Attribute VB_Name = "control_prog"

Public run_prog, run_option As Integer

Sub main()

'BMOD-STAR dictionary processing

'---------------------------------------

run_prog = 0
run_option = 0
Load launch_form

launch_form.Check1.Value = False
launch_form.Check4.Value = False

launch_form.Command1.Enabled = True
launch_form.Command2.Enabled = True

launch_form.Command1.Value = False
launch_form.Command2.Value = False

launch_form.Show

Load program_control
program_control.Command1.Enabled = True
program_control.Command2.Enabled = True
program_control.Command1.Value = False
program_control.Command2.Value = False

End Sub

Sub go_yahoo()

If run_option = 0 Then main

Dim path_source, path_distribution, dn As String
Dim output_files(30), input_files(20) As String
Dim tag_ct, table_ct, sf_ct As Integer
ReDim tag_char(1, 3) As String

ReDim tag_list(10000), sf_list(200), table_list(500) As String

'file_source = 0    '0 = BMOD (BMOD directory)
'file_source = 1    '1 = NEX (NEX directory)
file_source = 2    '2 = NMR-STAR (BMRB directory)

If file_source = 0 Then        'BMOD dictionary
    e_col_num = 28     'number of columns in the xlschem_ann.csv file
    dn = "BMOD"
    dictionary_name = "BMOD-STAR"

    pathmy = "\eldonulrich\bmrb\htdocs\dictionary\htmldocs\nmr_star\"
    path_source = "z:" + pathmy + dn + "\" + "source\"
    path_distribution = "z:" + pathmy + dn + "\" + "distribution\"

End If

If file_source = 1 Then        'NEX-STAR dictionary
    e_col_num = 28     'number of columns in the xlschem_ann.csv file
    dn = "NEX"
    dictionary_name = "NEX-STAR"

    pathmy = "\eldonulrich\bmrb\htdocs\dictionary\htmldocs\nmr_star\"
    path_source = "z:" + pathmy + dn + "\" + "source\"
    path_distribution = "z:" + pathmy + dn + "\" + "distribution\"

End If

If file_source = 2 Then        'NMR-STAR dictionary
    e_col_num = 28     'number of columns in the xlschem_ann.csv file
    dn = "NMR-STAR"
    dictionary_name = "NMR-STAR"

    pathmy = "\eldonulrich\bmrb\htdocs\dictionary\htmldocs\nmr_star\"
    path_source = "z:" + pathmy + dn + "\" + "source\"
    path_distribution = "z:" + pathmy + dn + "\" + "distribution\"

End If

If run_prog = 1 Then
    launch_form.Hide

input_files(2) = dn + "_xlschem_master.csv"                   'master manually edited file
input_files(3) = dn + "_super_grp_i.csv"
input_files(4) = dn + "_cat_grp_i.csv"
input_files(5) = dn + "_enumerations.txt"
input_files(6) = dn + "_category_desc.csv"
input_files(7) = dn + "_enum_dict.csv"
input_files(8) = dn + "_item_type_units.txt"
input_files(9) = dn + "_STAR_header.txt"

output_files(2) = dn + "_new_xlschem_master.csv"    'file generated from update command
                                                    'goes to source directory for handling
output_files(24) = dn + "_xlschem_ann.csv"

output_files(3) = dn + "_item_tbl_o.csv"
output_files(4) = dn + "_cat_grp_o.csv"
output_files(5) = dn + "_super_grp_o.csv"
output_files(6) = "STAR_dict.txt"
output_files(10) = dn + "_table_err.txt"
output_files(11) = dn + "_tag_err.txt"
output_files(15) = "fake_file.txt"
output_files(18) = dn + "_tag_validation.csv"
output_files(21) = "enum_dict.csv"
output_files(22) = dn + "_category_desc.csv"


'set up variables and arrays

spg_col_ct = 8: spg_row_ct = 0      'number of columns in the ADIT super group excel file
grp_col_ct = 28: grp_row_ct = 0     'number of columns in the ADIT group excel file


    If launch_form.Check1.Value = 1 Then            'Check dictionary excel file validation
        Debug.Print "Excel file integrity check"
        clear_fields

        ReDim grp_table(200, grp_col_ct) As String
        ReDim spg_table(20, spg_col_ct) As String

        Open path_source + output_files(10) For Output As 10
        Open path_source + output_files(11) For Output As 11
        
        'path_source = path_distribution_adit
        input_file = input_files(2)             'dn + "_xlschem_master.csv"
        
        syntax_check.Text1.Text = input_file
        syntax_check.Text2.Text = path_source
        syntax_check.Show
        syntax_check.Refresh
                
        load_adit_data path_source, input_files, spg_col_ct, spg_row_ct, spg_table, grp_col_ct, grp_row_ct, grp_table, file_source

        read_excel e_tag_ct, excel_tag_dat, excel_header_dat, path_source, input_file, e_col_num
        
        check_excel e_tag_ct, excel_tag_dat, excel_header_dat, e_col_num, spg_col_ct, spg_row_ct, spg_table, grp_col_ct, grp_row_ct, grp_table, path_source, file_source, dn
        
    End If
      
    If launch_form.Check4.Value = 1 Then                        'Update dictionary files from the new Excel file
        Debug.Print "Option #4:  Update from the Excel file"
        'path_source = path_distribution_adit
        
        ReDim tag_char(10000, 4) As String
        ReDim new_tag_dat(10000, e_col_num)
        ReDim grp_table(200, grp_col_ct) As String
        ReDim spg_table(20, spg_col_ct) As String
                
        tag_ct = 0
    
        input_file = input_files(2)     'input dictionary Excel file  dn + "_xlschem_master.csv"
        read_excel e_tag_ct, excel_tag_dat, excel_header_dat, path_source, input_file, e_col_num
                
            For i = 1 To e_tag_ct
                For j = 1 To e_col_num
                    new_tag_dat(i, j) = excel_tag_dat(i, j)
                    'Debug.Print new_tag_dat(i, j)
                Next j
                p = InStr(1, new_tag_dat(i, 7), ".")
                category1 = Left$(new_tag_dat(i, 7), p - 1)
                tag1 = Right$(new_tag_dat(i, 7), Len(new_tag_dat(i, 7)) - p)

                tag_char(i, 1) = tag1
                tag_char(i, 2) = category1
                tag_char(i, 3) = new_tag_dat(i, 2)
'               tag_char(i, 4) = new_tag_dat(i, 18)

            Next i
        
        tag_ct = e_tag_ct
                
        load_adit_data path_source, input_files, spg_col_ct, spg_row_ct, spg_table, grp_col_ct, grp_row_ct, grp_table, file_source
                
        set_tag_values tag_ct, tag_char, new_tag_dat, excel_header_dat, e_col_num, spg_col_ct, spg_row_ct, spg_table, grp_col_ct, grp_row_ct, grp_table
        
        order_tags tag_ct, new_tag_dat, e_col_num
        
        output_file = output_files(2)       'dn + "_new_xlschem_master.csv" is output file
        write_new_excel path_source, output_files, tag_ct, new_tag_dat, excel_header_dat, e_col_num, file_source
                           
        If file_source <> 2 Then write_fake_star path_distribution, output_files, tag_ct, new_tag_dat, e_col_num
                
        Close 2
        
        'read and write enumeration file
        Open path_distribution + output_files(21) For Output As 2       'dn + "_enum_dict.csv"

        run_option = 15
        read_star_enumerations tag_list, tag_ct, tag_char, table_list, table_ct, sf_list, sf_ct, path_source, input_files, run_option, tag_desc
        
        'write_enum_hdr path_source, output_files, tag_ct, new_tag_dat, e_col_num, excel_header_dat
                       
        Close
        
        lcase_flag = 0: full_dict_flag = 1 'Do not lower case tags and print all tags including the non-public tags
        Open path_distribution + output_files(6) For Output As 13       'dn + "_STAR_dict.txt"
        write_nmrstar_dict2 e_tag_ct, excel_tag_dat, tag_ct, tag_list, e_col_num, path_source, output_files, lcase_flag, full_dict_flag, dictionary_name, dn, input_files

        Close
        
        If file_source = 2 Then
            write_group_tables path_source, output_files, input_files, file_source
            write_fake_star path_distribution, output_files, tag_ct, new_tag_dat, e_col_num
        End If
        
        'FileCopy path_source + input_files(2), path_distribution + "xlschem_master.csv"
        
        FileCopy path_source + output_files(24), path_distribution + "xlschem_ann.csv"
        FileCopy path_source + input_files(3), path_distribution + "super_grp_i.csv"
        FileCopy path_source + input_files(4), path_distribution + "cat_grp_i.csv"
        FileCopy path_source + input_files(5), path_distribution + "enumerations.txt"
        FileCopy path_source + input_files(6), path_distribution + "category_desc.csv"
        'FileCopy path_source + output_files(2), path_distribution + "new_xlschem_master.csv"    'file generated from update command
                
        FileCopy path_distribution + output_files(21), path_source + input_files(7)     'dn+"_enum_dict.csv"
                        
        Debug.Print "tag count ="; tag_ct
        
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


