Attribute VB_Name = "control_prog"

Public run_prog, run_option As Integer

Sub main()

'DECLARE SUB dictionarychk (t$, token$(), linecount, tottoken%)
'NMR-STAR dictionary processing               (dict_control.bas  1/20/2002)

'---------------------------------------


run_prog = 0
run_option = 0
Load Intro_form
Intro_form.Check1.Value = False
Intro_form.Check2.Value = False
Intro_form.Check3.Value = False
Intro_form.Check4.Value = False
Intro_form.Check5.Value = False
Intro_form.Check6.Value = False

Intro_form.Command1.Enabled = True
Intro_form.Command2.Enabled = True
Intro_form.Command3.Enabled = True
Intro_form.Command1.Value = False
Intro_form.Command2.Value = False
Intro_form.Command3.Value = False
Intro_form.Show

Load program_control
program_control.Command1.Enabled = True
program_control.Command2.Enabled = True
program_control.Command1.Value = False
program_control.Command2.Value = False

End Sub

Sub go_yahoo()

If run_option = 0 Then main

Dim pathin, pathout As String
Dim output_files(20), input_files(20) As String
Dim tag_ct, table_ct, sf_ct As Integer
ReDim tag_char(1, 3) As String

ReDim tag_list(5000), sf_list(200), table_list(500) As String

'control variables

e_col_num = 53


'set up files for work or home

home_work = 0   '0 = laptop
'files for work

If home_work = 0 Then
    pathin = "c:\bmrb\htdocs\dictionary\htmldocs\nmr_star\adit_files\"
    
    pathout_adit = "c:\bmrb\htdocs\dictionary\htmldocs\nmr_star\adit_files\"
    
End If

If run_prog = 1 Then
    Intro_form.Hide


'Debug.Print "stop for now"
'Debug.Print

input_files(1) = "nmrstar3_edit.txt"
input_files(2) = "xlschem_ann.csv"
input_files(3) = "adit_super_grp_i.csv"
input_files(4) = "adit_cat_grp_i.csv"
input_files(5) = "enumerations.txt"
input_files(6) = "nmrstar3_sg_stub.csv"

output_files(1) = "adit_interface_dict2.txt"
output_files(2) = "new_xlschem.csv"
output_files(3) = "adit_item_tbl_o.csv"
output_files(4) = "adit_cat_grp_o.csv"
output_files(5) = "adit_super_grp_o.csv"
output_files(6) = ""
output_files(7) = "nmrstar3_anno.txt"
output_files(8) = "adit_enum_dtl.csv"
output_files(9) = "adit_enum_hdr.csv"
output_files(10) = "table_err.txt"
output_files(11) = "tag_err.txt"
output_files(12) = "adit_enum_ties.csv"
output_files(13) = "adit_man_over.csv"
output_files(14) = "nmrstar3.txt"
output_files(15) = "nmrstar3_fake.txt"
output_files(16) = "nmrstar_sg_dict.txt"
output_files(17) = ""
output_files(18) = ""

'set up variables and arrays

spg_col_ct = 8: spg_row_ct = 0      'number of columns in the ADIT super group excel file
grp_col_ct = 15: grp_row_ct = 0     'number of columns in the ADIT group excel file


    If Intro_form.Check1.Value = 1 Then
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
        
        load_adit_data pathin, input_files, spg_col_ct, spg_row_ct, spg_table, grp_col_ct, grp_row_ct, grp_table

        read_excel e_tag_ct, excel_tag_dat, excel_header_dat, pathin, input_file, e_col_num
        
        check_excel e_tag_ct, excel_tag_dat, excel_header_dat, e_col_num, spg_col_ct, spg_row_ct, spg_table, grp_col_ct, grp_row_ct, grp_table
        
        'read_star tag_list, tag_ct, tag_char, table_list, table_ct, sf_list, sf_ct, pathin, input_file, run_option, tag_desc

        'Debug.Print
    End If

    If Intro_form.Check3.Value = 1 Then
        Debug.Print "Write interface items"
        run_option = 3
        pathin = pathout_adit
        Intro_form.Hide
        'Tag_def.Show
        
        input_file = input_files(2)
        read_excel e_tag_ct, excel_tag_dat, excel_header_dat, pathin, input_file, e_col_num

        input_file = output_files(1)

        read_star tag_list, tag_ct, tag_char, table_list, table_ct, sf_list, sf_ct, pathin, input_file, run_option, tag_desc

   '     run_option = 13
        Open pathout_adit + output_files(1) For Append As 13
        update_interface e_tag_ct, excel_tag_dat, tag_ct, tag_list

   '     pathin = pathout_adit
   '     input_file = input_files(1)
   '     read_star tag_list, tag_ct, tag_char, table_list, table_ct, sf_list, sf_ct, pathin, input_file, run_option,tag_desc
        Close 13
    End If
    
    If Intro_form.Check4.Value = 1 Then
        Debug.Print "Update from the Excel file"
        pathin = pathout_adit
        
        ReDim tag_char(4000, 4) As String
        ReDim new_tag_dat(4000, e_col_num)
        ReDim grp_table(200, grp_col_ct) As String
        ReDim spg_table(20, spg_col_ct) As String
        
        write_group_tables pathin, output_files, input_files
        
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
        
        load_adit_data pathin, input_files, spg_col_ct, spg_row_ct, spg_table, grp_col_ct, grp_row_ct, grp_table
        
        set_tag_values tag_ct, tag_char, new_tag_dat, excel_header_dat, e_col_num, spg_col_ct, spg_row_ct, spg_table, grp_col_ct, grp_row_ct, grp_table
        
        order_tags tag_ct, new_tag_dat, e_col_num
        
        output_file = output_files(2)
        write_new_excel pathin, output_file, tag_ct, new_tag_dat, excel_header_dat, e_col_num
        
        write_new_adit_files pathin, output_files, tag_ct, new_tag_dat, e_col_num, excel_header_dat
        
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
        write_man_over pathin, output_files, tag_ct, new_tag_dat, e_col_num, grp_col_ct, grp_row_ct, grp_table
        Debug.Print "tag count ="; tag_ct
        
    End If
    
    If Intro_form.Check5.Value = 1 Then
        Debug.Print "Update keys in Excel file"
        pathin = pathout_adit
        input_file = output_files(3)
    
    
    End If
    If Intro_form.Check6.Value = 1 Then
        Debug.Print "Generate ADIT files"
        pathin = pathout_adit
        input_file = output_files(3)
    
    
    End If
    If Intro_form.Check7.Value = 1 Then
        Debug.Print "Update from the NMR-STAR file"
        pathin = pathout_adit
        input_file = input_files(1)
        ReDim tag_char(4000, 4) As String
        ReDim new_tag_dat(4000, e_col_num)
        ReDim grp_table(200, grp_col_ct) As String
        ReDim spg_table(20, spg_col_ct) As String
        
        tag_ct = 0
        read_star tag_list, tag_ct, tag_char, table_list, table_ct, sf_list, sf_ct, pathin, input_file, run_option, tag_desc
    
        input_file = input_files(2)
        read_excel e_tag_ct, excel_tag_dat, excel_header_dat, pathin, input_file, e_col_num
        
        compare_tags tag_ct, tag_char, e_tag_ct, excel_tag_dat, new_tag_dat, e_col_num
        
        load_adit_data pathin, input_files, spg_col_ct, spg_row_ct, spg_table, grp_col_ct, grp_row_ct, grp_table
        
        set_from_star tag_ct, tag_char, new_tag_dat, excel_header_dat, e_col_num, spg_col_ct, spg_row_ct, spg_table, grp_col_ct, grp_row_ct, grp_table
        
        order_tags tag_ct, new_tag_dat, e_col_num
        
        output_file = output_files(2)
        write_new_excel pathin, output_file, tag_ct, new_tag_dat, excel_header_dat, e_col_num
        
        Debug.Print "tag count ="; tag_ct
    
    End If
    If Intro_form.Check8.Value = 1 Then
        Debug.Print "Write a SG stub dictionary"
        run_option = 8
        pathin = pathout_adit
        Intro_form.Hide
        
        input_file = input_files(6)
        read_excel e_tag_ct, excel_tag_dat, excel_header_dat, pathin, input_file, e_col_num

        Open pathout_adit + output_files(16) For Output As 13
        write_sg_dict e_tag_ct, excel_tag_dat, tag_ct, tag_list

        Close 13
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


