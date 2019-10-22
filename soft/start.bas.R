Attribute VB_Name = "start"
'Create data dictionary from the deposition form    (nmr_star_dict   12/09/01)
'   the RDBMS dictionary, and the NMR-Dep CIF files
                               
Sub main()

Dim NMR_STAR_dict_path, NMR_STAR_dict_filename As String
Dim NMR_STAR_schema_filename, NMR_STAR_enum_filename As String
Dim new_table, old_table, output_enum_filename, help_output_table As String
Dim BMRB_path As String
Dim RDB_field_count As Integer

Dim tag_list(4000, 0 To 30) As String

ReDim dict_info(1, 1) As String

'Dim dictionary_path As String
'Dim dictionary_filename As String
'Dim RDB_dict_path As String
'Dim RDB_dict_filename As String
'Dim RDB_field_count As Integer
'Dim dep_value_ct As Integer
'Dim dep_info(1500, 50) As String

'fixed values used
RDB_field_count = 43

'paths and files used
NMR_STAR_dict_path = "c:\bmrb\htdocs\dictionary\htmldocs\nmr_star\adit_files\"
NMR_STAR_dict_filename = "nmrstar3_dict.txt"
NMR_STAR_enum_input_filename = "enumerations.txt"
'NMR_STAR_schema_filename = "harvest_schema.txt"
NMR_STAR_schema_filename = "nmrstar3.txt"

    'real set of tables
old_table = "xlschem_ann.csv"
super_grp_tbl_in = "adit_super_grp_i.csv"
group_tbl_in = "adit_cat_grp_i.csv"

    'test set of tables
'old_table = "test_xlschem_ann.csv"
'super_grp_tbl_in = "test_adit_super_grp_i.csv"
'group_tbl_in = "test_adit_cat_grp_i.csv"

'files generated
new_table = "new_xlschem.csv"
output_enum_filename = "enum_table.csv"
help_output_table = "help_table.csv"
super_grp_tbl_out = "adit_super_grp_o.csv"
group_tbl_out = "adit_cat_grp_o.csv"
item_tbl_out = "adit_item_tbl_o.csv"

'functions called and their order

Debug.Print
Debug.Print "Start program"
Debug.Print "Reading enumeration file"
Call read_enumerations(NMR_STAR_dict_path, NMR_STAR_enum_input_filename, output_enum_filename)

Debug.Print "Reading NMR-STAR schema file"
'Call read_BMRB_schema(NMR_STAR_dict_path, NMR_STAR_schema_filename, tag_list, tag_count, RDB_field_count, old_table, new_table)
'Debug.Print
'do not use  Debug.Print "Reading RDB dictionary"
'Call read_RDB_dict(RDB_dict_path, RDB_dict_filename, RDB_field_count, tag_count, tag_list)

'do not use Debug.Print "Reading NMR-Dep files"
'Call read_Dep(tag_list, tag_count, dep_info, dep_value_ct)

ReDim dict_info(1 To 4000, 1 To 7) As String
Debug.Print "Reading NMR-STAR dictionary"
'Call read_dictionary(NMR_STAR_dict_path, NMR_STAR_dict_filename, help_output_table, dict_info, dict_data_ct)

Debug.Print "Updating Excel file"
'Call tbl_update(NMR_STAR_dict_path, old_table, new_table, dict_info, dict_data_ct, RDB_field_count, super_grp_tbl_in, group_tbl_in)

'do not use Debug.Print "Writing NMR-STAR dictionary"
'Call write_dictionary(NMR_STAR_dict_path, dictionary_filename, new_table)

Debug.Print "Writing group files"
'Call write_grp_tables(NMR_STAR_dict_path, super_grp_tbl_in, super_grp_tbl_out, group_tbl_in, group_tbl_out, RDB_field_count, new_table, item_tbl_out)


Debug.Print
Debug.Print "Program finished"
End Sub


