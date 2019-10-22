Attribute VB_Name = "start"
'Manage the NMR-STAR schema and dictionary   (NMR_STAR_control     8/5/2000)

'Functions required:

' 1) syntax/semantic checks
'    a. data block
'    b. saveframe
'    c. loop
'    d. tag-value count
'    e. single quotes
'    f. double quotes
'    g. semicolons
'    h. duplicate tags in saveframe
'    i. saveframe/framecode exists

' 2) schema - dictionary comparison
' 3) write RDB control file
' 4) append dictionary
' 5) identify missing information in the dictionary
' 6) identify missing foreign keys in the schema
' 7) write html pages for dictionary website
' 8) write alphabetized list of unique tags
' 9) write list of saveframe/tag combinations
'10) check that enumerated tags have a list of enumerations
'11) check dictionary against the DDL file
'12) update dictionary with current saveframe, loop, and tag positions from the schema

Sub main()

Dim tag_count As Integer
Dim tag_list(1 To 2000, 1 To 25) As String

Dim dictionary_path, dictionary_file, schema_file, enumerations_file, ddl_file As String

dictionary_path = "d:\htdocs\dictionary\htmldocs\nmr_star\dictionary_files\"

schema_file = "harvest_schema.txt"
dictionary_file = "dictharvest.txt"
enumeration_file = "enumerations.txt"
ddl_file = "ddl30.txt"

Debug.Print "Checking dictionary for syntax errors"
Call syntax_checker(dictionary_path, dictionary_file)

Debug.Print "Comparing the schema contents with the dictionary"
Call schema_dictionary_comp(dictionary_path, schema_file, dictionary_file)

Debug.Print "Reading NMR-STAR schema file"
'Call read_BMRB(dict_path, schema_file, tag_list, tag_count)

Debug.Print "Reading NMR-Dep files"
'Call read_Dep(tag_list, tag_count, dep_info, dep_value_ct)

Debug.Print "Writing dictionary"
'Call write_dictionary(dictionary_path, dictionary_filename, tag_list, tag_count, dep_info, dep_value_ct)

Debug.Print
Debug.Print "Finished"
End Sub


