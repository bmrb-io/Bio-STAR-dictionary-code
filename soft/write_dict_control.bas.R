Attribute VB_Name = "write_dict"
' write_dict.bas
Sub write_dictionary(dictionary_path, dictionary_filename, tag_list, tag_count, dep_info, dep_value_ct)

Dim i As Integer, j As Integer, n As Integer, t As Integer, p As Integer
Dim print_flag As Integer, tag_print_count As Integer
Dim old_count As Integer

ReDim old_list(tag_count, 2) As String
tag_print_count = 0
Open dictionary_path + dictionary_filename For Output As 1
Open dictionary_path + "deposition_info.dict" For Output As 2

write_public_version old_list, tag_print_count, tag_list, tag_count, dep_info, dep_value_ct
Close 1
write_deposition_info_file old_list, tag_print_count, tag_list, tag_count, dep_info, dep_value_ct
Close 2
End Sub

Sub write_public_version(old_list, tag_print_count, tag_list, tag_count, dep_info, dep_value_ct)
Print #1, "#  NMR-STAR Dictionary Version 1.0"
Print #1,
Print #1, "data_NMR-STAR_Dictionary_v1.0"
Print #1,
Print #1, "   _NMR_STAR_supported_version      2.2"
Print #1, "   _NMR_STAR_dictionary_date        1999-07-03"
Print #1,
For i = 1 To tag_count
    For j = 1 To 7
        If j = 3 Then
            p = InStr(1, tag_list(i, 3), " ")
            If p > 0 Then
                tag_list(i, 3) = "'" + tag_list(i, 3) + "'"
            End If
            If tag_list(i, 3) = "" Then tag_list(i, 3) = "?"
        End If
        If tag_list(i, 6) = "" Then tag_list(i, 6) = "?"
    Next j
    
    print_flag = 1
    For j = 1 To i - 1
        If tag_list(i, 1) = tag_list(j, 1) Then
            print_flag = 0
            Exit For
        End If
    Next j
   ' p2 = InStr(1, tag_list(i, 1), "_label")
   ' If p2 = 0 Then print_flag = 0
    
 If print_flag = 1 Then
    tag_print_count = tag_print_count + 1
    Print #1, "save_" + tag_list(i, 1)
    Print #1, "  _Saveframe_category                 dictionary_item_definition"
    Print #1,
    Print #1, "  _Item_name                         "; "'"; tag_list(i, 1); "'"
    Print #1,
          
    Print #1, "  _Item_description"
    Print #1, ";"
    Print #1, "    ?"
    Print #1, ";"
    Print #1,
    
    set_NMR_STAR_type i, tag_list
       
    If tag_list(i, 6) > " " Then
        Print #1, "  _Item_type                              "; tag_list(i, 6)
    Else
        Print #1, "  _Item_type                               ?"
    End If
    Print #1, "  _Item_units                             ."
    Print #1, "  _Item_default_value                     ."
    If tag_list(i, 7) > " " Then
        Print #1, "  _Item_enumeration_framecode             $"; tag_list(i, 1); "_enumeration"
    End If
    If tag_list(i, 6) = "char" Then
        Print #1, "  _Item_maximum_character_length           "; tag_list(i, 3)
    End If
    Print #1, "  _Item_range_minimum                     ."
    Print #1, "  _Item_range_maximum                     ."
    Print #1, "  _Item_anchored                          no"
    Print #1, "  _Item_anchor_saveframe_category         ."
    Print #1,
    'Print #1, "  _Item_BMRB_DB_type                      "; tag_list(i, 3)
    'Print #1, "  _Item_BMRB_DB_column_name               ";
    Print #1,
    
    'Print #1, "  loop_"
    'Print #1, "     _Item_external_reference_file"
    'Print #1, "     _Item_external_reference_file_description"
    'Print #1,
    'Print #1, "   ?               ?"
    'Print #1,
    'Print #1, "  stop_"
    'Print #1,
    
    Print #1, "  loop_"
    Print #1, "     _Item_saveframe_category"
    Print #1, "     _Item_local_name"
    Print #1, "     _Item_NMR_STAR_mandatory_code"
    'Print #1, "     _Item_NMR_STAR_BMRB_provided"
    Print #1, "     _Item_NMR_STAR_conditional_mandatory_rule"
    Print #1, "     _Item_NMR_STAR_order"
    Print #1, "     _Item_NMR_STAR_loop_mandatory_code"
    Print #1, "     _Item_NMR_STAR_saveframe_loop_order"
    Print #1,

    For j = 1 To tag_count
        If tag_list(i, 1) = tag_list(j, 1) Then
            print_flag = 1
            If tag_list(j, 8) = "" Then tag_list(j, 8) = "?"
            If tag_list(j, 11) = "" Then tag_list(j, 11) = "."
            
            If tag_list(j, 2) > "   " Then
                Print #1, tag_list(j, 2)
            Else
                Print #1, "?"
            End If
            
            
            
            If tag_list(j, 1) > "   " Then
                Print #1, "     "; "'"; tag_list(j, 1); "'",
            Else
                Print #1, "?",
            End If
            
            If tag_list(j, 8) > "   " Then
                Print #1, "     "; tag_list(j, 8),
            Else
                Print #1, "?",
            End If
            
            'Print #1, "     ?",
            Print #1, "     ?"
            
            If tag_list(j, 10) > "   " Then
                Print #1, "     "; tag_list(j, 10),
            Else
                Print #1, "?",
            End If
            
            If tag_list(j, 5) > "   " Then
                Print #1, "     "; tag_list(j, 5),
            Else
                Print #1, "?",
            End If
            
            If tag_list(j, 11) > "   " Then
                Print #1, "     "; tag_list(j, 11)
            Else
                Print #1, "?"
            End If
            Print #1,
        End If
        test = LCase(tag_list(i, 1))
        p = InStr(1, tag_list(j, 1), test)
        
        If p > 2 Then
            print_flag = 1
            If tag_list(j, 8) = "" Then tag_list(j, 8) = "?"
            If tag_list(j, 11) = "" Then tag_list(j, 11) = "."
            
            If tag_list(j, 2) > "   " Then
                Print #1, tag_list(j, 2)
            Else
                Print #1, "?"
            End If
            
            If tag_list(j, 1) > "   " Then
                Print #1, "     "; "'"; tag_list(j, 1); "'",
            Else
                Print #1, "?",
            End If
            
            If tag_list(j, 8) > " " Then
                Print #1, "     "; tag_list(j, 8),
            Else
                Print #1, "?",
            End If
 
            'Print #1, "     ?",
            Print #1, "?"
                  
            If tag_list(j, 10) > " " Then
                Print #1, "     "; tag_list(j, 10),
            Else
                Print #1, "?",
            End If

                
            If tag_list(j, 5) > " " Then
                Print #1, "     "; tag_list(j, 5),
            Else
                Print #1, "?",
            End If
            
            If tag_list(j, 11) > " " Then
                Print #1, "     "; tag_list(j, 11)
            Else
                Print #1, "?"
            End If
            Print #1,
        End If
    Next j

    Print #1,
    Print #1, "  stop_"
        
    Print #1,
   
    If tag_list(i, 8) = "Cond-MANDATORY" Then
        Print #1, "  loop_"
        Print #1, "     _mandatory_dependency_type"    'exists; value_dependent; expression
        Print #1, "     _mandatory_dependency_value"
        Print #1, "    loop_"
        Print #1, "       _dependent_data_tag"          'name of the dependent tag
        Print #1, "       _dependent_saveframe_category"  'type of saveframe where the dependent tag is found
        Print #1,
        Print #1, "    ?      ?      ?      ?    "
        Print #1,
        Print #1, "    stop_"
        Print #1,
        Print #1, "  stop_"
        Print #1,
    End If

    Print #1, "  loop_"
    Print #1, "     _Item_examples"
    Print #1,
    If tag_list(i, 7) > "" Then Print #1, "    ?"
    If tag_list(i, 7) = "" Then
        If tag_list(i, 9) = "" Then Print #1, "     ?"
        If tag_list(i, 9) > " " Then
            test2 = ""
            For j = 1 To Len(tag_list(i, 9))
                test = Mid$(tag_list(i, 9), j, 1)
                If test = ":" Then test = " "
                If test = ";" Then test = " "
                If test = "]" Then test = ""
                test2 = test2 + test
            Next j
            tag_list(i, 9) = test2
        
            Print #1, tag_list(i, 9)
        End If
    End If
    
            dep_value_found = 0
        For j = 1 To dep_value_ct
        'Debug.Print Val(dep_info(j, 1))
            If Val(dep_info(j, 1)) = i Then
                dep_value_found = 1
             'Debug.Print j, Val(dep_info(j, 2)), Val(dep_info(j, 3)), Val(dep_info(j, 4))
             'Debug.Print
                
                r1 = Val(dep_info(j, 2))
                r2 = Val(dep_info(j, 3))
                r3 = Val(dep_info(j, 4))
                
                Print #1,
                
                If r2 > 0 Then
                    'If r2 > 1 And Left$(dep_info(j, 5 + r2), 1) <> ";" Then Print #1, ";"
                    If r2 = 1 Then Print #1, "'";
                    For k = 5 + r1 To 5 + r1 + r2 - 1
                        Print #1, dep_info(j, k);
                        If r2 = 1 Then Print #1, "'";
                        Print #1,
                    Next k
                    'If r2 > 1 And Left$(dep_info(j, 5 + r1 + r2 - 1), 1) <> ";" Then Print #1, ";"
                    Print #1,
                End If
            End If
        Next j

    Print #1,
    Print #1, "  stop_"
    Print #1,
        Print #1, "#  DEPOSITION SYSTEM INFORMATION"
    Print #1,
    Print #1, "  loop_"
    Print #1, "     _Item_deposition_saveframe_category"
    Print #1, "     _Item_deposition_item_name"
    Print #1, "     _Item_deposition_prompt"
    Print #1, "     _Item_deposition_help_message"
    'Print #1, "     _Item_deposition_error_message"
    Print #1,
    Print #1, "     "; tag_list(i, 2)
    Print #1, "     '" + tag_list(i, 1) + "'"
    Print #1,
        'Debug.Print dep_value_ct
        dep_value_found = 0
        For j = 1 To dep_value_ct
        'Debug.Print Val(dep_info(j, 1))
            If Val(dep_info(j, 1)) = i Then
                dep_value_found = 1
             'Debug.Print j, Val(dep_info(j, 2)), Val(dep_info(j, 3)), Val(dep_info(j, 4))
             'Debug.Print
                
                r1 = Val(dep_info(j, 2))
                r2 = Val(dep_info(j, 3))
                r3 = Val(dep_info(j, 4))
                
                Print #1, "#  --PROMPT--"
                Print #1,
                
                If r1 = 0 Then
                    Print #1, "?"
                    Print #1,
                End If
                
                If r1 > 0 Then
                    If r1 > 1 And Left$(dep_info(j, 5), 1) <> ";" Then Print #1, ";"
                    If r1 = 1 Then Print #1, "'";
                    For k = 5 To 5 + r1 - 1
                        Print #1, dep_info(j, k);
                        If r1 = 1 Then Print #1, "'";
                        Print #1,
                    Next k
                    If r1 > 1 And Left$(dep_info(j, 5 + r1 - 1), 1) <> ";" Then Print #1, ";"
                    Print #1,
                End If
                
                Print #1, "# --HELP MESSAGE--"
                Print #1,
                
                If r3 = 0 Then
                    Print #1, "?"
                    Print #1,
                End If
                    
                If r3 > 0 Then
                    
                    If r3 = 1 Then Print #1, "'";
                    For k = 5 + (r1 + r2) To 5 + (r1 + r2) + r3 - 1
                        Print #1, dep_info(j, k);
                        If r3 = 1 Then Print #1, "'";
                        Print #1,
                    Next k
                    
                    Print #1,
                End If
                'Exit For
            End If
        Next j
        'Debug.Print
        If dep_value_found = 0 Then
            Print #1, "     ?"
            Print #1, "     ?"
        End If
    
    Print #1,
    Print #1, "  stop_"
    Print #1,
        Print #1, "#  BACKWARD COMPATIBILITY"
    Print #1,
    Print #1, "  loop_"
    Print #1, "     _Item_previous_tag_name"
    Print #1, "     _Item_previous_dictionary_version"
    
    Print #1,
    Print #1, "    " + "'" + tag_list(i, 1) + "'" + "      2.1"
    Print #1,
    Print #1, "  stop_"
    Print #1,

    Print #1, "save_"
    Print #1,
End If
    
    If tag_list(i, 7) > " " Then
        Print #1, "save_"; tag_list(i, 1); "_enumeration"
        Print #1, "   _Saveframe_category       enumeration"
        Print #1,
        
        Print #1, "  loop_"
        Print #1, "     _item_enumeration_value"
        Print #1, "     _item_enumeration_description"
        Print #1,
        If tag_list(i, 9) > " " Then
            test2 = ""
            For j = 1 To Len(tag_list(i, 9))
                test = Mid$(tag_list(i, 9), j, 1)
                If test = ":" Then test = " "
                If test = ";" Then test = "           ?" + Chr$(10) + Chr$(13)
                If test = "]" Then test = ""
                test2 = test2 + test
            Next j
            tag_list(i, 9) = test2
            If tag_list(i, 9) > "   " Then
                Print #1, tag_list(i, 9)
            Else
                Print #1, "?",
            End If
            Print #1,
            Print #1, "    ?"
        End If
        If tag_list(i, 9) < "    " Then
            Print #1, "    ?                          ?"
            Print #1,
        End If
        Print #1, "  stop_"
        Print #1,
        Print #1, "save_"
        Print #1,
    End If

Next i

    Print #1, "save_saveframe_category_list"
    Print #1, "  _Saveframe_category           dictionary_NMR_STAR_categories"
    Print #1,
    Print #1, "  loop_"
    Print #1, "     _NMR_STAR_saveframe_category"
    Print #1, "     _DB_table_name"
    Print #1,
        For i = 1 To tag_count
            print_flag = 1
            For j = 1 To old_count
                If tag_list(i, 2) = old_list(j, 1) Then print_flag = 0
            Next j
            If print_flag = 1 Then
                old_count = old_count + 1
                old_list(old_count, 1) = tag_list(i, 2)
                old_list(old_count, 2) = tag_list(i, 19)
            End If
        Next i
        alphasort_sf old_count, old_list
        For i = 1 To old_count
            Print #1, "    "; old_list(i, 1),
            Print #1, "    "; old_list(i, 2)
        Next i
    Print #1,
    Print #1, "  stop_"
    Print #1,
    Print #1, "save_"

Debug.Print
Debug.Print "total tag count = "; tag_count
Debug.Print
Debug.Print "unique tag count ="; tag_print_count
Debug.Print
Debug.Print "saveframe count = "; old_count
Debug.Print

End Sub

Sub write_deposition_info_file(old_list, tag_print_count, tag_list, tag_count, dep_info, dep_value_ct)

Print #2, "#  NMR-STAR Deposition Dictionary Version 1.0"
Print #2,
Print #2, "data_NMR-STAR_Deposition_Dictionary_v1.0"
Print #2,
Print #2, "   _NMR_STAR_supported_version      3.0"
Print #2, "   _NMR_STAR_dictionary_date        1999-02-15"
Print #2,

    Print #2, "#=========================================================="
    Print #2,
    For i = 1 To tag_count
    For j = 1 To 7
        If j = 3 Then
            p = InStr(1, tag_list(i, 3), " ")
            If p > 0 Then
                tag_list(i, 3) = "'" + tag_list(i, 3) + "'"
            End If
            If tag_list(i, 3) = "" Then tag_list(i, 3) = "?"
        End If
        If tag_list(i, 6) = "" Then tag_list(i, 6) = "?"
    Next j
    
    print_flag = 1
    For j = 1 To i - 1
        If tag_list(i, 1) = tag_list(j, 1) Then
            print_flag = 0
            Exit For
        End If
    Next j
    
 If print_flag = 1 Then
    tag_print_count = tag_print_count + 1
    Print #2, "save_" + tag_list(i, 1)
    Print #2, "  _Saveframe_category                 deposition_item_information"
    Print #2,
    Print #2, "  _Item_name                         "; "'"; tag_list(i, 1); "'"

    Print #2,
    Print #2, "  loop_"
    Print #2, "     _Item_deposition_saveframe_category"
    Print #2, "     _Item_deposition_item_name"
    Print #2, "     _Item_deposition_prompt"
    Print #2, "     _Item_deposition_help_message"
    Print #2, "     _Item_deposition_error_message"
    Print #2,
    Print #2, "     "; tag_list(i, 2)
    Print #2, "     '" + tag_list(i, 1) + "'"
    Print #2,
        'Debug.Print dep_value_ct
        dep_value_found = 0
        For j = 1 To dep_value_ct
        'Debug.Print Val(dep_info(j, 1))
            If Val(dep_info(j, 1)) = i Then
                dep_value_found = 1
             'Debug.Print j, Val(dep_info(j, 2)), Val(dep_info(j, 3)), Val(dep_info(j, 4))
             'Debug.Print
                
                r1 = Val(dep_info(j, 2))
                r2 = Val(dep_info(j, 3))
                r3 = Val(dep_info(j, 4))
                
                Print #2, "#  --PROMPT--"
                Print #2,
                
                If r1 = 0 Then
                    Print #2, "?"
                    Print #2,
                End If
                
                If r1 > 0 Then
                    If r1 > 1 And Left$(dep_info(j, 5), 1) <> ";" Then Print #2, ";"
                    If r1 = 1 Then Print #2, "'";
                    For k = 5 To 5 + r1 - 1
                        Print #2, dep_info(j, k);
                        If r1 = 1 Then Print #2, "'";
                        Print #2,
                    Next k
                    If r1 > 1 And Left$(dep_info(j, 5 + r1 - 1), 1) <> ";" Then Print #2, ";"
                    Print #2,
                End If
                
                Print #2, "# --HELP MESSAGE--"
                Print #2,
                
                If r3 = 0 Then
                    Print #2, "?"
                    Print #2,
                End If
                    
                If r3 > 0 Then
                    
                    If r3 = 1 Then Print #2, "'";
                    For k = 5 + (r1 + r2) To 5 + (r1 + r2) + r3 - 1
                        Print #2, dep_info(j, k);
                        If r3 = 1 Then Print #2, "'";
                        Print #2,
                    Next k
                    
                    Print #2,
                End If
                'Exit For
            End If
        Next j
        'Debug.Print
        If dep_value_found = 0 Then
            Print #2, "     ?"
            Print #2, "     ?"
        End If
    
    Print #2,
    Print #2, "  stop_"
    Print #2,
    Print #2, "save_"
    Print #2,
 End If
Next i

End Sub

Sub write_database_info()
    Print #1, "#============================================================"
    Print #1,
    Print #1, "#  DATABASE INFORMATION"
            If tag_list(i, 3) = "" Then tag_list(i, 3) = "?"
    Print #1,
    Print #1, "  _Item_DB_type               "; tag_list(i, 3)
    Print #1,
    If tag_list(i, 6) = "framecode" Then
        Print #1, "  loop_"
        Print #1, "     _Item_DB_framecode_target_category"    'for tags that require $xxx values, gives the possible places to look for source
        Print #1,
        Print #1, "     'saveframe category list goes here'"
        Print #1,
        Print #1, "  stop_"
        Print #1,
    End If

    Print #1, "  loop_"
    Print #1, "     _Item_DB_relation_name"
    Print #1, "     _Item_DB_attribute_name"
    Print #1, "     _Item_DB_form_sequence"
    Print #1, "     _Item_DB_mandatory_code"
    Print #1, "     _Item_DB_enumeration_ID"
    Print #1, "     _Item_DB_secondary_index_code"
    Print #1, "     _Item_DB_BMRB_internal_only_code"
    Print #1, "     _Item_DB_loop_location_mandatory_code"
    Print #1, "     _Item_DB_loop_location"
    Print #1, "     _Item_DB_foreign_table"
    Print #1, "     _Item_DB_foreign_attribute"
    Print #1,
        print_flag = 0
        For j = 1 To tag_count
            If tag_list(i, 1) = tag_list(j, 1) Then

                For k = 15 To 22
                    If tag_list(j, k) = "" Then tag_list(j, k) = "?"
                Next k
            
                If tag_list(j, 23) = "" Then tag_list(j, 23) = "?"
                If tag_list(j, 3) = "" Then tag_list(j, 3) = "?"
                If tag_list(j, 4) = "" Then tag_list(j, 4) = "?"
                If tag_list(j, 7) = "" Then tag_list(j, 7) = "?"
                If tag_list(j, 18) = "?" Then tag_list(j, 18) = "N"
                If tag_list(j, 19) = "?" Then
                    tag_list(j, 19) = tag_list(j, 2)
                    tag_list(j, 20) = tag_list(j, 1)
                    tag_list(j, 23) = tag_list(j, 10)
                    tag_list(j, 21) = tag_list(j, 5)
                    tag_list(j, 22) = tag_list(j, 11)
                    If tag_list(j, 8) = "MANDATORY" Then
                        tag_list(j, 4) = "NOT NULL"
                    Else
                        tag_list(j, 4) = "NULL"
                    End If
                End If
                If tag_list(j, 19) <> "?" Then
                    print_flag = 1
                    Print #1, tag_list(j, 19)
                    Print #1, "    "; tag_list(j, 20),
                    Print #1, "    "; tag_list(j, 23),
                    Print #1, "    "; "'" + tag_list(j, 4); "'",
                    Print #1, "    "; tag_list(j, 7),
                    Print #1, "    "; tag_list(j, 17),
                    Print #1, "    "; tag_list(j, 18),
                    Print #1, "    "; tag_list(j, 21),
                    Print #1, "    "; tag_list(j, 22),
                    Print #1, "    "; tag_list(j, 15),
                    Print #1, "    "; tag_list(j, 16)
                    Print #1,
                End If
            End If
            
            test = LCase(tag_list(i, 1))
            p = InStr(1, tag_list(j, 1), test)
        
            If p > 2 Then
                print_flag = 1
                For k = 15 To 22
                    If tag_list(j, k) = "" Then tag_list(j, k) = "?"
                Next k
            
                If tag_list(j, 23) = "" Then tag_list(j, 23) = "?"
                If tag_list(j, 3) = "" Then tag_list(j, 3) = "?"
                If tag_list(j, 4) = "" Then tag_list(j, 4) = "?"
                If tag_list(j, 7) = "" Then tag_list(j, 7) = "?"
        
                If tag_list(j, 19) <> "?" Then
                    Print #1, tag_list(j, 19)
                    Print #1, "    "; tag_list(j, 20),
                    Print #1, "    "; tag_list(j, 23),
                    Print #1, "    "; tag_list(j, 4),
                    Print #1, "    "; tag_list(j, 7),
                    Print #1, "    "; tag_list(j, 17),
                    Print #1, "    "; tag_list(j, 18),
                    Print #1, "    "; tag_list(j, 21),
                    Print #1, "    "; tag_list(j, 22),
                    Print #1, "    "; tag_list(j, 15),
                    Print #1, "    "; tag_list(j, 16)
                    Print #1,
        
                End If
         
            End If

        Next j
        If print_flag = 0 Then
            Print #1, "  ?   ?   ?   ?   ?   ?   ?   ?   ?   ?   ?"
            Print #1,
        End If

    Print #1, "  stop_"
    Print #1,
    Print #1, "save_"
    Print #1,
    Print #1,
 End If

Next i

Close

End Sub

Sub dictsetup():

   'list of tokens and saveframes that have a mandatory dependency
Open "e:\bmrb\docs\deposit\lsav.lst" For Output As 2
   'saveframes where data tag has a positive loop test
Open "e:\bmrb\docs\deposit\ploop.lst" For Output As 3
   'saveframes where data tag has a negative loop test
Open "e:\bmrb\docs\deposit\nloop.lst" For Output As 4

For i = 1 To n
    If mantest(i) = 1 Then
       For j = 1 To n
           If i <> j Then
              If savestr(saveframe(i)) = savestr(saveframe(j)) Then
                 If mantest(j) = 1 Then
                    Print #2, i, token2(i), savestr(saveframe(i)), token2(j)
                 End If
              End If
           End If
       Next j
    End If
    If looptest(i) = 1 Then
       Print #3, i, token2(i), savestr(saveframe(i))
    End If
    If looptest(i) = 0 Then
       Print #4, i, token2(i), savestr(saveframe(i))
    End If

Next i
Close 2
Close 3

End Sub

Sub set_NMR_STAR_type(i, tag_list)

    p = InStr(1, tag_list(i, 3), "VARCHAR")
    If p > 0 Then tag_list(i, 6) = "char"
    
    p = InStr(1, tag_list(i, 3), "NCHAR")
    If p > 0 Then tag_list(i, 6) = "char"
    
    p = InStr(1, tag_list(i, 3), "TEXT")
    If p > 0 Then tag_list(i, 6) = "TEXT"
    
    p = InStr(1, tag_list(i, 3), "DATETIME")
    If p > 0 Then tag_list(i, 6) = "date"
    
    p = InStr(1, tag_list(i, 3), "FLOAT")
    If p > 0 Then tag_list(i, 6) = "real"
    
    p = InStr(1, tag_list(i, 3), "INTEGER")
    If p > 0 Then tag_list(i, 6) = "integer"
    
    p = InStr(1, tag_list(i, 1), "_label")
    If p > 0 Then tag_list(i, 6) = "framecode"

 
End Sub
Sub alphasort_sf(old_count, old_list)

Dim i As Integer, j As Integer, numrec As Integer, Switch As Integer
Dim test As String

numrec = old_count
offset = numrec \ 2
Do While offset > 0
        limit = numrec - offset
        Do
                Switch = False
                For i = 1 To limit
                        If old_list(i, 1) > old_list(i + offset, 1) Then
                           
                                test1 = old_list(i, 1)
                                test2 = old_list(i, 2)
                                old_list(i, 1) = old_list(i + offset, 1)
                                old_list(i, 2) = old_list(i + offset, 2)
                                old_list(i + offset, 1) = test1
                                old_list(i + offset, 2) = test2
                           
                           Switch = i
                        End If
                Next i
                limit = Switch
        Loop While Switch
        offset = offset \ 2
Loop

End Sub

