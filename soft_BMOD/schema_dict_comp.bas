Attribute VB_Name = "schema_dict_comp"
Sub schema_dictionary_comp(dictionary_path, schema_file, dictionary_file)

Dim schema_tag_ct, schema_sf_ct, dict_tag_ct, dict_sf_ct As Integer
Dim schema_tag_list(4000, 3), dict_tag_list(3000, 3) As String
Dim schema_sf_list(200, 2), dict_sf_list(200, 2) As String

'get_schema_frame_tags dictionary_path, schema_file, schema_tag_ct, schema_tag_list, schema_sf_ct, schema_sf_list
get_dict_frame_tags dictionary_path, dictionary_file, dict_tag_ct, dict_tag_list, dict_sf_ct, dict_sf_list
'comp_schema_dict schema_tag_ct, schema_tag_list, schema_sf_ct, schema_sf_list

End Sub

Sub get_schema_frame_tags(dictionary_path, schema_file, schema_tag_ct, schema_tag_list, schema_sf_ct, schema_sf_list)

Open dictionary_path + schema_file For Input As 1

Dim tag_pair(3) As String
ReDim save_names(300) As String

Dim save_names_ct As Integer

EndDoc = 0
saveframe = 0: loop_flag = 0

data_block_flag = 0

tt = "": t = ""

While EndDoc = 0
                'define a text line
    dep_retrieve_line t, EndDoc
    original_line = t
    clean_line t

    If t > "" Then
        'determine data_block and global_block conditions
    
        data_block_check t, data_block_flag, data_block_name, data_count, saveframe_flag, loop_flag, loop_item_ct
    
        If data_block_flag = 1 Then
       
    'Check to see if the line or a portion of the line is a comment
    '# indicates line is purely a comment

            remove_comments t, semicolon_flag

    'define location - saveframe, loop, data_tag
    
            saveframe_check t, saveframe_flag, loop_flag, loop_item_ct, data_read_flag, saveframe_name, save_names, save_names_ct

            loop_check t, loop_flag, loop_item_ct, loop_tag_list, data_read_flag, label_change
        
            data_tag_check t, data_tag_flag, data_tag_name
   
    'get data_tag values for non-looped data tags
    
            If data_tag_flag = 1 Then
                schema_tag_ct = schema_tag_ct + 1
                schema_tag_list(schema_tag_ct, 1) = data_tag_name
                tag_pair(1) = data_tag_name
                tag_pair(2) = "": tag_pair(3) = ""
                        

                If loop_flag = 0 And data_tag_name = "_Saveframe_category" Then
                    get_value_from_line t, tag_pair
                     
    'set saveframe category value
                    If data_tag_name = "_Saveframe_category" Then
                        schema_sf_ct = schema_sf_ct + 1
                        schema_sf_list(schema_sf_ct, 1) = tag_pair(2)
                        saveframe_category_name = tag_pair(2)
                    End If
                End If
                schema_tag_list(schema_tag_ct, 2) = saveframe_category_name
                'Debug.Print schema_tag_ct, schema_tag_list(schema_tag_ct, 1), schema_tag_list(schema_tag_ct, 2)
            End If
        End If
        t = ""

    End If
Wend
Close 1

For i = 1 To schema_tag_ct
    If schema_tag_list(i, 1) <> "" Then
        For j = i + 1 To schema_tag_ct
            If schema_tag_list(i, 1) = schema_tag_list(j, 1) Then schema_tag_list(j, 1) = ""
        Next j
    End If
Next i

n = schema_tag_ct
schema_tag_ct = 0
For i = 1 To n
    If schema_tag_list(i, 1) <> "" Then
        schema_tag_ct = schema_tag_ct + 1
        schema_tag_list(schema_tag_ct, 1) = schema_tag_list(i, 1)
        schema_tag_list(schema_tag_ct, 2) = schema_tag_list(i, 2)
        'Debug.Print schema_tag_ct, schema_tag_list(schema_tag_ct, 1), schema_tag_list(schema_tag_ct, 2)
    End If
Next i
alpha_schema_tag_sort schema_tag_ct, schema_tag_list

For i = 1 To schema_tag_ct
    Debug.Print i, schema_tag_list(i, 1), schema_tag_list(i, 2)
Next i

ReDim save_names(1)

End Sub

Sub get_dict_frame_tags(dictionary_path, dictionary_file, dict_tag_ct, dict_tag_list, dict_sf_ct, dict_sf_list)

Open dictionary_path + dictionary_file For Input As 1

ReDim save_names(300) As String

Dim tag_pair(3), text_value(20) As String
Dim save_names_ct As Integer

EndDoc = 0
saveframe = 0: loop_flag = 0

data_block_flag = 0

tt = "": t = ""

While EndDoc = 0
                'define a text line
    dep_retrieve_line t, EndDoc
    original_line = t
    clean_line t

    If t > "" Then
        'determine data_block and global_block conditions
    
        data_block_check t, data_block_flag, data_block_name, data_count, saveframe_flag, loop_flag, loop_item_ct
    
        If data_block_flag = 1 Then
       
    'Check to see if the line or a portion of the line is a comment
    '# indicates line is purely a comment

            remove_comments t, semicolon_flag

    'define location - saveframe, loop, data_tag
    
            saveframe_check t, saveframe_flag, loop_flag, loop_item_ct, data_read_flag, saveframe_name, save_names, save_names_ct

            loop_check t, loop_flag, loop_item_ct, loop_tag_list, data_read_flag, label_change
        
            data_tag_check t, data_tag_flag, data_tag_name
   
    'get data_tag values for non-looped data tags
    
            If data_tag_flag = 1 Then
                execute = 0
                tag_pair(1) = data_tag_name
                tag_pair(2) = "": tag_pair(3) = ""
                If loop_flag = 0 Then
                    While tag_pair(3) = ""
                        get_value_from_line t, tag_pair
                        If tag_pair(1) > " " And tag_pair(2) > " " Then tag_pair(3) = "y"
                        If tag_pair(3) = "" And semicolon_flag = 0 Then
                            dep_retrieve_line t, EndDoc
                            'Debug.Print t
                            'Debug.Print
                            semicolon_check t, semicolon_flag, total_data_ct, data_count, line_tot, loop_flag, semiflag
                            If semicolon_flag = 1 Then
                                text_value(1) = t: line_tot = 1
                                dep_retrieve_text_tag_value text_value, line_tot, EndDoc
                                tag_pair(3) = "y"
                                'For i = 1 To line_tot
                                '    Debug.Print text_value(i)
                                'Next i
                                'Debug.Print
                                semicolon_flag = 0
                            End If
                        End If
                    Wend
                     
    'set saveframe category value
                    If tag_pair(1) = "_Saveframe_category" Then
                        saveframe_category_name = tag_pair(2)
                    End If
                    If tag_pair(1) = "_Tag_name" Then
                        dict_tag_ct = dict_tag_ct + 1
                        dict_tag_list(dict_tag_ct, 1) = tag_pair(2)
                    End If
                End If
            
    'store the data tags in the current loop
                If loop_flag = 1 Then
                    If loop_item_ct = 0 Then
                        ReDim loop_text_value(1 To 11, 1 To 50) As String
                        ReDim loop_value(500) As String

                    End If
                    loop_item_ct = loop_item_ct + 1
                    loop_tag_list(loop_item_ct) = tag_pair(1)
                    'Debug.Print loop_item_ct, loop_tag_list(loop_item_ct)
                    'Debug.Print
                End If
            End If
 
            If data_tag_flag = 0 Then
                If loop_flag = 1 Then
                    semicolon_check t, semicolon_flag, total_data_ct, data_count, line_tot, loop_flag, semiflag

                    If semicolon_flag = 0 And t > "" Then
                'Debug.Print t
                        single_quote_check t, data_value, error_report
                        If data_value = "" Then double_quote_check t, data_value, error_report
                        If data_value = "" Then
                            If t <> Chr$(34) + Chr$(34) Then
                                If t <> "''" Then
                                    data_value = t
                                End If
                            End If
                        End If
                        loop_value_ct = loop_value_ct + 1
                        loop_value(loop_value_ct) = data_value
                    'Debug.Print loop_value_ct, data_value, original_line
                        data_value = ""
                        If saveframe_category_name = "saveframe_characteristics" Then
                            If loop_value_ct = 1 Then
                                dict_sf_ct = dict_sf_ct + 1
                                dict_sf_list(dict_sf_ct, 1) = loop_value(1)
                                Debug.Print dict_sf_ct, dict_sf_list(dict_sf_ct, 1)
                            End If
                        End If
                    End If
                    If semicolon_flag = 1 Then
                        text_value(1) = t: line_tot = 1
                        dep_retrieve_text_tag_value text_value, line_tot, EndDoc
                        loop_value_ct = loop_value_ct + 1
                 
                        For i = 1 To line_tot
                            loop_text_value(loop_value_ct, i) = text_value(i)
                            'Debug.Print text_value(i)
                        Next i
                        semicolon_flag = 0
                    End If
                End If
            End If

        End If
        t = ""
    End If
Wend
Close 1

End Sub

Sub comp_schema_dict(schema_tag_ct, schema_tag_list, schema_sf_ct, schema_sf_list)

End Sub
Sub alpha_schema_tag_sort(schema_tag_ct, schema_tag_list)

Dim i As Integer, j As Integer, numrec As Integer, Switch As Integer
Dim test As String

numrec = schema_tag_ct
offset = numrec \ 2
Do While offset > 0
        limit = numrec - offset
        Do
                Switch = False
                For i = 1 To limit
                        If schema_tag_list(i, 1) > schema_tag_list(i + offset, 1) Then
                           For j = 1 To 3
                                test = schema_tag_list(i, j)
                                schema_tag_list(i, j) = schema_tag_list(i + offset, j)
                                schema_tag_list(i + offset, j) = test
                           Next j
                           Switch = i
                        End If
                Next i
                limit = Switch
        Loop While Switch
        offset = offset \ 2
Loop

End Sub

