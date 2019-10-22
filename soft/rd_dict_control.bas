Attribute VB_Name = "rdBMRB_dict"
' rdbmrb_dict.bas

Sub read_BMRB(NMR_STAR_dict_path, NMR_STAR_dict_filename, tag_list, tag_count)

Dim i As Integer, stringflag As Integer
Dim semicolon_flag As Integer
Dim EndDoc As Integer, data_tag_flag As Integer
Dim q As Integer, loop_flag As Integer, data_block_flag As Integer
Dim semiflag As Integer
Dim t As String, tt As String, p3 As String, pp As String
Dim p2 As Integer
Dim saveframe_category_name As String
Dim original_line As String, data_block_name As String
Dim save_names(1 To 100) As String
Dim tag_pair(1 To 3) As String
Dim add_examples_flag As Integer

Open NMR_STAR_dict_path + NMR_STAR_dict_filename For Input As 1

EndDoc = 0
loop_item_ct = 0
saveframe = 0: loop_flag = 0: seqflag = 0: line_tot = 0
stringflag = 0: total_data_ct = 0
semicolon_flag = 0: semiflag = 0
tag_count = 0

data_block_flag = 0

tt = "": t = ""

While EndDoc = 0
                'define a text line
    retrieve_line t, EndDoc
'    Debug.Print t
'    Debug.Print

    original_line = t
    
    clean_line t

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
            add_examples_flag = 0
            execute = 0
            tag_pair(1) = data_tag_name
            tag_pair(2) = "": tag_pair(3) = ""
                        

            If loop_flag = 0 Then
'                While tag_pair(3) = ""
                    get_value_from_line t, tag_pair
                    If tag_pair(1) > " " And tag_pair(2) > " " Then tag_pair(3) = "y"
'                    If tag_pair(3) = "" And semicolon_flag = 0 Then
'                        retrieve_line t, EndDoc
                        'Debug.Print t
                        'Debug.Print
'                    End If
'                 Wend
                     
    'set saveframe category value
                If tag_pair(1) = "_Saveframe_category" Then
                    saveframe_category_name = tag_pair(2)
                End If

            End If
            
            tag_count = tag_count + 1
            
            tag_list(tag_count, 1) = tag_pair(1)
            tag_list(tag_count, 2) = saveframe_category_name
            If loop_flag = 1 Then
                loop_item_ct = loop_item_ct + 1
                tag_list(tag_count, 5) = "Y"        'loop mandatory true
                tag_list(tag_count, 11) = Str$(loop_item_ct * 10)   'loop position
            End If
            If loop_flag = 0 Then tag_list(tag_count, 5) = "N"
            
            tag_list(tag_count, 10) = Str$(tag_count * 10)   'position in full NMR-STAR file
            
            p = InStr(1, original_line, "MANDATORY")
            p1 = InStr(1, original_line, "Cond-MANDATORY")
            
            If p > 1 Then tag_list(tag_count, 8) = "MANDATORY"
            If p1 > 1 Then tag_list(tag_count, 8) = "Cond-MANDATORY"
            
            p = InStr(1, original_line, "[e.g.:")
            If p > 0 Then
                tag_list(tag_count, 9) = Right$(original_line, Len(original_line) - (p + 4))
                p1 = InStr(p + 1, original_line, "]")
                If p1 = 0 Then add_examples_flag = 1
            End If

            'Debug.Print tag_count, tag_list(tag_count, 1), tag_list(tag_count, 2)
            
        End If
 
        If add_examples_flag = 1 And data_tag_flag = 0 Then
            p = InStr(1, original_line, "#")
            If p > 0 Then
                new_line = Right$(original_line, Len(original_line) - (p + 1))
                tag_list(tag_count, 9) = tag_list(tag_count, 9) + new_line
                p1 = InStr(1, original_line, "]")
                If p1 > 1 Then add_examples_flag = 0
            End If
        End If
           
'APPLICATION SPECIFIC CODE BEGINS AT THIS POINT

        
    End If
    t = ""
Wend
Close 1

alphasort tag_count, tag_list

End Sub

Sub retrieve_line(t, EndDoc)

Dim tt As String
t = ""
EndDoc = 0: tt = " "
While Asc(tt) <> 10
    tt = Input$(1, 1)
    If EOF(1) = -1 Then EndDoc = 1
    If Asc(tt) <> 10 Then
        If Asc(tt) <> 13 Then
            t = t + tt
        End If
    End If
    If EndDoc = 1 Then tt = Chr$(10)
Wend

End Sub
Sub retrieve_text_tag_value(text_value, line_tot)

Dim p As Integer, p1 As Integer
p1 = 1: t = ""
While p <> 1
    retrieve_line t, EndDoc
    'Debug.Print t
    'Debug.Print
    p = InStr(1, t, ";")
    'If p <> 1 Then
        p1 = p1 + 1
        text_value(p1) = t
    'End If
    t = ""
Wend
line_tot = p1

End Sub




Sub value_in_line_search(t, data_value, error_report)
    'search for data values with white space and no quotes of any kind

p = InStr(1, t, "_")
If p = 1 Then
    p1 = InStr(1, t, " ")
    ln = Len(t)
    If p1 > 0 Then data_value = Right$(t, ln - p1)
End If
If p = 0 Then data_value = t
data_value = Trim$(data_value)
If Len(data_value) > 0 Then
    If Left$(data_value, 1) <> Chr$(34) Then
        If Left$(data_value, 1) <> "'" Then
            p = InStr(1, data_value, " ")
            If p > 0 Then
                'Print "Value with non-quoted white space"
                'Print data_value
                'Print "line number = "; linecount%
                'Print
                'Print #6, "Value with non-quoted white space"
                'Print #6, data_value
                'Print #6, "line number = "; linecount
                'Print #6,
                error_report = 1
            End If
        End If
    End If
End If

End Sub


Sub read_star_data_rows(t, datact, data)

Dim i As Integer, ln As Integer, whitesp As Integer, quote1 As Integer
Dim p As String

ReDim dataloc(60) As Integer
ReDim data(60) As String

ln = Len(t)
datact = 0: quote1 = 0
If ln > 0 Then
   whitesp = 1
   For i = 1 To ln
       p = Mid$(t, i, 1)
       If p <> Chr$(32) Or quote1 > 0 Then
          If p <> Chr$(9) Or quote1 > 0 Then
             If Asc(p) <> 10 And Asc(p) <> 13 Then
                If p > "" Then
                   If whitesp = 1 Then
                      datact = datact + 1
                      dataloc(datact) = i
                      If p = Chr$(39) And quote1 = 0 Then quote1 = 1
                      If p = Chr$(34) And quote1 = 0 Then quote1 = 2
                   End If
                   If whitesp = 0 Then
                      If p = Chr$(39) And quote1 = 1 Then quote1 = 0
                      If p = Chr$(34) And quote1 = 2 Then quote1 = 0
                   End If
                   data(datact) = data(datact) + p
                   whitesp = 0
                End If
             End If
          End If
       End If
       If p = Chr$(32) Or p = Chr$(9) Then
          If quote1 = 0 Then
             If whitesp = 0 Then whitesp = 1
          End If
       End If
   Next i
End If

End Sub




Public Sub parse_sequences(text_value, line_tot, BMRB_res_seq4a)

Dim first_parse(10) As String
Dim query_seq(5, 3) As String, match_seq(5, 3) As String, subject_seq(5, 3) As String

query1 = 0

For i = 1 To line_tot
    text_value(i) = Trim$(text_value(i))
    'Debug.Print text_value(i)
    p = InStr(1, text_value(i), "Query")
    If p > 0 Then
        query1 = query1 + 1
        n = 0
        ln = Len(text_value(i))
        For j = p + 6 To ln
            h = Mid$(text_value(i), j, 1)
            If h <> " " Then h1 = h1 + h
            If h = " " And h1 <> "" Then
                n = n + 1
                If n = 1 Then start_line = j
                If n = 3 Then end_line = j - (Len(h1) + 1)
                first_parse(n) = h1
                'Debug.Print n, first_parse(n)
                h1 = ""
            End If
        Next j
        n = n + 1: first_parse(n) = h1
        If n = 3 Then end_line = j - (Len(h1) + 1)
        h1 = ""
            
        query_seq(query1, 1) = Mid$(text_value(i), start_line + 1, end_line - (start_line + 1))
        match_seq(query1, 1) = Mid$(text_value(i + 1), start_line + 1, end_line - (start_line + 1))
        subject_seq(query1, 1) = Mid$(text_value(i + 2), start_line + 1, end_line - (start_line + 1))
            
        query_seq(query1, 2) = first_parse(1)
        query_seq(query1, 3) = first_parse(3)
    End If
Next i
For i = 1 To query1
    'Debug.Print query_seq(i, 1)
    If Len(query_seq(i, 1)) = 60 Then
        query_string = query_string + query_seq(i, 1)
        match_string = match_string + match_seq(i, 1)
        subject_string = subject_string + subject_seq(i, 1)
    End If
Next i
If Len(query_seq(query1, 1)) < 60 Then query_string = query_string + query_seq(query1, 1)
If Len(match_seq(query1, 1)) < 60 Then match_string = match_string + query_seq(query1, 1)
If Len(subject_seq(query1, 1)) < 60 Then subject_string = subject_string + query_seq(query1, 1)
'Debug.Print query_string
'Debug.Print match_string
'Debug.Print subject_string
'Debug.Print
ln = Len(query_string)
For i = 1 To ln
    BMRB_res_seq4a(4, query_seq(1, 2) + i - 1) = Mid$(query_string, i, 1)
    BMRB_res_seq4a(5, query_seq(1, 2) + i - 1) = Mid$(match_string, i, 1)
    BMRB_res_seq4a(6, query_seq(1, 2) + i - 1) = Mid$(subject_string, i, 1)
Next i

End Sub
Sub alphasort(tag_count, tag_list)

Dim i As Integer, j As Integer, numrec As Integer, Switch As Integer
Dim test As String

numrec = tag_count
offset = numrec \ 2
Do While offset > 0
        limit = numrec - offset
        Do
                Switch = False
                For i = 1 To limit
                        If tag_list(i, 1) > tag_list(i + offset, 1) Then
                           For j = 1 To 25
                                test = tag_list(i, j)
                                tag_list(i, j) = tag_list(i + offset, j)
                                tag_list(i + offset, j) = test
                           Next j
                           Switch = i
                        End If
                Next i
                limit = Switch
        Loop While Switch
        offset = offset \ 2
Loop

End Sub
