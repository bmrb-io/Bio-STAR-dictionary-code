Attribute VB_Name = "rdBMRB_schema"
' rdbmrb_dict.bas
' read in data from the NMR-STAR schema file in text format

Sub read_BMRB_schema(NMR_STAR_dict_path, NMR_STAR_schema_filename, tag_list, tag_count, RDB_field_count, old_table, new_table)

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

Dim old_tbl_ct As Integer
Open NMR_STAR_dict_path + old_table For Input As 1

old_tbl_ct = 0

ReDim old_table_dat(1 To 1, RDB_field_count)

For old_tbl_ct = 1 To 1
    For i = 1 To RDB_field_count
        Input #1, old_table_dat(old_tbl_ct, i)
        'Debug.Print i, old_tbl_ct, old_table_dat(old_tbl_ct, i)
    Next i
    'Debug.Print
Next old_tbl_ct
Close 1

Open NMR_STAR_dict_path + NMR_STAR_schema_filename For Input As 1
Open NMR_STAR_dict_path + new_table For Output As 12

EndDoc = 0
loop_item_ct = 0
saveframe = 0: loop_flag = 0: seqflag = 0: line_tot = 0
stringflag = 0: total_data_ct = 0
semicolon_flag = 0: semiflag = 0
tag_count = 0
taglist_ct = 0: tag_old_pos = 0

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
        If loop_flag = 0 Then taglist_ct = 0
        If loop_flag = 1 And taglist_ct = 0 Then
            lfmanDBtable = lfmanDBtable + 1
            lfmanDBtable_string = Trim(Str$(lfmanDBtable))
            ln = Len(lfmanDBtable_string)
            For j5 = ln To 6
                lfmanDBtable_string = "0" + lfmanDBtable_string
            Next j5
        End If
        If data_tag_flag = 1 Then
            manDBcolumn = manDBcolumn + 1
            manDBcolumn_string = Trim(Str$(manDBcolumn))
            ln = Len(manDBcolumn_string)
            For j5 = ln To 3
                manDBcolumn_string = "0" + manDBcolumn_string
            Next j5
            
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
                    sfmanDBtable = sfmanDBtable + 1
                    sfmanDBtable_string = Trim(Str$(sfmanDBtable))
                    ln = Len(sfmanDBtable_string)
                    For j5 = ln To 2
                        sfmanDBtable_string = "0" + sfmanDBtable_string
                    Next j5
                End If

            End If
            
            tag_count = tag_count + 1
            
            tag_list(tag_count, 0) = saveframe_name
            tag_list(tag_count, 1) = tag_pair(1)
            tag_list(tag_count, 2) = saveframe_category_name
'            p = InStr(1, tag_pair(1), "saveframe_ID")
'            If p > 0 Then
'                Debug.Print tag_pair(1)
'                tag_pair(1) = Left$(tag_pair(1), p - 1) + saveframe_category_name + "_ID"
'                Debug.Print tag_pair(1), saveframe_category_name
'                Debug.Print
'            End If
'            p = InStr(1, tag_pair(1), "_Saveframe_ID")
'            If p > 0 Then
'                ln3 = Len(saveframe_category_name)
'                tag_pair(1) = "_" + Chr$(Asc(Left$(saveframe_category_name, 1)) - 32) + Right$(saveframe_category_name, ln3 - 1) + "_ID"
'                Debug.Print tag_pair(1), saveframe_category_name
'                Debug.Print
'            End If
            ln = Len(tag_pair(1))
            If loop_flag = 1 Then
                loop_item_ct = loop_item_ct + 1
                tag_list(tag_count, 5) = "Y"        'loop mandatory true
                tag_list(tag_count, 11) = Str$(loop_item_ct * 10)   'loop position
                If loop_item_ct = 1 Then
                    dbmantable = saveframe_category_name + tag_pair(1)
                    tag_list(tag_count, 20) = dbmantable
                    tag_list(tag_count, 21) = Right$(tag_pair(1), ln - 1)
                End If
                If loop_item_ct > 1 Then
                    tag_list(tag_count, 20) = dbmantable
                    tag_list(tag_count, 21) = Right$(tag_pair(1), ln - 1)
                End If
            End If
            If loop_flag = 0 Then
                tag_list(tag_count, 5) = "N"
                tag_list(tag_count, 20) = saveframe_category_name
                tag_list(tag_count, 21) = Right$(tag_pair(1), ln - 1)
            End If
            
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
If tag_count > last_tag_count Then
    If tag_list(tag_count, 5) = "Y" Then
        taglist_ct = taglist_ct + 1
        tag_pos = taglist_ct * 10
    End If
    If loop_flag = 0 Then
        tag_old_pos = tag_old_pos + 1
        tag_pos = tag_old_pos * 10
        taglist_ct = 0
    End If
    For j2 = 1 To RDB_field_count
        If old_table_dat(1, j2) = "Dictionary sequence" Then Print #12, Trim(Str((tag_count * 10)));
        If old_table_dat(1, j2) = "SFCategory" Then Print #12, tag_list(tag_count, 2);
        If old_table_dat(1, j2) = "TagName" Then Print #12, tag_list(tag_count, 1);
        If old_table_dat(1, j2) = "ManDBTableName" Then Print #12, tag_list(tag_count, 20);
        If old_table_dat(1, j2) = "ManDBColumnName" Then Print #12, tag_list(tag_count, 21);
        If old_table_dat(1, j2) = "Loopflag" Then Print #12, tag_list(tag_count, 5);
        If old_table_dat(1, j2) = "Seq" Then Print #12, tag_pos;

        If j2 < RDB_field_count Then Print #12, ",";
        If j2 = RDB_field_count Then Print #12, "?"
    Next j2
    
    'Debug.Print tag_list(tag_count, 0); ","; tag_list(tag_count, 1); ","; tag_list(tag_count, 2); ","; tag_count * 10
    'Debug.Print
    last_tag_count = tag_count

        End If
       
    End If
    t = ""
Wend
Close 1

Close 12
'Debug.Print
'alphasort tag_count, tag_list
Open "c:\bmrb\htdocs\dictionary\htmldocs\nmr_star\dictionary_files\sch_tag_list.txt" For Output As 12
For i = 1 To tag_count
    Print #12, tag_list(i, 2); ",";
    Print #12, tag_list(i, 1)
Next i
Close 12
Open "c:\bmrb\htdocs\dictionary\htmldocs\nmr_star\dictionary_files\sch_tag_dup_list.txt" For Output As 12
Print #12, "List of duplicate tags"
Print #12,
For i = 1 To tag_count
    For j = 1 To tag_count
        If i <> j Then
        If tag_list(i, 1) = tag_list(j, 1) And tag_list(i, 2) = tag_list(j, 2) Then
            Print #12, tag_list(i, 2); ",";
            Print #12, tag_list(i, 1)
        End If
        End If
    Next j
Next i
Close 12

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

Public Sub data_block_check(t, data_block_flag, data_block_name, data_count, saveframe_flag, loop_flag, loop_item_ct)

p = InStr(1, t, "data_")
If p = 1 Then
    data_block_flag = 1: data_count = 0: saveframe_flage = 0: loop_flag = 0
    ln = Len(t)
    data_block_name = Right$(t, (ln - 5))
End If
        
End Sub
Sub remove_comments(t, semicolon_flag)
        
p = InStr(1, t, "#")
p1 = InStr(1, t, ";")
If semicolon_flag = 0 Or p1 = 1 Then
    If p = 1 Then t = ""
    If p > 1 Then
        t = Left$(t, p - 1)
        t = RTrim$(t)
        stringflag = 0
        For i = 1 To p - 1
            If Mid$(t, i, 1) > " " Then stringflag = 1
        Next i
        If stringflag = 0 Then t = ""
    End If
End If

End Sub
Sub saveframe_check(t, saveframe_flag, loop_flag, loop_item_ct, data_read_flag, saveframe_name, save_names, save_names_ct)

p = InStr(1, t, "save_")
If p > 0 Then
    saveframe_flag = 1
    loop_flag = 0: loop_item_ct = 0: data_read_flag = 0
    ln = Len(t)
    If ln = 5 Then saveframe_flag = 0
    If saveframe_flag = 1 Then
        p2 = InStr(p, t, " ")
        If p2 < p Then saveframe_name = Mid$(t, p + 5, (ln - (p - 4)))
        If p2 > p Then saveframe_name = Mid$(t, p + 5, (p2 - (p - 4)))
        save_names_ct = save_names_ct + 1
        save_names(save_names_ct) = saveframe_name
    End If
End If

End Sub
Sub loop_check(t, loop_flag, loop_item_ct, loop_tag_list, data_read_flag, label_change)

p = InStr(1, t, "loop_")
If p > 0 Then
    loop_flag = 1           'loop condition = true
    loop_item_ct = 0         'loop item counter set to 0
    ReDim loop_tag_list(15)     'field position array
    t = ""
    data_read_flag = 0
End If

p = InStr(1, t, "stop_")
If p > 0 Then                   'loop condition = false
    loop_flag = 0
    t = ""
End If

End Sub
Sub data_tag_check(t, data_tag_flag, data_tag_name)

p = 1: p1 = 1: ln = Len(t)

While p <> 0
    p = InStr(p1, t, "_")
    p2 = InStr(p1, t, " _")

    data_tag_flag = 0
    If p = 1 Then data_tag_flag = 1
    If p > 1 And p2 = p - 1 Then data_tag_flag = 1
    If data_tag_flag = 1 Then
        q = 0       'check to see if data tag is enclosed in a single quote
        For i = 1 To p
            If Mid$(t, i, 1) = "'" Then q = q + 1
        Next i
        If q > 0 And q Mod 2 <> 0 Then data_tag_flag = 0
        q = 0       'check to see if data tag is enclosed in a double quote
        For i = 1 To p
            If Mid$(t, i, 1) = Chr$(31) Then q = q + 1
        Next i
        If q > 0 And q Mod 2 <> 0 Then data_tag_flag = 0
    End If
    If data_tag_flag = 0 And p < ln Then p1 = p + 1
    If data_tag_flag = 0 And p = ln Then p = 0
    If data_tag_flag = 1 Then
        p2 = InStr(p, t, " ")
        If p2 < p Then data_tag_name = Right$(t, (ln - p) + 1)
        If p2 > p Then data_tag_name = Mid$(t, p, p2 - p)
        p = 0
    End If
Wend

End Sub
Sub semicolon_check(t, semicolon_flag, total_data_ct, data_count, line_tot, loop_flag, semiflag)

p = InStr(1, t, ";")
If p = 1 Then
    If semicolon_flag = 0 Then
        semicolon_flag = 1
        semiflag = 0
        If loop_flag = 1 Then total_data_ct = total_data_ct + 1
    End If
    If semicolon_flag = 1 And semiflag = 1 Then
        semicolon_flag = 0
        data_count = data_count + 1
    End If
    If semicolon_flag = 1 Then
        semiflag = 1: line_tot = 0
    End If
End If

End Sub

Sub get_value_from_line(t, tag_pair)

Dim data_value As String
Dim error_report As Integer

error_report = 2
data_value = ""

value_in_line_search t, data_value, error_report

If Left$(data_value, 1) = "'" Then
    single_quote_check t, data_value, error_report
End If

If Left$(data_value, 1) = Chr$(34) Then
    double_quote_check t, data_value, error_report
End If

If error_report = 0 Then
    data_value = Mid$(data_value, 2, ln - 2)
End If

tag_pair(2) = data_value

End Sub

Sub single_quote_check(t, data_value, error_report)

ReDim pt(20)
p = 1: p1 = 0: ptc = 0
While p > 0
    p = InStr(p1 + 1, t, "'")
    If p > 0 Then
        ptc = ptc + 1
        pt(ptc) = p
    End If
p1 = p
Wend

If ptc > 0 Then
    p2 = InStr(1, t, Chr$(34))         'check to see if ' is enclosed by "
    If p2 = 0 Or p2 > pt(1) Then
        If ptc Mod 2 <> 0 Then          'only looks for even number of single quotes
            'Print "ERROR in single quote count"
            'Print t
            'Print "line number = "; linecount
            'Print
            'Print #6, "ERROR in single quote count"
            'Print #6, t
            'Print #6, "line number = "; linecount
            'Print #6,
            errorct = 1
            error_report = 1
        End If
    End If
End If

End Sub

Sub double_quote_check(t, data_value, error_report)

Dim pt(20) As Integer

p = 1: p1 = 0: ptc = 0

While p > 0
    p = InStr(p1 + 1, t, Chr$(34))
    If p > 0 Then
        ptc = ptc + 1
        pt(ptc) = p
    End If
    p1 = p
Wend
If ptc > 0 Then
    p2 = InStr(1, t, Chr$(31))             'check to see if ' is enclosed by "
    If p2 = 0 Or p2 > pt(1) Then
        If ptc Mod 2 <> 0 Then       'checks only for even number of double quotes
            'Print "ERROR in double quote count"
            'Print t
            'Print "line number = "; linecount
            'Print
            'Print #6, "ERROR in double quote count"
            'Print #6, t
            'Print #6, "line number = "; linecount
            'Print #6,
            errorct = 1
            error_report = 1
        End If
    End If
End If

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

Sub clean_line(t)

Dim i As Integer, ln As Integer
Dim pp As String, p3 As String

ln = Len(t): pp = ""       'convert all tabs to spaces
If ln > 0 Then
    For i = 1 To ln
        p3 = Mid$(t, i, 1)
        If p3 = Chr$(9) Then p3 = Space$(5)
        pp = pp + p3
    Next i
    t = pp
End If
t = Trim$(t)             'remove leading and trailing spaces

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
