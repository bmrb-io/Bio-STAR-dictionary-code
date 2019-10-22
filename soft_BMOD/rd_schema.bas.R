Attribute VB_Name = "read_schema"
Sub dep_retrieve_line(t, EndDoc)

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
Sub dep_retrieve_text_tag_value(text_value, line_tot, EndDoc)

Dim p As Integer, p1 As Integer

p1 = 1: t = ""
While p <> 1
    dep_retrieve_line t, EndDoc
    'Debug.Print t
    'Debug.Print
    p = InStr(1, t, ";")
    'If p <> 1 Then
        p1 = p1 + 1
        text_value(p1) = t
    'End If
    t = ""
    If EndDoc = 1 Then p = 1
Wend
line_tot = p1
'Debug.Print "ENDOFTEXT"

End Sub
Public Sub global_block_check(t, global_block_flag, data_count, saveframe_flag, loop_flag, loop_item_ct)

p = InStr(1, t, "global_")
If p = 1 Then
    global_block_flag = 1
    data_count = 0: saveframe_flage = 0: loop_flag = 0
End If
        
End Sub
Public Sub data_block_check(t, data_block_flag, data_block_name, data_count, saveframe_flag, loop_flag, loop_item_ct)

p = InStr(1, t, "data_")
If p = 1 Then
    data_block_flag = 1: data_count = 0: saveframe_flag = 0: loop_flag = 0
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

error_report = 0
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
If error_report = 0 And pt(1) = 1 Then
    data_value = Mid$(t, pt(1) + 1, (pt(2) - pt(1)) - 1)
End If
End Sub

Sub double_quote_check(t, data_value, error_report)

Dim pt(20) As Integer
error_report = 0
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
If error_report = 0 And pt(1) = 1 Then
    data_value = Mid$(t, pt(1) + 1, (pt(2) - pt(1)) - 1)
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

Sub read_Dep(tag_list, tag_count, dep_info, dep_value_ct)

Dim i As Integer, loop_item_ct As Integer, stringflag As Integer
Dim global_block_flag As Integer, semicolon_flag As Integer
Dim line_count As Integer, EndDoc As Integer, data_tag_flag As Integer
Dim q As Integer, loop_flag As Integer, data_block_flag As Integer
Dim total_data_ct As Integer, semiflag As Integer
Dim t As String, tt As String, p3 As String, pp As String
Dim line_tot As Integer, p2 As Integer
Dim saveframe_category_name As String
Dim original_line As String, data_block_name As String

ReDim save_names(65) As String
Dim save_names_ct As Integer

ReDim loop_tag_list(15) As String
ReDim loop_value(15) As String

Dim tag_pair(1 To 3) As String

ReDim loopdatatag(30) As String
Dim text_value(50) As String
ReDim loop_text_value(11, 50) As String
Dim tag_use(1500)

Open "d:\bmrb\projects\software\NMR_STAR_Dict\testdir.lst" For Input As 2
While EOF(2) <> -1
    Input #2, filename$
    p = InStr(1, filename$, " ")
    filename$ = Left$(filename$, p - 1) + ".dep"
    Debug.Print filename$
    
    Open "d:\bmrb\projects\software\datadep\autous\" + filename$ For Input As 1

EndDoc = 0
loop_item_ct = 0
saveframe = 0: loop_flag = 0: seqflag = 0: line_tot = 0
stringflag = 0: total_data_ct = 0
semicolon_flag = 0: semiflag = 0

global_block_flag = 0
data_block_flag = 0

tt = "": t = ""

While EndDoc = 0
                'define a text line
    dep_retrieve_line t, EndDoc
'    Debug.Print t
'    Debug.Print

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
    
        'saveframe_check t, saveframe_flag, loop_flag, loop_item_ct, data_read_flag, saveframe_name, save_names, save_names_ct

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

            End If
            
    'store the data tags in the current loop
            If loop_flag = 1 Then
                If loop_item_ct = 0 Then
                    ReDim loop_text_value(1 To 11, 1 To 50) As String
                    ReDim loop_value(15) As String
                    prompt_line_ct = 0
                    example_line_ct = 0
                    help_line_ct = 0

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
                End If
                If semicolon_flag = 1 Then
                    text_value(1) = t: line_tot = 1
                    dep_retrieve_text_tag_value text_value, line_tot, EndDoc
                    loop_value_ct = loop_value_ct + 1
                    If loop_value_ct = 6 Then example_line_ct = line_tot
                    If loop_value_ct = 8 Then help_line_ct = line_tot
                    For i = 1 To line_tot
                        loop_text_value(loop_value_ct, i) = text_value(i)
                        'Debug.Print text_value(i)
                    Next i
                    semicolon_flag = 0
                End If
            End If
        End If
           
        If loop_value_ct = loop_item_ct Then
        'Debug.Print loop_value(11)
            p = InStr(1, loop_value(11), ";")
            If p > 2 Then
                nmr_dep_eq = Left$(loop_value(11), p - 1)
            Else
                nmr_dep_eq = ""
            End If
            If nmr_dep_eq > "" Then
            For i = 1 To tag_count
                If tag_list(i, 1) = nmr_dep_eq Then
                If tag_use(i) = 0 Then
                    tag_use(i) = 1
                    'Debug.Print i, tag_list(i, 1)
                    If loop_value(4) > " " Then prompt_line_ct = 1
                    If loop_value(6) > " " Then example_line_ct = 1
                    If loop_value(8) > " " Then help_line_ct = 1
                    
                    dep_value_ct = dep_value_ct + 1
                    dep_info(dep_value_ct, 1) = Str$(i)
                    dep_info(dep_value_ct, 2) = Str$(prompt_line_ct)
                    dep_info(dep_value_ct, 3) = Str$(example_line_ct)
                    dep_info(dep_value_ct, 4) = Str$(help_line_ct)
                    'Debug.Print dep_info(dep_value_ct, dep_line_ct)
                    
                    dep_line_ct = 5
                    
                        For k = 1 To prompt_line_ct
                            If prompt_line_ct = 1 Then dep_info(dep_value_ct, dep_line_ct) = loop_value(4)
                            If prompt_line_ct > 1 Then
                                dep_info(dep_value_ct, dep_line_ct) = loop_text_value(4, k)
                            End If
                            'Debug.Print dep_info(dep_value_ct, dep_line_ct)
                            'Debug.Print
                            dep_line_ct = dep_line_ct + 1
                        Next k
                        For k = 1 To example_line_ct
                            If example_line_ct = 1 Then dep_info(dep_value_ct, dep_line_ct) = loop_value(6)
                            If example_line_ct > 1 Then
                                dep_info(dep_value_ct, dep_line_ct) = loop_text_value(6, k)
                            End If
                            'Debug.Print dep_info(dep_value_ct, dep_line_ct)
                            'Debug.Print
                            dep_line_ct = dep_line_ct + 1
                        Next k
                        For k = 1 To help_line_ct
                            If help_line_ct = 1 Then dep_info(dep_value_ct, dep_line_ct) = loop_value(8)
                            If help_line_ct > 1 Then
                                dep_info(dep_value_ct, dep_line_ct) = loop_text_value(8, k)
                            End If
                            'Debug.Print dep_info(dep_value_ct, dep_line_ct)
                            dep_line_ct = dep_line_ct + 1
                            
                        Next k
                    'Debug.Print
                    prompt_line_ct = 0
                    example_line_ct = 0
                    help_line_ct = 0
                    Exit For
                End If
                End If
            Next i
            prompt_line_ct = 0
            example_line_ct = 0
            help_line_ct = 0

            End If
            loop_value_ct = 0
        End If
    End If
  End If
    t = ""
Wend
Close 1
Wend
Close 2
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
Debug.Print query_string
Debug.Print match_string
Debug.Print subject_string
Debug.Print
ln = Len(query_string)
For i = 1 To ln
    BMRB_res_seq4a(4, query_seq(1, 2) + i - 1) = Mid$(query_string, i, 1)
    BMRB_res_seq4a(5, query_seq(1, 2) + i - 1) = Mid$(match_string, i, 1)
    BMRB_res_seq4a(6, query_seq(1, 2) + i - 1) = Mid$(subject_string, i, 1)
Next i

End Sub
