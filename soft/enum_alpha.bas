Attribute VB_Name = "enum_alpha"

Sub main()        'alphabetize enumerations   5/22/05

ReDim enum_list(500, 2) As String
Dim file_ext1, file_ext As String

file_ext = ""
file_ext1 = Date
ln = Len(file_ext1)
For i = 1 To ln
    t = Mid$(file_ext1, i, 1)
    If t <> "/" Then file_ext = file_ext + t
Next i
file_path = "c:\bmrb\htdocs\dictionary\htmldocs\nmr_star\adit_files\"
file_path2 = "c:\bmrb\htdocs\dictionary\htmldocs\nmr_star\old_dictionary_files\"
'FileCopy file_path + "enumerations.txt", file_path2 + "enumerations_" + file_ext + ".txt"

Open "c:\bmrb\htdocs\dictionary\htmldocs\nmr_star\adit_files\enumerations.txt" For Input As 21
While EOF(21) <> -1
    t = Input$(1, 21)
    t1 = Asc(t)
    If t1 <> 10 And t1 <> 13 Then tt = tt + t
    If t1 = 10 Then
        p = InStr(1, tt, "save__")
        If p <> 0 Then
            ln = Len(tt)
            enum_ct = enum_ct + 1
            If p = 1 Then enum_list(enum_ct, 1) = Trim(Right$(tt, ln - 5))
            If p = 2 Then enum_list(enum_ct, 1) = Trim(Right$(tt, ln - 6))
       
        End If
        tt = ""
    End If
Wend
Close 21

enum_list(0, 1) = "aaaaaaaaaaa"
max_ct = 0
While max_ct < enum_ct
max_id = 0
For i = 0 To enum_ct
'    For j = 1 To enum_ct
        If enum_list(i, 2) = "" Then
'        If enum_list(j, 2) = "0" Then
            If enum_list(i, 1) < enum_list(max_id, 1) Then max_id = i
'        End If
        End If
'    Next j
Next i
    max_ct = max_ct + 1
    enum_list(max_id, 2) = Trim(Str(max_ct))
    'Debug.Print max_ct, max_id, enum_list(max_id, 1), enum_list(max_id, 2)
    'Debug.Print
Wend

Open "c:\bmrb\htdocs\dictionary\htmldocs\nmr_star\adit_files\enumerations_alpha.txt" For Output As 2

For i = 1 To enum_ct

Open "c:\bmrb\htdocs\dictionary\htmldocs\nmr_star\adit_files\enumerations.txt" For Input As 21
print_flg = 1
Print #2,
While EOF(21) <> -1
    t = Input$(1, 21)
    t1 = Asc(t)
    If t1 <> 10 And t1 <> 13 Then tt = tt + t
    If t1 = 10 Then
        p = InStr(1, tt, "save__")
        If p <> 0 Then
            print_flg = 0
            ln = Len(tt)
            If p = 1 Then enum_val = Trim(Right$(tt, ln - 5))
            If p = 2 Then enum_val = Trim(Right$(tt, ln - 6))
            For j = 1 To enum_ct
                If enum_list(j, 2) = Trim(Str(i)) Then
                    If enum_val = enum_list(j, 1) Then
                        print_flg = 2
                        Debug.Print i, j, enum_val
                        'Debug.Print
                    End If
                End If
            Next j
        End If
        
        If i = 1 And print_flg = 1 Then Print #2, tt
        If i > 0 And print_flg = 2 Then Print #2, tt
        If Trim(tt) = "save_" Then print_flg = 0
        If Trim(tt) = "#save_" Then print_flg = 0
        tt = ""
    End If
Wend
Print #2,
Close 21
Next i
Close
FileCopy file_path + "enumerations_alpha.txt", file_path + "enumerations.txt"

End Sub




