Attribute VB_Name = "rdmmdict"
Sub main()

'Open "c:\bmrb\htdocs\dictionary\htmldocs\mmcif\cif_mm.20001109.dic" For Input As 1
'Open "c:\bmrb\htdocs\dictionary\htmldocs\mmcif\mmcif_taglist.csv" For Output As 2

Open "z:" + "\eldonulrich" + "\bmrb\htdocs\dictionary\htmldocs\mmcif\pdbx_exch\mmcif_pdbx_v5_next_20150814.dic" For Input As 1
Open "z:" + "\eldonulrich" + "\bmrb\htdocs\dictionary\htmldocs\mmcif\pdbx_exch\pdbx_internal_20150814.csv" For Output As 2
Open "z:" + "\eldonulrich" + "\bmrb\htdocs\dictionary\htmldocs\mmcif\pdbx_exch\ddl_dict_tags_20150814.csv" For Output As 3

'Open "c:\bmrb\htdocs\dictionary\htmldocs\mmcif\rcsb_local_dic" For Input As 1
'Open "c:\bmrb\htdocs\dictionary\htmldocs\mmcif\rcsb_local_taglist.csv" For Output As 2

'Open "c:\bmrb\htdocs\dictionary\htmldocs\mmcif\iims.dic" For Input As 1
'Open "c:\bmrb\htdocs\dictionary\htmldocs\mmcif\iims_taglist.csv" For Output As 2

Dim dict_tag(500) As String
Dim tagcat_list(10000), tagfield_list(10000), tag_list(10000) As String
tag_list_ct = 0
tag_ct = 0
ln_ct = 0
Debug.Print "Start"

Dim t, tt, tag As String
While EOF(1) <> -1
    tt = Input(1, 1)
    t1 = Asc(tt)
    If t1 <> 10 And t1 <> 13 Then t = t + tt
    If t1 = 10 Or t1 = 13 Then
        ln_ct = ln_ct + 1
        If Left$(t, 6) = "save__" Then
            i = i + 1
            tag_list_ct = tag_list_ct + 1
            ln = Len(t)
            tag = Right$(t, ln - 5)
            Trim (tag)
            Debug.Print i, tag
            Debug.Print
            p = InStr(1, tag, ".")
            tagcat = Left$(tag, p - 1)
            ln = Len(tagcat)
            tagcat = Right$(tagcat, ln - 1)
            ln = Len(tag)
            tagfield = Right$(tag, ln - p)
            tagcat_list(tag_list_ct) = tagcat
            tagfield_list(tag_list_ct) = tagfield
            tag_list(tag_list_ct) = tag
            For j = 1 To tag_list_ct - 1
                If tag_list(j) = tag Then tag_list_ct = tag_list_ct - 1
            Next j
        End If
        ln = Len(t)
        For i1 = 1 To ln
            If Mid$(t, i1, 1) <> " " Then
                t = Right$(t, ln - (i1 - 1))
                Exit For
            End If
        Next i1
        'Debug.Print t
        'Debug.Print
        If Left$(t, 1) = ";" And no_tag = 1 Then no_tag = 0: t = ""
        If Left$(t, 1) = ";" Then no_tag = 1
        If Left$(t, 1) = "_" And no_tag = 0 Then
            p = InStr(1, t, " ")
            p1 = InStr(1, t, Chr$(9))
            p2 = InStr(1, t, "'")
            If p > 0 Then test_tag = Left$(t, p - 1)
            If p1 > 0 Then test_tag = Left$(t, p1 - 1)
            'If p2 > 0 Then test_tag = Left$(t, p2 - 2)
            If p = 0 And p1 = 0 Then test_tag = t
            'Debug.Print ln_ct, test_tag
            'Debug.Print
            hit = 0
            For i1 = 1 To tag_ct
                If test_tag = dict_tag(i1) Then
                    hit = 1
                    Exit For
                End If
            Next i1
            If hit = 0 Then
                tag_ct = tag_ct + 1
                dict_tag(tag_ct) = test_tag
                Debug.Print ln_ct, tag_ct, test_tag
                'Debug.Print
            End If
        End If
        t = ""
    End If
Wend

For i = 1 To tag_ct
    Print #3, dict_tag(i)
    Debug.Print i, dict_tag(i)
Next i

For i = 1 To tag_list_ct
    Print #2, tagcat_list(i) + "," + tagfield_list(i) + "," + tag_list(i);
    Print #2,
Next i
Close

Debug.Print "Finished"
End
End Sub
