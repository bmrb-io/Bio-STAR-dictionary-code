Attribute VB_Name = "entry_tag_insert"

Sub main()

Open "c:\bmrb\htdocs\dictionary\htmldocs\nmr_star\adit_files\xlschem_ann2.csv" For Input As 1
Open "c:\bmrb\htdocs\dictionary\htmldocs\nmr_star\adit_files\xlschem_ann3.csv" For Output As 2

Dim org_tags(4000, 82) As String
Dim new_tags(4000, 82) As String
Dim tag_ct, i, j, i1 As Integer

tag_ct = 0
While EOF(1) <> -1
    tag_ct = tag_ct + 1
    For i = 1 To 82
        Input #1, org_tags(tag_ct, i)
    Next i
    Debug.Print tag_ct, org_tags(tag_ct, 9)
    'Debug.Print
Wend

i1 = 0
For i = 1 To tag_ct
    i1 = i1 + 1
        For j = 1 To 82
            new_tags(i1, j) = org_tags(i, j)
        Next j
    
    p = InStr(1, org_tags(i, 53), "Pointer to '_")
    If p > 0 Then
        p = InStr(1, org_tags(i, 53), "'")
        If p > 0 Then
            p1 = InStr(p + 1, org_tags(i, 53), "'")
            ln = Len(org_tags(i, 53))
            If p1 > 0 Then tag1 = Mid$(org_tags(i, 53), p + 1, p1 - p)
            
            p2 = InStr(1, tag1, ".")
            p3 = InStr(1, tag1, "'")
            table1 = Mid$(tag1, 2, p2 - 2)
            column1 = Mid$(tag1, p2 + 1, p3 - (p2 + 1))
            new_tags(i, 38) = table1
            new_tags(i, 39) = column1
            new_tags(i, 64) = "_" + table1 + "." + column1
            Debug.Print i, table1, column1
            'Debug.Print
        End If
    End If

'    If org_tags(i, 34) <> "Y" Then                  'used to insert _Entry.ID tags
'        i1 = i1 + 1
'        For j = 1 To 80
'            new_tags(i1, j) = org_tags(i, j)
'        Next j
'    End If
'    If org_tags(i, 34) = "Y" Then
'        If org_tags(i, 4) = "R" Then
'            i1 = i1 + 1
'            For j = 1 To 80
'                new_tags(i1, j) = org_tags(i, j)
'            Next j
'        End If
'        If org_tags(i, 4) = "T6" Then
'            i1 = i1 + 1
'            For j = 1 To 80
'                new_tags(i1, j) = org_tags(i, j)
'            Next j
'            i1 = i1 + 1
'            For j = 1 To 80
'                new_tags(i1, j) = org_tags(i, j)
'            Next j
'            p = InStr(1, org_tags(i, 9), ".")
'            new_tags(i1, 9) = Left$(org_tags(i, 9), p) + "Entry_ID"
'            new_tags(i1, 27) = "Entry ID"
'            new_tags(i1, 28) = "CHAR(12)"
'            new_tags(i1, 32) = "EntryID"
'            new_tags(i1, 34) = "N"
'            new_tags(i1, 36) = ""
'            new_tags(i1, 38) = "Entry"
'            new_tags(i1, 39) = "_Entry.ID"
'            new_tags(i1, 53) = "Pointer to _Entry.ID"
'            new_tags(i1, 62) = "_Entry.ID"
'            new_tags(i1, 64) = "_Entry.ID"
'            new_tags(i1, 65) = "R"
'            new_tags(i1, 66) = "R"
'            new_tags(i1, 67) = "R"
'            new_tags(i1, 68) = "R"
'            new_tags(i1, 69) = "R"
'            new_tags(i1, 70) = "R"
'            new_tags(i1, 79) = "?"
'            new_tags(i1, 80) = "?"
'        End If
'    End If
Next i
For i = 1 To i1
    For j = 1 To 81
        Print #2, new_tags(i, j) + ",";
    Next j
    Print #2, new_tags(i, 82)
Next i
Debug.Print "Finished"
        
End Sub
