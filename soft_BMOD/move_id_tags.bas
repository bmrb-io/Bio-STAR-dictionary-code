Attribute VB_Name = "move_id_tags"
Sub main()

Debug.Print "Start"

Dim tag_item(5000, 86) As String
ReDim cat_items(200, 86) As String

Open "c:\bmrb\htdocs\dictionary\htmldocs\nmr_star\adit_files\xlschem_ann_abs_newest.csv" For Input As 1
Open "c:\bmrb\htdocs\dictionary\htmldocs\nmr_star\adit_files\xlschem_ann_abs_newest-2.csv" For Output As 2

While EOF(1) <> -1
    row_ct = row_ct + 1
    For i = 1 To 86
        Input #1, tag_item(row_ct, i)
    Next i
Wend
Close 1

For i = 1 To 4
    For j = 1 To 85
        Print #2, tag_item(i, j) + ",";
    Next j
    Print #2, tag_item(i, 86)
Next i

i = 5: cont = 1
While cont = 1
    cat_ct = 0
    ReDim cat_items(200, 86) As String
    
    While tag_item(i, 79) = tag_item(i + 1, 79)
        cat_ct = cat_ct + 1
        For j = 1 To 86
            cat_items(cat_ct, j) = tag_item(i, j)
        Next j
        i = i + 1
    Wend
    cat_ct = cat_ct + 1
    For j = 1 To 86
        cat_items(cat_ct, j) = tag_item(i, j)
    Next j
    For k = 1 To cat_ct
        If cat_items(k, 86) <> "Y" Then
            For m = 1 To 85
                Print #2, cat_items(k, m) + ",";
            Next m
            Print #2, cat_items(k, 86)
        End If
    Next k
    For k = 1 To cat_ct
        If cat_items(k, 86) = "Y" Then
            For m = 1 To 85
                Print #2, cat_items(k, m) + ",";
            Next m
            Print #2, cat_items(k, 86)
        End If
    Next k
    i = i + 1
    If i = row_ct Then cont = 0
Wend
For m = 1 To 85
    Print #2, tag_item(row_ct, m) + ",";
Next m
Print #2, tag_item(row_ct, 86)

Close 2
Debug.Print "Finished"
    
    
End Sub
