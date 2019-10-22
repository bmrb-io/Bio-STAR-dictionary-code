Attribute VB_Name = "nmrcif_match_update"
Sub main()

Open "c:\bmrb\htdocs\dictionary\htmldocs\mmcif\pdbx_exch\pdbx_taglist_edit.csv" For Input As 1
Open "c:\bmrb\htdocs\dictionary\htmldocs\nmr_star\development_adit_files\adit_input\new_nmr_cif_match.csv" For Output As 2
Open "c:\bmrb\htdocs\dictionary\htmldocs\nmr_star\development_adit_files\adit_input\nmr_cif_match.csv" For Input As 3

Dim tags(4000, 10) As String
Dim old_tags(4000, 10) As String
Dim t As String

Dim i, ct, j As Integer

ct = 0
While EOF(1) <> -1
    ct = ct + 1
    For i = 3 To 5
        Input #1, tags(ct, i)
    Next i
    Input #1, t
Wend
ct2 = 0
While EOF(3) <> -1
    ct2 = ct2 + 1
    For i = 1 To 8
        Input #3, old_tags(ct2, i)
    Next i
Wend

For i = 1 To ct
    For j = 1 To ct2
        If tags(i, 5) = old_tags(j, 5) Then
            tags(i, 1) = old_tags(j, 1)
            tags(i, 2) = old_tags(j, 2)
            tags(i, 6) = old_tags(j, 6)
            tags(i, 7) = old_tags(j, 7)
            tags(i, 8) = old_tags(j, 8)
            old_tags(j, 10) = "Y"
            Exit For
        End If
    Next j
Next i

For i = 1 To 3
    Print #2, old_tags(i, 1) + "," + old_tags(i, 2) + "," + old_tags(i, 3) + "," + old_tags(i, 4) + "," + old_tags(i, 5) + "," + old_tags(i, 6) + "," + old_tags(i, 7) + "," + old_tags(i, 8)
Next i
For i = 1 To ct
    Print #2, tags(i, 1) + "," + tags(i, 2) + "," + tags(i, 3) + "," + tags(i, 4) + "," + tags(i, 5) + "," + tags(i, 6) + "," + tags(i, 7) + "," + tags(i, 8)
Next i
Print #2, old_tags(ct2, 1) + "," + old_tags(ct2, 2) + "," + old_tags(ct2, 3) + "," + old_tags(ct2, 4) + "," + old_tags(ct2, 5) + "," + old_tags(ct2, 6) + "," + old_tags(ct2, 7) + "," + old_tags(ct2, 8)

Debug.Print "Finished"
Close
End Sub
