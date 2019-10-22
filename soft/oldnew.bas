Attribute VB_Name = "oldnew"
Sub main()
Dim old(4000, 81) As String
Dim newf(4000, 81) As String

Open "c:\bmrb\htdocs\dictionary\htmldocs\nmr_star\adit_files\xlschem_ann_old.csv" For Input As 1
ct = 0
While EOF(1) <> -1
    ct = ct + 1
    For i = 1 To 81
        Input #1, old(ct, i)
    Next i
Wend
Close 1

Open "c:\bmrb\htdocs\dictionary\htmldocs\nmr_star\adit_files\xlschem_ann2.csv" For Input As 1
ctn = 0
While EOF(1) <> -1
    ctn = ctn + 1
    For i = 1 To 81
        Input #1, newf(ctn, i)
    Next i
Wend
Close 1

Open "c:\bmrb\htdocs\dictionary\htmldocs\nmr_star\adit_files\xlschem_ann_newold.csv" For Output As 1
Open "c:\bmrb\htdocs\dictionary\htmldocs\nmr_star\adit_files\lost.csv" For Output As 2

For i = 1 To ct
    If i < 5 Then
        For j = 1 To 80
            Print #1, old(i, j) + ",";
        Next j
        Print #1, old(i, 81)
    End If
    If i > 4 And i < ct Then
        found = 0
        For j = 1 To ctn
            If LCase(old(i, 9)) = LCase(newf(j, 9)) Then
                found = 1
                For k = 10 To 81
                    old(i, k) = newf(j, k)
                Next k
                j5 = j
                Exit For
            End If
        Next j
        If found = 0 Then Print #2, old(i, 9)
        If Left$(old(i, 9), 3) = "_pH" Then
            ln = Len(old(i, 9))
            old(i, 9) = "_PH" + Right$(old(i, 9), ln - 3)
        End If
        For j = 1 To 80
            Print #1, old(i, j) + ",";
        Next j
        Print #1, old(i, 81)
        If found = 1 Then
        p = InStr(1, old(i, 9), ".Sf_ID")
        If p > 0 Then
            If old(i, 4) = "T6" Then
                If old(i, 2) <> "entry_information" Then
                    For j = 1 To 80
                        Print #1, newf(j5 + 1, j) + ",";
                    Next j
                    Print #1, old(j5 + 1, 81)
                End If
            End If
        End If
        End If
    End If
    If i = ct Then
        For j = 1 To 80
            Print #1, old(i, j) + ",";
        Next j
        Print #1, old(i, 81)
    End If
Next i

End Sub
