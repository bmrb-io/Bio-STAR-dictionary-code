Attribute VB_Name = "Module1"
'Create data dictionary from the deposition form    (nmrdic.bas   7/23/2000)

                               
Sub main()

Dim printout, date1 As String

Dim noglobal, nolabel, totaltokens, totalsavfrm As Integer


printout = "c:\bmrb\projects\datadict\nmrdict.tst"


date1 = "(9/30/97)"
totaltokens = 4000       'value for expected number of tokens in the file
totalsavfrm = 200       'value for expected number of saveframes in the file

noglobal = 0             '0 = ignore up to global
                          '1 = instruction section does not exist

nolabel = 0              '0 = include labels on save frame statements
                          '1 = leave labels off of save frame statements

 readdraft noglobal, nolabel, totaltokens, totalsavfrm, date1, printout
'GOSUB dictsetup
'GOSUB dictprint2


End Sub


Sub readdraft(noglobal, nolabel, totaltokens, totalsavfrm, date1, printout)



Open "c:\bmrb\htdocs\harvest.txt" For Input As 1
'OPEN "k:\depdraft.96" FOR INPUT AS 1
Open "c:\bmrb\projects\datadict\sfcat.lst" For Output As 4
Open "c:\bmrb\projects\datadict\dictdt.lst" For Output As 5

ReDim token(totaltokens) As String
ReDim savestr(totalsavfrm)

Dim t, tt As String
Dim t1, p0, p1, i, j, k, n, n1, linect As Integer
Dim offset, numrec As Integer

outputfile = "c:\bmrb\htdocs\harvest_tokens.lst"

n = 0: test2 = 0: save = 0: charct = 0: linect = 0

While EOF(1) <> -1
    t = Input(1, 1)
    t1 = Asc(t)
    If t1 <> 10 Then
        tt = tt + t
        If t1 <> 32 Then charct = charct + 1
    End If
    If t1 = 10 Then
        linect = linect + 1

        tt = Trim(tt)
        

            'locate key characters and statements
        p0 = InStr(1, tt, "_")
                p1 = InStr(1, tt, "'")
                p2 = InStr(1, tt, "data_")
                p3 = InStr(1, tt, "#")
                p4 = InStr(1, tt, "<")
                p5 = InStr(1, tt, "save_")
                p6 = InStr(1, tt, "loop_")
                p7 = InStr(1, tt, "stop_")
                p8 = InStr(1, tt, "e.g.:")
                p9 = InStr(1, tt, "#%!")
                p10 = InStr(1, tt, "#%&")
                p11 = InStr(1, tt, "#%*")

                If p2 = 1 Or noglobal = 1 Then 'locate beginning of file defined by global_
                        If p3 = 0 Then test2 = 1
                End If
         If p3 <> 1 Then
                If test2 = 1 Then

                If p5 = 1 Then                  'locate save_
                    test3 = 1
                    'PRINT "Hi": SLEEP
                    If Len(tt) > 5 Then 'is a save_xxx
                        savestr_num = savestr_num + 1
                        save = 1
                        'SLEEP
                    Else: saveend = saveend + 1: save = 0

                    End If
                End If

                p = InStr(1, tt, "_Saveframe_category")
                If p > 0 Then
                        p1 = InStr(p, tt, " ")
                        ln = Len(tt)
                        savestr(savestr_num) = Right$(tt, ln - p1)
                        savestr(savestr_num) = Trim(savestr(savestr_num))

                        Print #4, savestr_num, savestr(savestr_num)
                        Debug.Print savestr(savestr_num)
                        'SLEEP 1
                End If

        'Determine if a loop condition exists

                p = InStr(1, tt, "loop_")
                If p = 1 Then
                        loopflag = 1           'loop condition = true
                        loop1 = loop1 + 1       'loop counter
                End If
                p = InStr(1, tt, "stop_")
                If p = 1 Then
                        loopflag = 0
                        loopend = loopend + 1
                End If

                
'locate valid saveframe types for data tags that request a framecode as a value
 
'locate any examples
                If p8 > 0 Then

                End If
 
                If p0 = 1 Then
                If Len(tt) > 1 Then
                        mandatory = 0
                        p3 = InStr(1, tt, "MANDATORY")
                        If p3 > 1 Then mandatory = 1
                        t2 = Asc(Mid$(tt, p0, 1))
                                ln = Len(tt)
                                endtok = 1: tt2 = ""
                                For i = p0 To ln
                                        t = Mid$(tt, i, 1)
                                        t2 = Asc(t)
                                        If t2 = 32 Then endtok = 0
                                        If t2 = 9 Then endtok = 0
                                        If t2 = 13 Then endtok = 0
                                        If t2 = 10 Then endtok = 0
                                        If endtok = 1 Then tt2 = tt2 + t
                                        endtokstr = i - 1
                                        If endtok = 0 Then Exit For
                                Next i
                                        
                                If tt2 > "" Then
                                If Len(tt2) > 1 Then
                                   If Right$(tt2, 1) <> "'" Then
                                        n = n + 1
                                        token(n) = tt2
                                        Debug.Print "token = "; token(n), n, mandatory; "   "; loopflag
                                        'SLEEP 1
                                        Write #5, n, savestr_num, token(n), mandatory, loopflag
                                        tt2 = ""
                                   End If
                                End If
                                End If
                End If
                End If
                End If
                tt = ""
                tt2 = ""
                charct = 0
        End If
        tt = ""
        End If

Wend
Close 1
Close 4
Close 5


For i = 1 To n
        For j = 1 To n
                If i <> j Then
                        If token(i) = token(j) Then token(j) = ""
                End If
        Next j
Next i

n1 = 0
For i = 1 To n
        If token(i) <> "" Then
                n1 = n1 + 1
                token(n1) = token(i)
        End If
Next i
n = n1



'alphasort
numrec = n
offset = numrec \ 2
Do While offset > 0
        limit = numrec - offset
        Do
                test = 0
                For i = 1 To limit
                        If token(i) > token(i + offset) Then
                           temp_token = token(i)
                           token(i + offset) = token(i)
                           token(i) = temp_token
                           test = i
                        End If
                Next i
                limit = test
        Loop While test > 0
        offset = offset \ 2
Loop
'RETURN

Open outputfile For Output As 2
Print #2, "Alphabetized tokens from BMRB deposition form " + date1; ""
Print #2,
For i = 1 To n
        Print #2, Space$(6); i, token(i)
        'IF outputfile = "scrn:" THEN SLEEP 1
Next i
If outputfile = "lpt1:" Then Print #2, Chr$(12);
Close 2
tottoken = n

Open "c:\bmrb\projects\datadict\dictdt.lst" For Input As 1
ReDim saveframe(totaltokens), token2(totaltokens)
ReDim looptest(totaltokens), mantest(totaltokens)
n = 0
While EOF(1) <> -1
        n = n + 1
        Input #1, t, saveframe(n), token2(n), mantest(n), looptest(n)
Wend
Close

For i = 1 To n
        For j = 1 To n
                If i <> j Then
                        If saveframe(i) = saveframe(j) Then
                        If token2(i) = token2(j) Then
                                token(j) = "": token2(j) = ""
                        End If
                        End If
                End If
        Next j
Next i

n1 = 0
For i = 1 To n
        If token2(i) <> "" Then
                n1 = n1 + 1
                saveframe(n1) = saveframe(i)
                token2(n1) = token2(i)
                looptest(n1) = looptest(i)
                mantest(n1) = mantest(i)
        End If
Next i
n = n1

Open "c:\bmrb\projects\datadict\dictdt.lst" For Output As 1
For j = 1 To savestr_num
    For k = 1 To tottoken
        For i = 1 To n
            If saveframe(i) = j Then
            If token2(i) = token(k) Then
               Write #1, i, saveframe(i), token2(i), looptest(i), mantest(i)
               Debug.Print i, savestr(saveframe(i)), token2(i)
               'SLEEP
            End If
            End If
        Next i
    Next k
Next j
Close 1

End Sub



Sub dictprint2()


Cls
'OPEN "c:\bmrb\projects\datadict\dict.tst" FOR OUTPUT AS 1
Open "scrn:" For Output As 1
Open "c:\bmrb\projects\datadict\dictdt.lst" For Input As 2
i = 0
While EOF(2) <> -1
        i = i + 1
      Input #2, t, saveframe(i), token2(i), looptest(i), mantest(i)
Wend
Close 2
n = i

Print #1, "NMR-STAR Dictionary Version 2.0"
Print #1,
Print #1, "data_NMR-STAR_Dictionary_v2.0"
Print #1,
Print #1,
For i = 1 To n
    If saveframe(i) <> saveframe(i - 1) Then
       For j = 1 To Len(savestr(saveframe(i))) + 6
           Print #1, "#";
       Next j
       Print #1,
       Print #1, "#  " + savestr(saveframe(i)) + "  #"
       For j = 1 To Len(savestr(saveframe(i))) + 6
           Print #1, "#";
       Next j
       Print #1,
       Print #1,
       Print #1,
       SLEEP
    End If

    Print #1, "save_" + savestr(saveframe(i)) + "." + token2(i)
    Print #1,
    Print #1, "_Saveframe_category                "; savestr(saveframe(i))
    Print #1, "_Tag                               "; "'"; token2(i); "'"
    Print #1, "_Description"
    Print #1, ";"
    Print #1, "    @"
    Print #1, ";"
    Print #1,
    If token2(i) = "_Saveframe_category" Then
       Print #1, "_Required_mandatory                yes"
       Print #1, "_Conditional_mandatory             no"
    End If
    If mantest(i) = 0 And token2(i) <> "_Saveframe_category" Then
       Print #1, "_Required_mandatory                no"
       Print #1, "_Conditional_mandatory             no"
    End If
    If mantest(i) = 1 Then
       Print #1, "_Required_mandatory                @"
       Print #1, "_Conditional_mandatory             yes"
    End If
    Print #1, "_Data_type                         char"
    If token2(i) = "_Saveframe_category" Then
       Print #1, "_Defined_fixed_value               "; savestr(saveframe(i))
    End If
    Print #1, "_Minimum_value                     @"
    Print #1, "_Maximum_value                     @"
    p = InStr(1, token2(i), "label")
    If p > 0 Then
       Print #1, "_Value_is_saveframe_code           yes"
       Print #1, "   loop_"
       Print #1, "      _Value_from_saveframe_category     "
       Print #1, "        @"
       Print #1,
       Print #1, "   stop_"
       Print #1,
    End If
    Print #1, "_Enumerated                        @"
    Print #1, "_Enumeration_dependent             @"
    If looptest(i) = 1 Then Print #1, "_Loop_location_mandatory           yes"
    If looptest(i) = 0 Then Print #1, "_Loop_location_mandatory           no"
    SLEEP
    Print #1, "_Error_message"
    Print #1, ";"
    Print #1, "     @"
    Print #1, ";"
    Print #1,
    Print #1, "_Help_message"
    Print #1, ";"
    Print #1, "     @"
    Print #1, ";"
    Print #1,
    Print #1, "save_"
    Print #1,
    Print #1,
'    SLEEP
Next i
Close
End Sub





Sub dictprint()

Open printout For Output As 3
For i = 1 To n1

Print #3, "save" + token(i)
Print #3,
Print #3, "     _name              " + token(i)

Print #3, "     _description"
Print #3, ";"
Print #3, "    @    "
Print #3,
Print #3, ";"

Print #3,
Print #3, "     _data_type         " + "char"
Print #3,
Print #3, "  loop_"
Print #3, "     _mmCIF_equivalent_tag   "
Print #3,
For j = 1 To n
        If token(i) = token2(j) Then
           Print #3, savestr(saveframe(j)) + "." + Right$(token(i), Len(token(i)) - 1)
        End If
Next j
Print #3,
Print #3, "  stop_"
Print #3,
Print #3, "  loop_"
Print #3, "     _mandatory"
Print #3, "     _mandatory_saveframe_category"
For j = 1 To n
        If token(i) = token2(j) Then
                If mantest(j) = 1 Then Print #3, "          true    ";
                If mantest(j) = 0 Then Print #3, "          false   ";
                Print #3, savestr(saveframe(j))
        End If
Next j
Print #3, "  stop_"
Print #3,
If Label% = 1 Then
        Print #3, "     _saveframe_category_accepted"
End If
Print #3,

Print #3, "  loop_"
Print #3, "     _example"
Print #3, "       @"
Print #3,
Print #3, "  stop_"
Print #3,

Print #3, "  loop_"
Print #3, "     _valid_value"
Print #3, "      @"
Print #3,
Print #3, "  stop_"
Print #3,

'PRINT #3, "  loop_"
'PRINT #3, "     _mandatory_dependency_type"
'PRINT #3, "     _mandatory_dependency_value"
'PRINT #3,
'PRINT #3,
'PRINT #3, "      loop_"
'PRINT #3, "         _dependent_data_tag"
'PRINT #3, "         _dependent_saveframe_category"
'PRINT #3,
'PRINT #3,
'PRINT #3, "      stop_"
'PRINT #3, "  stop_"
'PRINT #3,

'PRINT #3, "  loop_"
'PRINT #3, "     _loop_requirement_type"            'true;  false
'IF looptest(i) = 1 THEN PRINT #3, "          true" ELSE PRINT #3, "          false"
'PRINT #3,
'PRINT #3, "      loop_"
'PRINT #3, "         _loop_dependent_saveframe_category"
'IF looptest(i) = 1 THEN PRINT #3, savestr(saveframe(i))
'PRINT #3,
'PRINT #3, "      stop_"
'PRINT #3, "  stop_"
'PRINT #3,

'PRINT #3, "  loop_"
'PRINT #3, "     _dictionary_version"
'PRINT #3, "     _alias"
'PRINT #3, "   @                @"
'PRINT #3,
'PRINT #3, "  stop_"
'PRINT #3,

'PRINT #3, "  loop_"
'PRINT #3, "     _equivalent_description_data_tag"
'PRINT #3,
'PRINT #3, "        @ "
'PRINT #3, "  stop_"
'PRINT #3,

'PRINT #3, "  loop_"                         'required for chemical shift data tags
'PRINT #3, "     _minimum_valid_value"
'PRINT #3, "     _maximum_valid_value"
'PRINT #3,
'PRINT #3,
'PRINT #3, "      loop_"
'PRINT #3, "         _required_data_tag"
'PRINT #3, "         _required_data_tag_value"
'PRINT #3,
'PRINT #3,
'PRINT #3, "      stop_"
'PRINT #3, "  stop_"

'PRINT #3,
Print #3, "save_"
Print #3,
Print #3,

'SLEEP

Next i

Close 3
End Sub




Sub dictsetup()


   'list of tokens and saveframes that have a mandatory dependency
Open "c:\bmrb\docs\deposit\lsav.lst" For Output As 2
   'saveframes where data tag has a positive loop test
Open "c:\bmrb\docs\deposit\ploop.lst" For Output As 3
   'saveframes where data tag has a negative loop test
Open "c:\bmrb\docs\deposit\nloop.lst" For Output As 4

For i = 1 To n
    If mantest(i) = 1 Then
       For j = 1 To n
           If i <> j Then
              If savestr(saveframe(i)) = savestr(saveframe(j)) Then
                 If mantest(j) = 1 Then
                    Print #2, i, token2(i), savestr(saveframe(i)), token2(j)
                 End If
              End If
           End If
       Next j
    End If
    If looptest(i) = 1 Then
       Print #3, i, token2(i), savestr(saveframe(i))
    End If
    If looptest(i) = 0 Then
       Print #4, i, token2(i), savestr(saveframe(i))
    End If

Next i
Close 2
Close 3
End Sub

