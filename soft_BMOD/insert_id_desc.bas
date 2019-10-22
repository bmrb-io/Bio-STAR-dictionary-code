Attribute VB_Name = "insert_id_desc"

Sub main()

Dim nmrfields(4000, 120) As String
Dim idlist(50, 2) As String
Dim list_ct, ct, i, j, p As Integer

Open "c:\bmrb\htdocs\dictionary\htmldocs\nmr_star\adit_files\xlschem_ann.csv" For Input As 1
Open "c:\bmrb\htdocs\dictionary\htmldocs\nmr_star\adit_files\xlschem_ann_idtest.csv" For Output As 2

field_ct = 80
row_ct = 4
For i = 1 To 4                              ' load header rows from excel file
    For j = 1 To field_ct
        Input #1, nmrfields(i, j)
    Next j
Next i

While EOF(1) <> -1                          ' load tag rows
    row_ct = row_ct + 1
    For j = 1 To field_ct
        Input #1, nmrfields(row_ct, j)
    Next j
    'Debug.Print e_tag_ct, excel_tag_dat(e_tag_ct, 1),
    'Debug.Print excel_tag_dat(e_tag_ct, 80)
    'Debug.Print
Wend


load_idlist list_ct, idlist

For i = 5 To row_ct - 1
    p = InStr(1, nmrfields(i, 9), ".")
    For j = 1 To list_ct
        
        p1 = InStr(p, nmrfields(i, 9), idlist(j, 1))
        If p1 > 0 Then
            If nmrfields(i, 53) = "?" Then nmrfields(i, 53) = (idlist(j, 2))
        End If
    Next j
Next i

For i = 1 To row_ct
    For j = 1 To 79
        Print #2, nmrfields(i, j) + ",";
    Next j
    Print #2, nmrfields(i, 80)
Next i

Debug.Print "Finished"
End Sub


Sub load_idlist(list_ct, idlist)

list_ct = 31
idlist(1, 1) = "Entry_ID":                   idlist(1, 2) = "Pointer to '_Entry.ID'"
idlist(2, 1) = "Entity_ID":                  idlist(2, 2) = "Pointer to '_Entity.ID'"
idlist(3, 1) = "Comp_index_ID":              idlist(3, 2) = "Pointer to '_Entity_comp_index.ID'"
idlist(4, 1) = "Seq_ID":                     idlist(4, 2) = "Pointer to '_Entity_poly_seq.Num'"
idlist(5, 1) = "Comp_ID":                    idlist(5, 2) = "Pointer to '_Chem_comp.ID'"
idlist(6, 1) = "Atom_ID":                    idlist(6, 2) = "Pointer to '_Chem_comp_atom.Atom_ID'"
idlist(7, 1) = "Entry_atom_ID":              idlist(7, 2) = "Pointer to '_Atom.Entry_atom_ID'"
idlist(8, 1) = "comp_index_ID":              idlist(8, 2) = "Pointer to '_Entity_comp_index.ID'"
idlist(9, 1) = "seq_ID":                     idlist(9, 2) = "Pointer to '_Entity_poly_seq.Num'"
idlist(10, 1) = "comp_ID":                   idlist(10, 2) = "Pointer to '_Chem_comp.ID'"
idlist(11, 1) = "atom_ID":                   idlist(11, 2) = "Pointer to '_Chem_comp_atom.Atom_ID'"
idlist(12, 1) = "entry_atom_ID":             idlist(12, 2) = "Pointer to '_Atom.Entry_atom_ID'"
idlist(13, 1) = "entity_ID":                 idlist(13, 2) = "Pointer to '_Entity.ID'"
idlist(14, 1) = "entity_assembly_ID":        idlist(14, 2) = "Pointer to '_Entity_assembly.ID'"
idlist(15, 1) = "Entity_assembly_ID":        idlist(15, 2) = "Pointer to '_Entity_assembly.ID'"
idlist(16, 1) = "Method_ID":                 idlist(16, 2) = "Pointer to '_Method.ID'"
idlist(17, 1) = "Software_label":            idlist(17, 2) = "Pointer to a saveframe of the category 'software'"
idlist(18, 1) = "Method_label":              idlist(18, 2) = "Pointer to a saveframe of the category 'method'"
idlist(19, 1) = "Sample_label":              idlist(19, 2) = "Pointer to a saveframe of the category 'sample'"
idlist(20, 1) = "NMR_spec_expt_label":       idlist(20, 2) = "Pointer to a saveframe of the category 'NMR_spectrometer_expt'"
idlist(21, 1) = "Citation_ID":               idlist(21, 2) = "Pointer to '_Citation.ID'"
idlist(22, 1) = "Citation_label":            idlist(22, 2) = "Pointer to a saveframe of the category 'citation'"
idlist(23, 1) = "Entity_label":              idlist(23, 2) = "Pointer to a saveframe of the category 'entity'"
idlist(24, 1) = "entity_label":              idlist(24, 2) = "Pointer to a saveframe of the category 'entity'"
idlist(25, 1) = "Chem_comp_label":           idlist(25, 2) = "Pointer to a saveframe of the category 'chem_comp'"
idlist(26, 1) = "chem_comp_label":           idlist(26, 2) = "Pointer to a saveframe of the category 'chem_comp'"
idlist(27, 1) = "Atom_type":                 idlist(27, 2) = "Standard symbol used to define the atom element type."
idlist(28, 1) = "atom_type":                 idlist(28, 2) = "Standard symbol used to define the atom element type."
idlist(29, 1) = "Sf_category":               idlist(29, 2) = "Category definition for the information content of the saveframe"
idlist(30, 1) = "Sf_framecode":              idlist(30, 2) = "A label for the saveframe that describes in very brief terms the information contained in the saveframe."
idlist(31, 1) = "Sample_state":              idlist(31, 2) = "Physical state of the sample either anisotropic or isotropic."






End Sub
