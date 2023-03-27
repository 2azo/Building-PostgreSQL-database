Sub H_RunAllMacros(wb As Workbook)
    
    'Application.ScreenUpdating = False
    Call Tables_names
    Call A_projects
    Call B_experiments
    Call C_measu_steps
    Call D_processing_steps
    Call E_MaterialAdditionSteps
    Call F_slurryMaterial
    G_DeleteSheets wb
    'Application.ScreenUpdating = True
End Sub

Sub A_projects()
'
' 1.projects Macro
'
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Name = "1.projects"
    ActiveCell.FormulaR1C1 = "project_name"
    Range("A2").Select
    ActiveCell.FormulaR1C1 = "=Schlickerherstellung!R[-1]C[3]"
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "notes"
    Range("B2").Select
    Range("A1:B2").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone
    
    Range("A1:B2").Select
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("$A$1:$B$2"), , xlYes).Name = _
        "project"
    
End Sub

Sub B_experiments()
'
' experiments Macro
'

'
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Name = "2.experiments"
    ActiveCell.FormulaR1C1 = "experiment_name"
    Range("A2").Select
    Selection.FormulaR1C1 = "=Schlickerherstellung!R[-1]C[1]"
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "project_name"
    Range("B2").Select
    Sheets("1.projects").Select
    Range("A2").Select
    Sheets("2.experiments").Select
    Selection.FormulaR1C1 = "=1.projects!RC[-1]"
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "experiment_date"
    Range("C2").Select
    Selection.FormulaR1C1 = "=Schlickerherstellung!R[-1]C[3]"
    Range("D1").Select
    ActiveCell.FormulaR1C1 = "required_mass_g"
    Range("D2").Select
    Selection.FormulaR1C1 = "=Schlickerherstellung!R[1]C[-2]"
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "required_solid_contents_percentage"
    Range("E2").Select
    Selection.FormulaR1C1 = "=Schlickerherstellung!R[3]C[-3]"
    Range("F1").Select
    ActiveCell.FormulaR1C1 = "mixing_tool"
    Range("F2").Select
    Selection.FormulaR1C1 = "=Schlickerherstellung!R[35]C"
    Range("G1").Select
    ActiveCell.FormulaR1C1 = "mixer"
    Range("G2").Select
    Selection.FormulaR1C1 = "=Schlickerherstellung!R[35]C[-5]"
    Range("A1:G2").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone
    Range("F13").Select
    
    Range("A1:G2").Select
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("$A$1:$G$2"), , xlYes).Name = _
        "experiment"
    
End Sub

Sub C_measu_steps()
'
' test_meas_steps Macro
'

'
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Name = "3.meas.steps"

    ActiveCell.FormulaR1C1 = "measurement_step_id"
    Range("A2").Select
    ActiveCell.Formula2R1C1 = "=measurement_steps_id"

    Range("B1").Select
    ActiveCell.FormulaR1C1 = "measurement_step_number"
    'Range("B2").Select
    'ActiveCell.Formula2R1C1 = "=measurement_after_proces_number"
    'Range("B3").Select
    'Sheets("3.meas.steps").Select
    
    'test
    Range("B2").Select
    ActiveCell.FormulaR1C1 = "1"
    Range("B3").Select
    ActiveCell.FormulaR1C1 = "2"
    Range("B2:B3").Select
    Selection.AutoFill Destination:=Range("B2:B6"), Type:=xlFillDefault

    Range("C1").Select
    ActiveCell.FormulaR1C1 = "experiment_name"
    Range("C2").Select
    ActiveCell.FormulaR1C1 = "=experiment_name"
    Range("C2").Select
    Selection.AutoFill Destination:=Range("C2:C6"), Type:=xlFillDefault

    Range("D1").Select
    ActiveCell.FormulaR1C1 = "project_name"
    Range("D2").Select
    ActiveCell.FormulaR1C1 = "=project_name"
    Range("D2").Select
    Selection.AutoFill Destination:=Range("D2:D6"), Type:=xlFillDefault

    Range("E1").Select
    ActiveCell.FormulaR1C1 = "viscosity_high_1_over_s"
    Range("E2").Select
    ActiveCell.Formula2R1C1 = "=Vico_high"

    Range("F1").Select
    ActiveCell.FormulaR1C1 = "viscosity_low_1000_over_s"
    Range("F2").Select
    ActiveCell.Formula2R1C1 = "=Visco_low"

    Range("G1").Select
    ActiveCell.FormulaR1C1 = "grindometer_mu_m"
    Range("G2").Select
    ActiveCell.Formula2R1C1 = "=Grindo"

    Range("H1").Select
    ActiveCell.FormulaR1C1 = "solid_contents_percentage"
    Range("H2").Select
    ActiveCell.Formula2R1C1 = "=Solid_content"

    Range("I1").Select
    ActiveCell.FormulaR1C1 = "temperature_celsius"
    Range("I2").Select
    ActiveCell.Formula2R1C1 = "=Temperature"

    Range("J1").Select
    ActiveCell.FormulaR1C1 = "notes"
    Range("J2").Select
    ActiveCell.Formula2R1C1 = "=Measurement_notes"

    Range("A1:J6").Select
    Application.CutCopyMode = False

    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False

    ActiveSheet.ListObjects.Add(xlSrcRange, Range("$A$1:$J$6"), , xlYes).Name = _
        "Table25"
    Range("Table25[#All]").Select
    ActiveSheet.ListObjects("Table25").Name = "measurement_step"
    
    ' deleting exrtra rows
    
    On Error Resume Next
    Columns("E").SpecialCells(xlBlanks).EntireRow.Delete
    
End Sub

Sub D_processing_steps()
'
' test_processing_step Macro
'

'
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Name = "4.proces.steps"
    Sheets("4.proces.steps").Select
    
    ActiveCell.FormulaR1C1 = "processing_step_id"
    Range("A2").Select
    ActiveCell.Formula2R1C1 = "=processing_step_id"
    
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "experiment_name"
    Range("B2").Select
    ActiveCell.FormulaR1C1 = "=experiment_name"
    Range("B2").Select
    Selection.AutoFill Destination:=Range("B2:B5")
    
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "project_name"
    
    Range("C2").Select
    ActiveCell.FormulaR1C1 = "=project_name"
    Range("C2").Select
    Selection.AutoFill Destination:=Range("C2:C5")
    Range("C2:C5").Select
    ActiveCell.FormulaR1C1 = "=project_name"
    
    Range("D1").Select
    ActiveCell.FormulaR1C1 = "measurement_step_id"
    
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "processing_step_number"
    Range("E2").Select
    ActiveCell.FormulaR1C1 = "1"
    Range("E3").Select
    ActiveCell.FormulaR1C1 = "2"
    Range("E2:E3").Select
    Selection.AutoFill Destination:=Range("E2:E5"), Type:=xlFillDefault
    
    Range("D2").Select
    'ActiveCell.FormulaR1C1 = _
    '    "=IF([@[processing_step_number]] =measurement_steps[@[measurement_step_number]],measurement_steps[@[measurement_step_id]])"
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[1]=measurement_step[@[measurement_step_number]],measurement_step[@[measurement_step_id]])"
    Range("D3").Select
    Range("D2").Select
    Selection.AutoFill Destination:=Range("D2:D5"), Type:=xlFillDefault

    Range("F1").Select
    ActiveCell.FormulaR1C1 = "description"
    Range("F2").Select
    ActiveCell.Formula2R1C1 = "=Description"
    
    'Sheets("Schlickerherstellung").Select
    'Range("B38:B42").Select
    'Application.CutCopyMode = False
    'ActiveSheet.ListObjects.Add(xlSrcRange, Range("$B$38:$B$42"), , xlYes).Name = _
    '    "Description"
    'Range("Table14[[#All],[Beschreibung]]").Select
    'ActiveSheet.ListObjects("Table14").Name = "Description"

    Sheets("4.proces.steps").Select
    Range("F2").Select
    ActiveCell.Formula2R1C1 = "=Description"
    
    Range("G1").Select
    ActiveCell.FormulaR1C1 = "mixing_speed_1_rpm"
    
    Range("G2").Select
    ActiveCell.Formula2R1C1 = "=Speed1"
    
    Range("H1").Select
    ActiveCell.FormulaR1C1 = "mixing_speed_2_rpm"
    
    Range("H2").Select
    ActiveCell.Formula2R1C1 = "=Speed2"
    
    Range("I1").Select
    ActiveCell.FormulaR1C1 = "mixing_time_minutes"
    
    Range("I2").Select
    ActiveCell.Formula2R1C1 = "=Time"
    
    Range("J1").Select
    ActiveCell.FormulaR1C1 = "sieve_size_mu_m"""
    
    Range("K1").Select
    ActiveCell.FormulaR1C1 = "partial_pressure_mbar"
    
    Range("L1").Select
    ActiveCell.FormulaR1C1 = "notes"
    
    Range("L2").Select
    ActiveCell.Formula2R1C1 = "=Processing_notes"
    
    Range("A1:L5").Select
    Range("L1").Activate
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("$A$1:$L$5"), , xlYes).Name = _
        "processing_step"
    Range("processing_step[#All]").Select
    With Selection
        .HorizontalAlignment = xlCenter
    End With
        
End Sub

Sub E_MaterialAdditionSteps()
'
' test_material_addition_steps Macro
'

'
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Name = "5.mater.add.steps"
    
    ActiveCell.FormulaR1C1 = "material_addition_step_id"
    Range("A2").Select
    ActiveCell.Formula2R1C1 = "=material_addition_step_id"
    
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "experiment_name"
    Range("B2").Select
    ActiveCell.FormulaR1C1 = "=experiment_name"
    Range("B2").Select
    Selection.AutoFill Destination:=Range("B2:B7")
    
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "project_name"
    Range("C2").Select
    ActiveCell.FormulaR1C1 = "=project_name"
    Range("C2").Select
    Selection.AutoFill Destination:=Range("C2:C7")
    
    ' wrong, has to be "processing_step_id", fixed at the end of the macro
    Range("D1").Select
    ActiveCell.FormulaR1C1 = "processing_step_number"
    Range("D2").Select
    ' =IF(D11='4.proces.steps'!$E$2;'4.proces.steps'!$A$2;IF(D11='4.proces.steps'!$E$3;'4.proces.steps'!$A$3; IF(D11='4.proces.steps'!$E$4;'4.proces.steps'!$A$4; IF(D11='4.proces.steps'!$E$5;'4.proces.steps'!$A$5))))
    ActiveCell.FormulaR1C1 = _
        "=IF(ISNUMBER(FIND(Schlickerherstellung!R[24]C[-3],Schlickerherstellung!R39C1)),1,IF(ISNUMBER(FIND(Schlickerherstellung!R[24]C[-3],Schlickerherstellung!R40C1)),2,IF(ISNUMBER(FIND(Schlickerherstellung!R[24]C[-3],Schlickerherstellung!R41C1)),3,IF(ISNUMBER(FIND(Schlickerherstellung!R[24]C[-3],Schlickerherstellung!R42C1)),4,IF(ISNUMBER(FIND(Schlickerherstellung!R[24]C[-3" & _
        "],Schlickerherstellung!R43C1)),5)))))" & _
        ""
    Range("D2").Select
    Selection.AutoFill Destination:=Range("D2:D7")
    
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "slurry_material_id"
    
    Range("F1").Select
    ActiveCell.FormulaR1C1 = "material_addition_step_number"
    Range("F2").Select
    ActiveCell.FormulaR1C1 = "1"
    Range("F3").Select
    ActiveCell.FormulaR1C1 = "2"
    Range("F2:F3").Select
    Selection.AutoFill Destination:=Range("F2:F7"), Type:=xlFillDefault
    
    Range("G1").Select
    ActiveCell.FormulaR1C1 = "material_mass"
    Range("G2").Select
    ActiveCell.Formula2R1C1 = "=Material_add_mass_order"
    
    Range("A1:G7").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues
    Range("A1:G6").Select
    Application.CutCopyMode = False
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("$A$1:$G$6"), , xlYes).Name = _
        "Table32"
    Range("Table32[#All]").Select
    ActiveSheet.ListObjects("Table32").Name = "material_addition_step"
    
    ' deleting exrtra rows
    
    On Error Resume Next
    Columns("G").SpecialCells(xlBlanks).EntireRow.Delete

    ' Replacing processing_step_number with processing_step_id
    
    ' steps:
    ' move the  processing_step_number column to D10
    ' adding column processing_step_id in D1
    ' apply the formula in D2 to D6
    ' =IF(D11='4.proces.steps'!$E$2;'4.proces.steps'!$A$2;IF(D11='4.proces.steps'!$E$3;'4.proces.steps'!$A$3; IF(D11='4.proces.steps'!$E$4;'4.proces.steps'!$A$4; IF(D11='4.proces.steps'!$E$5;'4.proces.steps'!$A$5))))
    ' delete the processing_step_number column
    
    Range("material_addition_step[[#All],[processing_step_number]]").Select
    Selection.Cut
    Range("D10").Select
    ActiveSheet.Paste
    Range("D10").Select
    ActiveCell.FormulaR1C1 = "processing_step_number"
    Range("material_addition_step[[#Headers],[Column1]]").Select
    ActiveCell.FormulaR1C1 = "processing_step_id"
    Range("D2").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(R[9]C='4.proces.steps'!R2C5,'4.proces.steps'!R2C1,IF(R[9]C='4.proces.steps'!R3C5,'4.proces.steps'!R3C1, IF(R[9]C='4.proces.steps'!R4C5,'4.proces.steps'!R4C1, IF(R[9]C='4.proces.steps'!R5C5,'4.proces.steps'!R5C1))))"
    Range("material_addition_step[processing_step_id]").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Range("D10:D15").Select
    Selection.Delete Shift:=xlUp
    
    
    
End Sub

Sub F_slurryMaterial()
'
' SlurryMaterial Macro
'
'
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Name = "6.slurry.mater."

    ActiveCell.FormulaR1C1 = "slurry_material_id"
    Range("A2").Select
    ActiveCell.Formula2R1C1 = "=slurry_material_id"
    
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "experiment_name"
    Range("B2").Select
    Selection.FormulaR1C1 = "=experiment_name"
    Selection.AutoFill Destination:=Range("B2:B6")
    Range("B2:B6").Select
    
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "project_name"
    Range("C2").Select
    Selection.FormulaR1C1 = "=project_name"
    Selection.AutoFill Destination:=Range("C2:C6")
    Range("C2:C6").Select
    
    Range("D1").Select
    ActiveCell.FormulaR1C1 = "material_addition_step_id"
    
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "slurry_material_number"
    Range("E2").Select
    ActiveCell.FormulaR1C1 = "1"
    Range("E3").Select
    ActiveCell.FormulaR1C1 = "2"
    Range("E2:E3").Select
    Selection.AutoFill Destination:=Range("E2:E6"), Type:=xlFillDefault
    Range("E2:E6").Select
    
    Range("F1").Select
    ActiveCell.FormulaR1C1 = "material_name"
    Range("F2").Select
    Selection.Formula2R1C1 = "=Slurry_materials_names"
    
    Range("G1").Select
    ActiveCell.FormulaR1C1 = "percentage"
    Range("G2").Select
    ActiveCell.Formula2R1C1 = "=Slurry_materials_percentage"
    
    Range("H1").Select
    ActiveCell.FormulaR1C1 = "density_gram_over_cupic_cm"
    Range("H2").Select
    ActiveCell.Formula2R1C1 = "=Slurry_materials_densiry"
    
    Range("I1").Select
    ActiveCell.FormulaR1C1 = "material_function"
    
    Range("J1").Select
    ActiveCell.FormulaR1C1 = "material_type"
    Range("J2").Select
    ActiveCell.Formula2R1C1 = "=Slurry_material_type"
    
    Range("K1").Select
    ActiveCell.FormulaR1C1 = "concentration"
    Range("K2").Select
    ActiveCell.Formula2R1C1 = "=Slurry_materials_concentration"
    
    Range("L1").Select
    ActiveCell.FormulaR1C1 = "solved_in"
    Range("L2").Select
    ActiveCell.Formula2R1C1 = "=Solved_in"
    
    Range("A1:L6").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    Sheets("5.mater.add.steps").Select
    Range("E2").Select
    Selection.FormulaR1C1 = _
        "=INDEX('6.slurry.mater.'!R2C1:R6C1, MATCH(Schlickerherstellung!R[24]C[-3], '6.slurry.mater.'!R2C6:R6C6,0))"
    'Range("E2:E3").Select
    'Selection.AutoFill Destination:=Range("E2:E6"), Type:=xlFillDefault
    
    Sheets("6.slurry.mater.").Select
    Range("A1:L6").Select
    Range("L1").Activate
    Application.CutCopyMode = False
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("$A$1:$L$6"), , xlYes).Name = _
        "Table33"
    Range("Table33[#All]").Select
    ActiveSheet.ListObjects("Table33").Name = "slurry_material"
    
    Sheets("6.slurry.mater.").Select
    Range("D2").Select
    ActiveCell.FormulaR1C1 = _
        "=IF('6.slurry.mater.'!RC[-3]='5.mater.add.steps'!R2C5,'5.mater.add.steps'!R2C1,IF('6.slurry.mater.'!RC[-3]='5.mater.add.steps'!R3C5,'5.mater.add.steps'!R3C1,IF('6.slurry.mater.'!RC[-3]='5.mater.add.steps'!R4C5,'5.mater.add.steps'!R4C1,IF('6.slurry.mater.'!RC[-3]='5.mater.add.steps'!R5C5,'5.mater.add.steps'!R5C1,IF('6.slurry.mater.'!RC[-3]='5.mater.add.steps'!R6C5,'5" & _
        ".mater.add.steps'!R6C1)))))" & _
        ""
    Range("slurry_material[material_addition_step_id]").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    Sheets("5.mater.add.steps").Select
    Range("E2:E6").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    
End Sub

Sub G_DeleteSheets(wb As Workbook)
    Dim ws As Worksheet
    
    ' Loop over each sheet in the workbook
    For Each ws In wb.Worksheets
        ' Check if the sheet name matches any of the target names
        Select Case ws.Name
            Case "Arbeitsauftrag", "Regression", "Hilfstabelle", "Kalandrieren", "Beschichtung", "Kalibrierung", "QM", "Schlickerherstellung", "Incremental"
                ' Delete the sheet
                Application.DisplayAlerts = False ' Suppress the confirmation message
                ws.Unprotect
                ws.Delete
                Application.DisplayAlerts = True ' Re-enable the confirmation message
        End Select
    Next ws
End Sub

Sub Tables_names()
'
' Tables_names Macro
'

'
    Range("B1").Select
    Sheets("Schlickerherstellung").Select
    ActiveSheet.Unprotect
    
    ActiveWorkbook.Names.Add Name:="experiment_name", RefersToR1C1:= _
        "=Schlickerherstellung!R1C2"
        
    Range("D1").Select
    ActiveWorkbook.Names.Add Name:="project_name", RefersToR1C1:= _
        "=Schlickerherstellung!R1C4"
        
    Range("B3").Select
    ActiveWorkbook.Names.Add Name:="required_mass_g", RefersToR1C1:= _
        "=Schlickerherstellung!R3C2"
        
    Range("B5").Select
    ActiveWorkbook.Names.Add Name:="required_solid_contents", RefersToR1C1:= _
        "=Schlickerherstellung!R5C2"
        
    Range("F37").Select
    ActiveWorkbook.Names.Add Name:="mixing_tool", RefersToR1C1:= _
        "=Schlickerherstellung!R37C6"
        
    Range("B37").Select
    ActiveWorkbook.Names.Add Name:="mixer", RefersToR1C1:= _
        "=Schlickerherstellung!R37C2"
        
    Range("E8:G8").Select
    Selection.Copy
    Range("A12").Select
    ActiveSheet.Paste
    Range("A7:A12").Select
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("$A$7:$A$12"), , xlYes).Name = _
        "Table2"
    Range("Table2[[#All],[Name]]").Select
    ActiveSheet.ListObjects("Table2").Name = "Slurry_materials_names"
    
    Range("B7:B12").Select
    Application.CutCopyMode = False
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("$B$7:$B$12"), , xlYes).Name = _
        "Table3"
    Range("Table3[[#All],[Anteil  '[m%']]]").Select
    ActiveSheet.ListObjects("Table3").Name = "Slurry_materials_percentage"
    
    Range("C7:C12").Select
    Application.CutCopyMode = False
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("$C$7:$C$12"), , xlYes).Name = _
        "Table4"
    Range("Table4[[#All],[Dichte '[g/cm³']]]").Select
    ActiveSheet.ListObjects("Table4").Name = "Slurry_materials_densiry"
    
    Range("D7:D12").Select
    Application.CutCopyMode = False
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("$D$7:$D$12"), , xlYes).Name = _
        "Table5"
    Range("Table5[[#All],[Hersteller/" & Chr(10) & "Lieferant]]").Select
    ActiveSheet.ListObjects("Table5").Name = "Slurry_materials_function"
    
    Range("B16:B21").Select
    Application.CutCopyMode = False
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("$B$16:$B$21"), , xlYes).Name = _
        "Table6"
    Range("Table6[[#All],[Zugabe]]").Select
    ActiveSheet.ListObjects("Table6").Name = "Slurry_material_type"
    
    Range("C16:C21").Select
    Application.CutCopyMode = False
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("$C$16:$C$21"), , xlYes).Name = _
        "Table7"
    Range("Table7[[#All],[Konzen-tration]]").Select
    ActiveSheet.ListObjects("Table7").Name = "Slurry_materials_concentration"
    
    Range("D16:D21").Select
    Application.CutCopyMode = False
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("$D$16:$D$21"), , xlYes).Name = _
        "Table8"
    Range("Table8[[#All],[gelöst in]]").Select
    ActiveSheet.ListObjects("Table8").Name = "Solved_in"
    
    Range("A25:A30").Select
    Application.CutCopyMode = False
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("$A$25:$A$30"), , xlYes).Name = _
        "Table9"
    Range("Table9[[#All],[Nr]]").Select
    ActiveSheet.ListObjects("Table9").Name = "Material_add_step_number"
    
    Range("B25:B30").Select
    Application.CutCopyMode = False
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("$B$25:$B$30"), , xlYes).Name = _
        "Table10"
    Range("Table10[[#All],[Name]]").Select
    ActiveSheet.ListObjects("Table10").Name = "material_add_names_order"
    
    Range("C25:C30").Select
    Application.CutCopyMode = False
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("$C$25:$C$30"), , xlYes).Name = _
        "Table11"
    Range("Table11[[#All],[Masse]]").Select
    ActiveSheet.ListObjects("Table11").Name = "Material_add_mass_order"
    
    ActiveWindow.SmallScroll Down:=6
    Range("A38:A42").Select
    Application.CutCopyMode = False
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("$A$38:$A$42"), , xlYes).Name = _
        "Table12"
    Range("Table12[[#All],[nach Zugabe]]").Select
    ActiveSheet.ListObjects("Table12").Name = "Processing_after_add_step"
    
    Range("B38:C42").Select
    Selection.UnMerge
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("$B$38:$B$42"), , xlYes).Name = _
        "Description"
    
    Range("D38:D42").Select
    Application.CutCopyMode = False
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("$D$38:$D$42"), , xlYes).Name = _
        "Table15"
    Range("Table15[[#All],[Drehzahl 1 '[U/min']]]").Select
    ActiveSheet.ListObjects("Table15").Name = "Speed1"
    
    Range("E38:E42").Select
    Application.CutCopyMode = False
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("$E$38:$E$42"), , xlYes).Name = _
        "Table16"
    Range("Table16[[#All],[Drehzahl 2 '[U/min']]]").Select
    ActiveSheet.ListObjects("Table16").Name = "Speed2"
    
    Range("F38:F42").Select
    Application.CutCopyMode = False
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("$F$38:$F$42"), , xlYes).Name = _
        "Table17"
    Range("Table17[[#All],[Zeit '[min']]]").Select
    ActiveSheet.ListObjects("Table17").Name = "Time"
    
    Range("G38:G42").Select
    Application.CutCopyMode = False
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("$G$38:$G$42"), , xlYes).Name = _
        "Table18"
    Range("Table18[[#All],[Kommentar]]").Select
    ActiveSheet.ListObjects("Table18").Name = "Processing_notes"
    
    Sheets("QM").Select
    ActiveSheet.Unprotect
    Range("A9:A13").Select
    Application.CutCopyMode = False
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("$A$9:$A$13"), , xlYes).Name = _
        "Table19"
    Range("Table19[[#All],[nach Arbeitsschritt]]").Select
    ActiveSheet.ListObjects("Table19").Name = "measurement_after_proces_number"
    
    Range("B9:B13").Select
    Application.CutCopyMode = False
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("$B$9:$B$13"), , xlYes).Name = _
        "Table20"
    Range("Table20[[#All],[Viskosität '[1/s']]]").Select
    ActiveSheet.ListObjects("Table20").Name = "Vico_high"
    
    Range("C9:C13").Select
    Application.CutCopyMode = False
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("$C$9:$C$13"), , xlYes).Name = _
        "Table21"
    Range("Table21[[#All],[Viskosität '[1000/s']]]").Select
    ActiveSheet.ListObjects("Table21").Name = "Visco_low"
    
    Range("D9:D13").Select
    Application.CutCopyMode = False
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("$D$9:$D$13"), , xlYes).Name = _
        "Table22"
    Range("Table22[[#All],[Grindometer '[µm']]]").Select
    ActiveSheet.ListObjects("Table22").Name = "Grindo"
    
    Range("E9:E13").Select
    Application.CutCopyMode = False
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("$E$9:$E$13"), , xlYes).Name = _
        "Table23"
    Range("Table23[[#All],[FG '[m%']]]").Select
    ActiveSheet.ListObjects("Table23").Name = "Solid_content"
    
    Range("F9:F13").Select
    Application.CutCopyMode = False
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("$F$9:$F$13"), , xlYes).Name = _
        "Table24"
    Range("Table24[[#All],[Schlickertemperatur '[°C']]]").Select
    ActiveSheet.ListObjects("Table24").Name = "Temperature"
    
    Range("G9:G13").Select
    Application.CutCopyMode = False
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("$G$9:$G$13"), , xlYes).Name = _
        "Table25"
    Range("Table25[[#All],[Sonstiges]]").Select
    ActiveSheet.ListObjects("Table25").Name = "Measurement_notes"
    
End Sub



' test succeeded
Sub IncrementalSheet(Path)
    ' Dim Path As String
    Dim FileNameNew As String
    Dim wbNew As Workbook
    Dim wsNew As Worksheet
    
    ' set the path to the folder containing the Excel files
    ' Path = "C:\Users\mou95504\Desktop\Test\"
    Static measurement_counter As Integer
    measurement_counter = 0
    processing_counter = 0
    material_addition_step_counter = 0
    slurry_material_counter = 0
    

    ' loop through all files in the folder
    FileNameNew = Dir(Path & "*.xlsx")
    ' Application.ScreenUpdating = False
    Do While FileNameNew <> ""
        measurement_counter = measurement_counter + 1
        processing_counter = processing_counter + 1
        material_addition_step_counter = material_addition_step_counter + 1
        slurry_material_counter = slurry_material_counter + 1
        
        Set wbNew = Workbooks.Open(Path & FileNameNew)

        Sheets.Add After:=ActiveSheet
        ActiveSheet.Name = "Incremental"
        
        ' measurement_counter
        Range("A1").Select
        ActiveCell.FormulaR1C1 = "measurement_step_id"
        Range("A2").Select
        Selection.FormulaR1C1 = measurement_counter
        Range("A3").Select
        Selection.FormulaR1C1 = measurement_counter + 1
        Range("A2:A3").Select
        Selection.AutoFill Destination:=Range("A2:A6"), Type:=xlFillDefault
        Range("A2:A6").Select
        measurement_counter = Range("A6")

        ' processing_counter
        Range("B1").Select
        ActiveCell.FormulaR1C1 = "processing_step_id"
        Range("B2").Select
        Selection.FormulaR1C1 = processing_counter
        Range("B3").Select
        Selection.FormulaR1C1 = processing_counter + 1
        Range("B2:B3").Select
        Selection.AutoFill Destination:=Range("B2:B5"), Type:=xlFillDefault
        Range("B2:B5").Select
        processing_counter = Range("B5")
        
        ' material_addition_step_counter
        Range("C1").Select
        ActiveCell.FormulaR1C1 = "material_addition_step_id"
        Range("C2").Select
        Selection.FormulaR1C1 = material_addition_step_counter
        Range("C3").Select
        Selection.FormulaR1C1 = material_addition_step_counter + 1
        Range("C2:C3").Select
        Selection.AutoFill Destination:=Range("C2:C7"), Type:=xlFillDefault
        Range("C2:C7").Select
        material_addition_step_counter = Range("C7")
        
        ' slurry_material_counter
        Range("D1").Select
        ActiveCell.FormulaR1C1 = "slurry_material_id"
        Range("D2").Select
        Selection.FormulaR1C1 = slurry_material_counter
        Range("D3").Select
        Selection.FormulaR1C1 = slurry_material_counter + 1
        Range("D2:D3").Select
        Selection.AutoFill Destination:=Range("D2:D6"), Type:=xlFillDefault
        Range("D2:D4").Select
        slurry_material_counter = Range("D6")

        ' adding tables
        Range("A1:A6").Select
        ActiveSheet.ListObjects.Add(xlSrcRange, Range("$A$1:$A$6"), , xlYes).Name = _
        "measurement_steps_id"
        
        Range("B1:B5").Select
        ActiveSheet.ListObjects.Add(xlSrcRange, Range("$B$1:$B$5"), , xlYes).Name = _
        "processing_step_id"
        
        Range("C1:C7").Select
        ActiveSheet.ListObjects.Add(xlSrcRange, Range("$C$1:$C$7"), , xlYes).Name = _
        "material_addition_step_id"
        
        Range("D1:D6").Select
        ActiveSheet.ListObjects.Add(xlSrcRange, Range("$D$1:$D$6"), , xlYes).Name = _
        "slurry_material_id"

        wbNew.Close SaveChanges:=True
        
        ' move to the next file in the folder
        FileNameNew = Dir()
    Loop
    ' Application.ScreenUpdating = True
End Sub



Sub RunMacroInAllFiles()
    Application.ScreenUpdating = False

    Dim Path As String
    Dim FileName As String
    Dim wb As Workbook
    Dim ws As Worksheet
    
    ' set the path to the folder containing the Excel files
    Path = "C:\Users\mou95504\Desktop\Test\"
    
    ' adding the incremental sheet first
    Call IncrementalSheet(Path)
    
    ' loop through all files in the folder
    FileName = Dir(Path & "*.xlsx")

    ' then looping throght all files
    Do While FileName <> ""
        Set wb = Workbooks.Open(Path & FileName)
        
        ' run the "RunAllMacros" macro in the current workbook"
        H_RunAllMacros wb
        
        
        ' save and close the workbook
        wb.Close SaveChanges:=True
        
        ' move to the next file in the folder
        FileName = Dir()
    Loop
    Application.ScreenUpdating = True
End Sub















