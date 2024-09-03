Attribute VB_Name = "ConvertISMDownload"
Public Sub Create_ISM_Tables()
Attribute Create_ISM_Tables.VB_ProcData.VB_Invoke_Func = "E\n14"
    Application.ScreenUpdating = False

    Convert_Sheets

    Application.ScreenUpdating = True

End Sub

Private Sub Convert_Sheets()
    Dim Sheetname As String
    Dim Concept As String
    Dim Extras As String
    Dim HideWorksheet As Boolean
  
    Dim WS As Worksheet
    
    For Each WS In ThisWorkbook.Worksheets
    
        If WS.Visible = xlSheetHidden Then
            WS.Visible = xlSheetVisible
            HideWorksheet = True
          End If
    
        Sheetname = WS.Name
        WS.Select
       
''Calls the 'Convert_Sheet' Sub if the SheetName matches one of the ISM Sheet Names listed
''If the sheet is Class/Attribute type returns the value of True
        If Sheetname = "ISM Attributes" Then
            Convert_Sheet_to_Table Sheetname
        ElseIf Sheetname = "ISM Functional Classes" Then
            Convert_Sheet_to_Table Sheetname
        ElseIf Sheetname = "ISM Functional Class Attributes" Then
            Convert_Sheet_to_Table Sheetname
        ElseIf Sheetname = "ISM Functional Class Naming Tpl" Then
            Convert_Sheet_to_Table Sheetname
        ElseIf Sheetname = "ISM Physical Classes" Then
            Convert_Sheet_to_Table Sheetname
        ElseIf Sheetname = "ISM Physical Class Attributes" Then
            Convert_Sheet_to_Table Sheetname
        ElseIf Sheetname = "ISM Physical Class Naming Tpl" Then
            Convert_Sheet_to_Table Sheetname
        ElseIf Sheetname = "ISM Document Classes" Then
            Convert_Sheet_to_Table Sheetname
        ElseIf Sheetname = "ISM Document Class Attributes" Then
            Convert_Sheet_to_Table Sheetname
        ElseIf Sheetname = "ISM Document Class Naming Tpl" Then
            Convert_Sheet_to_Table Sheetname
        ElseIf Sheetname = "ISM General Classes" Then
            Convert_Sheet_to_Table Sheetname
        ElseIf Sheetname = "ISM General Class Attributes" Then
            Convert_Sheet_to_Table Sheetname
        ElseIf Sheetname = "ISM General Class Naming Tpl" Then
            Convert_Sheet_to_Table Sheetname
        ElseIf Sheetname = "ISM UoM Units" Then
            Convert_Sheet_to_Table Sheetname
        ElseIf Sheetname = "ISM UoM Classes" Then
            Convert_Sheet_to_Table Sheetname
        ElseIf Sheetname = "ISM UoM Class Units" Then
            Convert_Sheet_to_Table Sheetname
        ElseIf Sheetname = "ISM Enumerations" Then
            Convert_Sheet_to_Table Sheetname
        ElseIf Sheetname = "ISM N&N Elements" Then
            Convert_Sheet_to_Table Sheetname
        ElseIf Sheetname = "ISM N&N Templates" Then
            Convert_Sheet_to_Table Sheetname
        ElseIf Sheetname = "ISM N&N Template Elements" Then
            Convert_Sheet_to_Table Sheetname
        ElseIf Sheetname = "ISM Maturity Levels" Then
            Convert_Sheet_to_Table Sheetname
        ElseIf Sheetname = "ISM Life Cycle Types" Then
            Convert_Sheet_to_Table Sheetname
        End If
        
        If HideWorksheet = True Then
            WS.Visible = xlSheetHidden
            HideWorksheet = False
        End If
        
    Next

''Calls the Private Sub to add validation to the ISM sheets.
''This cannot be added until all sheets are processed and all listobject items are created
    Extra_Validation
    
    Worksheets("ISM Class Library Header").Select
    Range("B2").Select
    
End Sub


Private Sub Convert_Sheet_to_Table(Sheetname As String)

    Dim LastRow As Long
    Dim LastDataRow As Long
    Dim RN As Long ''Row Number
    
    Dim LastColumn As Integer
    Dim amp As Integer '' used to identify if a sheetname has an Ampresand
    
    Dim TableName As String
    Dim SortID As String
    
    Dim TRange As Range
    Dim CR As Range ''Cell Range
    
    Dim Tbl As ListObject
    
    ThisWorkbook.ActiveSheet.Cells.ClearFormats
    If ActiveSheet.AutoFilterMode = True Then ActiveSheet.AutoFilterMode = False
      
    
''Sets value of the Last row to define row limit for this funtion only
    LastRow = Cells(Rows.Count, "B").End(xlUp).Row
''    Debug.Print "Last Row " & LastRow
    
''Sets value of the Last Column to define Column limit for this funtion only
    LastColumn = Cells(1, Columns.Count).End(xlToLeft).Column
    
''Sets the range for the to be created table based on A1 to Lastrow/LastColumn
    Set TRange = Range(Cells(1, 1), Cells(LastRow, LastColumn))
    
''If There is no data in the sheet, this will exit the Sub, skipping the formatting, This prevents errors when loading into ISM later
    If LastRow = 1 Then Exit Sub
    
''Creates the table name based on the Sheet name without spaces
    TableName = Replace(Sheetname, " ", "_")
    amp = InStr(TableName, "&")
    If amp >= 1 Then
        TableName = Replace(TableName, "&", "")
    End If

''Creates a String value for Sort ID to be used when sorting the newly created table
    SortID = TableName & "[[#All],[ID]]"
    
''Checks to see if the Rows are Grouped and removes Group if true
    For RN = 2 To LastRow
        If ActiveSheet.Rows(RN).OutlineLevel > 1 Then
            Selection.Rows.ClearOutline
        End If
    Next RN
    
''Checks to see if an autofilter is applied on the sheet and removes it if true
    If ActiveSheet.FilterMode Then
        ActiveSheet.AutoFilter
    End If

''Removes blank rows, or rows where only Column A has a value
    LastDataRow = LastRow
      
    For RN = 2 To LastRow
        Set CR = ActiveSheet.Cells(RN, 2)
        If IsEmpty(CR) = True Then
            Rows(RN).Delete
            If RN < LastRow - 1 Then
                RN = RN - 1
                LastRow = LastRow - 1
                LastDataRow = LastDataRow - 1
            End If
        End If
        
        If RN = LastDataRow Or RN > LastDataRow Then
            Exit For
        End If
        
''        Debug.Print "RN: " & RN
''        Debug.Print "Last Data Row: " & LastDataRow
    Next
    
''Creates the table for the data in the sheet. The range is set by TRange and name is set by TName
''After creation it will sort the data by ID.
    TRange.Select

    Application.CutCopyMode = False

On Error GoTo TableExists
    
    ActiveSheet.ListObjects.Add(xlSrcRange, TRange, , xlYes) _
        .Name = TableName
    
    Set Tbl = ActiveSheet.ListObjects(TableName)
    
    Tbl.TableStyle = "TableStyleMedium9"
    Tbl.Sort.SortFields.Add Key:=Range(SortID), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal

    With Tbl.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    Tbl.ShowAutoFilter = True

''Selects all cells in Activesheet and sets the style to "Normal". This removes all previous formatting save for
''the table style
    Cells.Select
    Selection.Style = "Normal"
    
    Range("A2").Select

''Format Cells to: 1) Wrap data in Cell, 20 Center Vertically and 3) Autofit
    Cells.Select
        Cells.Select
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Cells.EntireColumn.AutoFit
    
TableExists:
    Exit Sub

End Sub

Private Sub Extra_Validation()
    Dim Sheetname As String
    Dim Concept As String
    Dim Extras As String
    
    Dim WS As Worksheet
     
    For Each WS In ThisWorkbook.Worksheets
    
        If WS.Visible = xlSheetHidden Then
            WS.Visible = xlSheetVisible
            HideWorksheet = True
          End If
    
        Sheetname = WS.Name
        WS.Select
        
''Calls the 'Convert_Sheet' Sub if the SheetName matches one of the ISM Sheet Names listed
        If Sheetname = "ISM Attributes" Then
            Concept = "All"
            Extras = "Attributes"
            Add_Extra_Validation Concept, Sheetname, Extras
        ElseIf Sheetname = "ISM Functional Classes" Then
            Concept = "Functional"
            Extras = "Class"
            Add_Extra_Validation Concept, Sheetname, Extras
        ElseIf Sheetname = "ISM Functional Class Attributes" Then
            Concept = "Functional"
            Extras = "ClassAtt"
            Add_Extra_Validation Concept, Sheetname, Extras
        ElseIf Sheetname = "ISM Functional Class Naming Tpl" Then
            Concept = "Functional"
            Extras = "ClassName"
            Add_Extra_Validation Concept, Sheetname, Extras
        ElseIf Sheetname = "ISM Physical Classes" Then
            Concept = "Physical"
            Extras = "Class"
            Add_Extra_Validation Concept, Sheetname, Extras
        ElseIf Sheetname = "ISM Physical Class Attributes" Then
            Concept = "Physical"
            Extras = "ClassAtt"
            Add_Extra_Validation Concept, Sheetname, Extras
        ElseIf Sheetname = "ISM Physical Class Naming Tpl" Then
            Concept = "Physical"
            Extras = "ClassName"
            Add_Extra_Validation Concept, Sheetname, Extras
        ElseIf Sheetname = "ISM Document Classes" Then
            Concept = "Document"
            Extras = "Class"
            Add_Extra_Validation Concept, Sheetname, Extras
        ElseIf Sheetname = "ISM Document Class Attributes" Then
            Concept = "Document"
            Extras = "ClassAtt"
            Add_Extra_Validation Concept, Sheetname, Extras
        ElseIf Sheetname = "ISM Document Class Naming Tpl" Then
            Concept = "Document"
            Extras = "ClassName"
            Add_Extra_Validation Concept, Sheetname, Extras
        ElseIf Sheetname = "ISM General Classes" Then
            Concept = "General"
            Extras = "Class"
            Add_Extra_Validation Concept, Sheetname, Extras
        ElseIf Sheetname = "ISM General Class Attributes" Then
            Concept = "General"
            Extras = "ClassAtt"
            Add_Extra_Validation Concept, Sheetname, Extras
        ElseIf Sheetname = "ISM General Class Naming Tpl" Then
            Concept = "General"
            Extras = "ClassName"
            Add_Extra_Validation Concept, Sheetname, Extras
        ElseIf Sheetname = "ISM UoM Units" Then
            Concept = "All"
            Extras = ""
            Add_Extra_Validation Concept, Sheetname, Extras
        ElseIf Sheetname = "ISM UoM Classes" Then
            Concept = "All"
            Extras = ""
            Add_Extra_Validation Concept, Sheetname, Extras
        ElseIf Sheetname = "ISM UoM Class Units" Then
            Concept = "All"
            Extras = ""
            Add_Extra_Validation Concept, Sheetname, Extras
        ElseIf Sheetname = "ISM Enumerations" Then
            Concept = "All"
            Extras = ""
            Add_Extra_Validation Concept, Sheetname, Extras
        ElseIf Sheetname = "ISM N&N Elements" Then
            Concept = "All"
            Extras = "NNElement"
            Add_Extra_Validation Concept, Sheetname, Extras
        ElseIf Sheetname = "ISM N&N Templates" Then
            Concept = "All"
            Extras = ""
            Add_Extra_Validation Concept, Sheetname, Extras
        ElseIf Sheetname = "ISM N&N Template Elements" Then
            Concept = "All"
            Extras = ""
            Add_Extra_Validation Concept, Sheetname, Extras
        ElseIf Sheetname = "ISM Maturity Levels" Then
            Concept = "All"
            Extras = ""
            Add_Extra_Validation Concept, Sheetname, Extras
        End If
        
        If HideWorksheet = True Then
            WS.Visible = xlSheetHidden
            HideWorksheet = False
        End If
    Next

''Calls the Private Sub to add validation to the ISM sheet.
''This cannot be added until the other objects (Enum, UoM Etc.) are created - which happens after the creation of the Attribute table

    
    Worksheets("ISM Class Library Header").Select
    Range("B2").Select
    
End Sub


Private Sub Add_Extra_Validation(Concept As String, Sheetname As String, Extras As String)
    
    Dim Tbl As ListObject
    Dim TableName As String
    
    Dim IDObjectType As String
    
''If the Formattinghas been skipped (see Convert_Sheet_to_Table) then there will be no table. this will raise an error, meaning that this sub will be skipped.
On Error GoTo NoTable
            
    TableName = ActiveSheet.ListObjects(1).Name
    Set Tbl = ActiveSheet.ListObjects(TableName)
    
    ''Depending on the value of "Extras" Calls the appropriate subs to add sheet validations
    If Extras = "Attributes" Then
        Add_Enum_Validation_to_Attribute Tbl, TableName
        Add_UOM_Validation_to_Attribute Tbl, TableName
    
    ElseIf Extras = "ClassAtt" Then
            IDObjectType = "Attribute"
        Add_Duplicate_Check Tbl, TableName
        Add_ClassID_Check Concept, Tbl, TableName
        Add_ID_Check IDObjectType, Tbl, TableName
        Add_Attribute_Description_Check Tbl, TableName
        ''Add_Attribute_Primary_Concept_Check Tbl, TableName
    
    ElseIf Extras = "Class" Then
        Add_Extends_Validation_to_Class Concept, Tbl, TableName
    
    ElseIf Extras = "ClassName" Then
            IDObjectType = "Name"
        Add_ID_Check IDObjectType, Tbl, TableName
        Add_ClassID_Check Concept, Tbl, TableName
    
    ElseIf Extras = "NNElement" Then
        Add_Attribute_Validation_of_Source Tbl, TableName
    
    End If

NoTable:
    Exit Sub

End Sub

Private Sub Add_Duplicate_Check(Tbl As ListObject, TableName As String)

    Dim SortID As String
    Dim LastColumn As Integer
      
    SortID = TableName & "[[#All],[Duplicate Check]]"
    
    LastColumn = Cells(1, Columns.Count).End(xlToLeft).Column
    
''Checks if Column has already been created - if so exits the Sub and continues
    If Tbl.ListColumns(LastColumn).Name = "Duplicate Check" Then GoTo ColumnExists
    LastColumn = LastColumn + 1


''Creates the "Duplicate Check" Column at the end of the Table
    Cells(1, LastColumn).Select
    ActiveCell.FormulaR1C1 = "Duplicate Check"
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 10498160
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
''Adds the Formula '=[Class_Id].[Id]' to the first cell of the 'Duplicate Check' Column
''utilising the behaviour of the Tables - MS Excel will automatically populate the remaining rows
    Cells(2, LastColumn).Select
    ActiveCell.FormulaR1C1 = "=[Class_Id] & ""."" & [Id]"
    
''Resorts the data based on 'Duplicate Check' Column
    Tbl.Sort.SortFields.Clear
    Tbl.Sort.SortFields.Add Key:=Range(SortID), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With Tbl.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
''Adds duplicate conditional formatting
    Columns(LastColumn).Select
    Selection.FormatConditions.AddUniqueValues
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    Selection.FormatConditions(1).DupeUnique = xlDuplicate
    With Selection.FormatConditions(1).Font
        .Color = -16383844
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13551615
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    
    Columns(LastColumn).AutoFit
    
ColumnExists:
    Cells(1, LastColumn).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 10498160
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
    Columns(LastColumn).AutoFit
    
    Exit Sub

End Sub

Private Sub Add_ClassID_Check(Concept As String, Tbl As ListObject, TableName As String)
    
    Dim NewColNum As Integer
    Dim NewColCheck As Integer
        
    Dim DocClassIDCheck As String, FuncClassIDCheck As String, GenClassIDCheck As String, PhysClassIDCheck As String
    
    DocClassIDCheck = "=IF([Class_Id] = """",""ERROR"",IF(VLOOKUP([Class_Id],ISM_Document_Classes[Id],1,FALSE)=[Class_Id],VLOOKUP([Class_Id],ISM_Document_Classes[[Id]:[Name]],2,FALSE),""ERROR""))"
    FuncClassIDCheck = "=IF([Class_Id] = """",""ERROR"",IF(VLOOKUP([Class_Id],ISM_Functional_Classes[Id],1,FALSE)=[Class_Id],VLOOKUP([Class_Id],ISM_Functional_Classes[[Id]:[Name]],2,FALSE),""ERROR""))"
    GenClassIDCheck = "=IF([Class_Id] = """",""ERROR"",IF(VLOOKUP([Class_Id],ISM_General_Classes[Id],1,FALSE)=[Class_Id],VLOOKUP([Class_Id],ISM_General_Classes[[Id]:[Name]],2,FALSE),""ERROR""))"
    PhysClassIDCheck = "=IF([Class_Id] = """",""ERROR"",IF(VLOOKUP([Class_Id],ISM_Physical_Classes[Id],1,FALSE)=[Class_Id],VLOOKUP([Class_Id],ISM_Physical_Classes[[Id]:[Name]],2,FALSE),""ERROR""))"
            
''Adds the column 'Class Name' to the left of "Class_ID"
    NewColNum = Tbl.ListColumns("Class_ID").Index
    NewColCheck = NewColNum - 1
    
''Checks if Column has already been created - if so exits the Sub and continues
    If Tbl.ListColumns(NewColCheck).Name = "Class Name" Then GoTo ColumnExists
            
    Tbl.ListColumns.Add(NewColNum).Name = "Class Name"
    
    Cells(1, NewColNum).Select
        With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = 10498160
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
  
''Adds the Formula, based on the defined concept to the first cell of the 'Class Name' Column
''utilising the behaviour of the Tables - MS Excel will automatically populate the remaining rows
    Cells(2, NewColNum).Select
    
    If Concept = "Document" Then
        ActiveCell.FormulaR1C1 = DocClassIDCheck
    ElseIf Concept = "Functional" Then
        ActiveCell.FormulaR1C1 = FuncClassIDCheck
    ElseIf Concept = "General" Then
        ActiveCell.FormulaR1C1 = GenClassIDCheck
    ElseIf Concept = "Physical" Then
        ActiveCell.FormulaR1C1 = PhysClassIDCheck
    End If
    
    Columns(NewColNum).EntireColumn.AutoFit
    
ColumnExists:
    Cells(1, NewColNum).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 10498160
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
     Columns(NewColNum).EntireColumn.AutoFit
    
    Exit Sub
    
End Sub

Private Sub Add_ID_Check(IDObjectType As String, Tbl As ListObject, TableName As String)

    Dim IDFormula As String
    Dim ColName As String
    
    Dim NewColNum As Integer
    
    If IDObjectType = "Attribute" Then
        IDFormula = "=IF([Class_Id] = """",""ERROR"",IF(VLOOKUP([Id],ISM_Attributes[Id],1,FALSE)=[Id],VLOOKUP([Id],ISM_Attributes[[Id]:[Name]],2,FALSE),""ERROR""))"
        ColName = "Attribute Name"
    ElseIf IDObjectType = "Name" Then
        IDFormula = "=IF([Class_Id] = """",""ERROR"",IF(VLOOKUP([Id],ISM_NN_Templates[Id],1,FALSE)=[Id],VLOOKUP([Id],ISM_NN_Templates[[Id]:[Name]],2,FALSE),""ERROR""))"
        ColName = "Naming Template"
    Else
        Exit Sub
    End If

    NewColNum = Tbl.ListColumns("ID").Index
    NewColNum = NewColNum + 1
    
''Checks if Column has already been created - if so exits the Sub and continues
    If Tbl.ListColumns(NewColNum).Name = ColName Then GoTo ColumnExists
    
    Tbl.ListColumns.Add(NewColNum).Name = ColName
    
    Cells(1, NewColNum).Select
        With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = 10498160
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With

     Cells(2, NewColNum).Select
     ActiveCell.FormulaR1C1 = IDFormula

    Columns(NewColNum).EntireColumn.AutoFit

ColumnExists:
    Cells(1, NewColNum).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 10498160
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
    Columns(NewColNum).EntireColumn.AutoFit
    
    Exit Sub

End Sub

Private Sub Add_Attribute_Description_Check(Tbl As ListObject, TableName As String)

    Dim NewColNum As Integer

    NewColNum = Tbl.ListColumns("Attribute Name").Index
    NewColNum = NewColNum + 1
    
''Checks if Column has already been created - if so exits the Sub and continues
    If Tbl.ListColumns(NewColNum).Name = "Attribute Description" Then GoTo ColumnExists
    
    Tbl.ListColumns.Add(NewColNum).Name = "Attribute Description"
    
    Cells(1, NewColNum).Select
        With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = 10498160
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With

     Cells(2, NewColNum).Select
     ActiveCell.FormulaR1C1 = _
            "=IF([Class_Id] = """",""ERROR"",IF(VLOOKUP([Id],ISM_Attributes[Id],1,FALSE)=[Id],IF(VLOOKUP([Id],ISM_Attributes[[Id]:[Description]],3,FALSE)="""","""",VLOOKUP([Id],ISM_Attributes[[Id]:[Description]],3,FALSE)),""ERROR""))"

    Columns(NewColNum).EntireColumn.AutoFit
    
ColumnExists:
    Cells(1, NewColNum).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 10498160
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
    Columns(NewColNum).EntireColumn.AutoFit
    
    Exit Sub

End Sub

'' This Sub is only to be used on Goliath when identifiying "Primary Concept" this is an extension specific to the Goliath Class Library, or variants created by GTC

Private Sub Add_Attribute_Primary_Concept_Check(Tbl As ListObject, TableName As String)

    Dim NewColNum As Integer

    NewColNum = Tbl.ListColumns("Attribute Description").Index
    NewColNum = NewColNum + 1
    
''Checks if Column has already been created - if so exits the Sub and continues
    If Tbl.ListColumns(NewColNum).Name = "Primary Concept" Then GoTo ColumnExists
        
    Tbl.ListColumns.Add(NewColNum).Name = "Primary Concept"
    
    Cells(1, NewColNum).Select
        With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = 10498160
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With

     Cells(2, NewColNum).Select
     ActiveCell.FormulaR1C1 = _
            "=IF([Class_Id] = """",""ERROR"",IF(VLOOKUP([Id],ISM_Attributes[Id],1,FALSE)=[Id],VLOOKUP([Id],ISM_Attributes[[Id]:[GCCL_CatDet_Schema:GCCL_CatDef_Attr_PrimaryConcept]],21,FALSE),""ERROR""))"

    Columns(NewColNum).EntireColumn.AutoFit
    
ColumnExists:
    Cells(1, NewColNum).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 10498160
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
    Columns(NewColNum).EntireColumn.AutoFit
    
    Exit Sub

End Sub

Private Sub Add_Enum_Validation_to_Attribute(Tbl As ListObject, TableName As String)

On Error GoTo Ignore

    Dim NewColNum As Integer

    NewColNum = Tbl.ListColumns("ValidationType").Index
    NewColNum = NewColNum + 1
    
''Checks if Column has already been created - if so exits the Sub and continues
    If Tbl.ListColumns(NewColNum).Name = "Enumeration Validation" Then GoTo ColumnExists
    
    Tbl.ListColumns.Add(NewColNum).Name = "Enumeration Validation"
    
    Cells(1, NewColNum).Select
        With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = 10498160
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With

     Cells(2, NewColNum).Select
     ActiveCell.FormulaR1C1 = _
            "=IFERROR(IF([ValidationType]=""Enumeration"",IF(VLOOKUP([ValidationRule],ISM_Enumerations[Id],1,FALSE)=[ValidationRule],""Enum Valid"",""ERROR""),""""),""ERROR"")"
    
    Columns(NewColNum).EntireColumn.AutoFit
    
ColumnExists:
    Cells(1, NewColNum).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 10498160
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
    Columns(NewColNum).EntireColumn.AutoFit
    
    Exit Sub

Ignore:
'' This evaluates to see if the column has been created and if so deletes it as the reference does not exist

    Dim EnumVCol As Integer
    
    EnumVCol = Tbl.ListColumns("Enumeration Validation").Index
    
    If UoMVCol > 0 Then
        Tbl.ListColumns(UoMVCol).Delete
    End If

End Sub

Private Sub Add_UOM_Validation_to_Attribute(Tbl As ListObject, TableName As String)

On Error GoTo Ignore
    
    Dim NewColNum As Integer

    NewColNum = Tbl.ListColumns("UomRequired").Index
    NewColNum = NewColNum + 1
    
''Checks if Column has already been created - if so exits the Sub and continues
    If Tbl.ListColumns(NewColNum).Name = "UoM Validation" Then GoTo ColumnExists
    
    Tbl.ListColumns.Add(NewColNum).Name = "UoM Validation"
    
    Cells(1, NewColNum).Select
        With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 10498160
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With

     Cells(2, NewColNum).Select
     ActiveCell.FormulaR1C1 = _
            "=IFERROR(IF([UomRequired]=""True"",IF([UomClassId]="""",""UOM Class Needed!"",IF(VLOOKUP([UomClassId],ISM_UoM_Classes[Id],1,FALSE)=[UomClassId],""UoM Valid"",""ERROR"")),""""),""ERROR"")"

    Columns(NewColNum).EntireColumn.AutoFit
    
ColumnExists:
    Cells(1, NewColNum).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 10498160
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
    Columns(NewColNum).EntireColumn.AutoFit
    
    Exit Sub

Ignore:
'' This evaluates to see if the column has been created and if so deletes it as the reference does not exist

    Dim UoMVCol As Integer
    
    UoMVCol = Tbl.ListColumns("UoM Validation").Index
    
    If UoMVCol > 0 Then
        Tbl.ListColumns(UoMVCol).Delete
    End If
    
End Sub

Private Sub Add_Extends_Validation_to_Class(Concept As String, Tbl As ListObject, TableName As String)

    Dim NewColNum As Integer
    Dim Func_Extends As String, Phys_Extends As String, Doc_Extends As String, Gen_Extends As String
    
    Func_Extends = "=IF([Extends]="""",""Is this the Root Class?"",IF(VLOOKUP([Extends],ISM_Functional_Classes[[Id]:[Name]],1,FALSE)=[Extends],VLOOKUP([Extends],ISM_Functional_Classes[[Id]:[Name]],2,FALSE),""ERROR""))"
    Phys_Extends = "=IF([Extends]="""",""Is this the Root Class?"",IF(VLOOKUP([Extends],ISM_Physical_Classes[[Id]:[Name]],1,FALSE)=[Extends],VLOOKUP([Extends],ISM_Physical_Classes[[Id]:[Name]],2,FALSE),""ERROR""))"
    Doc_Extends = "=IF([Extends]="""",""Is this the Root Class?"",IF(VLOOKUP([Extends],ISM_Document_Classes[[Id]:[Name]],1,FALSE)=[Extends],VLOOKUP([Extends],ISM_Document_Classes[[Id]:[Name]],2,FALSE),""ERROR""))"
    Gen_Extends = "=IF([Extends]="""",""Is this the Root Class?"",IF(VLOOKUP([Extends],ISM_General_Classes[[Id]:[Name]],1,FALSE)=[Extends],VLOOKUP([Extends],ISM_General_Classes[[Id]:[Name]],2,FALSE),""ERROR""))"


    NewColNum = Tbl.ListColumns("Extends").Index
    NewColNum = NewColNum + 1
    
''Checks if Column has already been created - if so exits the Sub and continues
    If Tbl.ListColumns(NewColNum).Name = "Extends Name" Then GoTo ColumnExists
    
    Tbl.ListColumns.Add(NewColNum).Name = "Extends Name"
    
    Cells(1, NewColNum).Select
        With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = 10498160
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With

''Adds the Formula, based on the defined concept to the first cell of the 'Class Name' Column
''utilising the behaviour of the Tables - MS Excel will automatically populate the remaining rows
    Cells(2, NewColNum).Select

    If Concept = "Document" Then
        ActiveCell.FormulaR1C1 = Doc_Extends
    ElseIf Concept = "Functional" Then
        ActiveCell.FormulaR1C1 = Func_Extends
    ElseIf Concept = "General" Then
        ActiveCell.FormulaR1C1 = Gen_Extends
    ElseIf Concept = "Physical" Then
        ActiveCell.FormulaR1C1 = Phys_Extends
    End If
        
    Columns(NewColNum).EntireColumn.AutoFit
    
ColumnExists:
    Cells(1, NewColNum).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 10498160
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
    Columns(NewColNum).EntireColumn.AutoFit
    
    Exit Sub


End Sub

Private Sub Add_Attribute_Validation_of_Source(Tbl As ListObject, TableName As String)

On Error GoTo Ignore
    
    Dim NewColNum As Integer

    NewColNum = Tbl.ListColumns("Source").Index
    NewColNum = NewColNum + 1
    
''Checks if Column has already been created - if so exits the Sub and continues
    If Tbl.ListColumns(NewColNum).Name = "Attribute Name" Then GoTo ColumnExists
    
    Tbl.ListColumns.Add(NewColNum).Name = "Attribute Name"
    
    Cells(1, NewColNum).Select
        With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 10498160
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With

     Cells(2, NewColNum).Select
     ActiveCell.FormulaR1C1 = _
            "=IFERROR(IF(VLOOKUP([Source],ISM_Attributes[Id],1,FALSE)=[Source],VLOOKUP([Source],ISM_Attributes[[Id]:[Name]],2,FALSE),""ERROR""),""ERROR"")"

    Columns(NewColNum).EntireColumn.AutoFit
    
ColumnExists:
    Cells(1, NewColNum).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 10498160
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
    Columns(NewColNum).EntireColumn.AutoFit
    
    Exit Sub

Ignore:
'' This evaluates to see if the column has been created and if so deletes it as the reference does not exist

    Dim AttVCol As Integer
    
    AttVCol = Tbl.ListColumns("Attribute Name").Index
    
    If AttVCol > 0 Then
        Tbl.ListColumns(UoMVCol).Delete
    End If
     
   
End Sub
