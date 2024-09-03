Attribute VB_Name = "Create_Discipline_files"
Public Sub Create_Discipline_Files()
    Dim CL_Name As String
    Dim Concept As String
            
    Application.ScreenUpdating = False

    Worksheets("ISM Class Library Header").Select
    CL_Name = Range("C2").Value
   
    Create_Lookups
    
    Add_Disc
    
    Define_Disciplines CL_Name
    
    Delete_Disc
    
    Application.DisplayAlerts = False
    Worksheets("Lookups").Delete
    Application.DisplayAlerts = True
    
    Worksheets("ISM Class Library Header").Select
    Range("B2").Select
    
    ActiveWorksheet.Save
    
    Application.ScreenUpdating = True
    
    MsgBox ("All finished")
    
End Sub


Private Sub Create_Lookups()
Attribute Create_Lookups.VB_ProcData.VB_Invoke_Func = " \n14"

Dim Lookup As Worksheet

Dim Func_Tbl As ListObject
Dim Phys_Tbl As ListObject
Dim Att_Tbl As ListObject


Dim LastRow As Integer
''To define the Last row of a specific column use (where 1 is the specific Column):
''LastRow = Cells(Rows.Count, 1).End(xlUp).Row

''Creates the Lookups Sheet
    With Lookup
        Set Lookup = Sheets.Add
        Lookup.Name = "Lookups"
    End With
    
''Copys the data from Functional Classses to Lookups Sheet
    Range("ISM_Functional_Classes[[#All],[ID]:[Name]]").Copy _
        Destination:=Worksheets("Lookups").Range("B1")
    Range("ISM_Functional_Classes[[#All],[Extends]]").Copy _
        Destination:=Worksheets("Lookups").Range("D1")
    Range("ISM_Functional_Classes[[#All],[_Action]]").Copy _
        Destination:=Worksheets("Lookups").Range("E1")

''Copys the data from Physical Classses to Lookups Sheet
    Range("ISM_Physical_Classes[[#All],[ID]:[Name]]").Copy _
        Destination:=Worksheets("Lookups").Range("G1")
    Range("ISM_Physical_Classes[[#All],[Extends]]").Copy _
        Destination:=Worksheets("Lookups").Range("I1")
    Range("ISM_Physical_Classes[[#All],[_Action]]").Copy _
        Destination:=Worksheets("Lookups").Range("J1")
    
''Copys the data from Attributes to Lookups Sheet
    Range("ISM_Attributes[[#All],[ID]:[Description]]").Copy _
        Destination:=Worksheets("Lookups").Range("L1")
    Range("ISM_Attributes[[#All],[_Action]]").Copy _
        Destination:=Worksheets("Lookups").Range("O1")
    
    ActiveSheet.Cells.ClearFormats
    
''Create Lookup_Func_Class table
    LastRow = Cells(Rows.Count, 2).End(xlUp).Row
    
    With Func_Tbl
        Set Func_Tbl = ActiveSheet.ListObjects.Add(xlSrcRange, Range(Cells(1, 2), Cells(LastRow, 5)), , xlYes)
            Func_Tbl.Name = "Lookup_Func_Classes"
            Func_Tbl.TableStyle = "TableStyleLight14"
    End With
    
''Create Lookup_Phys_Class table
    LastRow = Cells(Rows.Count, 7).End(xlUp).Row
    
    With Phys_Tbl
        Set Phys_Tbl = ActiveSheet.ListObjects.Add(xlSrcRange, Range(Cells(1, 7), Cells(LastRow, 10)), , xlYes)
            Phys_Tbl.Name = "Lookup_Phys_Classes"
            Phys_Tbl.TableStyle = "TableStyleLight12"
    End With
    
''Create Lookup_Attributes table
    LastRow = Cells(Rows.Count, 12).End(xlUp).Row
    
    With Att_Tbl
        Set Att_Tbl = ActiveSheet.ListObjects.Add(xlSrcRange, Range(Cells(1, 12), Cells(LastRow, 15)), , xlYes)
            Att_Tbl.Name = "Lookup_Attributes"
            Att_Tbl.TableStyle = "TableStyleLight10"
    End With
    
    Range("A1").Select
End Sub

Private Sub Add_Disc()

    Dim Concept As String
   
    Concept = "Functional"
    Add_Disc_CA Concept
    
    Concept = "Physical"
    Add_Disc_CA Concept

End Sub

Private Sub Add_Disc_CA(Concept As String)

'' Add Discipline to the Funtional and Physical Class Attribute tables

Dim NewColNum As Integer
Dim CNRange As String
Dim DiscRange As String
Dim Disc_Formula As String
Dim Tbl As ListObject

''Sets Variables based on Concept = Functional
    If Concept = "Functional" Then
        Worksheets("ISM Functional Class Attributes").Select
        
        CNRange = "ISM_Functional_Class_Attributes[Class Name]"
        DiscRange = "ISM_Functional_Class_Attributes[[#Headers],[Class Discipline]]"
        Disc_Formula = "=VLOOKUP([Class_Id],ISM_Functional_Classes[[Id]:[Discipline]],4,FALSE)"
        
        Set Tbl = ActiveSheet.ListObjects("ISM_Functional_Class_Attributes")

''Sets Variables based on Concept = Physical
    ElseIf Concept = "Physical" Then
        Worksheets("ISM Physical Class Attributes").Select
        
        CNRange = "ISM_Physical_Class_Attributes[Class Name]"
        DiscRange = "ISM_Physical_Class_Attributes[[#Headers],[Class Discipline]]"
        Disc_Formula = "=VLOOKUP([Class_Id],ISM_Physical_Classes[[Id]:[Discipline]],4,FALSE)"

        Set Tbl = ActiveSheet.ListObjects("ISM_Physical_Class_Attributes")
        
    End If
        
    NewColNum = Range(CNRange).Column
    Tbl.ListColumns.Add(NewColNum).Name = "Class Discipline"
    
    Range(DiscRange).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 10498160
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
    Cells(2, NewColNum).Select
    ActiveCell.FormulaR1C1 = Disc_Formula
End Sub

Private Sub Define_Disciplines(CL_Name As String)

Dim ColNo As Integer
Dim RowNo As Integer
Dim LastRow As Integer
Dim DiscLoop As Integer

Dim DiscValue As String
Dim DiscName As String

Dim AddDisc As Boolean

''Creates Collection
Dim Disciplines As Collection
Set Disciplines = New Collection

''Adds Discipline values to Discipline Collection from ISM Functional Class Attributes sheet
    Worksheets("ISM Functional Class Attributes").Select

    ColNo = Range("ISM_Functional_Class_Attributes[Class Discipline]").Column
    LastRow = Cells(Rows.Count, ColNo).End(xlUp).Row
    
    For RowNo = 2 To LastRow
        DiscValue = Cells(RowNo, ColNo).Value
        If Disciplines.Count = 0 Then
            Disciplines.Add DiscValue
        Else
            AddDisc = True
            
            For DiscLoop = 1 To Disciplines.Count
                If DiscValue = Disciplines(DiscLoop) Or DiscValue = "0" Then
                    AddDisc = False
                End If
            Next
            
            If AddDisc = True Then
                Disciplines.Add DiscValue
            End If
        
        End If
    Next

''Adds Discipline values to Discipline Collection from ISM Physical Class Attributes sheet
    Worksheets("ISM Physical Class Attributes").Select
    
    ColNo = Range("ISM_Physical_Class_Attributes[Class Discipline]").Column
    LastRow = Cells(Rows.Count, ColNo).End(xlUp).Row
    
    For RowNo = 2 To LastRow
        DiscValue = Cells(RowNo, ColNo).Value
        If Disciplines.Count = 0 Then
            Disciplines.Add DiscValue
        Else
            AddDisc = True
            
            For DiscLoop = 1 To Disciplines.Count
                If DiscValue = Disciplines(DiscLoop) Or DiscValue = "0" Then
                    AddDisc = False
                End If
            Next
            
            If AddDisc = True Then
                Disciplines.Add DiscValue
            End If
        
        End If
    Next

''Creates individual discipline sheets based on Disciplines in Collection
    For DiscLoop = 1 To Disciplines.Count
        DiscName = Disciplines(DiscLoop)
''        Debug.Print DiscName
        
        Create_Discipline_Worksheet CL_Name, DiscName
    
    Next
    
End Sub

Private Sub Create_Discipline_Worksheet(CL_Name As String, DiscName As String)

Dim FileName As String
Dim FilePath As String

Dim LookupSheet As Worksheet
Dim CLFCASheet As Worksheet
Dim CLPCASheet As Worksheet
Dim LUFCASheet As Worksheet
Dim LUPCASheet As Worksheet

Dim CLFCATbl As ListObject
Dim CLPCATbl As ListObject
Dim LUFCATbl As ListObject
Dim LUPCATbl As ListObject

Dim ClassLibraryWB As Workbook
Dim DiscCAWB As Workbook
    
    Set ClassLibraryWB = ActiveWorkbook
    
    Set LookupSheet = ClassLibraryWB.Sheets("Lookups")
    Set CLFCASheet = ClassLibraryWB.Sheets("ISM Functional Class Attributes")
    Set CLPCASheet = ClassLibraryWB.Sheets("ISM Physical Class Attributes")
    
    Set CLFCATbl = CLFCASheet.ListObjects("ISM_Functional_Class_Attributes")
    Set CLPCATbl = CLPCASheet.ListObjects("ISM_Physical_Class_Attributes")

    FileName = CL_Name & " (" & DiscName & ")"
    FilePath = Application.ActiveWorkbook.Path & "/"
''    Debug.Print FilePath
    
    FileName = Replace(FileName, "/", " and ")
    
    Set DiscCAWB = Workbooks.Add
    DiscCAWB.SaveAs FilePath & FileName & ".xlsx"
    
    LookupSheet.Copy before:=DiscCAWB.Sheets(1)
    
''    Set LUPCASheet = Worksheets.Add
''    LUPCASheet.Name = "ISM Physical Class Attributes"
  
    Create_FuncClassAtt_Sheet ClassLibraryWB, DiscCAWB, CLFCASheet, DiscName
    Create_PhysClassAtt_Sheet ClassLibraryWB, DiscCAWB, CLFCASheet, DiscName
    
    Application.DisplayAlerts = False
    DiscCAWB.Sheets("Sheet1").Delete
    Application.DisplayAlerts = True
    DiscCAWB.Save
    DiscCAWB.Close

End Sub
    

Private Sub Delete_Disc()

    Delete_Disc_CA "Functional"
    Delete_Disc_CA "Physical"

End Sub

Private Sub Delete_Disc_CA(Concept As String)

Dim Tbl As ListObject

    If Concept = "Functional" Then
        Worksheets("ISM Functional Class Attributes").Select
        Set Tbl = ActiveSheet.ListObjects("ISM_Functional_Class_Attributes")
    
    ElseIf Concept = "Physical" Then
        Worksheets("ISM Physical Class Attributes").Select
        Set Tbl = ActiveSheet.ListObjects("ISM_Physical_Class_Attributes")
    
    End If
    
    Tbl.ListColumns("Class Discipline").Delete

End Sub

Private Sub Create_FuncClassAtt_Sheet(ClassLibraryWB As Workbook, DiscCAWB As Workbook, CLFCASheet As Worksheet, DiscName As String)

Dim LUFCASheet As Worksheet

Dim CLFCATbl As ListObject
Dim LUFCATbl As ListObject

Dim LastRow As Integer
Dim LastCol As Integer

    DiscCAWB.Activate
    Set LUFCASheet = Worksheets.Add
    LUFCASheet.Name = "ISM Functional Class Attributes"
    
    Set CLFCATbl = CLFCASheet.ListObjects("ISM_Functional_Class_Attributes")
    
    With CLFCATbl.Range
        .AutoFilter field:=2, Criteria1:=DiscName
        .SpecialCells(xlCellTypeVisible).Copy _
            Destination:=LUFCASheet.Range("A1")
    End With
    
    CLFCASheet.ShowAllData
    
    ActiveSheet.Cells.ClearFormats

    LastRow = Cells(Rows.Count, 2).End(xlUp).Row
    LastCol = Cells(1, Columns.Count).End(xlToLeft).Column
    
    With LUFCATbl
        Set LUFCATbl = LUFCASheet.ListObjects.Add(xlSrcRange, Range(Cells(1, 1), Cells(LastRow, LastCol)), , xlYes)
        LUFCATbl.Name = "LU_Functional_Class_Attributes"
        LUFCATbl.TableStyle = "TableStyleMedium9"
    End With
    
    LUFCATbl.ListColumns("Class Discipline").Delete
    LUFCATbl.ListColumns("Class Name").Delete
    LUFCATbl.ListColumns("Attribute Name").Delete
    LUFCATbl.ListColumns("Attribute Description").Delete
    LUFCATbl.ListColumns("Duplicate Check").Delete
    
    Add_Duplicate_Check LUFCATbl
    Add_ClassID_Check "Functional", LUFCATbl
    Add_ID_Check LUFCATbl
    Add_Attribute_Description_Check LUFCATbl
    Add_Validate_Extends "Functional", LUFCATbl
    Add_Inheritance_Check LUFCATbl
    
    LUFCASheet.Cells.WrapText = True
    LUFCASheet.Columns.AutoFit
    LUFCASheet.Rows.AutoFit
    
End Sub

Private Sub Create_PhysClassAtt_Sheet(ClassLibraryWB As Workbook, DiscCAWB As Workbook, CLPCASheet As Worksheet, DiscName As String)

Dim LUPCASheet As Worksheet

Dim CLPCATbl As ListObject
Dim LUPCATbl As ListObject

Dim LastRow As Integer
Dim LastCol As Integer

    DiscCAWB.Activate
    Set LUPCASheet = Worksheets.Add
    LUPCASheet.Name = "ISM Physical Class Attributes"
    
    Set CLPCATbl = CLPCASheet.ListObjects("ISM_Functional_Class_Attributes")
    
    With CLPCATbl.Range
        .AutoFilter field:=2, Criteria1:=DiscName
        .SpecialCells(xlCellTypeVisible).Copy _
            Destination:=LUPCASheet.Range("A1")
    End With
    
    CLPCASheet.ShowAllData
    
    ActiveSheet.Cells.ClearFormats

    LastRow = Cells(Rows.Count, 2).End(xlUp).Row
    LastCol = Cells(1, Columns.Count).End(xlToLeft).Column
    
    With LUPCATbl
        Set LUPCATbl = LUPCASheet.ListObjects.Add(xlSrcRange, Range(Cells(1, 1), Cells(LastRow, LastCol)), , xlYes)
        LUPCATbl.Name = "LU_Physical_Class_Attributes"
        LUPCATbl.TableStyle = "TableStyleMedium9"
    End With
    
    LUPCATbl.ListColumns("Class Discipline").Delete
    LUPCATbl.ListColumns("Class Name").Delete
    LUPCATbl.ListColumns("Attribute Name").Delete
    LUPCATbl.ListColumns("Attribute Description").Delete
    LUPCATbl.ListColumns("Duplicate Check").Delete
    
    Add_Duplicate_Check LUPCATbl
    Add_ClassID_Check "Functional", LUPCATbl
    Add_ID_Check LUPCATbl
    Add_Attribute_Description_Check LUPCATbl
    Add_Validate_Extends "Functional", LUPCATbl
    Add_Inheritance_Check LUPCATbl
    
    LUPCASheet.Cells.WrapText = True
    LUPCASheet.Columns.AutoFit
    LUPCASheet.Rows.AutoFit
    
    LUPCASheet.Move After:=Worksheets("ISM Functional Class Attributes")
    
End Sub


Private Sub Add_Duplicate_Check(Tbl As ListObject)

Dim SortID As String
Dim TableName As String
Dim LastColumn As Integer
      
    SortID = TableName & "[[#All],[Duplicate Check]]"
    TableName = Tbl.Name
    
    LastColumn = Cells(1, Columns.Count).End(xlToLeft).Column + 1

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
    
    Columns(LastColumn).EntireColumn.AutoFit
    
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

End Sub


Private Sub Add_ClassID_Check(Concept As String, Tbl As ListObject)
    
    Dim NewColNum As Integer
    
    Dim FuncClassIDCheck As String, PhysClassIDCheck As String
    
    FuncClassIDCheck = "=IF([Class_Id] = """",""ERROR"",IF(VLOOKUP([Class_Id],Lookup_Func_Classes[Id],1,FALSE)=[Class_Id],VLOOKUP([Class_Id],Lookup_Func_Classes,2,FALSE),""ERROR""))"
    PhysClassIDCheck = "=IF([Class_Id] = """",""ERROR"",IF(VLOOKUP([Class_Id],Lookup_Phys_Classes[Id],1,FALSE)=[Class_Id],VLOOKUP([Class_Id],Lookup_Phys_Classes,2,FALSE),""ERROR""))"
            
''Adds the column 'Class Name' to the left of "Class_ID"
    NewColNum = Tbl.ListColumns("Class_ID").Index
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
    
    If Concept = "Functional" Then
        ActiveCell.FormulaR1C1 = FuncClassIDCheck
    ElseIf Concept = "Physical" Then
        ActiveCell.FormulaR1C1 = PhysClassIDCheck
    Else
    End If
    
    Columns(NewColNum).EntireColumn.AutoFit
    
End Sub

Private Sub Add_ID_Check(Tbl As ListObject)

    Dim IDFormula As String
    Dim ColName As String
    
    Dim NewColNum As Integer
    
    IDFormula = "=IF([Class_Id] = """",""ERROR"",IF(VLOOKUP([Id],Lookup_Attributes[Id],1,FALSE)=[Id],VLOOKUP([Id],Lookup_Attributes,2,FALSE),""ERROR""))"
    ColName = "Attribute Name"

    NewColNum = Tbl.ListColumns("ID").Index
    NewColNum = NewColNum + 1
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

End Sub

Private Sub Add_Attribute_Description_Check(Tbl As ListObject)

    Dim NewColNum As Integer

    NewColNum = Tbl.ListColumns("Attribute Name").Index
    NewColNum = NewColNum + 1
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
            "=IF([Class_Id] = """",""ERROR"",IF(VLOOKUP([Id],Lookup_Attributes[Id],1,FALSE)=[Id],VLOOKUP([Id],Lookup_Attributes,3,FALSE),""ERROR""))"

End Sub

Private Sub Add_Validate_Extends(Concept As String, Tbl As ListObject)

    Dim NewColNum As Integer
    Dim ValFormula As String
    
    If Concept = "Functional" Then
        ValFormula = "=VLOOKUP([Class_Id],Lookup_Func_Classes,3,FALSE)"
    ElseIf Concept = "Physical" Then
        ValFormula = "=VLOOKUP([Class_Id],Lookup_Func_Classes,3,FALSE)"
    End If

    NewColNum = Tbl.ListColumns("_Action").Index
    NewColNum = NewColNum + 1
    Tbl.ListColumns.Add(NewColNum).Name = "Validate Extends"
    
    Cells(1, NewColNum).Select
        With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = 10498160
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With

     Cells(2, NewColNum).Select
     ActiveCell.FormulaR1C1 = ValFormula

End Sub

Private Sub Add_Inheritance_Check(Tbl As ListObject)

    Dim NewColNum As Integer
    Dim InhFormula As String
    
    InhFormula = "=IF(COUNTIF([Duplicate Check],[Validate Extends]&"".""&[Id])>=1,""Inherited"","""")"

    NewColNum = Tbl.ListColumns("Validate Extends").Index
    NewColNum = NewColNum + 1
    Tbl.ListColumns.Add(NewColNum).Name = "Inheritance Check"
    
    Cells(1, NewColNum).Select
        With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = 10498160
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With

     Cells(2, NewColNum).Select
     ActiveCell.FormulaR1C1 = InhFormula

End Sub
