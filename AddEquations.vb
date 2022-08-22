'Add in to help with equation editing for Office 365 (2019)
'By George Keeling and on my blog at www.general-relativity.net
'Possibly documented at https://www.general-relativity.net/search/label/Tools
'Feel free to use, copy, modify and give away but not for commerce. Please acknowledge me

'Entry points:
'	EquationTable2 - Create a table with 2 columns suitable for numbered equations
'	EquationTable3 - Create a table with 3 columns suitable for numbered equations 
'	InsertEquationInLine - Insert inline equation
'	InsertEquationNewLine - Insert unnumbered equation on New line

Module AddEquations
    Public BelowTable As Boolean        'insertion point for new equation
    'Change these to adjust table layout for equations to suit your taste.
    Public EquationFontSize As Single          '**** Equation font size in points
    Public TableHeight As Single             '**** Table Height in cm - 1.29 leaves a slight gap
    ' above and below equations, avoids lots of blank lines
    Public EquationColumnWidth As Single     'column width of equation number column
    Public prevScrollPos As Long      'to preserve vertical scroll during equation insertion

    Sub EquationTable2()
        ' Create a table with 2 columns suitable for numbered equations of form
        '   x = sin a / cos a   (9)
        '      x = tan a       (10)
        Dim SaveFontSize As Single
        Dim TableWidth As Single        'Table width in points. Make sure that table uses whole width
        Dim ThisRow As Long

        Dim Selection As Word.Selection
        Dim objUndo As Word.UndoRecord
        objUndo = Globals.EquationAcc.Application.UndoRecord
        objUndo.StartCustomRecord("EquationTable2")

        Selection = Globals.EquationAcc.Application.ActiveDocument.ActiveWindow.Selection

        SaveFontSize = Selection.Font.Size
        TableWidth = CreateTable(2)
        If TableWidth = 0 Then
            objUndo.EndCustomRecord()
            Exit Sub
        End If

        'Get in first cell
        ThisRow = Selection.Tables(1).Rows.Count
        Selection.Tables(1).Cell(ThisRow, 1).Select()

        If Not BelowTable Then
            Selection.Columns.PreferredWidth = TableWidth - CentimetersToPoints(EquationColumnWidth)

            Selection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
        End If

        InsertEquation(Selection, "=", Word.WdOMathJc.wdOMathJcCenterGroup)

        'second cell
        Selection.Tables(1).Cell(ThisRow, 2).Select()
        Call FinishOff(SaveFontSize)
        objUndo.EndCustomRecord()
    End Sub
    Sub EquationTable3()
        ' Create a table with 3 columns suitable for numbered equations of form
        ' x = sin a / cos a      (9)
        '   = tan a             (10)
        Dim objUndo As Word.UndoRecord

        objUndo = Globals.EquationAcc.Application.UndoRecord
        objUndo.StartCustomRecord("EquationTable3")

        EquationTable3a(3.38)      'Not sure where that number came from!

        objUndo.EndCustomRecord()
    End Sub
    Function EquationTable3a(Col1Width As Single) As Boolean
        'Col1Width = column 1 width
        Dim SaveFontSize
        Dim ThisRow As Long
        Dim TableWidth As Single        'Table width in points. Make sure that table uses whole width
        Dim Selection As Word.Selection

        Selection = Globals.EquationAcc.Application.ActiveDocument.ActiveWindow.Selection
        EquationTable3a = False
        SaveFontSize = Selection.Font.Size

        TableWidth = CreateTable(3)
        If TableWidth = 0 Then Exit Function

        'Get in first cell
        ThisRow = Selection.Tables(1).Rows.Count
        Selection.Tables(1).Cell(ThisRow, 1).Select()

        If Not BelowTable Then
            Selection.Tables(1).Columns(1).PreferredWidth = CentimetersToPoints(Col1Width)
            Selection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
        End If

        'Insert right justified equation
        InsertEquation(Selection, "x", Word.WdOMathJc.wdOMathJcRight)
        'Selection.OMaths(1).Justification = Word.WdOMathJc.wdOMathJcRight did not work out here!

        'Second cell
        Selection.Tables(1).Cell(ThisRow, 2).Select()

        If Not BelowTable Then
            Selection.Columns.PreferredWidth = TableWidth - CentimetersToPoints(Col1Width + EquationColumnWidth)
            Selection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft
        End If

        'Insert left justified equation
        Selection.MoveLeft(Word.WdUnits.wdCharacter, 1)
        InsertEquation(Selection, "=", Word.WdOMathJc.wdOMathJcLeft)

        'Third cell
        Selection.Tables(1).Cell(ThisRow, 3).Select()
        Call FinishOff(SaveFontSize)
        EquationTable3a = True
    End Function
    Sub InsertEquationInLine()
        'Insert inline equation with font EquationFont.
        InsertEquationNoNumber(Word.WdOMathType.wdOMathInline)
    End Sub
    Sub InsertEquationNewLine()
        'Insert unnumbered equation on new line with font EquationFont.
        InsertEquationNoNumber(Word.WdOMathType.wdOMathDisplay)
    End Sub
    Sub InsertEquationNoNumber(DisplayType As Word.WdOMathType)
        'Insert unnumbered inline or display equation with required font
        Dim objUndo As Word.UndoRecord
        Dim Selection As Word.Selection
        Dim objRange As Word.Range

        Selection = Globals.EquationAcc.Application.ActiveDocument.ActiveWindow.Selection

        If Selection.OMaths.Count > 0 Then
            TensorError("The insertion point cannot be in an equation." + Chr(10) _
                + "Move the insertion point please.")
            Exit Sub
        End If
        objUndo = Globals.EquationAcc.Application.UndoRecord
        objUndo.StartCustomRecord("InsertEquationInLine")
        If DisplayType = Word.WdOMathType.wdOMathInline Then
            Selection.TypeText(" ")
            Selection.MoveLeft(Word.WdUnits.wdCharacter, 1)
        End If
        objRange = Selection.Range
        Selection.OMaths.Add(objRange)

        Selection.OMaths(1).Range.Font.Size = EquationFontSize
        Selection.OMaths(1).Type = DisplayType
        objUndo.EndCustomRecord()
    End Sub
    Function CreateTable(Columns As Integer) As Single
        'Create table with Columns (Columns = 2 or 3) columns
        'returns width of table in points, 0 if error
        'Ends execution if insertion point was invalid
        'These rules apply to inserting a table in MS-Word. We follow them with additions
        'Insertion point on line below table extends table with same characteristics.
        '   public BelowTable remembers this
        '   This leads to nasty dependencies on BelowTable throughout creation od new row / table
        'Insertion point on line above table has new table separated from table below
        'Insertion point in text splits text
        Dim LineStyle As Word.WdLineStyle
        Dim Selection As Word.Selection
        Dim ActiveDoc As Word.Document
        Dim LinesMoved As Integer

        Selection = Globals.EquationAcc.Application.ActiveDocument.ActiveWindow.Selection
        ActiveDoc = Globals.EquationAcc.Application.ActiveDocument

        CreateTable = 0
        If Selection.Information(Word.WdInformation.wdWithInTable) Or (Selection.OMaths.Count > 0) Then
            TensorError("The insertion point cannot be in a table or equation." & vbLf &
                "Move the insertion point please.")
            Exit Function
        End If

        prevScrollPos = ActiveDoc.ActiveWindow.VerticalPercentScrolled

        LinesMoved = Selection.MoveUp(Word.WdUnits.wdLine, 1)           'Necessary in case selection at top of doc
        BelowTable = Selection.Information(Word.WdInformation.wdWithInTable)
        If BelowTable Then
            If (Selection.Tables(1).Columns.Count = Columns) Then
                'Create a row like the row above
                Selection.MoveRight(Word.WdUnits.wdCell, 2)
                If Columns = 3 Then Selection.MoveRight(Word.WdUnits.wdCell)
                CreateTable = Selection.Tables(1).PreferredWidth
            Else
                'trying to add wrong kind of row, so insert blank line
                BelowTable = False      'not copying from table above
                Selection.MoveDown(Word.WdUnits.wdLine, 1)
                Selection.TypeParagraph()
                Selection.MoveUp(Word.WdUnits.wdLine, 1)
            End If
        End If
        If Not BelowTable Then
            'Create a new table
            Selection.MoveDown(Word.WdUnits.wdLine, LinesMoved)
            'for mysterious reasons the movedown sometimes just moves into the lower index of a variable (bug 7) so
            If Selection.Information(Word.WdInformation.wdWithInTable) Then
                Selection.MoveDown(Word.WdUnits.wdLine, 1)
            End If
            LineStyle = Word.WdLineStyle.wdLineStyleNone     '  wdLineStyleNone / wdLineStyleSingle
            'Last 2 parameters of below were DefaultTableBehavior:=wdWord9TableBehavior,
            'AutoFitBehavior:=wdAutoFitFixed
            ActiveDoc.Tables.Add(Selection.Range, 1, Columns,
                                 Word.WdDefaultTableBehavior.wdWord9TableBehavior, Word.WdAutoFitBehavior.wdAutoFitFixed)

            'Mystery here. Removed three VBA lines below. 
            'If Selection.Tables(1).Style <> "Table Grid" Then
            '   Selection.Tables(1).Style = "Table Grid"
            'End If
            'Selection.Tables(1).Style is not a string and gets error

            'Dim oStyle As Word.Style
            'oStyle = Selection.Tables(1).Style
            'oStyle.NameLocal 
            'got a name but
            'Selection.Tables(1).Style.NameLocal
            'does not work

            With Selection.Tables(1)
                .Rows.HeightRule = Word.WdRowHeightRule.wdRowHeightAtLeast
                .Rows.Height = CentimetersToPoints(TableHeight)
                .ApplyStyleHeadingRows = True
                .ApplyStyleLastRow = False
                .ApplyStyleFirstColumn = True
                .ApplyStyleLastColumn = False
                .ApplyStyleRowBands = True
                .ApplyStyleColumnBands = False
                .Borders(Word.WdBorderType.wdBorderTop).LineStyle = LineStyle
                .Borders(Word.WdBorderType.wdBorderLeft).LineStyle = LineStyle
                .Borders(Word.WdBorderType.wdBorderBottom).LineStyle = LineStyle
                .Borders(Word.WdBorderType.wdBorderRight).LineStyle = LineStyle
                .Borders(Word.WdBorderType.wdBorderVertical).LineStyle = LineStyle
                .Borders(Word.WdBorderType.wdBorderDiagonalDown).LineStyle = Word.WdLineStyle.wdLineStyleNone
                .Borders(Word.WdBorderType.wdBorderDiagonalUp).LineStyle = Word.WdLineStyle.wdLineStyleNone
                .PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPoints
                CreateTable = .PreferredWidth
            End With
        End If
    End Function
    Sub InsertEquation(Selection As Word.Selection, Content As String, Justification As Word.WdOMathJc)
        'Justification = Word.WdOMathJc.xxxx
        Dim objRange As Word.Range
        Selection.Font.Size = EquationFontSize
        Selection.MoveRight(Word.WdUnits.wdCharacter, 1)   'wiggle cursor
        Selection.MoveLeft(Word.WdUnits.wdCharacter, 1)
        'Remainder replaces WordBasic.EquationEdit from VBA.  
        objRange = Selection.Range
        Selection.OMaths.Add(objRange)
        objRange.Text = Content
        Selection.OMaths(1).Justification = Justification       'was Selection.OMaths(1).ParentOMath.Justification
        Selection.Delete(Word.WdUnits.wdCharacter, 1)       'Removes Type Equation Here
        Selection.MoveRight(Word.WdUnits.wdCharacter, 1)
    End Sub
    Sub FinishOff(FontSize)
        'Do the last column which contains formula number. That cell should be selected
        'Also vertically align central (useful if formula gets tall with fractions or matrices etc)
        Dim Selection As Word.Selection
        Dim ActiveDocument As Word.Document
        Selection = Globals.EquationAcc.Application.ActiveDocument.ActiveWindow.Selection
        ActiveDocument = Globals.EquationAcc.Application.ActiveDocument

        If Not BelowTable Then
            Selection.Font.Size = FontSize
            Selection.Columns.PreferredWidth = CentimetersToPoints(EquationColumnWidth)
            Selection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
        End If

        Call InsertEquationNumber()

        If Not BelowTable Then
            Selection.Tables(1).Select()
            Selection.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
            Selection.MoveLeft(Word.WdUnits.wdCharacter, 1)
        Else
            Selection.MoveLeft(Word.WdUnits.wdCharacter, 4)
        End If

        Call ClearFindParameters()
        ActiveDocument.ActiveWindow.VerticalPercentScrolled = prevScrollPos
    End Sub
    Sub InsertEquationNumber()
        'Insert Equation Number. (1) or previous +1
        'Equation number cell of table (at right)should be selected
        Dim EquationNumber As Integer
        Dim Bookmark As String
        Dim Selection As Word.Selection
        Dim ActiveDocument As Word.Document
        Selection = Globals.EquationAcc.Application.ActiveDocument.ActiveWindow.Selection
        ActiveDocument = Globals.EquationAcc.Application.ActiveDocument

        Bookmark = "xxMostRecentEquation"
        Selection.MoveLeft(Word.WdUnits.wdCharacter, 1)
        ActiveDocument.Bookmarks.Add(Bookmark, Selection.Range)
        'was
        'With ActiveDocument.Bookmarks.Add(Bookmark, Selection.Range)
        '.DefaultSorting = wdSortByName
        '.ShowHidden = False
        'End With

        EquationNumber = 1
        Do While FindWildStopUp("\([1-9]*\)")
            If Selection.Information(Word.WdInformation.wdWithInTable) Then
                EquationNumber = Val(Mid(Selection.Text, 2)) + 1
                Exit Do
            End If
        Loop

        Selection.GoTo(Word.WdGoToItem.wdGoToBookmark,,, Bookmark)
        Selection.Text = "(" + CStr(EquationNumber) + ")"
    End Sub
    Function CentimetersToPoints(cm As Single) As Single
        Return Globals.EquationAcc.Application.CentimetersToPoints(cm)
    End Function
End Module