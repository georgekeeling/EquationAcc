'Add in to help with equation editing for Office 365 (2019)
'By George Keeling and on my blog at www.general-relativity.net
'Possibly documented at https://www.general-relativity.net/search/label/Tools
'Feel free to use, copy, modify and give away but not for commerce. Please acknowledge me

'Entry points:
'	SelectTable, BordersAll, BordersNone, BordersOutside, Point8cmTableRows - table functions

Module TableActions
    Sub SelectTable()
        Dim Selection As Word.Selection
        Selection = Globals.EquationAcc.Application.ActiveDocument.ActiveWindow.Selection
        If Not Selection.Information(Word.WdInformation.wdWithInTable) Then Exit Sub
        Selection.Tables(1).Select()
    End Sub
    Sub BordersAll()
        Dim Selection As Word.Selection
        Dim ActiveDocument As Word.Document
        Dim Options As Word.Options
        Dim Bookmark As String

        Selection = Globals.EquationAcc.Application.ActiveDocument.ActiveWindow.Selection
        ActiveDocument = Globals.EquationAcc.Application.ActiveDocument
        Options = Globals.EquationAcc.Application.Options

        If Not Selection.Information(Word.WdInformation.wdWithInTable) Then Exit Sub

        prevScrollPos = ActiveDocument.ActiveWindow.VerticalPercentScrolled

        Bookmark = "xxMostRecentEquation"
        With ActiveDocument.Bookmarks
            .Add(Bookmark, Selection)
            .DefaultSorting = Word.WdBookmarkSortBy.wdSortByName
            .ShowHidden = False
        End With

        Selection.Tables(1).Select()
        With Selection.Borders(Word.WdBorderType.wdBorderTop)
            .LineStyle = Options.DefaultBorderLineStyle
            .LineWidth = Options.DefaultBorderLineWidth
            .Color = Options.DefaultBorderColor
        End With
        With Selection.Borders(Word.WdBorderType.wdBorderLeft)
            .LineStyle = Options.DefaultBorderLineStyle
            .LineWidth = Options.DefaultBorderLineWidth
            .Color = Options.DefaultBorderColor
        End With
        With Selection.Borders(Word.WdBorderType.wdBorderBottom)
            .LineStyle = Options.DefaultBorderLineStyle
            .LineWidth = Options.DefaultBorderLineWidth
            .Color = Options.DefaultBorderColor
        End With
        With Selection.Borders(Word.WdBorderType.wdBorderRight)
            .LineStyle = Options.DefaultBorderLineStyle
            .LineWidth = Options.DefaultBorderLineWidth
            .Color = Options.DefaultBorderColor
        End With
        On Error GoTo OneRow
        With Selection.Borders(Word.WdBorderType.wdBorderHorizontal)
            .LineStyle = Options.DefaultBorderLineStyle
            .LineWidth = Options.DefaultBorderLineWidth     ' error here on table with one row
            .Color = Options.DefaultBorderColor
        End With
OneRow:
        With Selection.Borders(Word.WdBorderType.wdBorderVertical)
            .LineStyle = Options.DefaultBorderLineStyle
            .LineWidth = Options.DefaultBorderLineWidth
            .Color = Options.DefaultBorderColor
        End With
        Selection.GoTo(Word.WdGoToItem.wdGoToBookmark,, , Bookmark)
        ActiveDocument.ActiveWindow.VerticalPercentScrolled = prevScrollPos
    End Sub
    Sub BordersNone()
        Dim Selection As Word.Selection
        Dim ActiveDocument As Word.Document
        Dim Options As Word.Options
        Dim Bookmark As String

        Selection = Globals.EquationAcc.Application.ActiveDocument.ActiveWindow.Selection
        ActiveDocument = Globals.EquationAcc.Application.ActiveDocument
        Options = Globals.EquationAcc.Application.Options

        If Not Selection.Information(Word.WdInformation.wdWithInTable) Then Exit Sub

        prevScrollPos = ActiveDocument.ActiveWindow.VerticalPercentScrolled

        Bookmark = "xxMostRecentEquation"
        With ActiveDocument.Bookmarks
            .Add(Bookmark, Selection.Range)
            .DefaultSorting = Word.WdBookmarkSortBy.wdSortByName
            .ShowHidden = False
        End With

        BordersNone2(Selection)

        Selection.GoTo(Word.WdGoToItem.wdGoToBookmark,, , Bookmark)
        ActiveDocument.ActiveWindow.VerticalPercentScrolled = prevScrollPos
    End Sub
    Sub BordersNone2(Selection As Word.Selection)
        Selection.Tables(1).Select()
        Selection.Borders(Word.WdBorderType.wdBorderTop).LineStyle = Word.WdLineStyle.wdLineStyleNone
        Selection.Borders(Word.WdBorderType.wdBorderLeft).LineStyle = Word.WdLineStyle.wdLineStyleNone
        Selection.Borders(Word.WdBorderType.wdBorderBottom).LineStyle = Word.WdLineStyle.wdLineStyleNone
        Selection.Borders(Word.WdBorderType.wdBorderRight).LineStyle = Word.WdLineStyle.wdLineStyleNone
        Selection.Borders(Word.WdBorderType.wdBorderHorizontal).LineStyle = Word.WdLineStyle.wdLineStyleNone
        Selection.Borders(Word.WdBorderType.wdBorderVertical).LineStyle = Word.WdLineStyle.wdLineStyleNone
        Selection.Borders(Word.WdBorderType.wdBorderDiagonalDown).LineStyle = Word.WdLineStyle.wdLineStyleNone
        Selection.Borders(Word.WdBorderType.wdBorderDiagonalUp).LineStyle = Word.WdLineStyle.wdLineStyleNone
    End Sub
    Sub BordersOutside()
        Dim Selection As Word.Selection
        Dim ActiveDocument As Word.Document
        Dim Options As Word.Options
        Dim Bookmark As String
        Dim LineWidth As Word.WdLineWidth

        Selection = Globals.EquationAcc.Application.ActiveDocument.ActiveWindow.Selection
        If Not Selection.Information(Word.WdInformation.wdWithInTable) Then Exit Sub
        Selection = Globals.EquationAcc.Application.ActiveDocument.ActiveWindow.Selection
        ActiveDocument = Globals.EquationAcc.Application.ActiveDocument
        Options = Globals.EquationAcc.Application.Options

        prevScrollPos = ActiveDocument.ActiveWindow.VerticalPercentScrolled

        LineWidth = Word.WdLineWidth.wdLineWidth150pt
        Bookmark = "xxMostRecentEquation"
        With ActiveDocument.Bookmarks
            .Add(Bookmark, Selection.Range)
            .DefaultSorting = Word.WdBookmarkSortBy.wdSortByName
            .ShowHidden = False
        End With

        BordersNone2(Selection)

        With Selection.Borders(Word.WdBorderType.wdBorderTop)
            .LineStyle = Options.DefaultBorderLineStyle
            .LineWidth = LineWidth
            .Color = Options.DefaultBorderColor
        End With
        With Selection.Borders(Word.WdBorderType.wdBorderLeft)
            .LineStyle = Options.DefaultBorderLineStyle
            .LineWidth = LineWidth
            .Color = Options.DefaultBorderColor
        End With
        With Selection.Borders(Word.WdBorderType.wdBorderBottom)
            .LineStyle = Options.DefaultBorderLineStyle
            .LineWidth = LineWidth
            .Color = Options.DefaultBorderColor
        End With
        With Selection.Borders(Word.WdBorderType.wdBorderRight)
            .LineStyle = Options.DefaultBorderLineStyle
            .LineWidth = LineWidth
            .Color = Options.DefaultBorderColor
        End With
        Selection.GoTo(Word.WdGoToItem.wdGoToBookmark,, , Bookmark)
        ActiveDocument.ActiveWindow.VerticalPercentScrolled = prevScrollPos
    End Sub
    Sub Point8cmTableRows()
        'Set all selected rows in table to be 0.8cm / or back to big size
        Dim Selection As Word.Selection
        Selection = Globals.EquationAcc.Application.ActiveDocument.ActiveWindow.Selection
        On Error GoTo NotInTable
        If Selection.Rows.Height < (CentimetersToPoints(0.8) + 0.1) Then
            Selection.Rows.Height = CentimetersToPoints(TableHeight)
        Else
            Selection.Rows.Height = CentimetersToPoints(0.8)
        End If
        Exit Sub
NotInTable:
        Call TensorError("The insertion point must be in a table.")
    End Sub
End Module
