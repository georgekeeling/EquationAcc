'Add in to help with equation editing for Office 365 (2019)
'By George Keeling and on my blog at www.general-relativity.net
'Possibly documented at https://www.general-relativity.net/search/label/Tools
'Feel free to use, copy, modify and give away but not for commerce. Please acknowledge me

'Entry points:
'	PrepareForWeb0Async - From document create plain text for all text And latex equations. Suitable for web
'	RenumberEquationsCodeAsync - renumber all equations And references to them
'Also Contains
'	FindWildStop, SearchReplace, ClearFindParameters, FindWildStopUp - find functions
Imports System.Threading.Tasks

Module TransformersOther
    Public InlineDelimiter As String = "##"
    Public DisplayDelimiter As String = "$$"
    Public BraPadding As String = "\phantom {10000}"   'amount of padding in front of BRAcket

    'globals for progresss window
    Public gProgressName As String
    Public gMainDocName As String
    Public gProgressStep As Integer

    'declarations and globals for equation renumbering
    'constants for renumbering equations
    Public MagicChar1 As String = "л"           'Cyrillic small EL
    Public MagicChar2 As String = "м"           'Cyrillic next one
    Public Structure EquationRefNum
        Public Num As String       'Equation number including brackets
        Public Refs As Single      'References to this equation
    End Structure
    Public gEquations() As EquationRefNum
    Public gRefErrors As String 'referencing errors, if any
    Public ElapsedTime As New Diagnostics.Stopwatch
End Module
Module PrepareForWebFunctions
    Async Function PrepareForWeb0Async() As Task
        'From document create plain text for all text and latex equations. Suitable for web
        'Latex equations must be switched on in MS-Word.
        'Equations in table are treated specially
        'Things that MS produces that MathJax doesn't like. dealt with in PasteEquation
        '\sfrac    \frac
        'removed vbCr in equation output. Messier output but doesn't screw up blogger
        Dim SourceDoc As Word.Document
        Dim PasteDoc As Word.Document           'minimised document just for pasting in individual equations
        Dim TargetW As New OutputWindow()

        SourceDoc = Globals.EquationAcc.Application.ActiveDocument
        If Not CreateOutputWindow(SourceDoc, PasteDoc) Then Exit Function
        If Not LatexOn(PasteDoc) Then
            TensorError("Please switch Latex on " & vbCrLf & "in equation tab.")
            PasteDoc.Close(Word.WdSaveOptions.wdDoNotSaveChanges)
            Exit Function
        End If
        TargetW.Show()
        Await PrepareForWebAsync(SourceDoc, PasteDoc, TargetW)

        PasteDoc.Close(Word.WdSaveOptions.wdDoNotSaveChanges)

    End Function
    Async Function PrepareForWebAsync(SourceDoc As Word.Document, PasteDoc As Word.Document,
                                      TargetW As OutputWindow) As Task
        'Loop that goes equation by equation converting each to latex text
        Dim Equation As Word.OMath, EquationRange As Word.Range
        Dim EquationRangeStart As Long, EquationRangeEnd As Long, TextRangeStart As Long
        Dim SourceRange As Word.Range, TargetRange As Word.Range       'whole source / target story
        Dim nEquation As Integer
        Dim TextToCopy As String
        Dim PrevEquation As Word.WdOMathType
        Dim ResultsTitle As String

        ResultsTitle = TargetW.Text
        ElapsedTime.Restart()
        SourceRange = SourceDoc.Range
        TextRangeStart = 0
        TargetRange = PasteDoc.Range
        PrevEquation = Word.WdOMathType.wdOMathInline

        For nEquation = 1 To SourceDoc.OMaths.Count
            Equation = SourceDoc.OMaths(nEquation)
            EquationRange = Equation.Range
            EquationRangeStart = EquationRange.Start
            EquationRangeEnd = EquationRange.End

            If EquationRangeStart - TextRangeStart > 0 Then
                'insert text, remove initial CR if after display equation
                TextToCopy = SourceDoc.Range(TextRangeStart, EquationRangeStart).Text
                If (PrevEquation = Word.WdOMathType.wdOMathDisplay) And (Mid(TextToCopy, 1, 1) = vbCr) Then
                    TextToCopy = Mid(TextToCopy, 2)
                End If
                TargetW.ResultsBox.AppendText(TextToCopy)
            End If

            If EquationRange.Information(Word.WdInformation.wdWithInTable) Then
                If Not ProcessTable(nEquation, EquationRangeEnd, SourceDoc, PasteDoc, TargetW) Then Return
                PrevEquation = Word.WdOMathType.wdOMathDisplay
            Else
                'either inline or stand alone display
                EquationRange.Copy()

                If Equation.Type = Word.WdOMathType.wdOMathInline Then
                    'In line very simple
                    TargetW.ResultsBox.AppendText(InlineDelimiter)
                    If Not PasteEquation(EquationRange, PasteDoc, TargetW) Then Return
                    TargetW.ResultsBox.AppendText(InlineDelimiter)
                Else
                    'possibly add vbCr after and before DisplayDelimiter for clarity
                    TargetW.ResultsBox.AppendText(DisplayDelimiter)
                    If Not PasteEquation(EquationRange, PasteDoc, TargetW) Then Return
                    TargetW.ResultsBox.AppendText(DisplayDelimiter)
                End If
                PrevEquation = Equation.Type
            End If
            TextRangeStart = EquationRangeEnd
            TargetW.Text = $"{ResultsTitle} {ElapsedTime.Elapsed.Minutes:00}:{ElapsedTime.Elapsed.Seconds:00}"
            Await Task.Delay(1)
        Next
        'Insert any text at the end
        TextToCopy = SourceDoc.Range(TextRangeStart, SourceDoc.Range.End).Text
        If (PrevEquation = Word.WdOMathType.wdOMathDisplay) And (Mid(TextToCopy, 1, 1) = vbCr) Then
            TextToCopy = Mid(TextToCopy, 2)
        End If
        TargetW.ResultsBox.AppendText(TextToCopy)
        TargetW.Text &= " Finished"
        TensorMessage($"Finished in {ElapsedTime.Elapsed.Minutes:00}:{ElapsedTime.Elapsed.Seconds:00}")
    End Function
    Function LatexOn(WorkDoc As Word.Document) As Boolean
        'Tests if Word is in latex mode by creating equation, copying
        'and pasting it and seeing if result is latex. Done in fleetingly visible WorkDoc
        Dim Term As Word.OMathFunction, MyRange As Word.Range
        Dim objRange As Word.Range

        objRange = WorkDoc.Range()
        objRange.Text = "x"
        objRange.OMaths.Add(objRange)
        MyRange = objRange.OMaths(1).Range
        MyRange.End = MyRange.Start
        Term = objRange.OMaths(1).Functions.Add(MyRange, Word.WdOMathFunctionType.wdOMathFunctionFrac)
        Term.Frac.Num.Range.Text = "1"
        Term.Frac.Den.Range.Text = "2"
        objRange.OMaths(1).Range.Copy()
        objRange.OMaths(1).Range.Delete()
        objRange.PasteAndFormat(Word.WdRecoveryType.wdFormatPlainText)
        objRange.End -= 1
        If objRange.Text = "\frac{1}{2}x" Then LatexOn = True Else LatexOn = False
        objRange.Delete()
    End Function
    Function CreateOutputWindow(SourceDoc As Word.Document, ByRef TargetDoc As Word.Document) As Boolean
        'return false if aborted - very similar to ProgressStart
        Dim Answer As Integer, Question As String
        Dim ActiveDocument As Word.Document = Globals.EquationAcc.Application.ActiveDocument

        CreateOutputWindow = False
        If (ActiveDocument.Windows.Count > 1) Then
            Call MsgBox("Please close all but one window on document.", vbOKOnly)
            Exit Function
        End If

        If SourceDoc.Saved = False Then
            Question = "Do you want to save " + gMainDocName + " before starting?"
            Question = Question + vbCr + "Yes = Save before starting (safe)"
            Question = Question + vbCr + "No = Continue without saving (dangerous)"
            Question = Question + vbCr + "Cancel = Do nothing (safe but dull)"
            Answer = MsgBox(Question, vbYesNoCancel)
            If Answer = vbYes Then
                SourceDoc.Save()
            End If
            If Answer = vbCancel Then
                Exit Function
            End If
        End If

        CreateOutputWindow = True
        TargetDoc = Globals.EquationAcc.Application.Documents.Add(, ,
            Word.WdNewDocumentType.wdNewBlankDocument, False)       'False -> True for debugging
        TargetDoc.ActiveWindow.WindowState = Word.WdWindowState.wdWindowStateMinimize

    End Function
    Function ProcessTable(ByRef nEquation As Integer, ByRef EquationRangeEnd As Long, SourceDoc As Word.Document,
                      WorkDoc As Word.Document, TargetW As OutputWindow) As Boolean
        'we have found an equation in a table. Assume that it is a table of equations with equation numbers in last column
        'A table in latex begins with \begin{align}
        'contains equations separated by & which is the alignment code and \\ which signals a new line
        'and ends with \end{align}
        'The last column of each row usually contains an equation number in text.

        Dim nRows As Integer
        Dim nColumn As Integer, nRow As Integer
        Dim TheTable As Word.Table, TheCell As Word.Cell, TheRow As Word.Row
        Dim Equation As Word.OMath, EquationRange As Word.Range
        Dim EqNumber As String
        Dim HaveEquation As Boolean

        Equation = SourceDoc.OMaths(nEquation)
        EquationRange = Equation.Range
        TheTable = EquationRange.Tables(1)
        nRows = TheTable.Rows.Count

        TargetW.ResultsBox.AppendText("\begin{align}")  'could add vbCr as well
        For nRow = 1 To nRows
            nColumn = 1
            TheRow = TheTable.Rows(nRow)
            While nColumn <= TheRow.Cells.Count
                'Does cell contain text or equation?
                TheCell = TheTable.Cell(nRow, nColumn)
                HaveEquation = False        'Need this for case when equation table is last thing in document
                If nEquation <= SourceDoc.OMaths.Count Then
                    Equation = SourceDoc.OMaths(nEquation)
                    EquationRange = Equation.Range
                    If (Equation.Range.Start >= TheCell.Range.Start) And (Equation.Range.End <= TheCell.Range.End) Then
                        'equation is in cell
                        HaveEquation = True
                        If Not PasteEquation(EquationRange, WorkDoc, TargetW) Then Return False
                        nEquation += 1
                    End If
                End If
                If Not HaveEquation Then
                    'must be text or blank in cell. Is it last column?
                    If nColumn = TheRow.Cells.Count Then
                        EqNumber = Left(TheCell.Range.Text, Len(TheCell.Range.Text) - 2)  'remove odd charachters
                        While nColumn < TheTable.Columns.Count
                            TargetW.ResultsBox.AppendText("&")       'align with last cells in busiest row
                            nColumn += 1
                        End While
                        TargetW.ResultsBox.AppendText(BraPadding & EqNumber)
                    Else
                        'text
                        TargetW.ResultsBox.AppendText(TheCell.Range.Text)
                    End If
                End If
                nColumn += 1
                If nColumn <= TheRow.Cells.Count Then TargetW.ResultsBox.AppendText("&")
            End While
            TargetW.ResultsBox.AppendText("\nonumber")           'suppress automatic numbering (Physics Forums)
            If nRow < nRows Then TargetW.ResultsBox.AppendText("\\") 'could add vbCr after that
        Next
        TargetW.ResultsBox.AppendText("\end{align}")
        EquationRangeEnd = TheTable.Range.End
        nEquation -= 1
        Return True
    End Function
    Function PasteEquation(EquationRange As Word.Range, WorkDoc As Word.Document, TargetW As OutputWindow) As Boolean
        'copy / paste equation into target & do a bit of checking. Return false if error found
        Dim ConvertedString As String
        Dim Selection As Word.Selection = Globals.EquationAcc.Application.ActiveDocument.ActiveWindow.Selection
        'Check for things that we can't handle such as Delta from basic math which should be Greek delta

        WorkDoc.Range.Delete()
        EquationRange.Copy()
        WorkDoc.Range.PasteAndFormat(Word.WdRecoveryType.wdFormatPlainText)
        ConvertedString = WorkDoc.Range.Text
        ConvertedString = Mid(ConvertedString, 1, Len(ConvertedString) - 1) 'remove trailing &VbCr
        'replace problem latex
        ConvertedString = Replace(ConvertedString, "\sfrac", "\frac")       'replaces all occurences
        ConvertedString = Replace(ConvertedString, "\mathbit", "\mathbf")   'bold italic font!
        If InStr(ConvertedString, "∆") > 0 Then
            HighlightError(EquationRange, "∆ from Basic Math used." & vbCr &
                        "Word cannot create Latex." & vbCr &
                        "Change to ∆ from Greek Letters.")
            Return False
        End If
        TargetW.ResultsBox.AppendText(ConvertedString)
        Return True
    End Function
End Module
Module RenumberFunctions
    Async Function RenumberEquationsCodeAsync() As Task
        'renumber all equations and references to them
        Dim ExpectedDuration As Single      'in seconds
        Dim RefCount As Long
        Dim Selection As Word.Selection
        Dim ProgressW As New OutputWindow()
        Dim StartOK As Boolean


        ElapsedTime.Restart()
        ProgressW.Show()
        StartOK = Await ProgressStartAsync(ProgressW)
        If Not (StartOK) Then Exit Function
        Selection = Globals.EquationAcc.Application.ActiveDocument.ActiveWindow.Selection
        Selection.HomeKey(Word.WdUnits.wdStory)

        'trial and error on my computer
        ExpectedDuration = Globals.EquationAcc.Application.ActiveDocument.Paragraphs.Count * 0.08

        Await ProgressContAsync(ProgressW, "Estimated duration " + Format(ExpectedDuration, "#") + " seconds in 6 steps.")
        Await ProgressContAsync(ProgressW, "CLOSE THIS WINDOW TO ABORT." & vbCr & "BUT " + gMainDocName + " WILL BE IN A MESS. ")
        gRefErrors = ""
        Await FindOldEquationNumbersAsync(ProgressW)
        Await ProgressContAsync(ProgressW, "Found all equations ")
        Await FindOldEquationReferencesAsync(ProgressW)
        Await ProgressContAsync(ProgressW, "Found all equation references ")
        SearchReplace(MagicChar2, "(")
        If gRefErrors <> "" Then
            Await ProgressContAsync(ProgressW, "Errors found. Please rectify." + vbCr + gRefErrors)
        Else
            RefCount = Await ChangeEquationNumbersAsync(ProgressW)
            Await ProgressContAsync(ProgressW, "Finished OK. Changed " + CStr(UBound(gEquations)) + " equations numbers and " + CStr(RefCount) + " references")
        End If

        SearchReplace(MagicChar1, "(")
        SearchReplace(MagicChar2, "(")
        Selection.HomeKey(Word.WdUnits.wdStory)
        ClearFindParameters()
        If gRefErrors <> "" Then
            TensorError("Errors found. See box." + vbCr + "Please rectify.")
        Else
            TensorMessage($"Finished in {ElapsedTime.Elapsed.Minutes:00}:{ElapsedTime.Elapsed.Seconds:00}")
        End If
    End Function
    Async Function FindOldEquationNumbersAsync(ProgressW As OutputWindow) As Task
        'find all (*) that are in last column of table and add them to gEquations
        'Mark in text with MagicChar1
        Dim MaxEqnr As Single, NewEqNr As String
        MaxEqnr = 0
        Dim Selection As Word.Selection

        Selection = Globals.EquationAcc.Application.ActiveDocument.ActiveWindow.Selection
        Selection.HomeKey(Word.WdUnits.wdStory)
        Do While FindWildStop("\(*\)")
            'Selection.Information(wdWithInTable) is False on and after equation (3.3) in TMI.docx
            'so we must resort to other methods
            Await ProgressContSmallAsync(ProgressW)
            NewEqNr = Selection.Text
            Selection.MoveRight(Word.WdUnits.wdCharacter, 2)

            If Selection.IsEndOfRowMark Then
                'it is end of row. So we have found an equation number
                MaxEqnr = MaxEqnr + 1
                ReDim Preserve gEquations(0 To MaxEqnr)
                gEquations(MaxEqnr).Num = NewEqNr
                gEquations(MaxEqnr).Refs = 0
                'get back to the equation number
                Selection.Find.ClearFormatting()

                With Selection.Find
                    .Text = NewEqNr
                    .Forward = False
                    .Wrap = Word.WdFindWrap.wdFindAsk
                    .Format = False
                    .MatchCase = False
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                Selection.Find.Execute()
                'replace ( by magicchar1, so equation numbers are different ftom references
                Selection.MoveLeft(Word.WdUnits.wdCharacter, 1)
                Selection.Delete(Word.WdUnits.wdCharacter, 1)
                Selection.TypeText(MagicChar1)
            End If
        Loop
    End Function
    Async Function FindOldEquationReferencesAsync(ProgressW As OutputWindow) As Task
        'find all (*). If in gEquations, mark them with MagicChar2
        'look out for referenced duplicates in gEquations and 
        'potential references that will reference a new equation number
        Dim FoundRef As String
        Dim sFoundNumber As String
        Dim iFoundNumber As Integer
        Dim FoundRefCount As Single
        Dim NewEqNr As Single
        Dim Selection As Word.Selection

        Selection = Globals.EquationAcc.Application.ActiveDocument.ActiveWindow.Selection
        Selection.HomeKey(Word.WdUnits.wdStory)
        Do While FindWildStop("\(*\)")
            While Selection.OMaths.Count > 0
                'In an equation, can occur in symmetrisation operator or bad use of brackets
                Selection.Range.Start = Selection.OMaths(1).Range.End
                Selection.Range.End = Selection.Range.Start
                If Not FindWildStop("\(*\)") Then Exit Function
            End While
            Await ProgressContSmallAsync(ProgressW)
            FoundRef = Selection.Text
            FoundRefCount = 0
            NewEqNr = 1
            Do While NewEqNr <= UBound(gEquations)
                If gEquations(NewEqNr).Num = FoundRef Then
                    FoundRefCount += 1
                    gEquations(NewEqNr).Refs = gEquations(NewEqNr).Refs + 1
                End If
                NewEqNr += 1
            Loop
            If FoundRefCount = 0 Then
                'it was apparently not an equation, but if its a number which will refer to one of the
                'new equation numbers....
                sFoundNumber = Mid(FoundRef, 2, Len(FoundRef) - 2)
                iFoundNumber = MyCint(sFoundNumber) '********************
                If (InStr(sFoundNumber, ".") > 0) Or (InStr(sFoundNumber, ",") > 0) Then
                    iFoundNumber = 0
                End If

                If (iFoundNumber > 0) And (iFoundNumber <= UBound(gEquations)) And (Selection.OMaths.Count = 0) Then
                    'PROBLEM
                    'Selection.OMaths.Count = 0 means in an equation. Not a problem.
                    gRefErrors &= "Equation " + FoundRef + " is referenced. It does not exist." + vbCr
                End If
                Selection.MoveLeft(Word.WdUnits.wdCharacter, 1)
                Selection.MoveRight(Word.WdUnits.wdCharacter, 1)
            End If
            If FoundRefCount >= 1 Then
                'it was an equation reference (*) selected by FindWildStop("\(*\)")
                'following sequence should move selection back 1, delete (, insert MagicChar2
                'but sometimes, if the previous FindWildStop was in an equation it deletes too much
                'wiggling the cursor solves the problem!!!! Fucking VBA or Word?
                Selection.MoveLeft(Word.WdUnits.wdCharacter, 1)
                Selection.MoveRight(Word.WdUnits.wdCharacter, 1)
                Selection.MoveLeft(Word.WdUnits.wdCharacter, 1)
                Selection.Delete(Word.WdUnits.wdCharacter, 1)
                Selection.TypeText(MagicChar2)
            End If
            If FoundRefCount > 1 Then
                'the equation reference we found references more thatn one equation. PROBLEM
                'Note problem and carry on to see if there are more
                gRefErrors = gRefErrors + "Equation " + FoundRef + " is defined " + Format(FoundRefCount, "#") + " times and referenced one or more times." + vbCr
            End If
        Loop
    End Function
    Function MyCint(Input As String) As Integer
        'Need a little function here because cannot use on error in Async function
        'Apparently:
        '1,0 is converted to 10, 4.31 is converted to 4, other characters seem to give an error
        'this code may be locale dependent!
        On Error GoTo Result0
        Return CInt(Input)
Result0:
        Return 0
    End Function
    Async Function ChangeEquationNumbersAsync(ProgressW As OutputWindow) As Task(Of Long)
        Dim EqNr As Integer          'Peculiar variable type to chose! Probably meant integer
        Dim NewEqNr As String
        Dim RefCount As Long
        Dim Selection As Word.Selection

        Selection = Globals.EquationAcc.Application.ActiveDocument.ActiveWindow.Selection
        RefCount = 0
        Selection.HomeKey(Word.WdUnits.wdStory)
        SearchReplace(MagicChar2, "(")
        Selection.HomeKey(Word.WdUnits.wdStory)
        EqNr = 1
        Do While FindWildStop(MagicChar1 + "*\)")
            Await ProgressContSmallAsync(ProgressW)
            NewEqNr = (MagicChar2 + Format(EqNr, "#") + ")")
            Selection.TypeText(NewEqNr)
            If gEquations(EqNr).Refs > 0 Then
                RefCount += gEquations(EqNr).Refs
                Selection.HomeKey(Word.WdUnits.wdStory)
                SearchReplace(gEquations(EqNr).Num, NewEqNr)
                Selection.HomeKey(Word.WdUnits.wdStory)
            End If
            EqNr += 1
        Loop
        Selection.HomeKey(Word.WdUnits.wdStory)
        SearchReplace(MagicChar2, "(")
        Return RefCount
    End Function
    Async Function ProgressStartAsync(ProgressW As OutputWindow) As Task(Of Boolean)
        'return false if aborted very similar to CreateOutputWindow
        Dim Answer, Question As String
        Dim ActiveDocument As Word.Document

        ActiveDocument = Globals.EquationAcc.Application.ActiveDocument
        If (ActiveDocument.Windows.Count > 1) Then
            Call MsgBox("Please close all but one window on document.", vbOKOnly)
            Return False
        End If

        gMainDocName = ActiveDocument.Name
        If ActiveDocument.Saved = False Then
            Question = "Do you want to save " + gMainDocName + " before starting?"
            Question = Question + vbCr + "Yes = Save before starting"
            Question = Question + vbCr + "No = Continue without saving (dangerous)"
            Question = Question + vbCr + "Cancel = Abort program"
            Answer = MsgBox(Question, vbYesNoCancel)
            If Answer = vbYes Then
                ActiveDocument.Save()
            End If
            If Answer = vbCancel Then
                Return False
            End If
        Else
            Answer = vbYes
        End If

        gProgressStep = 0
        If Answer = vbNo Then
            Await ProgressContAsync(ProgressW, "Progress processing (unsaved) " + gMainDocName)
        Else
            Await ProgressContAsync(ProgressW, "Progress processing (saved) " + gMainDocName)
        End If
        Return True
    End Function
    Async Function ProgressContSmallAsync(ProgressW As OutputWindow) As Task
        ProgressW.Text = $"Progress {ElapsedTime.Elapsed.Minutes:00}:{ElapsedTime.Elapsed.Seconds:00}"
        ProgressW.ResultsBox.AppendText(".")
        Await Task.Delay(1)
    End Function
    Async Function ProgressContAsync(ProgressW As OutputWindow, Message As String) As Task
        Dim sStep As String

        ProgressW.Text = $"Progress {ElapsedTime.Elapsed.Minutes:00}:{ElapsedTime.Elapsed.Seconds:00}"
        gProgressStep += 1
        sStep = Format(gProgressStep, "#")
        If gProgressStep < 10 Then sStep = " " + sStep

        ProgressW.ResultsBox.AppendText(vbCr & sStep & ":  " & Message)
        Await Task.Delay(1)
    End Function
End Module
Module FindFunctions
    ' Find Functions **************************************************
    Function FindWildStop(FindText) As Boolean
        'returns true if FindText is found (in selection). Searches forward only. FindText is selected
        Dim Selection As Word.Selection
        Selection = Globals.EquationAcc.Application.ActiveDocument.ActiveWindow.Selection
        Selection.Find.ClearFormatting()

        With Selection.Find
            .Text = FindText
            .Replacement.Text = ""
            .Forward = True
            .Wrap = Word.WdFindWrap.wdFindStop
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = True
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        FindWildStop = Selection.Find.Execute
    End Function
    Sub SearchReplace(FindChar As String, ReplaceChar As String)
        ' Vanilla search replace on whole document
        Dim Selection As Word.Selection
        Selection = Globals.EquationAcc.Application.ActiveDocument.ActiveWindow.Selection
        Selection.Find.ClearFormatting()
        Selection.Find.Replacement.ClearFormatting()

        With Selection.Find
            .Text = FindChar
            .Replacement.Text = ReplaceChar
            .Forward = True
            .Wrap = Word.WdFindWrap.wdFindContinue
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute(Replace:=Word.WdReplace.wdReplaceAll)
    End Sub
    Sub ClearFindParameters()
        Dim Selection As Word.Selection
        Selection = Globals.EquationAcc.Application.ActiveDocument.ActiveWindow.Selection
        Selection.Find.ClearFormatting()
        Selection.Find.Replacement.ClearFormatting()

        With Selection.Find
            .Text = ""
            .Replacement.Text = ""
            .Forward = True
            .Wrap = Word.WdFindWrap.wdFindStop
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
    End Sub
    Function FindWildStopUp(FindText As String) As Boolean
        Dim Selection As Word.Selection
        Selection = Globals.EquationAcc.Application.ActiveDocument.ActiveWindow.Selection
        Selection.Find.ClearFormatting()

        With Selection.Find
            .Text = FindText
            .Forward = False
            .Wrap = Word.WdFindWrap.wdFindStop
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchAllWordForms = False
            .MatchSoundsLike = False
            .MatchWildcards = True
        End With
        FindWildStopUp = Selection.Find.Execute
    End Function
End Module