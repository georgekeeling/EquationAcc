'Add in to help with equation editing for Office 365 (2019)
'By George Keeling and on my blog at www.general-relativity.net
'Possibly documented at https://www.general-relativity.net/search/label/Tools
'Feel free to use, copy, modify and give away but not for commerce. Please acknowledge me
'
'Entry points
'	PickUpCoordinates - Read coordinates from equation
'	PickUpEitherMetric - read metric Or inverse metric from a selected matrix Or 
'       read metric from line element equation Like ds²=dr²+r²dθ². Inverse Of diagonal metric is calculated. Components are found from line element equation.
'	WriteMetrics - write metrics as matrix at insertion point in equation Or over selected part of equation
'	ClearMetrics - duh
'	WriteChristoffelSymbolsAsync -Write out all Christoffel symbols in three column equations. Need all metrics And coordinates to be loaded

Imports System.Threading.Tasks
Module Christoffel
    'Code concerning Christoffel symbols and metric (Christoffel group on ribbon)
    'Coordinates and metrics always from 1 to 4, which may be shown as 0-3 in GR
    Public gNcoords As Integer                  'nr of coords = 2,3,4
    'Next arrays were all (1 To 4) => indexed from 1 to 4. Have become (4) which is same as (0 To 4) => indexed from 1 to 4.
    Public gCoordinates(4) As String       'coordinates in use, start fro 0 with 4, otherwise from 1
    Public gMetric(4, 4) As String    'The  metric stored in 'linear' format
    Public gNmetricDimension As Integer         'Metric dimension 2,3,4 should be same as gNcoords.
    'Only use gMetric(0,x) if gNmetricDimension=4
    Public gInvMetric(4, 4) As String 'inverse metric  stored in 'linear' format
    Public gNInvmetricDimension As Integer      ' inverse Metric dimension. Should be same as others.

    Public gCoordinatesString As String         ' messages diplayed on ribon
    Public gMetricNCoordsString As String
    Public gInvMetricNCoordsString As String

    Public Sub PickUpCoordinates()
        'Read coordinates from equation. We just expect 2,3 or 4 charachters separated by commas
        'possibly inside brackets
        Dim InnerEquation As Word.OMath
        Dim MathTerm As Word.OMathFunction
        Dim iX As Integer, MyChar As String, iXcoords As Integer
        Dim Selection As Word.Selection

        Selection = Globals.EquationAcc.Application.ActiveDocument.ActiveWindow.Selection
        If Selection.OMaths.Count <> 1 Then GoTo BadEquation

        MathTerm = Selection.OMaths(1).Functions(1)
        If MathTerm.Type = Word.WdOMathFunctionType.wdOMathFunctionDelim Then
            'get into brackets
            InnerEquation = MathTerm.Delim.E.Item(1)
            If MathTerm.Delim.E.Count <> 1 Then GoTo BadEquation
            MathTerm = InnerEquation.Functions(1)     'was MathTerm.Delim.E(1).Functions(1)
        End If
        If MathTerm.Type <> Word.WdOMathFunctionType.wdOMathFunctionText Then GoTo BadEquation 'wdOMathFunctionNormalText or wdOMathFunctionText

        gCoordinatesString = ""
        For iXcoords = 1 To 4
            gCoordinates(iXcoords) = ""
        Next

        CleanString(MathTerm.Range, gCoordinatesString)
        iXcoords = 1
        For iX = 1 To Len(gCoordinatesString)
            MyChar = Mid(gCoordinatesString, iX, 1)
            If (MyChar <> ",") And (MyChar <> ")") And (MyChar <> "(") And (MyChar <> "{") And (MyChar <> "}") Then
                If iXcoords > UBound(gCoordinates) Then GoTo BadEquation
                gCoordinates(iXcoords) = MyChar
                iXcoords += 1
            End If
        Next

        gNcoords = iXcoords - 1
        If gNcoords < 2 Then GoTo BadEquation

        gCoordinatesString = ""
        For iXcoords = 1 To gNcoords
            gCoordinatesString &= gCoordinates(iXcoords)
            If iXcoords < 4 Then
                If gCoordinates(iXcoords + 1) <> "" Then
                    gCoordinatesString &= ","
                End If
            End If
        Next
        Globals.EquationAcc.MyRibbon.InvalidateControl("Coords")
        TensorMessage("Picked up coordinates " & gCoordinatesString)
        Exit Sub

BadEquation:
        gCoordinatesString = ""
        For iXcoords = 1 To 4
            gCoordinates(iXcoords) = ""
        Next
        TensorError("Cursor must be in equation" & vbCrLf &
                    "which must have 2-4 coordinates" & vbCrLf &
                    "and be of form x,y,z or (x,y)")
        gCoordinatesString = ""
        Globals.EquationAcc.MyRibbon.InvalidateControl("Coords")
    End Sub
    Sub PickUpEitherMetric(ByRef Dimension As Integer,
                           ByRef Metric(,) As String, ByRef ToDisplay As String, Which As String,
                           ByRef OMetric(,) As String, ByRef OToDisplay As String, OWhich As String)
        'read metric or inverse metric from a selected matrix OR 
        'read Metric from line element equation Like ds²=dr²+r²dθ².
        'Inverse of diagonal Is calculated.
        'Components are found from line element equation.

        'Dimension, Metric() for metric we want
        'ToDisplay is global for nxn result
        'Which = "Metric" or "Inv.metric"  Must be Id of control in ribbon
        'O parameters are same but of the opposite metric.
        Dim MathTerm As Word.OMathFunction
        Dim Row As Integer, Col As Integer
        Dim Message As String
        Dim Component As String, OppositeComponent As String
        Dim Selection As Word.Selection
        Dim Diagonal As Boolean                 'Iftrue we will calculate inverse metric.

        Selection = Globals.EquationAcc.Application.ActiveDocument.ActiveWindow.Selection
        If Selection.OMaths.Count <> 1 Then TensorError("Matrix should be selected.") : Exit Sub
        MathTerm = Selection.OMaths(1).Functions(1)

        If MathTerm.Type <> Word.WdOMathFunctionType.wdOMathFunctionMat Then
            If Metric Is gInvMetric Then
                TensorError("Matrix should be selected.")
                Return
            Else
                'check for metric in form ds²=dr²+r²dθ²
                Selection.Collapse(Word.WdCollapseDirection.wdCollapseStart) 'ensure we work on whole equation
                Message = LoadMetricNotMatrix0(Selection)
                If Message <> "" Then
                    TensorError(Message)
                    Return
                Else
                    Dimension = gNcoords
                    Diagonal = False
                    GoTo FinishUp
                End If
            End If
        End If
        If (MathTerm.Mat.Rows.Count < 2) Or (MathTerm.Mat.Rows.Count > 4) _
        Or (MathTerm.Mat.Cols.Count < 2) Or (MathTerm.Mat.Cols.Count > 4) _
        Or (MathTerm.Mat.Cols.Count <> MathTerm.Mat.Rows.Count) Then
            TensorError(Which & " matrix must be 2×2, 3×3 or 4×4")
            Return
        End If

        Dimension = MathTerm.Mat.Cols.Count

        For Row = 1 To Dimension
            For Col = 1 To Dimension
                Component = GetCleanLinear(MathTerm.Mat.Cell(Row, Col))
                'often have something like Component =""1" ". Clean it up
                If Row > Col Then
                    'in lower left part of metric. Use symmetry
                    OppositeComponent = GetCleanLinear(MathTerm.Mat.Cell(Col, Row))
                    If Component = "" Then
                        Component = OppositeComponent
                        SetLinear(MathTerm.Mat.Cell(Row, Col), Component)
                    Else
                        If Component <> OppositeComponent Then
                            TensorError(Which & " must be symmetric.")
                            Dimension = 0
                            Metric(1, 1) = ""
                            Exit Sub
                        End If
                    End If
                End If
                Metric(Row, Col) = Component
            Next
        Next

FinishUp:
        Diagonal = True
        For Row = 1 To Dimension
            For Col = 1 To Dimension
                If (Row <> Col) And (Metric(Row, Col) <> "0") Then Diagonal = False
            Next
        Next

        Message = Dimension & "×" & Dimension
        ToDisplay = Message
        Message = Message & " " & Which & " saved."
        If Diagonal Then
            gNmetricDimension = Dimension
            gNInvmetricDimension = Dimension
            If Not CalculateDiagonalInverse(Metric, OMetric) Then
                TensorError("Diagonal component in " & vbCr &
                            "diagonal metric cannot be 0.")
                ClearMetrics2()
                Return
            End If
            Message &= vbCr & OWhich & " calculated"
            OToDisplay = Dimension & "×" & Dimension
            Globals.EquationAcc.MyRibbon.InvalidateControl(OWhich)
        End If

        TensorMessage(Message)
        Globals.EquationAcc.MyRibbon.InvalidateControl(Which)
    End Sub
    Function LoadMetricNotMatrix0(Selection As Word.Selection) As String
        'wrapper for LoadMetricNotMatrix.
        'Puts equation at Selection into Tempdoc where it max be destroyed at leisure
        Dim OriginalEq As String
        Dim Message As String
        Dim TempDoc As Word.Document
        Dim objRange As Word.Range
        Dim Position As Long
        Position = Selection.Start      'Otherwise will return to beginninf of doc. Duh

        OriginalEq = GetLinear(Selection.OMaths(1))
        TempDoc = Globals.EquationAcc.Application.Documents.Add(, ,
            Word.WdNewDocumentType.wdNewBlankDocument, False)       'False -> True for debugging
        TempDoc.ActiveWindow.WindowState = Word.WdWindowState.wdWindowStateMinimize
        objRange = TempDoc.Range()
        objRange.Text = " "
        objRange.OMaths.Add(objRange)               'creates equation containing blank
        'objRange.OMaths(1).Range.Italic = 0         'will preserve italic / non-italic
        objRange.OMaths(1).Range.Text = OriginalEq  'inserts the text e.g 𝑥^2 
        objRange.OMaths(1).BuildUp()        'makes the equation pretty (but size 10 not 12). e.g. 𝑥²
        objRange.OMaths(1).Range.Bold = 0   'turn off bold. Italic is on everywhere

        Message = LoadMetricNotMatrix(objRange.OMaths(1))
        TempDoc.Close(Word.WdSaveOptions.wdDoNotSaveChanges)
        Selection.Start = Position

        Return Message
    End Function
    Function LoadMetricNotMatrix(Equation As Word.OMath) As String
        'check for metric in form ds²=dr²+r²dθ² (therefore not matrix)
        'ds²= part is optional. Everything to left of = is igonred
        'returns error message or "" if all OK
        Dim MathTerm As Word.OMathFunction, iMathTerm As Integer, iMathTerm0 As Integer, iMathTermBodge As Integer
        Dim SquareTerm As Word.OMathScrSup
        Dim TestString As String, FinalError As String
        Dim RHSEquation As String
        Dim Metric(4, 4) As String, Coords As String = ""    'proposed new metric and coordinates
        Dim Coordinate1 As String, Coordinate2 As String 'coordinates which we build up
        Dim iCoord1 As Integer, iCoord2 As Integer      'indexes into Coords and Metric
        Dim CarryForward As String, CarryBack As String, iDpos As Integer

        iMathTerm = 1
        While iMathTerm <= Equation.Functions.Count
            MathTerm = Equation.Functions(iMathTerm)
            If TextTerm(MathTerm) Then
                If InStr(MathTerm.Range.Text, "=") > 0 Then
                    Exit While
                End If
            End If
            iMathTerm += 1
        End While
DoRHS:
        CarryForward = ""
        If iMathTerm > Equation.Functions.Count Then
            iMathTerm = 0
        Else
            CleanString(MathTerm.Range, TestString)
            CarryForward = Mid(TestString, InStr(TestString, "=") + 1)
        End If
        'we are now on the first part of the equation that is of interest
        'remove stuff before =, including =
        While iMathTerm > 0
            Equation.Functions(1).Range.Text = ""       'Equation.Functions(1).Remove() crashes here and elsewhere
            iMathTerm -= 1
        End While
        InsertCarryText(Equation.Range, Equation.Range.Start, CarryForward)

        RHSEquation = GetLinear(Equation)         'we will need this for each metric component. It might get smaller
        iMathTerm = 1
        While iMathTerm <= Equation.Functions.Count
            MathTerm = Equation.Functions(iMathTerm)
            Coordinate1 = ""
            CarryForward = ""
            CarryBack = ""
            iMathTerm0 = iMathTerm
            iMathTermBodge = 0
            If TextTerm(MathTerm) Then
                'Looking for dxdy,dx² 
                CleanString(MathTerm.Range, TestString)
                iDpos = InStr(1, TestString, "d")
                If iDpos <> 0 Then 'its ...dxdy... or ...dx²
                    CarryBack = Mid(TestString, 1, iDpos - 1)
                    TestString = Mid(TestString, iDpos)
                    If (Len(TestString) >= 4) And (Mid(TestString, 3, 1) = "d") Then 'its dxdy...
                        Coordinate1 = Mid(TestString, 2, 1)
                        Coordinate2 = Mid(TestString, 4, 1)
                        If Len(TestString) > 4 Then CarryForward = Mid(TestString, 5)
                    End If
                    If Len(TestString) = 1 Then 'check for dx²
                        If iMathTerm = Equation.Functions.Count Then Return "d must precede component"
                        iMathTerm += 1
                        MathTerm = Equation.Functions(iMathTerm)
                        If MathTerm.Type = Word.WdOMathFunctionType.wdOMathFunctionScrSup Then
                            SquareTerm = MathTerm.ScrSup
                            If SquareTerm.Sup.Range.Text <> "2" Then Return "𝑑 must precede component²"
                            CleanString(SquareTerm.E.Range, TestString)
                            If Len(TestString) > 1 Then Return "square term too complex"
                            Coordinate1 = TestString
                            Coordinate2 = TestString
                            iMathTermBodge = 1
                        End If
                    End If
                    If Coordinate1 = "" Then Return "cannot recognise stuff after d"
                End If
            End If
            If MathTerm.Type = Word.WdOMathFunctionType.wdOMathFunctionScrSup Then
                'check for dx² where dx is the squared term
                SquareTerm = MathTerm.ScrSup
                If SquareTerm.Sup.Range.Text = "2" Then
                    CleanString(SquareTerm.E.Range, TestString)
                    If (Len(TestString) = 2) And (Mid(TestString, 1, 1) = "d") Then
                        Coordinate1 = Mid(TestString, 2, 1)
                        Coordinate2 = Coordinate1
                    End If
                End If
            End If
            If Coordinate1 <> "" Then
                'we have found two coordinates which may be the same e.g. dxdy or dx²
                'First record the coordinate if it is new and get its index
                'Then delete everything from the beginning of dxdy or dx² to the end.
                'That leaves the metric component which we save
                'Then reinstate equation (from RHSEquation) and delete dxdy or dx² and everything before it
                For iCoord1 = 1 To Len(Coords)
                    If Coordinate1 = Mid(Coords, iCoord1, 1) Then Exit For
                Next
                If iCoord1 > Len(Coords) Then Coords &= Coordinate1
                For iCoord2 = 1 To Len(Coords)
                    If Coordinate2 = Mid(Coords, iCoord2, 1) Then Exit For
                Next
                If iCoord2 > Len(Coords) Then Coords &= Coordinate2

                While Equation.Functions.Count >= iMathTerm0
                    'If Equation.Functions.Count = 1 Then Exit While  ' can do no more
                    If Equation.Functions.Count = 1 Then
                        'iMathTerm0 was 1
                        Equation.Functions(1).Range.Text = " "
                        Exit While
                    End If
                    'following decreases Equation.Functions.Count by 1, unless it is 1 already
                    'condition above is fix to prevent endless loop
                    Equation.Functions(Equation.Functions.Count).Range.Text = ""
                    'Note: Must be done from last function to first otherwise text cells can get merged
                End While
                InsertCarryText(Equation.Range, Equation.Range.End, CarryBack)
                TestString = GetLinear(Equation)
                While (Mid(TestString, 1, 1) = " ") Or (Mid(TestString, 1, 1) = "+")
                    TestString = Mid(TestString, 2)
                End While
                If TestString = "" Then TestString = "1"
                If Len(TestString) = 1 And IsStrMinus(TestString) Then TestString &= "1"
                Metric(iCoord1, iCoord2) = TestString
                If Mid(Metric(iCoord1, iCoord2), 1, 1) = "+" Then Metric(iCoord1, iCoord2) = Mid(Metric(iCoord1, iCoord2), 2)
                SetLinear(Equation, RHSEquation)
                iMathTerm0 += iMathTermBodge
                While iMathTerm0 > 0
                    If Equation.Functions.Count = 1 Then GoTo ErrorNone  'We are done. Avoid deleting equation entirely!
                    Equation.Functions(1).Range.Text = ""
                    iMathTerm0 -= 1
                End While
                InsertCarryText(Equation.Range, Equation.Range.Start, CarryForward)
                RHSEquation = GetLinear(Equation)
                iMathTerm = 0
            End If
            iMathTerm += 1
        End While

ErrorNone:
        'set up results 
        FinalError = FinalCheck(Metric, Coords)
        If FinalError <> "" Then Return FinalError
        gNcoords = Len(Coords)
        gNmetricDimension = gNcoords
        gMetricNCoordsString = gNcoords & "×" & gNcoords
        gCoordinatesString = ""
        For iCoord1 = 1 To gNcoords
            gCoordinates(iCoord1) = Mid(Coords, iCoord1, 1)
            gCoordinatesString &= gCoordinates(iCoord1)
            If iCoord1 < gNcoords Then gCoordinatesString &= ","
            For iCoord2 = 1 To gNcoords
                If Metric(iCoord1, iCoord2) = "" Then
                    gMetric(iCoord1, iCoord2) = "0"
                Else
                    gMetric(iCoord1, iCoord2) = Metric(iCoord1, iCoord2)
                End If
            Next
        Next
        Globals.EquationAcc.MyRibbon.InvalidateControl("Coords")
        Return ""
    End Function
    Sub InsertCarryText(InsertIn As Word.Range, Pos As Long, CarryText As String)
        Dim InsertPoint As Word.Range
        If CarryText = "" Then Return
        InsertPoint = InsertIn
        InsertPoint.Start = Pos
        InsertPoint.End = Pos
        InsertPoint.Text = CarryText
    End Sub
    Function TextTerm(MathTerm As Word.OMathFunction) As Boolean
        'see https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.word.wdomathfunctiontype?view=word-pia
        'I have no idea why there are three of these. In debug I have only seen wdOMathFunctionText. But who knows?
        'And I am tired of that long condition
        If (MathTerm.Type = Word.WdOMathFunctionType.wdOMathFunctionText) Or
                (MathTerm.Type = Word.WdOMathFunctionType.wdOMathFunctionNormalText) Or
                (MathTerm.Type = Word.WdOMathFunctionType.wdOMathFunctionLiteralText) Then
            Return True
        End If
        Return False
    End Function
    Function FinalCheck(ByRef Metric(,) As String, Coords As String) As String
        'if metric is diagonal then all diagonal components must be non zero
        'if metric is not diagonal then it must be symmetric. If one off diagonal term is absent,
        'use the other to fill it in  and divide both by 2.
        Dim Diagonal As Boolean = True
        Dim iCoord1 As Integer, iCoord2 As Integer
        Dim Dimension As Integer = Len(Coords)

        If Dimension < 2 Then
            Return "Equation does not appear" & vbCr & "to define metric."
        End If

        For iCoord1 = 1 To Dimension
            For iCoord2 = 1 To Dimension
                If iCoord1 <> iCoord2 Then
                    If Metric(iCoord1, iCoord2) <> "" Then Diagonal = False
                End If
            Next
        Next
        If Diagonal Then
            For iCoord1 = 1 To Dimension
                If Metric(iCoord1, iCoord1) = "" Then Return "Diagonal components" _
                   & vbCr & "cannot be zero."
                'I dont see how this could ever happen. Hey ho
            Next
            Return ""
        End If

        'so its's not diagonal
        For iCoord1 = 1 To Dimension
            For iCoord2 = 1 To Dimension
                If iCoord1 <> iCoord2 Then
                    If Metric(iCoord1, iCoord2) <> Metric(iCoord2, iCoord1) Then
                        If (Metric(iCoord1, iCoord2) = "") And (Metric(iCoord2, iCoord1) <> "") Then
                            CopyDivide(Metric(iCoord2, iCoord1), Metric(iCoord1, iCoord2))
                        End If
                        If (Metric(iCoord1, iCoord2) <> "") And (Metric(iCoord2, iCoord1) = "") Then
                            CopyDivide(Metric(iCoord1, iCoord2), Metric(iCoord2, iCoord1))
                        End If
                        If Metric(iCoord1, iCoord2) <> Metric(iCoord2, iCoord1) Then
                            'could not dorrect it
                            Return "Metric must be symmetric"
                        End If
                    End If
                End If
            Next
        Next
        Return ""
    End Function
    Sub CopyDivide(ByRef Component1 As String, ByRef Component2 As String)
        'Component1 has some formula in it, Component2 does not.
        'divide component1 by 2 then set COmponent2 to be the same
        'This is the case for a metric like ds² = dx² + 4ydxdy + x³dy²
        'which should have been pedantically written ds² = dx² + 2ydxdy + 2ydydx + x³dy²
        Dim Number As String = "", Half As String, iNumber As Integer, dHalf As Double
        Dim iX As Integer, iXstart As Integer

        'The  "-" and "−" below are different! Paste them into Notepad or Word
        If (Mid(Component1, 1, 1) = "-") Or (Mid(Component1, 1, 1) = "−") Then iXstart = 2 Else iXstart = 1
        iX = iXstart
        While iX <= Len(Component1)
            If InStr("0123456789", Mid(Component1, iX, 1)) = 0 Then Exit While
            Number &= Mid(Component1, iX, 1)
            iX += 1
        End While
        If Number = "" Then
            Half = "0.5"
        Else
            iNumber = CInt(Number)
            dHalf = iNumber / 2
            Half = dHalf
        End If
        If (dHalf = 1) And (Mid(Component1, iXstart + 1, 1) <> "/") Then Half = ""
        If iXstart = 2 Then
            Component1 = Mid(Component1, 1, 1) & Half & Mid(Component1, Len(Number) + 2)
        Else
            Component1 = Half & Mid(Component1, Len(Number) + 1)
        End If
        Component2 = Component1
    End Sub
    Public Sub WriteMetrics()
        'write metrics as matrix at insertion point in equation or over selected part of equation
        Dim Equation As Word.OMath, InsertPoint As Word.Range
        Dim MetricIndices As String
        Dim Selection As Word.Selection
        Dim objUndo As Word.UndoRecord

        Selection = Globals.EquationAcc.Application.ActiveDocument.ActiveWindow.Selection
        If Selection.OMaths.Count <> 1 Then TensorError("Please put cursor in equation.") : Exit Sub
        If (gMetricNCoordsString = "") And (gInvMetricNCoordsString = "") Then _
            TensorError("Please pick up metrics first.") : Exit Sub

        objUndo = Globals.EquationAcc.Application.UndoRecord
        objUndo.StartCustomRecord("Write Metric")

        Equation = Selection.OMaths(1)
        InsertPoint = Selection.Range

        If gNmetricDimension = 4 Then MetricIndices = ("μν") Else MetricIndices = "ij"
        InsertMetric(Equation.Functions, InsertPoint, MetricIndices)
        InsertMatrix(Equation, InsertPoint, gNmetricDimension, gMetric)

        Equation = Selection.OMaths(1)      'get back outside brackets, to top level
        InsertPoint.Start = Equation.Range.End
        InsertPoint.End = Equation.Range.End
        InsertText(InsertPoint, " , ")
        InsertInverseMetric(Equation.Functions, InsertPoint, MetricIndices)
        InsertMatrix(Equation, InsertPoint, gNmetricDimension, gInvMetric)

        TensorMessage("✓")
        objUndo.EndCustomRecord()
    End Sub
    Public Sub ClearMetrics()
        ClearMetrics2()
        Globals.EquationAcc.MyRibbon.Invalidate()
        TensorMessage("Cleared metrics and coordinates")
    End Sub
    Public Sub ClearMetrics2()
        Dim iX1 As Integer, iX2 As Integer
        For iX1 = 0 To 4
            gCoordinates(iX1) = ""
            For iX2 = 0 To 4
                gMetric(iX1, iX2) = ""
                gInvMetric(iX1, iX2) = ""
            Next
        Next
        gNcoords = 0
        gNmetricDimension = 0
        gNInvmetricDimension = 0
        gMetricNCoordsString = ""
        gInvMetricNCoordsString = ""
        gCoordinatesString = ""
    End Sub
    Async Function WriteChristoffelSymbolsAsync() As Task
        'Function WriteChristoffelSymbolsAsync()   ' Non-async useful for debugging
        'Write out all Christoffel symbols in three column equations.
        'Need all metrics and coordinates to be loaded
        Dim UpIndex As Integer, LeftIndex As Integer, RightIndex As Integer
        Dim CSymbols(4, 4, 4) As String      'calculated Christoffel symbols as linear text. Were 1 to 4. Now 0 to 4
        'use these to avoid recaculating symmetric coefficients
        Dim OppositeTerm As String
        Dim Selection As Word.Selection
        Dim IOff As Integer    'Index offset for display of Gamma indices if numbers

        If (gMetric(1, 1) = "") Or (gInvMetric(1, 1) = "" Or gCoordinates(1) = "") Then
            TensorError("Metrics and coordinates" & vbCr & "must be loaded.")
            Return
        End If
        If (gNcoords <> gNmetricDimension) Or (gNcoords <> gNInvmetricDimension) Then
            TensorError("Dimensions of metric," & vbCr & "inverse metric and" & vbCr & "coordinates must be same.")
            Return
        End If

        If gNcoords = 4 Then IOff = -1 Else IOff = 0
        Selection = Globals.EquationAcc.Application.ActiveDocument.ActiveWindow.Selection
        For UpIndex = 1 To gNmetricDimension
            For LeftIndex = 1 To gNmetricDimension
                For RightIndex = 1 To gNmetricDimension
                    OppositeTerm = CSymbols(UpIndex, RightIndex, LeftIndex) 'blank is OK
                    If Not WriteChristoffel(UpIndex, LeftIndex, RightIndex, OppositeTerm) Then Return
                    'Message below is never visible :-(
                    TensorProgress("Done " & MyStr(UpIndex + IOff) & "," & MyStr(LeftIndex + IOff) & "," & MyStr(RightIndex + IOff))
                    Await Task.Delay(1)
                    CSymbols(UpIndex, LeftIndex, RightIndex) = OppositeTerm     'Now not opposite term
                    Selection.MoveRight(Word.WdUnits.wdCell, 1) 'get into equation nr cell
                    Selection.MoveDown(Word.WdUnits.wdLine, 1)
                Next
            Next
            If gNmetricDimension > 2 Then Selection.InsertNewPage()
        Next
        TensorMessage("Written out all " & MyStr(gNmetricDimension * gNmetricDimension * gNmetricDimension) & vbCr &
                      "Christoffel coefficients.")
        Globals.EquationAcc.MyRibbon.ActivateTabMso("TabAddIns")
    End Function
    Function WriteChristoffel(UpIndex As Integer, LeftIndex As Integer, RightIndex As Integer,
                                    ByRef LinearResult As String) As Boolean
        'write out one Christoffel symbol given three indices =0,1,2,3
        'LinearResult contains precalulated symbol and calulated symbol is returned in it
        Dim Equation As Word.OMath, InsertPoint As Word.Range
        Dim TheTable As Word.Table
        Dim TheRow As Word.Row
        Dim UseCoords As Boolean  'Indicates whether to use coordinates or numbers for Gamma indices
        Dim IOff As Integer    'Index offset for display of Gamma indices if numbers
        Dim LinearTerm(3) As String 'Terms of expansion in Linear format. Create them, compare, then adjust and insert. Was (1 To 3)
        Dim Term As String      'The result as a linear term
        Dim UpIndexS As String, LeftIndexS As String, RightIndexS As String
        Dim Selection As Word.Selection

        Selection = Globals.EquationAcc.Application.ActiveDocument.ActiveWindow.Selection
        WriteChristoffel = False

        'create new numbered row, selection in last row (column?)
        If Not EquationTable3a(1.3) Then Exit Function
        WriteChristoffel = True
        TheTable = Selection.Tables(1)
        TheRow = TheTable.Rows(TheTable.Rows.Count)
        Equation = TheRow.Cells(1).Range.OMaths(1)
        Equation.Range.Text = ""
        InsertPoint = Equation.Range
        InsertPoint.Start = InsertPoint.End

        If (gCoordinates(1) = "x") Or (gCoordinates(2) = "x") Then UseCoords = False Else UseCoords = True
        If gNcoords = 4 Then IOff = -1 Else IOff = 0
        If UseCoords Then
            UpIndexS = gCoordinates(UpIndex)
            LeftIndexS = gCoordinates(LeftIndex)
            RightIndexS = gCoordinates(RightIndex)
        Else
            UpIndexS = MyStr(UpIndex + IOff)
            LeftIndexS = MyStr(LeftIndex + IOff)
            RightIndexS = MyStr(RightIndex + IOff)
        End If
        Call InsertGamma(Equation.Functions, InsertPoint, UpIndexS, LeftIndexS & RightIndexS)

        Equation = TheRow.Cells(2).Range.OMaths(1)
        If LinearResult <> "" Then
            'Must be in symmetric term (LeftIndex <> RightIndex) and have already calculated, so
            Equation.Range.Text = " "
            InsertPoint = Equation.Range
            InsertPoint.Start = InsertPoint.End
            InsertText(InsertPoint, "=")
            InsertGamma(Equation.Functions, InsertPoint, UpIndexS, RightIndexS & LeftIndexS)
            InsertText(InsertPoint, "=")
            Term = GetLinear(Equation)
            Term &= LinearResult
            SetLinear(Equation, Term)
            Exit Function
            'Avoided hard work!
        End If

        'Now the hard work, similar to ExpandChristoffel except we are using actual components
        LinearTerm(1) = CalculateChristoffelTerm(Equation, UpIndex, LeftIndex, -1, RightIndex)
        LinearTerm(2) = CalculateChristoffelTerm(Equation, UpIndex, RightIndex, LeftIndex, -1)
        LinearTerm(3) = CalculateChristoffelTerm(Equation, UpIndex, -1, LeftIndex, RightIndex)

        InsertPoint = Equation.Range
        InsertPoint.Start = InsertPoint.End

        'See if any terms are same or cancel and write the bit after the = sign
        If (LinearTerm(1) = "0") And (LinearTerm(2) = "0") And (LinearTerm(3) = "0") Then
            InsertText(InsertPoint, "0")
        ElseIf LinearTerm(1) = LinearTerm(3) Then
            '+L1-L3=0
            If LinearTerm(2) = "0" Then
                InsertText(InsertPoint, "0")
            Else
                InsertFraction(Equation.Functions, InsertPoint, "1", "2")
                InsertBrackets(Equation.Functions, InsertPoint, LinearTerm(2))
            End If
        ElseIf LinearTerm(2) = LinearTerm(3) Then
            '+L2-L3=0
            If LinearTerm(1) = "0" Then
                InsertText(InsertPoint, "0")
            Else
                InsertFraction(Equation.Functions, InsertPoint, "1", "2")
                InsertBrackets(Equation.Functions, InsertPoint, LinearTerm(1))
            End If
        ElseIf (LinearTerm(1) = LinearTerm(2)) And (LinearTerm(1) <> "0") Then
            'will cancel half at front
            Term = GetLinear(Equation)
            Term = Term & LinearTerm(1)
            SetLinear(Equation, Term)
            InsertPoint.Start = Equation.Range.End
            If LinearTerm(3) <> "0" Then
                InsertText(InsertPoint, "-")
                InsertFraction(Equation.Functions, InsertPoint, "1", "2")
                InsertBrackets(Equation.Functions, InsertPoint, LinearTerm(3))
            End If
        Else
            'no simplification. At least one term is non-zero. Need to add
            '  1/2 (L1+L2-L3)
            InsertFraction(Equation.Functions, InsertPoint, "1", "2")
            Equation = InsertBrackets(Equation.Functions, InsertPoint, "")
            Term = ""
            If LinearTerm(1) <> "0" Then
                Term = LinearTerm(1)
            End If
            If LinearTerm(2) <> "0" Then
                If Term <> "" Then Term = Term & "+" & LinearTerm(2) Else Term = LinearTerm(2)
            End If
            Term = GetLinear(Equation) & Term
            SetLinear(Equation, Term)
            InsertPoint.Start = Equation.Range.End
            If LinearTerm(3) <> "0" Then
                InsertText(InsertPoint, "-")
                InsertBrackets(Equation.Functions, InsertPoint, LinearTerm(3))
            End If
        End If

        Equation = TheRow.Cells(2).Range.OMaths(1)
        LinearResult = GetLinear(Equation)

        'Must postpone putting in = at start, so we can easily return the equation for possible future use
        InsertPoint = Equation.Range
        InsertPoint.End = InsertPoint.Start
        Call InsertText(InsertPoint, "=")
        If LeftIndex <> RightIndex Then
            'C symbols are symmetric in bottom two indices
            Call InsertGamma(Equation.Functions, InsertPoint, UpIndexS, RightIndexS & LeftIndexS)
            Call InsertText(InsertPoint, "=")
        End If
    End Function
    Function CalculateChristoffelTerm(Equation As Word.OMath, ImIx1 As Integer, DerIx As Integer, mIx1 As Integer, mIx2 As Integer) As String
        'calculate one term in three terms of C expansion
        'Equation is the work area where eventually the full expansion will go
        'consists of inverse metric component contracted with derivative of metric component = summation
        'ImIx1 = first index of inverse metric. Its second index is contracted
        'DerIx, mIx1 , mIx2 are indices of partial derivative, and metrix
        'one of DerIx, mIx1 , mIx2 is -1, that is the one to be contracted with the second index of the inverse metric
        'returns "0" or Linear version
        Dim SumIx As Integer        'summation index
        Dim DerWRT As String        'derivative with respect to
        Dim NewMathTerm As Word.OMathFunction
        Dim Term As String          'term as linear string, what we will return
        Dim Term1 As String         'intermediate term
        Dim iPoint As Word.Range
        Dim cDerIx As Integer, cmIx1 As Integer, cmIx2 As Integer       'copies of last three parameters
        Dim PlainComponent As String
        Dim FunIx As Integer, TimesIx As Long

        Term = ""
        cDerIx = DerIx
        cmIx1 = mIx1
        cmIx2 = mIx2
        Equation.Range.Text = " "
        Equation.Type = Word.WdOMathType.wdOMathDisplay
        iPoint = Equation.Range

        For SumIx = 1 To gNcoords
            If gInvMetric(ImIx1, SumIx) <> "0" Then
                'have inv metric component x derivative of metric component
                If DerIx = -1 Then cDerIx = SumIx
                If mIx1 = -1 Then cmIx1 = SumIx
                If mIx2 = -1 Then cmIx2 = SumIx
                DerWRT = gCoordinates(cDerIx)
                CleanStringString(gMetric(cmIx1, cmIx2), PlainComponent)
                If InStr(1, PlainComponent, DerWRT, vbTextCompare) <> 0 Then      'vbTextCompare or vbBinaryCompare?
                    'some char in operand of derivative contains charachter same as DerWRT
                    '(otherwise derivative=0 so we ignote this term)
                    'd/dx
                    Call InsertFraction(Equation.Functions, iPoint, gPartialDerivative, gPartialDerivative & gCoordinates(cDerIx))
                    'bracket containing metric. iPointneeds dragging along...
                    iPoint.Start = Equation.Range.End
                    iPoint.End = Equation.Range.End
                    Call InsertBrackets(Equation.Functions, iPoint, gMetric(cmIx1, cmIx2))

                    If gInvMetric(ImIx1, SumIx) = "1" Then
                        Term1 = ""
                    Else
                        Term1 = gInvMetric(ImIx1, SumIx) & "×"
                    End If
                    If Term <> "" Then
                        Term = Term & "+" & Term1 & GetLinear(Equation)
                    Else
                        Term = Term1 & GetLinear(Equation)
                    End If
                    'Call SetLinear(Equation, Term)      'debug only******************
                    Equation.Range.Text = " "            'ready for next term
                    Equation.Type = Word.WdOMathType.wdOMathDisplay
                End If
            End If
        Next

        If Term = "" Then
            Term = "0"
        Else
            'remove the times sign before the derivative
            Call SetLinear(Equation, Term)
            FunIx = 1
            While FunIx <= Equation.Functions.Count
                NewMathTerm = Equation.Functions(FunIx)
                If NewMathTerm.Type = Word.WdOMathFunctionType.wdOMathFunctionText Then 'wdOMathFunctionText or wdOMathFunctionNormalText??
                    If NewMathTerm.Range.Text = "×" Then
                        'was NewMathTerm.Remove() : FunIx -= 1
                        'but Remove crashes
                        NewMathTerm.Range.Text = ""
                    Else
                        TimesIx = InStr(1, NewMathTerm.Range.Text, "×", vbTextCompare)
                        If TimesIx = Len(NewMathTerm.Range.Text) Then
                            'times at end of string, just before partial derivative
                            NewMathTerm.Range.Text = Mid(NewMathTerm.Range.Text, 1, Len(NewMathTerm.Range.Text) - 1)
                        End If
                    End If
                End If
                FunIx += 1
            End While
            Term = GetLinear(Equation)
            Equation.Range.Text = " "            'ready for next term
            Equation.Type = Word.WdOMathType.wdOMathDisplay
        End If
        CalculateChristoffelTerm = Term
    End Function
    '**************************************
    'Calculations!
    '**************************************
    Function CalculateDiagonalInverse(ByRef Metric(,) As String, ByRef InvMetric(,) As String) As Boolean
        'Calculate inverse of diagonal Metric. That is reciprocal of each diagonal element.
        'Put result in InvMetric, NMetric which may be metric or inverse metric
        'Has to create TempDoc as equation playground
        'return false if had divide by 0
        Dim TempDoc As Word.Document
        Dim objRange As Word.Range
        Dim iRow, iCol As Integer
        Dim Selection As Word.Selection, Position As Long

        Selection = Globals.EquationAcc.Application.ActiveDocument.ActiveWindow.Selection
        Position = Selection.Start
        TempDoc = Globals.EquationAcc.Application.Documents.Add(, ,
            Word.WdNewDocumentType.wdNewBlankDocument, False)       'False -> True for debugging
        TempDoc.ActiveWindow.WindowState = Word.WdWindowState.wdWindowStateMinimize
        For iRow = 1 To gNmetricDimension
            For iCol = 1 To gNmetricDimension
                If iRow = iCol Then
                    objRange = TempDoc.Range()
                    objRange.Text = Metric(iRow, iCol)     'inserts the text e.g 𝑥^2 
                    objRange.OMaths.Add(objRange)       'turns that into equation, still e.g 𝑥^2 
                    objRange.OMaths(1).BuildUp()        'makes the equation pretty (but size 10 not 12). e.g. 𝑥²
                    If Not InvertOmath(objRange.OMaths(1)) Then
                        TempDoc.Close(Word.WdSaveOptions.wdDoNotSaveChanges)
                        Selection.Start = Position
                        Return False
                    End If    'takes reciprocal, e.g. 1/𝑥² 
                    InvMetric(iRow, iCol) = GetCleanLinear(objRange.OMaths(1))
                Else
                    InvMetric(iRow, iCol) = "0"
                End If

            Next
        Next
        TempDoc.Close(Word.WdSaveOptions.wdDoNotSaveChanges)
        Selection.Start = Position
        Return True
    End Function
    Function InvertOmath(Equation As Word.OMath) As Boolean
        'make equation reciprocal of self
        'Special cases, equation is
        'fraction - swap denominator and numerator
        '1 / 1 is 1!
        'Equation is xx to power of yy, yy goes to -yy NOT. Superscript might be an index.
        'NOT DONE trig functions
        Dim MathTerm As Word.OMathFunction
        Dim Term1 As String, Term2 As String
        Dim HadMinus As Boolean

        'Check if - sign at front. If there is, remove it
        Term1 = GetLinear(Equation)
        If IsEqMinus(Equation) Then
            HadMinus = True
            Call SetLinear(Equation, Mid(Term1, 2))
        Else
            HadMinus = False
        End If

        If Equation.Functions.Count = 1 Then
            MathTerm = Equation.Functions(1)

            '1 or -1
            If MathTerm.Range.Text = "1" Then
                If HadMinus Then Call SetLinear(Equation, "-" & GetLinear(Equation))
                Return True
            End If

            'Check 0
            If MathTerm.Range.Text = "0" Then
                HighlightError(MathTerm.Range, "Divide by zero.")
                Return False
            End If

            'fraction
            If MathTerm.Type = Word.WdOMathFunctionType.wdOMathFunctionFrac Then
                Term1 = GetLinear(MathTerm.Frac.Num)
                Term2 = GetLinear(MathTerm.Frac.Den)
                If (Term1 = "1") Or (IsStrMinus(Term1) And (Mid(Term1, 2) = "1")) Then
                    ' 1 or -1 in numerator
                    If IsStrMinus(Term1) Then HadMinus = Not HadMinus
                    Call SetLinearWithMinus(Equation, Term2, HadMinus)
                Else
                    'keep -sign on top
                    If IsEqMinus(MathTerm.Frac.Num) Then
                        Term1 = Mid(Term1, 2)
                        HadMinus = Not HadMinus
                    End If
                    If IsEqMinus(MathTerm.Frac.Den) Then
                        Term2 = Mid(Term2, 2)
                        HadMinus = Not HadMinus
                    End If
                    Call SetLinearWithMinus(MathTerm.Frac.Num, Term2, HadMinus)
                    Call SetLinear(MathTerm.Frac.Den, Term1)
                End If
                Return True
            End If
        End If

        'Normal case. Create fraction
        Term1 = GetLinear(Equation)
        Equation.Range.Text = ""
        Equation.Range.End = Equation.Range.Start
        MathTerm = Equation.Functions.Add(Equation.Range, Word.WdOMathFunctionType.wdOMathFunctionFrac)
        SetLinearWithMinus(MathTerm.Frac.Num, "1", HadMinus)
        SetLinear(MathTerm.Frac.Den, Term1)
        Return True
    End Function
    Sub SetLinearWithMinus(Equation As Word.OMath, ByVal Term As String, HadMinus As Boolean)
        If HadMinus Then Term = "-" & Term
        Call SetLinear(Equation, Term)
    End Sub
    Function IsEqMinus(Equation As Word.OMath) As Boolean
        'test if an equation is negative. Must be one or two terms starting with the dodgy minus
        '-2+3 must fail
        Dim Term As String
        IsEqMinus = False
        If Equation.Functions.Count > 2 Then Exit Function

        If Equation.Functions(1).Type = Word.WdOMathFunctionType.wdOMathFunctionText Then 'wdOMathFunctionNormalText or wdOMathFunctionText ??
            Term = Equation.Functions(1).Range.Text
            If Not IsStrMinus(Term) Then Exit Function      'does not start with -
            If InStr(2, Term, "+", vbTextCompare) <> 0 Then Exit Function  'contains a +
        Else
            Exit Function
        End If
        IsEqMinus = True
    End Function
    Function MyStr(Num As Integer) As String
        'like Str, but without leading space (which occurs for +ve Num)
        If Num >= 0 Then MyStr = Mid(Str(Num), 2) Else MyStr = Str(Num)
    End Function
    '**************************************
    'Functions for manipulating equations
    '**************************************
    Sub CleanStringString(InString As String, ByRef OutString As String)
        'Translage InString from Omath encoding to vanilla with no encoding (not bold, not italic)
        Dim iX As Integer
        OutString = ""
        For iX = 1 To Len(InString)
            OutString &= CleanChar(Mid(InString, iX, 1))
        Next
    End Sub
    Sub CleanString(InRange As Word.Range, ByRef OutString As String)
        'Translage InRange.Text from Omath encoding to vanilla with no encoding (not bold, not italic)
        'First attempt was roughly
        'InRange.Bold = False : InRange.Italic = False : OutString = InRange.Text
        'which should work, but it randomly doesn't!! So this elaborate bodge is the result
        'See "Char codes.docx" for notes on the very weird MS-math character encoding
        CleanStringString(InRange.Text, OutString)
    End Sub
    Function CleanChar(InChar As String) As String
        'InChar is one character with bold / italic / none. Returns one character eith none
        'Do tests on most likely first. Might be a millisecond quicker
        Dim GreekChar As String
        Dim XChar As Integer = AscW(InChar)

        If XChar = &HD835 Then Return ""            '50% is junk!
        If XChar = &H210E Then Return "h"                              'Italic h, Planck Surprise!
        If &H20 <= XChar And XChar <= &H2720 Then Return InChar        'Not Bold or Italic. Everything in Unicode ...
        'a greek and romans
        If &HDEFC <= XChar And XChar <= &HDF14 Then Return ChrW(XChar - &HDEFC + &H3B1)    'Italic α-ω
        If &HDC4E <= XChar And XChar <= &HDC67 Then Return ChrW(XChar - &HDC4E + &H61)     'Italic a-z
        If &HDC34 <= XChar And XChar <= &HDC4D Then Return ChrW(XChar - &HDC34 + &H41)     'Italic A-Z
        If &HDFCE <= XChar And XChar <= &HDFD7 Then Return ChrW(XChar - &HDFCE + &H30)     'Bold 0-9
        If &HDC00 <= XChar And XChar <= &HDC19 Then Return ChrW(XChar - &HDC00 + &H41)     'Bold A-Z
        If &HDC1A <= XChar And XChar <= &HDC33 Then Return ChrW(XChar - &HDC1A + &H61)     'bold a-z
        If &HDC68 <= XChar And XChar <= &HDC81 Then Return ChrW(XChar - &HDC68 + &H41)     'bold italic A-Z
        If &HDC82 <= XChar And XChar <= &HDC9B Then Return ChrW(XChar - &HDC82 + &H61)     'bold italic  a-z
        'greeks
        If &HDEE2 <= XChar And XChar <= &HDEFA Then Return ChrW(XChar - &HDEE2 + &H391)    'Italic Α-Ω
        If &HDEA8 <= XChar And XChar <= &HDEC0 Then Return ChrW(XChar - &HDEA8 + &H391)    'Bold Α-Ω
        If &HDEC2 <= XChar And XChar <= &HDEDA Then Return ChrW(XChar - &HDEC2 + &H3B1)    'Bold α-ω
        If &HDF1C <= XChar And XChar <= &HDF34 Then Return ChrW(XChar - &HDF1C + &H391)    'Bold Italic Α-Ω
        If &HDF36 <= XChar And XChar <= &HDF4E Then Return ChrW(XChar - &HDF36 + &H3B1)    'Bold Italic α-ω

        If GreekOddities(XChar - &HDEDD, GreekChar) Then Return GreekChar      'bold
        If GreekOddities(XChar - &HDF51, GreekChar) Then Return GreekChar      'bold italic
        If GreekOddities(XChar - &HDF17, GreekChar) Then Return GreekChar      'italic

        Return ""
    End Function
    Function GreekOddities(Offset As Integer, ByRef GreekChar As String) As Boolean
        'Deal with ϑ,ϕ,ϖ,ϱ which are encouraged by MS-math and are alternate θ,φ,π,ρ
        If Offset = 0 Then
            GreekChar = "ϑ"
            Return True
        End If
        If Offset = 2 Then
            GreekChar = "ϕ"
            Return True
        End If
        If Offset = 3 Then
            GreekChar = "ϱ"
            Return True
        End If
        If Offset = 4 Then
            GreekChar = "ϖ"
            Return True
        End If
        Return False
    End Function
    Public Function GetNormal(Equation As Word.OMath) As String
        'return normal text of equation - useful when you want to check for specific characters
        Dim Original As String
        Original = GetLinear(Equation)
        Equation.ConvertToNormalText()
        GetNormal = Equation.Range.Text
        Equation.Range.Text = Original
        Equation.BuildUp()
    End Function
    Public Sub SetLinear(Equation As Word.OMath, LinearEq As String)
        'Puts back equation found with get linear. For example converts 𝑥^2 to 𝑥²
        'Equation.Range.Italic = 0           'Ensures that no-italic and bold are preserved !!! SOMETIMES
        Equation.Range.Text = LinearEq
        Equation.BuildUp()
    End Sub
    Public Function GetLinear(Equation As Word.OMath) As String
        'return linear version of equation - useful when you want to put equation somewhere else
        If Equation.Range.Text = "" Then
            ' Equation.linearize does something horrid with empty equations
            GetLinear = ""
            Exit Function
        End If
        Equation.Linearize()
        GetLinear = Equation.Range.Text
        Equation.BuildUp()
    End Function
    Public Function GetCleanLinear(Equation As Word.OMath) As String
        'As GetLinear but cleans up results such as ""1" " or ""0" " which often seem to occur in matrices
        Dim Result As String, QuoteIx As Integer
        If Equation.Range.Text = "" Then
            ' Equation.linearize does something horrid with empty equations
            GetCleanLinear = ""
            Exit Function
        End If
        Equation.Linearize()
        Result = Equation.Range.Text
        If Mid(Result, 1, 1) = gDoubleQuote Then
            Result = Mid(Result, 2)     'remove first quote
            QuoteIx = InStr(1, Result, gDoubleQuote, vbTextCompare)
            Result = Mid(Result, 1, QuoteIx - 1) & Mid(Result, QuoteIx + 1)
            ' RTrim (Result) Does not work
            For QuoteIx = Len(Result) To 1 Step -1
                If Mid(Result, QuoteIx, 1) = " " Then
                    Result = Mid(Result, 1, QuoteIx - 1)
                Else
                    Exit For
                End If
            Next
            Equation.Range.Text = Result
        End If
        Equation.BuildUp()
        GetCleanLinear = Result
    End Function
End Module
