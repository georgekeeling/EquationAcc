'Add in to help with equation editing for Office 365 (2019)
'By George Keeling and on my blog at www.general-relativity.net
'Possibly documented at https://www.general-relativity.net/search/label/Tools
'Feel free to use, copy, modify and give away but not for commerce. Please acknowledge me

'Entry points
'  ExpandSymbols -Expand Christoffel symbol (Γ), Covariant derivative (∇) And / Or
'      Riemann tensor (R) where they occur in an equation.
'Also contains
'  InsertFunctions -Functions for inserting various things in equations And while keeping track of insertion point

Module TransformersExpand
    Dim gIndexesUsed(51) As Integer  'Count of usage of each index alpha at 1, omega at 25, a at 26, z at 51. Was (1 To 51) now (0 t0 51)
    Dim gGreekUsed As Boolean

    Public Sub ExpandSymbols()
        'Expand Christoffel symbol (Γ), Covariant derivative (∇) and / or
        'Riemann tensor (R) where they occur in an equation.
        Dim Equation As Word.OMath
        Dim Selection As Word.Selection
        Dim objUndo As Word.UndoRecord
        objUndo = Globals.EquationAcc.Application.UndoRecord
        Selection = Globals.EquationAcc.Application.ActiveDocument.ActiveWindow.Selection

        ZapIndexesUsed()
        If Selection.OMaths.Count <> 1 Then TensorError("Please put cursor in equation.") : Exit Sub
        objUndo.StartCustomRecord("Expand equation")
        Selection.Collapse(Word.WdCollapseDirection.wdCollapseStart)        'Must do this because functions do not quite work on selection
        Equation = Selection.OMaths(1)
        If Not FindIndexesUsed(Equation) Then
            objUndo.EndCustomRecord()
            TensorError("Index formatting too complex." & vbCrLf & "Suggest all italic, not bold.")
            Return
        End If
        If ExpandEquation(Equation) Then
            objUndo.EndCustomRecord()
            TensorMessage("Equation expanded.")
        Else
            objUndo.EndCustomRecord()           'error message was given
        End If

    End Sub
    Function ExpandEquation(ByVal Equation As Word.OMath) As Boolean
        Dim MathTerm As Word.OMathFunction, iMathTerm As Integer
        Dim InnerEquation As Word.OMath, iInnerEq As Integer
        Dim HaveMinus As String, NormalText As String, TermsInCD As Integer

        HaveMinus = ""
        ExpandEquation = True
        iMathTerm = 1
        While iMathTerm <= Equation.Functions.Count
            MathTerm = Equation.Functions(iMathTerm)
            If (MathTerm.Type = Word.WdOMathFunctionType.wdOMathFunctionText) Or
                (MathTerm.Type = Word.WdOMathFunctionType.wdOMathFunctionNormalText) Then
                'If last charachter is -
                If IsStrMinus(Mid(MathTerm.Range.Text, Len(MathTerm.Range.Text), 1)) Then HaveMinus = "-"
            End If

            If (MathTerm.Type = Word.WdOMathFunctionType.wdOMathFunctionScrSubSup) Then
                'Christoffel or Riemann
                NormalText = GetNormal(MathTerm.ScrSubSup.E)
                If NormalText = gCapGamma Then
                    'Have Christoffel symbol. Wierdly, in the debugger is displayed as G (VBA)
                    Call ExpandChristoffel(Equation.Functions, MathTerm)
                    iMathTerm += 2       'two terms added. (term in bracket = 1 term)
                ElseIf NormalText = "R" Then   'Letter R
                    'expand Riemann and ignore terms created
                    iMathTerm += ExpandRiemannTensor(Equation.Functions, MathTerm, HaveMinus)
                End If
            ElseIf MathTerm.Type = Word.WdOMathFunctionType.wdOMathFunctionScrSub Then
                NormalText = GetNormal(MathTerm.ScrSub.E)
                If NormalText = gNabla Then 'Nabla: Covariant derivative
                    'Do Covariant derivative term, next term and skip terms created.
                    If iMathTerm < Equation.Functions.Count Then
                        iMathTerm = iMathTerm + 1
                        TermsInCD = ExpandCovariantDerivative(Equation.Functions, MathTerm, Equation.Functions(iMathTerm), HaveMinus)
                        If TermsInCD = -1 Then Return False   'error
                        iMathTerm += TermsInCD
                    End If
                End If
            ElseIf MathTerm.Type = Word.WdOMathFunctionType.wdOMathFunctionDelim Then   'delimiter - brackets, need to recurse
                For iInnerEq = 1 To MathTerm.Delim.E.Count
                    InnerEquation = MathTerm.Delim.E.Item(iInnerEq)
                    Call ExpandEquation(InnerEquation)
                Next
            End If
            iMathTerm += 1
        End While
    End Function

    Sub ExpandChristoffel(Functions As Word.OMathFunctions, MathTerm As Word.OMathFunction)
        'Expand Christoffel symbol
        'Functions are the functions of the original equation, which we add to
        'Mathterm is the Christoffel symbol term
        Dim UpIndex As String, LeftIndex As String, RightIndex As String, DownIndices As String, DummyIndex As String
        Dim BracketFunctions As Word.OMathFunctions    'just like paramter of this function
        Dim BracketEq As Word.OMath
        Dim InsertionPoint As Word.Range

        UpIndex = GetNormal(MathTerm.ScrSubSup.Sup)
        DownIndices = GetNormal(MathTerm.ScrSubSup.Sub)
        LeftIndex = Mid(DownIndices, 1, 1)
        RightIndex = Mid(DownIndices, 2, 1)

        DummyIndex = GetNewDummyIndex()

        MathTerm.Range.Text = ""    'removes Christoffel symbol
        InsertionPoint = MathTerm.Range

        '1/2 x inverse metric
        Call InsertFraction(Functions, InsertionPoint, "1", "2")
        Call InsertInverseMetric(Functions, InsertionPoint, UpIndex & DummyIndex)

        'brackets
        BracketEq = InsertBrackets(Functions, InsertionPoint, "")
        BracketFunctions = BracketEq.Functions

        'first + term
        Call InsertPartial(BracketEq.Functions, InsertionPoint, LeftIndex)
        Call InsertMetric(BracketEq.Functions, InsertionPoint, DummyIndex & RightIndex)

        'second + term
        Call InsertText(InsertionPoint, "+")
        Call InsertPartial(BracketEq.Functions, InsertionPoint, RightIndex)
        Call InsertMetric(BracketEq.Functions, InsertionPoint, LeftIndex & DummyIndex)

        'third - term
        Call InsertText(InsertionPoint, "-")
        Call InsertPartial(BracketEq.Functions, InsertionPoint, DummyIndex)
        Call InsertMetric(BracketEq.Functions, InsertionPoint, LeftIndex & RightIndex)

    End Sub

    Function ExpandCovariantDerivative(Functions As Word.OMathFunctions, CdTerm As Word.OMathFunction,
                    MathTerm As Word.OMathFunction, Sign As String) As Integer
        'Expand Covariant Derivative
        'Functions are the functions of the original equation, which we add to
        'CdTerm is the covariant derivative operator
        'Mathterm is the term after the covariant derivative term, the operand, a tensor with up or down indices or both
        'Sign is sign before the covariant derivative operator
        'We do not (yet) do an operand in brackets or deal with metric compatibility
        'returns number of OMathFunctions inserted or -1 if error
        'first do all up indices, then all down indices
        Dim Tensor As String, TensorType As Word.WdOMathFunctionType
        Dim UpIndexes As String, DownIndexes As String, iIndexString As Integer
        Dim TermIndexes As String
        Dim DummyIndex As String, CdIndex As String
        Dim NewMathTerm As Word.OMathFunction
        Dim InsertionPoint As Word.Range, OMFsCreated As Integer

        OMFsCreated = 0
        ExpandCovariantDerivative = 0

        TensorType = MathTerm.Type
        If TensorType = Word.WdOMathFunctionType.wdOMathFunctionScrSub Then
            CleanString(MathTerm.ScrSub.Sub.Range, DownIndexes)
            CleanString(MathTerm.ScrSub.E.Range, Tensor)
        ElseIf TensorType = Word.WdOMathFunctionType.wdOMathFunctionScrSup Then
            CleanString(MathTerm.ScrSup.Sup.Range, UpIndexes)
            CleanString(MathTerm.ScrSup.E.Range, Tensor)
        ElseIf TensorType = Word.WdOMathFunctionType.wdOMathFunctionScrSubSup Then
            CleanString(MathTerm.ScrSubSup.Sub.Range, DownIndexes)
            CleanString(MathTerm.ScrSubSup.Sup.Range, UpIndexes)
            CleanString(MathTerm.ScrSubSup.E.Range, Tensor)
        ElseIf TensorType = Word.WdOMathFunctionType.wdOMathFunctionText Then
            'scalar, just change symbol
            CdTerm.ScrSub.E.Range.Text = gPartialDerivative    'change nabla to partial derivative
            Exit Function
        Else
            Call HighlightError(MathTerm.Range, "Operand of covariant derivative" & vbCrLf & "must be tensor.")
            Return -1
        End If

        DummyIndex = GetNewDummyIndex()
        CleanString(CdTerm.ScrSub.Sub.Range, CdIndex)
        'CdIndex = ChrW(ToUnicode(Mid(CdTerm.ScrSub.Sub.Range.Text, 2, 1)))

        CdTerm.ScrSub.E.Range.Text = gPartialDerivative    'change nabla to partial derivative

        'Now add + term for each up index and - term for each down index
        InsertionPoint = MathTerm.Range
        InsertionPoint.Collapse(Word.WdCollapseDirection.wdCollapseEnd)

        If (TensorType = Word.WdOMathFunctionType.wdOMathFunctionScrSup) Or
            (TensorType = Word.WdOMathFunctionType.wdOMathFunctionScrSubSup) Then
            'add + terms
            For iIndexString = 1 To Len(UpIndexes)
                If Mid(UpIndexes, iIndexString, 1) <> " " Then
                    'insert Christoffel and trensor replacing that charachter by dummy index
                    TermIndexes = Mid(UpIndexes, 1, iIndexString - 1) & DummyIndex & Mid(UpIndexes, iIndexString + 1)
                    InsertPlusMinus(InsertionPoint, Sign, "+")
                    'insert Christoffel
                    InsertGamma(Functions, InsertionPoint, Mid(UpIndexes, iIndexString, 1), CdIndex & DummyIndex)
                    OMFsCreated += 3
                    'insert modified tensor
                    NewMathTerm = InsertFunction(Functions, InsertionPoint, TensorType)
                    If (TensorType = Word.WdOMathFunctionType.wdOMathFunctionScrSup) Then
                        NewMathTerm.ScrSup.E.Range.Text = Tensor
                        NewMathTerm.ScrSup.Sup.Range.Text = TermIndexes
                    Else
                        'upper and lower indices
                        NewMathTerm.ScrSubSup.E.Range.Text = Tensor
                        NewMathTerm.ScrSubSup.Sup.Range.Text = TermIndexes
                        NewMathTerm.ScrSubSup.Sub.Range.Text = DownIndexes
                    End If
                End If
            Next
        End If

        If (TensorType = Word.WdOMathFunctionType.wdOMathFunctionScrSub) Or
            (TensorType = Word.WdOMathFunctionType.wdOMathFunctionScrSubSup) Then
            'add - terms
            For iIndexString = 1 To Len(DownIndexes)
                If Mid(DownIndexes, iIndexString, 1) <> " " Then
                    'insert Christoffel and trensor replacing that charachter by dummy index
                    TermIndexes = Mid(DownIndexes, 1, iIndexString - 1) & DummyIndex & Mid(DownIndexes, iIndexString + 1)
                    InsertPlusMinus(InsertionPoint, Sign, "-")
                    'insert Christoffel
                    InsertGamma(Functions, InsertionPoint, DummyIndex, CdIndex & Mid(DownIndexes, iIndexString, 1))
                    OMFsCreated = OMFsCreated + 3
                    'insert modified tensor
                    NewMathTerm = InsertFunction(Functions, InsertionPoint, TensorType)
                    If (TensorType = Word.WdOMathFunctionType.wdOMathFunctionScrSub) Then
                        NewMathTerm.ScrSub.E.Range.Text = Tensor
                        NewMathTerm.ScrSub.Sub.Range.Text = TermIndexes
                    Else
                        'upper and lower indices
                        NewMathTerm.ScrSubSup.E.Range.Text = Tensor
                        NewMathTerm.ScrSubSup.Sup.Range.Text = UpIndexes
                        NewMathTerm.ScrSubSup.Sub.Range.Text = TermIndexes
                    End If
                End If
            Next

        End If
        ExpandCovariantDerivative = OMFsCreated
    End Function

    Function ExpandRiemannTensor(Functions As Word.OMathFunctions, MathTerm As Word.OMathFunction, Sign As String) As Integer
        'Expand Riemann Tensor
        'Functions are the functions of the original equation, which we add to
        'Mathterm is the Riemann tensor term
        'Sign is sign before it
        'Returns number of OMathFunctions inserted. 0 = error and problem is highlighted
        Dim InsertionPoint As Word.Range
        Dim UpIndex As String, DownIndexes As String, StringI As Integer, DummyIndex As String
        Dim DownIx(0 To 3) As String, ArrayIx As Integer
        Dim NewMathTerm As Word.OMathFunction

        If MathTerm.Type <> Word.WdOMathFunctionType.wdOMathFunctionScrSubSup Then
            ExpandRiemannTensor = 0
            Exit Function
        End If
        CleanString(MathTerm.ScrSubSup.Sup.Range, UpIndex)
        If Len(UpIndex) <> 1 Then
            MathTerm.ScrSubSup.Sup.Range.HighlightColorIndex = Word.WdColorIndex.wdRed
            ExpandRiemannTensor = 0
            Exit Function
        End If
        Call CleanString(MathTerm.ScrSubSup.Sub.Range, DownIndexes)
        ArrayIx = 1
        For StringI = 1 To Len(DownIndexes)
            If Mid(DownIndexes, StringI, 1) <> " " Then
                If ArrayIx = 4 Then Exit For
                DownIx(ArrayIx) = Mid(DownIndexes, StringI, 1)
                ArrayIx = ArrayIx + 1
            End If
        Next
        If ArrayIx <> 4 Then
            MathTerm.ScrSubSup.Sub.Range.HighlightColorIndex = Word.WdColorIndex.wdRed
            ExpandRiemannTensor = 0
            Exit Function
        End If

        DummyIndex = GetNewDummyIndex()
        MathTerm.Range.Text = ""    'removes Riemann
        InsertionPoint = MathTerm.Range

        'insert partial, Christoffel
        InsertPartial(Functions, InsertionPoint, DownIx(2))
        InsertGamma(Functions, InsertionPoint, UpIndex, DownIx(3) & DownIx(1))

        'insert partial, Christoffel
        InsertPlusMinus(InsertionPoint, Sign, "-")
        InsertPartial(Functions, InsertionPoint, DownIx(3))
        InsertGamma(Functions, InsertionPoint, UpIndex, DownIx(2) & DownIx(1))

        'insert Christoffel, Christoffel
        InsertPlusMinus(InsertionPoint, Sign, "+")
        InsertGamma(Functions, InsertionPoint, UpIndex, DownIx(2) & DummyIndex)
        InsertGamma(Functions, InsertionPoint, DummyIndex, DownIx(3) & DownIx(1))

        'insert Christoffel, Christoffel
        InsertPlusMinus(InsertionPoint, Sign, "-")
        InsertGamma(Functions, InsertionPoint, UpIndex, DownIx(3) & DummyIndex)
        InsertGamma(Functions, InsertionPoint, DummyIndex, DownIx(2) & DownIx(1))

        ExpandRiemannTensor = 10     'Always insert 10 OMathFunctions

    End Function
    Sub ZapIndexesUsed()
        'values preserved in gIndexesUsed from one call to another ...
        Dim iIU As Integer
        For iIU = 0 To UBound(gIndexesUsed)
            gIndexesUsed(iIU) = 0
        Next
        gGreekUsed = False
    End Sub
    Function FindIndexesUsed(ByVal Equation As Word.OMath) As Boolean
        'Find all tensor indexes used in equation, so that we do not use them as dummy variables.
        'only search for lower case italic roman and greek!
        Dim MathTerm As Word.OMathFunction
        Dim InnerEquation As Word.OMath
        Dim iInnerEq As Integer

        For Each MathTerm In Equation.Functions
            If MathTerm.Type = Word.WdOMathFunctionType.wdOMathFunctionScrSubSup Then   'super and sub script
                If Not CheckIndexes(MathTerm.ScrSubSup.Sub.Range) Then Return False
                If Not CheckIndexes(MathTerm.ScrSubSup.Sup.Range) Then Return False
            End If
            If MathTerm.Type = Word.WdOMathFunctionType.wdOMathFunctionScrSup Then   'super script
                If Not CheckIndexes(MathTerm.ScrSup.Sup.Range) Then Return False
            End If
            If MathTerm.Type = Word.WdOMathFunctionType.wdOMathFunctionScrSub Then   'sub script
                If Not CheckIndexes(MathTerm.ScrSub.Sub.Range) Then Return False
            End If
            If MathTerm.Type = Word.WdOMathFunctionType.wdOMathFunctionDelim Then   'delimiter - brackets, need to recurse
                For iInnerEq = 1 To MathTerm.Delim.E.Count
                    InnerEquation = MathTerm.Delim.E.Item(iInnerEq)
                    If Not FindIndexesUsed(InnerEquation) Then Return False
                Next
            End If
        Next
        Return True
    End Function
    Function CheckIndexes(ByVal IndexRange As Word.Range) As Boolean
        'Check all the indexes in a string which may contain spaces.
        Dim iText As Integer, UniCode As Integer, IndexString As String

        CheckIndexes = True
        CleanString(IndexRange, IndexString)

        For iText = 1 To Len(IndexString)
            UniCode = AscW(Mid(IndexString, iText, 1))
            If UniCode = &H3D1 Then UniCode = &H3B8         'alternate theta
            If UniCode = &H3D5 Then UniCode = &H3C6         'alternate phi, same as phi
            If UniCode = &H3D6 Then UniCode = &H3C0         'alternate pi
            If UniCode = &H3F1 Then UniCode = &H3C1         'alternate rho,
            If UniCode < &H20 Or UniCode > &H3D3 Then Return False  'CleanString fucked up!

            'Now nearly have correct Unicode char in UniCode which we will use to index into gIndexesUsed
            'Greek Α=H391 Ω=H3A9 α=H3B1, ω=H3C9 ,ϑ=H3D1, ϕ=H3D5
            'Roman A=H41 Z=H5A a=H61, z=H7A
            If (UniCode >= &H61) And (UniCode <= &H7A) Then
                'It's roman a-z
                UniCode = UniCode - &H60 + 25
            ElseIf (UniCode >= &H3B1) And (UniCode <= &H3C9) Then
                'alpha to omega
                UniCode -= &H3B0
                gGreekUsed = True
            Else
                UniCode = 0
            End If

            If UniCode > 0 Then gIndexesUsed(UniCode) += 1
        Next
    End Function
    Function GetNewDummyIndex() As String
        'get new dummy index.  greek if any greek in list, otherwise roman.
        'Start with mu,nu,lambda,,rho, sigma,tau kappa
        Dim iIU As Integer
        If gGreekUsed Then
            If gIndexesUsed(12) = 0 Then
                gIndexesUsed(12) = 1
                Return ChrW(&H3BC)
            End If
            If gIndexesUsed(13) = 0 Then
                gIndexesUsed(13) = 1
                Return ChrW(&H3BD)
            End If
            If gIndexesUsed(11) = 0 Then
                gIndexesUsed(11) = 1
                Return ChrW(&H3BB)
            End If
            If gIndexesUsed(17) = 0 Then
                gIndexesUsed(17) = 1
                Return ChrW(&H3C1)
            End If
            If gIndexesUsed(19) = 0 Then
                gIndexesUsed(19) = 1
                Return ChrW(&H3C3)
            End If
            If gIndexesUsed(20) = 0 Then
                gIndexesUsed(20) = 1
                Return ChrW(&H3C4)
            End If
            If gIndexesUsed(10) = 0 Then
                gIndexesUsed(10) = 1
                Return ChrW(&H3BA)
            End If
            'No obvious spare ones found. Start at alpha and end at z (oops, in worst case could be roman)
            For iIU = 1 To UBound(gIndexesUsed)
                If gIndexesUsed(iIU) = 0 Then
                    gIndexesUsed(iIU) = 1
                    If iIU <= 25 Then
                        Return ChrW(iIU + &H3B0)      'greek
                    Else
                        Return ChrW(iIU + &H60 - 25)      'bad luck roman
                    End If
                End If
            Next
        Else
            'roman just do a to z then alpha to omega
            For iIU = 26 To 51
                If gIndexesUsed(iIU) = 0 Then
                    gIndexesUsed(iIU) = 1
                    Return ChrW(iIU + &H60 - 25)
                End If
            Next
            For iIU = 1 To 25
                If gIndexesUsed(iIU) = 0 Then
                    gIndexesUsed(iIU) = 1
                    Return ChrW(iIU + &H3B0)
                End If
            Next
        End If
        'very bad luck, no indices left
        Return ChrW(&H2605)  'black star
    End Function
    Sub InsertPlusMinus(InsertionPoint As Word.Range, Sign1 As String, Sign2 As String)
        'Insert plus or minus sepending on ...
        Dim ActualSign As String
        If Sign1 = "-" Then
            If Sign2 = "-" Then ActualSign = "+" Else ActualSign = "-"
        Else
            ActualSign = Sign2
        End If
        Call InsertText(InsertionPoint, ActualSign)
    End Sub
End Module
Module CalculationFunctions
    Function IsStrMinus(Term As String) As Boolean
        'Aug 2022, after experiment with TestSomething7, we have
        If Len(Term) = 0 Then Return False
        If Asc(Mid(Term, 1, 1)) = 45 Then Return True
        Return False
        'from VBA we had ...
        'something smells about - signs
        IsStrMinus = False
        If Len(Term) = 0 Then Exit Function
        If (Mid(Term, 1, 1) = "-") Or (Asc(Mid(Term, 1, 1)) = 45) Then IsStrMinus = True
    End Function
End Module
Module InsertFunctions
    '**************************************
    'Functions for inserting various things in equations and while keeping track of insertion point
    '**************************************
    Public Function InsertBrackets(ByRef Functions As Word.OMathFunctions,
                                   ByRef InsertionPoint As Word.Range, ByRef LinearEq As String) As Word.OMath
        'Insert brackets at InsertionPoint and put LinearEq, built up, inside brackets, if not ""
        'Move InsertionPoint to inside bracket
        'return omath object inside bracket
        Dim NewMathTerm As Word.OMathFunction
        Dim InnerEq As Word.OMath

        NewMathTerm = InsertFunction(Functions, InsertionPoint, Word.WdOMathFunctionType.wdOMathFunctionDelim)
        InnerEq = NewMathTerm.Delim.E.Item(1)
        If LinearEq <> "" Then
            InnerEq.Range.Text = LinearEq
            InnerEq.BuildUp()
        End If
        InsertionPoint.Start = InnerEq.Range.Start
        InsertionPoint.End = InnerEq.Range.Start
        Return InnerEq
    End Function
    Public Sub InsertFraction(ByRef Functions As Word.OMathFunctions, ByRef InsertionPoint As Word.Range, Num As String, Den As String)
        Dim NewMathTerm As Word.OMathFunction
        NewMathTerm = InsertFunction(Functions, InsertionPoint, Word.WdOMathFunctionType.wdOMathFunctionFrac)
        NewMathTerm.Frac.Num.Range.Text = Num
        NewMathTerm.Frac.Den.Range.Text = Den
        NewMathTerm.Frac.Type = Word.WdOMathFracType.wdOMathFracBar
    End Sub
    Public Sub InsertMetric(ByRef Functions As Word.OMathFunctions, ByRef InsertionPoint As Word.Range, Indices As String)
        Dim NewMathTerm As Word.OMathFunction
        NewMathTerm = InsertFunction(Functions, InsertionPoint, Word.WdOMathFunctionType.wdOMathFunctionScrSub)
        NewMathTerm.ScrSub.E.Range.Text = "g"
        NewMathTerm.ScrSub.Sub.Range.Text = Indices
    End Sub
    Public Sub InsertInverseMetric(ByRef Functions As Word.OMathFunctions, ByRef InsertionPoint As Word.Range, Indices As String)
        Dim NewMathTerm As Word.OMathFunction
        NewMathTerm = InsertFunction(Functions, InsertionPoint, Word.WdOMathFunctionType.wdOMathFunctionScrSup)
        NewMathTerm.ScrSup.E.Range.Text = "g"
        NewMathTerm.ScrSup.Sup.Range.Text = Indices
    End Sub
    Public Sub InsertGamma(ByRef Functions As Word.OMathFunctions, ByRef InsertionPoint As Word.Range, UpIndex As String, DownIndices As String)
        'Christoffel symbol
        Dim NewMathTerm As Word.OMathFunction
        NewMathTerm = InsertFunction(Functions, InsertionPoint, Word.WdOMathFunctionType.wdOMathFunctionScrSubSup)
        NewMathTerm.ScrSubSup.E.Range.Text = gCapGamma
        NewMathTerm.ScrSubSup.E.Range.Italic = False
        NewMathTerm.ScrSubSup.Sup.Range.Text = UpIndex
        NewMathTerm.ScrSubSup.Sub.Range.Text = DownIndices
    End Sub
    Public Sub InsertPartial(ByRef Functions As Word.OMathFunctions, ByRef InsertionPoint As Word.Range, DownIndex As String)
        'Partial derivative
        Dim NewMathTerm As Word.OMathFunction

        NewMathTerm = InsertFunction(Functions, InsertionPoint, Word.WdOMathFunctionType.wdOMathFunctionScrSub)
        NewMathTerm.ScrSub.E.Range.Text = gPartialDerivative
        NewMathTerm.ScrSub.Sub.Range.Text = DownIndex
    End Sub
    Public Function InsertFunction(ByRef Functions As Word.OMathFunctions, ByRef InsertionPoint As Word.Range, FuncType As Word.WdOMathFunctionType) As Word.OMathFunction
        'Add a function in Functions at InsertionPoint and move InsertionPoint to after the function, ready for next
        Dim NewFunction As Word.OMathFunction
        NewFunction = Functions.Add(InsertionPoint, FuncType)
        InsertionPoint = NewFunction.Range
        InsertionPoint.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
        Return NewFunction
    End Function
    Sub InsertText(ByRef InsertionPoint As Word.Range, MyText As String)
        'Inserts MyText at InsertionPoint and moves InsertionPoint to after that text
        'so this is very similar to
        'Equation.Functions.Add(InsertionPoint, wdOMathFunctionNormalText) .... which does not work!!
        'many thankss to jpl for this. https://www.msofficeforums.com/word-vba/31587-vba-omath-object.html
        InsertionPoint.Text = MyText
        InsertionPoint.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
    End Sub
    Sub InsertMatrix(Equation As Word.OMath, ByRef InsertPoint As Word.Range, IDim As Integer, Matrix(,) As String)
        'insert =(M)
        'where M is a square matrix dimension IDim with elements given in Matrix
        Dim oMatrix As Word.OMathFunction

        InsertText(InsertPoint, "=")
        Equation = InsertBrackets(Equation.Functions, InsertPoint, "")

        oMatrix = Equation.Functions.Add(Equation.Range, Word.WdOMathFunctionType.wdOMathFunctionMat)
        'now oMatrix is the matrix inside a bracket...
        For iRow = 2 To IDim
            oMatrix.Mat.Cols.Add()
            oMatrix.Mat.Rows.Add()
        Next
        For iRow = 1 To IDim
            For iCol = 1 To IDim
                SetLinear(oMatrix.Mat.Cell(iRow, iCol), Matrix(iRow, iCol))
            Next
        Next
        InsertPoint = Equation.Range
        InsertPoint.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
    End Sub
End Module