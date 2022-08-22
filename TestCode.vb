Imports System.Diagnostics      'for Debug 
#If DEBUG Then
Module TestCode
    'test code in here which may be driven off BtnTest if it is visible! 
    Public Sub TestSomething()
        'Can I activate another ribbon? Can I see what their ID's are? N0t Yet!
        Dim OtherRibbon As Office.IRibbonUI
        Dim Test As String
        Dim iRibbons As Integer = Globals.Ribbons.Count

        For Each OtherRibbon In Globals.Ribbons
            Test = OtherRibbon.ToString()
        Next
        'Globals.EquationAcc.MyRibbon.ActivateTabMso("TabAddIns")

    End Sub
    Public Sub TestSomething10()
        'italic and non-italic in equation 
        Dim Selection As Word.Selection
        Dim objRange As Word.Range
        Dim Equation As Word.OMath

        Selection = Globals.EquationAcc.Application.ActiveDocument.ActiveWindow.Selection
        Selection.TypeText(" ")
        Selection.MoveLeft(Word.WdUnits.wdCharacter, 1)
        objRange = Selection.Range
        Selection.OMaths.Add(objRange)
        Equation = Selection.OMaths(1)
        Equation.Type = Word.WdOMathType.wdOMathInline
        'Default  Equation.Range.Italic is -1. If left at -1, equation will be all italics
        Equation.Range.Italic = 0  'Now equation can be mix of italics and non-italics. 
        Equation.Range.Text = "〖𝑑𝑠〗^2=〖d𝑥〗^2"
        'Now Equation.Range.Italic = 9999999
        Equation.BuildUp()
    End Sub
    Public Sub TestSomething9()
        Dim Selection As Word.Selection
        Dim Equation As Word.OMath
        Dim EquationS As String
        Dim Action As Integer = 1       'change value of option in debugger
        Dim lTest As Long

        Selection = Globals.EquationAcc.Application.ActiveDocument.ActiveWindow.Selection
        Equation = Selection.OMaths(1)          'equation is d𝑧𝑑𝑥
        If Action = 1 Then
            Equation.Linearize()
            EquationS = Equation.Range.Text
            lTest = Equation.Range.Font.Italic
            Equation.Range.Font.Italic = Word.WdConstants.wdToggle
            lTest = Equation.Range.Font.Italic
            Equation.BuildUp()
        End If
        If Action = 2 Then
            Equation.Linearize()
            EquationS = Equation.Range.Text
            Equation.Range.Text = EquationS
            Equation.BuildUp()                  'equation is d𝑧𝑑𝑥
        End If
        If Action = 3 Then
            TestSomething8()
        End If

    End Sub
    Public Sub TestSomething8()
        'Copy equation using Getlinear / SetLinear. Copy equation is all italics
        'Problem is in SetLinear: Equation.Range.Text = LinearEq
        'only italics arrive in Equation.Range.Text
        Dim Selection As Word.Selection
        Dim Equation As Word.OMath
        Dim EquationS As String
        Dim FieldCodes As Boolean

        Selection = Globals.EquationAcc.Application.ActiveDocument.ActiveWindow.Selection
        Equation = Selection.OMaths(1)          'equation is d𝑧𝑑𝑥
        EquationS = GetLinear(Equation)         'EquationS is "d𝑧𝑑𝑥"
        FieldCodes = Equation.Range.TextRetrievalMode.IncludeFieldCodes
        Selection.MoveDown(Word.WdUnits.wdLine, 1)
        InsertEquationNewLine()
        Selection = Globals.EquationAcc.Application.ActiveDocument.ActiveWindow.Selection
        Equation = Selection.OMaths(1)
        SetLinear(Equation, EquationS)          'New equation is "d𝑧𝑑𝑥". After fix in SetLinear 
        Return

        Equation.Range.Italic = 0
        Equation.Range.Text = EquationS
        Equation.BuildUp()
        Return
        'and this is no better
        Equation.Range.FormattedText.Text = EquationS
        Equation.BuildUp()
    End Sub
    Public Sub TestSomething7()
        'Diagnosing the difference between the two ways of doing minus. See Function IsStrMinus
        Dim Selection As Word.Selection
        Dim Result As Boolean
        Selection = Globals.EquationAcc.Application.ActiveDocument.ActiveWindow.Selection
        Selection.Range.Text = "chr (45) is " & Chr(45) & vbCrLf
        Selection.MoveDown(Word.WdUnits.wdLine, 1)
        'copy three lines below into Word or Notepad to see the differnce 
        Selection.Range.Text = "asc('-') is " & Asc("-") & vbCrLf
        Selection.MoveDown(Word.WdUnits.wdLine, 1)
        Selection.Range.Text = "other asc('−') is " & Asc("−") & vbCrLf
        Selection.MoveDown(Word.WdUnits.wdLine, 1)
        Result = IsStrMinus("")
    End Sub
    Public Sub TestSomething6()
        'Unexpected result for integer divide
        Dim dHalf As Double, iNumber As Integer, iHalf As Integer
        For iNumber = 1 To 11
            iHalf = iNumber / 2
            Debug.WriteLine("'Half " & iNumber & " = " & iHalf)
        Next
        'result is
        'Half 1 = 0
        'Half 2 = 1
        'Half 3 = 2
        'Half 4 = 2
        'Half 5 = 2
        'Half 6 = 3
        'Half 7 = 4
        'Half 8 = 4
        'Half 9 = 4
        'Half 10 = 5
        'Half 11 = 6
        For iNumber = 1 To 11
            dHalf = iNumber / 2
            Debug.WriteLine("'Half " & iNumber & " = " & dHalf)
        Next
        'result is
        'Half 1 = 0.5
        'Half 2 = 1
        'Half 3 = 1.5
        'Half 4 = 2
        'Half 5 = 2.5
        'Half 6 = 3
        'Half 7 = 3.5
        'Half 8 = 4
        'Half 9 = 4.5
        'Half 10 = 5
        'Half 11 = 5.5

    End Sub

    Public Sub TestSomething5()
        'testing InsertMatrix
        Dim Selection As Word.Selection
        Dim objRange As Word.Range
        Selection = Globals.EquationAcc.Application.ActiveDocument.ActiveWindow.Selection

        objRange = Selection.Range
        objRange.Text = " "
        objRange.OMaths.Add(objRange)       'creates equation
        InsertMatrix(objRange.OMaths(1), objRange, gNmetricDimension, gMetric)
    End Sub
    Public Sub TestSomething4()
        'convert 1,1 component into its reciprocal at insertion point
        'add linear version below
        Dim Selection As Word.Selection
        Dim objRange As Word.Range
        Selection = Globals.EquationAcc.Application.ActiveDocument.ActiveWindow.Selection

        objRange = Selection.Range
        objRange.Text = gMetric(1, 1)       'inserts the text e.g 𝑥^2 
        objRange.OMaths.Add(objRange)       'creates equation of text in nasty format: also 𝑥^2
        objRange.OMaths(1).BuildUp()        'makes the equation pretty (but size 10 not 12). e.g. 𝑥²
        InvertOmath(objRange.OMaths(1))     'takes reciprocal
        Selection.MoveDown(Word.WdUnits.wdLine, 1)
        Selection.TypeText(GetCleanLinear(objRange.OMaths(1)))
    End Sub
    Public Sub TestSomething3()
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
        gMetricNCoordsString = ""
        gInvMetricNCoordsString = ""
        gCoordinatesString = ""
        Globals.EquationAcc.MyRibbon.Invalidate()
        TensorMessage("Cleared metrics and coordinates")
    End Sub
    Public Sub TestSomething2()
        'Take selection, go down a line, write out character code for selection
        Dim Selection As Word.Selection
        Dim MyString As String, iX As Integer, iChar As Integer

        MyString = ""
        Selection = Globals.EquationAcc.Application.ActiveDocument.ActiveWindow.Selection
        For iX = 1 To Len(Selection.Range.Text)
            iChar = AscW(Mid(Selection.Range.Text, iX, 1))
            If iChar <> &HD835 Then MyString &= Hex(iChar) & "," 'Skip D835's which are often every other character
        Next

        Selection.MoveDown(Word.WdUnits.wdLine)
        Selection.Range.Text = MyString
    End Sub
    Public Sub TestSomething1()
        'Use CleanString on selection. Write cleaned string on line below
        Dim Selection As Word.Selection
        Dim MyString As String

        Selection = Globals.EquationAcc.Application.ActiveDocument.ActiveWindow.Selection
        CleanString(Selection.Range, MyString)
        Selection.MoveDown(Word.WdUnits.wdLine)
        Selection.Range.Text = MyString

    End Sub
End Module
#End If
