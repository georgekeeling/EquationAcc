'Add in to help with equation editing for Office 365 (2019)
'By George Keeling and on my blog at www.general-relativity.net
'Possibly documented at https://www.general-relativity.net/search/label/Tools
'Feel free to use, copy, modify and give away but not for commerce. Please acknowledge me

'Entry points
'	StrikeOutCross1, StrikeOutBLTR, StrikeOutTLBR - Adornment functions
'	SetColour, InsertCrossTick - Adornment functions
'	StartAll - Initialisation
'Also contains
'	CleanString - Translage from Omath encoding to vanilla with no encoding
'	SetLinear, GetLinear - translating from Omath to linear math And back
'	TensorError - error messages And other messages

Module Adornments
    Public Const gCapGamma As String = "Γ"
    Public Const gPartialDerivative = "∂"
    Public Const gNabla As String = "∇"
    Public Const gDoubleQuote As String = """"

    Public gLogErrors As String = ""                    'error log for current operation
    Public gBeeps As Boolean = True
    Public gTLightState As Byte = 2     '1 = red, 2 = none , 3 = green, 4 = amber
    Sub StrikeOutCross1()
        Dim Selection As Word.Selection
        Dim objUndo As Word.UndoRecord

        Selection = Globals.EquationAcc.Application.ActiveDocument.ActiveWindow.Selection
        If Not InEquation(Selection) Then
            Exit Sub
        End If
        objUndo = Globals.EquationAcc.Application.UndoRecord
        objUndo.StartCustomRecord("Strike out")
        With Selection.OMaths(1).Functions.Add(Selection.Range, Word.WdOMathFunctionType.wdOMathFunctionBorderBox)
            .BorderBox.HideTop = True
            .BorderBox.HideBot = True
            .BorderBox.HideLeft = True
            .BorderBox.HideRight = True
            .BorderBox.StrikeH = False
            .BorderBox.StrikeV = False
            .BorderBox.StrikeBLTR = True
            .BorderBox.StrikeTLBR = True
        End With
        objUndo.EndCustomRecord()
    End Sub
    Sub StrikeOutBLTR()
        Dim Selection As Word.Selection
        Dim objUndo As Word.UndoRecord

        Selection = Globals.EquationAcc.Application.ActiveDocument.ActiveWindow.Selection
        If Not InEquation(Selection) Then
            Exit Sub
        End If
        objUndo = Globals.EquationAcc.Application.UndoRecord
        objUndo.StartCustomRecord("Strike out")
        With Selection.OMaths(1).Functions.Add(Selection.Range, Word.WdOMathFunctionType.wdOMathFunctionBorderBox)
            .BorderBox.HideTop = True
            .BorderBox.HideBot = True
            .BorderBox.HideLeft = True
            .BorderBox.HideRight = True
            .BorderBox.StrikeH = False
            .BorderBox.StrikeV = False
            .BorderBox.StrikeBLTR = True
            .BorderBox.StrikeTLBR = False
        End With
        objUndo.EndCustomRecord()
    End Sub
    Sub StrikeOutTLBR()
        Dim Selection As Word.Selection
        Dim objUndo As Word.UndoRecord

        Selection = Globals.EquationAcc.Application.ActiveDocument.ActiveWindow.Selection
        If Not InEquation(Selection) Then
            Exit Sub
        End If
        objUndo = Globals.EquationAcc.Application.UndoRecord
        objUndo.StartCustomRecord("Strike out")
        With Selection.OMaths(1).Functions.Add(Selection.Range, Word.WdOMathFunctionType.wdOMathFunctionBorderBox)
            .BorderBox.HideTop = True
            .BorderBox.HideBot = True
            .BorderBox.HideLeft = True
            .BorderBox.HideRight = True
            .BorderBox.StrikeH = False
            .BorderBox.StrikeV = False
            .BorderBox.StrikeBLTR = False
            .BorderBox.StrikeTLBR = True
        End With
        objUndo.EndCustomRecord()
    End Sub
    Sub SetColour(RGBcol As Integer)
        'set selected text to colour. Black = RGB(0, 0, 0), red = RGB(255, 0, 0),
        'green = 5287936 = hex 50B000 
        Dim oWin As Word.Window
        oWin = Globals.EquationAcc.Application.ActiveDocument.ActiveWindow
        oWin.Selection.Range.Font.TextColor.RGB = RGBcol
    End Sub
    Sub InsertCrossTick(RGBcol As Integer, Charachter As Integer)
        Dim Selection As Word.Selection
        Dim objUndo As Word.UndoRecord

        objUndo = Globals.EquationAcc.Application.UndoRecord
        objUndo.StartCustomRecord("Ticky")

        Selection = Globals.EquationAcc.Application.ActiveDocument.ActiveWindow.Selection

        Selection.TypeText(Text:=" ")
        Selection.MoveLeft(Word.WdUnits.wdCharacter, Count:=1)
        Selection.InsertSymbol(Charachter, "MS Gothic", True)
        Selection.MoveLeft(Word.WdUnits.wdCharacter, Count:=1, Word.WdMovementType.wdExtend)
        'oSelection.MoveLeft(oUnit, 1,)
        SetColour(RGBcol)
        Selection.MoveRight(Word.WdUnits.wdCharacter, Count:=1)
        objUndo.EndCustomRecord()
    End Sub
    Sub StartAll()
        gCoordinates(1) = ""        'Blank in these indicates undefined.
        gMetric(1, 1) = ""
        gInvMetric(1, 1) = ""
        EquationColumnWidth = 1.38
        EquationFontSize = 12
        TableHeight = 1.29
    End Sub
    Function InEquation(Selection As Word.Selection) As Boolean
        If Selection.OMaths.Count = 0 Then
            TensorError("Selection must be in equation.")
            InEquation = False
            Exit Function
        End If
        InEquation = True
    End Function
    '***********************************************
    'Stuff for messages in ribbon
    Sub ClearError()
        UpdateMessageBox(2, "")
    End Sub
    Public Sub TensorMessage(Message As String)
        UpdateMessageBox(3, Message)
        If gBeeps Then Beep()
    End Sub
    Public Sub TensorProgress(Message As String)
        UpdateMessageBox(3, Message)
    End Sub

    Public Sub HighlightError(ErrorRange As Word.Range, ErrorMessage As String)
        'Highlight an error and, maybe, log error with ErrorMessage somewhere
        ErrorRange.HighlightColorIndex = Word.WdColorIndex.wdYellow
        TensorError(ErrorMessage)
    End Sub
    Sub TensorError(Message As String)
        UpdateMessageBox(1, Message)
        If gBeeps Then Beep()
    End Sub
    Sub UpdateMessageBox(TLightState As Byte, Message As String)
        gTLightState = TLightState
        gLogErrors = Message
        Globals.EquationAcc.MyRibbon.InvalidateControl("EqMessage")
        Globals.EquationAcc.MyRibbon.InvalidateControl("TrafficLight")
    End Sub
End Module

