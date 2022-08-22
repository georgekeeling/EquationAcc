Imports Microsoft.Office.Interop.Word

Public Class EquationAcc
    Public MyRibbon As Office.IRibbonUI
    Public SelChars As Integer
    Public InEquation As Boolean
    Public InTable As Boolean
    Private Sub ThisAddIn_Startup() Handles Me.Startup
        StartAll()
    End Sub

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown

    End Sub
    Protected Overrides Function CreateRibbonExtensibilityObject() As Microsoft.Office.Core.IRibbonExtensibility
        'Added in as instructed in Ribbon.vb
        Return New Ribbon()
    End Function
    Private Sub Application_WindowSelectionChange(Sel As Selection) Handles Application.WindowSelectionChange
        'Causes havoc if there are breakpoints in this or Application_WindowActivate
        SelChars = Len(Sel.Range.Text)
        If Sel.OMaths.Count > 0 Then
            InEquation = True
        Else
            InEquation = False
        End If
        InTable = Sel.Information(Word.WdInformation.wdWithInTable)
        If MyRibbon IsNot Nothing Then
            'sometimes MyRibbon is nothing. Not sure why
            MyRibbon.Invalidate()       'fires off getEnabled/onGetEnabledX functions (twice?). Does not fire loadImage
        End If
    End Sub
    Private Sub Application_WindowActivate(Doc As Document, Wn As Window) Handles Application.WindowActivate
        'get menu enablement right when we start.
        'Causes havoc if there are breakpoints in this or Application_WindowSelectionChange
        Application_WindowSelectionChange(Wn.Selection)
    End Sub
End Class
