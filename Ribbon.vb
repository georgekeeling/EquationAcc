'TO DO:  Follow these steps to enable the Ribbon (XML) item:

'1: Copy the following code block into the ThisAddin, ThisWorkbook, or ThisDocument class.

'Protected Overrides Function CreateRibbonExtensibilityObject() As Microsoft.Office.Core.IRibbonExtensibility
'    Return New Ribbon()
'End Function

'2. Create callback methods in the "Ribbon Callbacks" region of this class to handle user
'   actions, such as clicking a button. Note: if you have exported this Ribbon from the
'   Ribbon designer, move your code from the event handlers to the callback methods and
'   modify the code to work with the Ribbon extensibility (RibbonX) programming model.

'3. Assign attributes to the control tags in the Ribbon XML file to identify the appropriate callback methods in your code.

'For more information, see the Ribbon XML documentation in the Visual Studio Tools for Office Help.

Imports Microsoft.Office.Core
Imports System.Drawing          'needed for Bitmap
Imports System.Diagnostics      'for Debug (not used now)
Imports System.Collections      'needed for DictionaryEntry
Imports System.Globalization    'needed for CultureInfo
Imports System.Reflection       'needed for Assembly
Imports System.Resources        'needed for ResourceManager, ResourceSet
Imports stdole                  'needed for IPictureDisp (not used now)

<Runtime.InteropServices.ComVisible(True)>
Public Class Ribbon
    Implements Office.IRibbonExtensibility
    ReadOnly ErrorBitmap As New Bitmap(My.Resources._Error)
    ReadOnly TLightGreen As New Bitmap(My.Resources.TLightGreen)
    ReadOnly TLightRed As New Bitmap(My.Resources.TLightRed)
    ReadOnly TLightAmber As New Bitmap(My.Resources.TLightAmber)
    ReadOnly ResMan As New ResourceManager("EquationAcc.Resources1", Assembly.GetExecutingAssembly)
    Public Sub New()
    End Sub
    Public Function GetCustomUI(ByVal ribbonID As String) As String Implements Office.IRibbonExtensibility.GetCustomUI
        Return GetResourceText("EquationAcc.Ribbon.xml")
    End Function

#Region "Ribbon Callbacks"
    'Create callback methods here. For more information about adding callback methods,
    'visit https://go.microsoft.com/fwlink/?LinkID=271226
    'one place that takes you to is
    'https://docs.microsoft.com/en-us/previous-versions/office/developer/office-2007/aa722523(v=office.12)/ 
    'which contans the signatures of all callback functions about halfway down
    Public Sub Ribbon_Load(ByVal ribbonUI As Office.IRibbonUI)
        'This is called on load bexause in XML we have onLoad="Ribbon_Load"
        Globals.EquationAcc.MyRibbon = ribbonUI
    End Sub
    Public Function GetButtonImages(imageId As String) As Bitmap  'was IPictureDisp as specified in documenation 
        'Assigns images to buttons and groups. Gets called once at start
        'to enumerate resources
        'see https://stackoverflow.com/questions/4656883/how-enumerate-resources-inside-a-resx-file-programmatically
        Dim ResSet As ResourceSet
        ResSet = ResMan.GetResourceSet(CultureInfo.CurrentCulture, True, True)
        For Each Item As DictionaryEntry In ResSet
            'Debug.WriteLine(item.Key)       'appears in window from Debug / windows / output
            If Item.Key = imageId Then
                Return Item.Value
            End If
        Next
        Return ErrorBitmap
    End Function

    'Functions to update messages and other info 
    Public Function ShowMetric(control As IRibbonControl) As String
        Return gMetricNCoordsString
    End Function
    Public Function ShowInvMetric(control As IRibbonControl) As String
        'Debug.Write("ShowInvMetric ") : Debug.WriteLine(gInvMetricNCoordsString)
        Return gInvMetricNCoordsString
    End Function
    Public Function ShowCoords(control As IRibbonControl) As String
        Return gCoordinatesString
    End Function
    Public Function UpdateMessage(control As IRibbonControl) As String
        Return gLogErrors
    End Function
    Public Function SetTrafficLight(control As IRibbonControl) As Bitmap
        If gTLightState = 1 Then Return TLightRed
        If gTLightState = 3 Then Return TLightGreen
        If gTLightState = 4 Then Return TLightAmber
        Return Nothing
    End Function
    Public Sub BeepsCheck(control As IRibbonControl, pressed As Boolean)
        gBeeps = pressed
    End Sub
    Function BeepsSet(control As IRibbonControl) As Boolean
        'tempInt += 1
        Return gBeeps
        'if you have a breakpoint here it gets calles interminably
        'if you don't it doesnt as you can see by commenting in tempInt
    End Function

    'Functions for enabling / disabling buttons depending on selection. Triggered by Application_WindowSelectionChange 
    Public Function onGetEnabled1(control As IRibbonControl) As Boolean
        'numbered equations: btnInsEq2 / btnInsEq3; write metrics / Christoffels
        If Globals.EquationAcc.InEquation Or Globals.EquationAcc.InTable Then Return False
        Return True
    End Function
    Public Function onGetEnabled2(control As IRibbonControl) As Boolean
        'unnumbered equations: btnInsEqInLine / btnInsEqNewLine
        If Globals.EquationAcc.InEquation Then Return False
        Return True
    End Function
    Public Function onGetEnabled3(control As IRibbonControl) As Boolean
        'colour lettering
        If Globals.EquationAcc.SelChars = 0 Then Return False
        Return True
    End Function
    Public Function onGetEnabled4(control As IRibbonControl) As Boolean
        'strike out in equation 
        If Globals.EquationAcc.SelChars = 0 Or (Not Globals.EquationAcc.InEquation) Then Return False
        Return True
    End Function
    Public Function onGetEnabled5(control As IRibbonControl) As Boolean
        'table functions
        Return Globals.EquationAcc.InTable
    End Function
    Public Function onGetEnabled6(control As IRibbonControl) As Boolean
        'Expand equation, pick up coordinates,, metrics
        Return Globals.EquationAcc.InEquation
    End Function

    'Functions that respond to buttons. BeepsCheck might have been here but its not.
    Public Sub BtnInsEq2_Click(control As IRibbonControl)
        ClearError()
        EquationTable2()
    End Sub
    Public Sub BtnInsEq3_Click(control As IRibbonControl)
        ClearError()
        EquationTable3()
    End Sub
    Public Sub BtnInsEqInLine_Click(control As IRibbonControl)
        ClearError()
        InsertEquationInLine()
    End Sub
    Public Sub BtnInsEqNewLine_Click(control As IRibbonControl)
        ClearError()
        InsertEquationNewLine()
    End Sub
    Public Sub BtnInsertTick_Click(control As IRibbonControl)
        ClearError()
        InsertCrossTick(5287936, 10003)
    End Sub
    Public Sub ButtonInsertCross_Click(control As IRibbonControl)
        ClearError()
        InsertCrossTick(RGB(255, 0, 0), 10008)
    End Sub
    Public Sub BtnTextRed_Click(control As IRibbonControl)
        ClearError()
        SetColour(RGB(255, 0, 0))
    End Sub
    Public Sub BtnTextGreen_Click(control As IRibbonControl)
        ClearError()
        SetColour(5287936)
    End Sub
    Public Sub BtnTextBlack_Click(control As IRibbonControl)
        ClearError()
        SetColour(RGB(0, 0, 0))
    End Sub
    Public Sub BtnStrikeCross_Click(control As IRibbonControl)
        ClearError()
        StrikeOutCross1()
    End Sub
    Public Sub BtnStrikeSlash_Click(control As IRibbonControl)
        ClearError()
        StrikeOutBLTR()
    End Sub
    Public Sub BtnStrikeBackSlash_Click(control As IRibbonControl)
        ClearError()
        StrikeOutTLBR()
    End Sub
    Public Sub BtnTableBordersOne_Click(control As IRibbonControl)
        ClearError()
        BordersOutside()
    End Sub
    Public Sub BtnTableRowsToggle_Click(control As IRibbonControl)
        ClearError()
        Point8cmTableRows()
    End Sub
    Public Sub BtnTableBordersNone_Click(control As IRibbonControl)
        ClearError()
        BordersNone()
    End Sub
    Public Sub BtnTableBordersAll_Click(control As IRibbonControl)
        ClearError()
        BordersAll()
    End Sub
    Public Sub BtnTableSelect_Click(control As IRibbonControl)
        ClearError()
        SelectTable()
    End Sub
    Public Async Sub BtnWriteLatex_Click(control As IRibbonControl)
        ClearError()
        Await PrepareForWeb0Async()
    End Sub
    Public Async Sub BtnRenumber_Click(control As IRibbonControl)
        ClearError()
        Await RenumberEquationsCodeAsync()
    End Sub
    Public Sub BtnExpand_Click(control As IRibbonControl)
        ClearError()
        ExpandSymbols()
    End Sub
#If DEBUG Then
    Public Sub BtnTest(control As IRibbonControl)
        TestSomething()
    End Sub
#End If
    Public Sub BtnCoordinates_Click(control As IRibbonControl)
        ClearError()
        PickUpCoordinates()
    End Sub
    Public Sub BtnMetric_Click(control As IRibbonControl)
        ClearError()
        Call PickUpEitherMetric(gNmetricDimension, gMetric, gMetricNCoordsString, "Metric",
                                                   gInvMetric, gInvMetricNCoordsString, "Inv.Metric")
    End Sub
    Public Sub BtnInvMetric_Click(control As IRibbonControl)
        ClearError()
        PickUpEitherMetric(gNInvmetricDimension, gInvMetric, gInvMetricNCoordsString, "Inv.Metric",
                                                 gMetric, gMetricNCoordsString, "Metric")
    End Sub
    Public Sub BtnWriteMetrics_Click(control As IRibbonControl)
        ClearError()
        WriteMetrics()
    End Sub
    Public Async Sub BtnWriteChris_Click(control As IRibbonControl)
        ClearError()
        Await WriteChristoffelSymbolsAsync()
        'WriteChristoffelSymbolsAsync()
    End Sub
    Public Sub BtnClearMetric_Click(control As IRibbonControl)
        ClearError()
        ClearMetrics()
    End Sub
#End Region

#Region "Helpers"
    Private Shared Function GetResourceText(ByVal resourceName As String) As String
        Dim asm As Reflection.Assembly = Reflection.Assembly.GetExecutingAssembly()
        Dim resourceNames() As String = asm.GetManifestResourceNames()
        For i As Integer = 0 To resourceNames.Length - 1
            If String.Compare(resourceName, resourceNames(i), StringComparison.OrdinalIgnoreCase) = 0 Then
                Using resourceReader As IO.StreamReader = New IO.StreamReader(asm.GetManifestResourceStream(resourceNames(i)))
                    If resourceReader IsNot Nothing Then
                        Return resourceReader.ReadToEnd()
                    End If
                End Using
            End If
        Next
        Return Nothing
    End Function
#End Region

End Class
