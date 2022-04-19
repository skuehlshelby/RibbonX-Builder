Imports System.Reflection
Imports Microsoft.Office.Core
Imports RibbonFactory
Imports RibbonFactory.Controls
Imports RibbonFactory.ControlInterfaces
Imports RibbonFactory.Controls
Imports RibbonFactory.Enums
Imports RibbonFactory.Utilities.Testing
Imports stdole

Public MustInherit Class RibbonTestBase
    Implements IRibbonExtensibility
    Implements ICreateCallbacks

    Protected Ribbon As Ribbon
    Protected OfficeHost As OfficeHostAppilcation
    Protected ReadOnly ImageCache As IDictionary(Of String, IPictureDisp) = New Dictionary(Of String, IPictureDisp)
    Protected ReadOnly ControlGenerator As RandomControlGenerator

    Public Sub New()
        ControlGenerator = New RandomControlGenerator(Me)
    End Sub

    <TestInitialize()>
    Public Overridable Sub Initialize()
        Ribbon = CreateRibbon()
        OfficeHost = New OfficeHostAppilcation(Me)
        Ribbon.AssignRibbonUI(OfficeHost)
    End Sub

    Protected MustOverride Function CreateRibbon() As Ribbon

    <TestMethod>
    Public Sub ContainsNoNullValues()
        For Each control As RibbonElement In Ribbon
            Dim controlType As Type = control.GetType()

            For Each propertyInfo As PropertyInfo In controlType.GetProperties()
                If Not String.Equals(propertyInfo.Name, NameOf(RibbonElement.Tag)) Then
                    Assert.IsNotNull(propertyInfo.GetValue(control), $"Property '{propertyInfo.Name}' on {control.ID} is null.")
                End If
            Next
        Next
    End Sub

    <TestMethod>
    Public Sub ProducesLegalRibbonX()
        Dim errors As RibbonErrorLog = Ribbon.GetErrors()

        Debug.WriteLine(Ribbon.RibbonX)

        Assert.IsTrue(errors.None, String.Join(Environment.NewLine, errors.Errors))
    End Sub

    <TestMethod>
    Public Sub PropertiesAreMappedCorrectly()
        For Each dictionary As KeyValuePair(Of RibbonElement, Dictionary(Of String, Object)) In ControlGenerator.GetCreationLogs()
            Dim control As RibbonElement = dictionary.Key
            Dim controlType As Type = control.GetType()

            For Each propertyValuePair As KeyValuePair(Of String, Object) In dictionary.Value
                Dim propertyInfo As PropertyInfo = controlType.GetProperty(propertyValuePair.Key)

                If propertyInfo IsNot Nothing Then 'The control has this property
                    Dim valueAtCreationTime As Object = propertyValuePair.Value
                    Dim currentValue As Object = propertyInfo.GetValue(control)

                    Assert.AreEqual(valueAtCreationTime, currentValue, $"Property '{propertyInfo.Name}' on {control.ID} is '{currentValue}', but '{valueAtCreationTime}' was expected instead.")
                End If
            Next
        Next
    End Sub

    <TestCleanup()>
    Public Overridable Sub Cleanup()
        Ribbon = Nothing
        ControlGenerator.ClearLogs()
    End Sub

    Protected Function MakeRibbonWithOneTabAndOneGroup(ParamArray controls As RibbonElement()) As Ribbon
        Return ControlGenerator.MakeRibbon(ControlGenerator.MakeTab(ControlGenerator.MakeGroup(controls)))
    End Function


#Region "IRibbonExtensibility Members"

    Private Function GetCustomUI(ribbonID As String) As String Implements IRibbonExtensibility.GetCustomUI
        Return Ribbon.RibbonX
    End Function

#End Region

#Region "ICreateCallbacks"

    Public Sub OnLoad(ribbonUI As IRibbonUI) Implements ICreateCallbacks.OnLoad
        Ribbon.AssignRibbonUI(ribbonUI)
    End Sub

    Public Function LoadImage(imageID As String) As IPictureDisp Implements ICreateCallbacks.LoadImage
        Return ImageCache(imageID)
    End Function

    Public Function GetEnabled(control As IRibbonControl) As Boolean Implements ICreateCallbacks.GetEnabled
        Return Ribbon.GetElement(Of IEnabled)(control.Id).Enabled
    End Function

    Public Function GetVisible(control As IRibbonControl) As Boolean Implements ICreateCallbacks.GetVisible
        Return Ribbon.GetElement(Of IVisible)(control.Id).Visible
    End Function

    Public Function GetLabel(control As IRibbonControl) As String Implements ICreateCallbacks.GetLabel
        Return Ribbon.GetElement(Of ILabel)(control.Id).Label
    End Function

    Public Function GetShowLabel(control As IRibbonControl) As Boolean Implements ICreateCallbacks.GetShowLabel
        Return Ribbon.GetElement(Of IShowLabel)(control.Id).ShowLabel
    End Function

    Public Function GetDescription(control As IRibbonControl) As String Implements ICreateCallbacks.GetDescription
        Return Ribbon.GetElement(Of IDescription)(control.Id).Description
    End Function

    Public Function GetScreenTip(control As IRibbonControl) As String Implements ICreateCallbacks.GetScreenTip
        Return Ribbon.GetElement(Of IScreenTip)(control.Id).ScreenTip
    End Function

    Public Function GetSuperTip(control As IRibbonControl) As String Implements ICreateCallbacks.GetSuperTip
        Return Ribbon.GetElement(Of ISuperTip)(control.Id).SuperTip
    End Function

    Public Function GetKeyTip(control As IRibbonControl) As KeyTip Implements ICreateCallbacks.GetKeyTip
        Return Ribbon.GetElement(Of IKeyTip)(control.Id).KeyTip
    End Function

    Public Function GetSize(control As IRibbonControl) As ControlSize Implements ICreateCallbacks.GetSize
        Return Ribbon.GetElement(Of ISize)(control.Id).Size
    End Function

    Public Function GetImage(control As IRibbonControl) As IPictureDisp Implements ICreateCallbacks.GetImage
        Return Ribbon.GetElement(Of IImage)(control.Id).Image.AsIPictureDisp()
    End Function

    Public Function GetShowImage(control As IRibbonControl) As Boolean Implements ICreateCallbacks.GetShowImage
        Return Ribbon.GetElement(Of IShowImage)(control.Id).ShowImage
    End Function

    Public Function GetPressed(control As IRibbonControl) As Boolean Implements ICreateCallbacks.GetPressed
        Return Ribbon.GetElement(Of IPressed)(control.Id).Checked
    End Function

    Public Function GetText(control As IRibbonControl) As String Implements ICreateCallbacks.GetText
        Return Ribbon.GetElement(Of IText)(control.Id).Text
    End Function

    Public Function GetItemCount(control As IRibbonControl) As Integer Implements ICreateCallbacks.GetItemCount
        Return Ribbon.GetContainer(Of Item)(control.Id).Count
    End Function

    Public Function GetItemID(control As IRibbonControl, index As Integer) As String Implements ICreateCallbacks.GetItemID
        Return Ribbon.GetContainerItem(control.Id, index).ID
    End Function

    Public Function GetItemImage(control As IRibbonControl, index As Integer) As IPictureDisp Implements ICreateCallbacks.GetItemImage
        Return Ribbon.GetContainerItem(control.Id, index).Image.AsIPictureDisp()
    End Function

    Public Function GetItemLabel(control As IRibbonControl, index As Integer) As String Implements ICreateCallbacks.GetItemLabel
        Return Ribbon.GetContainerItem(control.Id, index).Label
    End Function

    Public Function GetItemScreenTip(control As IRibbonControl, index As Integer) As String Implements ICreateCallbacks.GetItemScreenTip
        Return Ribbon.GetContainerItem(control.Id, index).ScreenTip
    End Function

    Public Function GetItemSuperTip(control As IRibbonControl, index As Integer) As String Implements ICreateCallbacks.GetItemSuperTip
        Return Ribbon.GetContainerItem(control.Id, index).SuperTip
    End Function

    Public Function GetItemHeight(control As IRibbonControl) As Integer Implements ICreateCallbacks.GetItemHeight
        Return Ribbon.GetElement(Of IItemDimensions)(control.Id).ItemHeight
    End Function

    Public Function GetItemWidth(control As IRibbonControl) As Integer Implements ICreateCallbacks.GetItemWidth
        Return Ribbon.GetElement(Of IItemDimensions)(control.Id).ItemWidth
    End Function

    Public Function GetSelectedItemID(control As IRibbonControl) As String Implements ICreateCallbacks.GetSelectedItemID
        Return Ribbon.GetElement(Of ISelect)(control.Id).Selected.ID
    End Function

    Public Function GetSelectedItemIndex(control As IRibbonControl) As Integer Implements ICreateCallbacks.GetSelectedItemIndex
        Return Ribbon.GetElement(Of ISelect)(control.Id).SelectedItemIndex
    End Function

    Public Sub OnAction(control As IRibbonControl) Implements ICreateCallbacks.OnAction
        Ribbon.GetElement(Of IClickable)(control.Id).Click()
    End Sub

    Public Sub OnChange(control As IRibbonControl, text As String) Implements ICreateCallbacks.OnChange
        Ribbon.GetElement(Of IText)(control.Id).Text = text
    End Sub

    Public Sub OnSelectionChange(control As IRibbonControl, selectedId As String, selectedIndex As Integer) Implements ICreateCallbacks.OnSelectionChange
        Ribbon.GetElement(Of ISelect)(control.Id).SelectedItemIndex = selectedIndex
    End Sub

    Public Sub OnButtonToggle(control As IRibbonControl, pressed As Boolean) Implements ICreateCallbacks.OnButtonToggle
        Ribbon.GetElement(Of IPressed)(control.Id).Checked = pressed
    End Sub

#End Region


End Class