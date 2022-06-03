Imports System.Reflection
Imports RibbonX.ComTypes.Microsoft.Office.Core
Imports RibbonX
Imports RibbonX.Controls
Imports RibbonX.Enums
Imports RibbonX.ComTypes.StdOle

<TestClass()>
Public MustInherit Class TestBase
    Implements ICreateCallbacks

    Protected Function CreateSimpleRibbon(control As RibbonElement) As Ribbon
        Return New Ribbon(New Tab(New Group(control)))
    End Function

    <TestMethod>
    Public MustOverride Sub NullTemplate_NoThrow()

    <TestMethod>
    Public MustOverride Sub NullConfiguration_NoThrow()

    <TestMethod>
    Public MustOverride Sub ProducesLegalRibbonX()

    <TestMethod>
    Public MustOverride Sub ContainsNoNullValuesByDefault()

    Protected Sub ContainsNoNullValuesByDefaultHelper(Of T As RibbonElement)(control As T)
        For Each propertyInfo As PropertyInfo In GetType(T).GetProperties()
            If Not String.Equals(propertyInfo.Name, NameOf(RibbonElement.Tag)) Then
                Assert.IsNotNull(propertyInfo.GetValue(control), $"Property '{propertyInfo.Name}' on {control.ID} is null.")
            End If
        Next
    End Sub

    <TestMethod>
    Public MustOverride Sub PropertiesAreMappedCorrectly()

    <TestMethod>
    Public MustOverride Sub PropertiesWithoutCallbacksCannotBeModified()

    <TestMethod>
    Public MustOverride Sub PropertiesWithCallbacksCanBeModified()

    <TestMethod>
    Public MustOverride Sub TemplatePropertiesAreCopiedToNewControl()

    Public Sub OnLoad(ribbon As IRibbonUI) Implements ICreateCallbacks.OnLoad
        Throw New NotSupportedException()
    End Sub

    Public Function LoadImage(imageID As String) As IPictureDisp Implements ICreateCallbacks.LoadImage
        Throw New NotSupportedException()
    End Function

    Public Function GetEnabled(control As IRibbonControl) As Boolean Implements ICreateCallbacks.GetEnabled
        Throw New NotSupportedException()
    End Function

    Public Function GetVisible(control As IRibbonControl) As Boolean Implements ICreateCallbacks.GetVisible
        Throw New NotSupportedException()
    End Function

    Public Function GetLabel(control As IRibbonControl) As String Implements ICreateCallbacks.GetLabel
        Throw New NotSupportedException()
    End Function

    Public Function GetShowLabel(control As IRibbonControl) As Boolean Implements ICreateCallbacks.GetShowLabel
        Throw New NotSupportedException()
    End Function

    Public Function GetDescription(control As IRibbonControl) As String Implements ICreateCallbacks.GetDescription
        Throw New NotSupportedException()
    End Function

    Public Function GetScreenTip(control As IRibbonControl) As String Implements ICreateCallbacks.GetScreenTip
        Throw New NotSupportedException()
    End Function

    Public Function GetSuperTip(control As IRibbonControl) As String Implements ICreateCallbacks.GetSuperTip
        Throw New NotSupportedException()
    End Function

    Public Function GetKeyTip(control As IRibbonControl) As KeyTip Implements ICreateCallbacks.GetKeyTip
        Throw New NotSupportedException()
    End Function

    Public Function GetSize(control As IRibbonControl) As ControlSize Implements ICreateCallbacks.GetSize
        Throw New NotSupportedException()
    End Function

    Public Function GetImage(control As IRibbonControl) As IPictureDisp Implements ICreateCallbacks.GetImage
        Throw New NotSupportedException()
    End Function

    Public Function GetShowImage(control As IRibbonControl) As Boolean Implements ICreateCallbacks.GetShowImage
        Throw New NotSupportedException()
    End Function

    Public Function GetPressed(control As IRibbonControl) As Boolean Implements ICreateCallbacks.GetPressed
        Throw New NotSupportedException()
    End Function

    Public Function GetText(control As IRibbonControl) As String Implements ICreateCallbacks.GetText
        Throw New NotSupportedException()
    End Function

    Public Function GetItemCount(control As IRibbonControl) As Integer Implements ICreateCallbacks.GetItemCount
        Throw New NotSupportedException()
    End Function

    Public Function GetItemID(control As IRibbonControl, index As Integer) As String Implements ICreateCallbacks.GetItemID
        Throw New NotSupportedException()
    End Function

    Public Function GetItemImage(control As IRibbonControl, index As Integer) As IPictureDisp Implements ICreateCallbacks.GetItemImage
        Throw New NotSupportedException()
    End Function

    Public Function GetItemLabel(control As IRibbonControl, index As Integer) As String Implements ICreateCallbacks.GetItemLabel
        Throw New NotSupportedException()
    End Function

    Public Function GetItemScreenTip(control As IRibbonControl, index As Integer) As String Implements ICreateCallbacks.GetItemScreenTip
        Throw New NotSupportedException()
    End Function

    Public Function GetItemSuperTip(control As IRibbonControl, index As Integer) As String Implements ICreateCallbacks.GetItemSuperTip
        Throw New NotSupportedException()
    End Function

    Public Function GetItemHeight(control As IRibbonControl) As Integer Implements ICreateCallbacks.GetItemHeight
        Throw New NotSupportedException()
    End Function

    Public Function GetItemWidth(control As IRibbonControl) As Integer Implements ICreateCallbacks.GetItemWidth
        Throw New NotSupportedException()
    End Function

    Public Function GetSelectedItemID(control As IRibbonControl) As String Implements ICreateCallbacks.GetSelectedItemID
        Throw New NotSupportedException()
    End Function

    Public Function GetSelectedItemIndex(control As IRibbonControl) As Integer Implements ICreateCallbacks.GetSelectedItemIndex
        Throw New NotSupportedException()
    End Function

    Public Sub OnAction(control As IRibbonControl) Implements ICreateCallbacks.OnAction
        Throw New NotSupportedException()
    End Sub

    Public Sub OnChange(control As IRibbonControl, text As String) Implements ICreateCallbacks.OnChange
        Throw New NotSupportedException()
    End Sub

    Public Sub OnSelectionChange(control As IRibbonControl, selectedId As String, selectedIndex As Integer) Implements ICreateCallbacks.OnSelectionChange
        Throw New NotSupportedException()
    End Sub

    Public Sub OnButtonToggle(control As IRibbonControl, pressed As Boolean) Implements ICreateCallbacks.OnButtonToggle
        Throw New NotSupportedException()
    End Sub

End Class
