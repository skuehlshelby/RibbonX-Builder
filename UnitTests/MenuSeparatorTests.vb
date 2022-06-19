Imports RibbonX.Builders
Imports RibbonX.Controls
Imports RibbonX.Controls.BuiltIn

<TestClass>
Public Class MenuSeparatorTests
    Inherits TestBase

    Private Const TITLE As String = "My Separator Title"

    <TestMethod>
    Public Overrides Sub NullTemplate_NoThrow()
        Dim control As MenuSeparator = New MenuSeparator(template:=Nothing)
    End Sub

    <TestMethod>
    Public Overrides Sub NullConfiguration_NoThrow()
        Dim control As MenuSeparator = New MenuSeparator(config:=Nothing)
    End Sub

    <TestMethod>
    Public Overrides Sub ProducesLegalRibbonX()
        Assert.That.ValidRibbonXIsProduced(MenuTests.BuildMenu(MenuControls.From(ButtonTests.BuildReadonlyButton()).Add(BuildSeparator()).Add(ButtonTests.BuildButtonII())))
    End Sub

    <TestMethod>
    Public Overrides Sub ContainsNoNullValuesByDefault()
        Assert.That.NoPropertiesAreNull(New MenuSeparator())
    End Sub

    <TestMethod>
    Public Overrides Sub PropertiesAreMappedCorrectly()
        Dim control As MenuSeparator = BuildReadonlySeparator()

        Assert.AreEqual(control.Title, TITLE)
    End Sub

    <TestMethod>
    Public Overrides Sub PropertiesWithoutCallbacksCannotBeModified()
        Assert.That.PropertiesCannotBeModified(BuildReadonlySeparator())
    End Sub

    <TestMethod>
    Public Overrides Sub PropertiesWithCallbacksCanBeModified()
        Dim control As MenuSeparator = BuildSeparatorII()

        Assert.That.PropertyMayBeModified(control, Function(c) c.Title, String.Empty)
    End Sub

    <TestMethod>
    Public Overrides Sub TemplatePropertiesAreCopiedToNewControl()
        Dim control As MenuSeparator = BuildReadonlySeparatorII()

        Assert.That.SharedPropertiesAreEqual(control, New DropDown(template:=control))
    End Sub

    Public Shared Function BuildReadonlySeparator() As MenuSeparator
        Return New MenuSeparator(config:=Sub(b) b.InsertAfter(Excel.About).WithTitle(TITLE))
    End Function

    Public Shared Function BuildReadonlySeparatorII() As MenuSeparator
        Return New MenuSeparator(config:=Sub(b) b.InsertAfter(Excel.About).WithTitle(TITLE))
    End Function

    Public Shared Function BuildSeparator() As MenuSeparator
        Return New MenuSeparator(config:=Sub(b) b.InsertBefore(Excel.About).WithTitle(TITLE, AddressOf GetTitle))
    End Function

    Public Shared Function BuildSeparatorII() As MenuSeparator
        Return New MenuSeparator(config:=Sub(b) b.InsertBefore(Excel.About).WithTitle(TITLE, AddressOf GetTitle))
    End Function
End Class

