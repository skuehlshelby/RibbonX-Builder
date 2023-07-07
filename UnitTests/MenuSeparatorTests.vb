Imports RibbonX
Imports RibbonX.Controls.BuiltIn

<TestClass>
Public Class MenuSeparatorTests
    Inherits TestBase

    Private Const TITLE As String = "My Separator Title"

    <TestMethod>
    Public Overrides Sub NullTemplate_NoThrow()
        Dim control As IMenuSeparator = RxApi.MenuSeparator(Sub(b) b.FromTemplate(Nothing))
    End Sub

    <TestMethod>
    Public Overrides Sub NullConfiguration_NoThrow()
        Dim control As IMenuSeparator = RxApi.MenuSeparator(Nothing)
    End Sub

    <TestMethod>
    Public Overrides Sub ProducesLegalRibbonX()
        Assert.That.ValidRibbonXIsProduced(MenuTests.BuildMenu(ButtonTests.BuildReadonlyButton(), BuildSeparator(), ButtonTests.BuildButtonII()))
    End Sub

    <TestMethod>
    Public Overrides Sub ContainsNoNullValuesByDefault()
        Assert.That.NoPropertiesAreNull(RxApi.MenuSeparator())
    End Sub

    <TestMethod>
    Public Overrides Sub PropertiesAreMappedCorrectly()
        Dim control As IMenuSeparator = BuildReadonlySeparator()

        Assert.AreEqual(control.Title, TITLE)
    End Sub

    <TestMethod>
    Public Overrides Sub PropertiesWithoutCallbacksCannotBeModified()
        Assert.That.PropertiesCannotBeModified(BuildReadonlySeparator())
    End Sub

    <TestMethod>
    Public Overrides Sub PropertiesWithCallbacksCanBeModified()
        Dim control As IMenuSeparator = BuildSeparatorII()

        Assert.That.PropertyMayBeModified(control, Function(c) c.Title, String.Empty)
    End Sub

    <TestMethod>
    Public Overrides Sub TemplatePropertiesAreCopiedToNewControl()
        Dim control As IMenuSeparator = BuildReadonlySeparatorII()

        Assert.That.SharedPropertiesAreEqual(control, RxApi.DropDown(Sub(b) b.FromTemplate(control)))
    End Sub

    Public Shared Function BuildReadonlySeparator() As IMenuSeparator
        Return RxApi.MenuSeparator(Sub(b) b.InsertAfter(Excel.About).WithTitle(TITLE))
    End Function

    Public Shared Function BuildReadonlySeparatorII() As IMenuSeparator
        Return RxApi.MenuSeparator(Sub(b) b.InsertAfter(Excel.About).WithTitle(TITLE))
    End Function

    Public Shared Function BuildSeparator() As IMenuSeparator
        Return RxApi.MenuSeparator(Sub(b) b.InsertBefore(Excel.About).WithTitle(TITLE, AddressOf GetTitle))
    End Function

    Public Shared Function BuildSeparatorII() As IMenuSeparator
        Return RxApi.MenuSeparator(Sub(b) b.InsertBefore(Excel.About).WithTitle(TITLE, AddressOf GetTitle))
    End Function
End Class

