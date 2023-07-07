Imports RibbonX
Imports RibbonX.Controls.BuiltIn

<TestClass>
Public Class SeparatorTests
    Inherits TestBase

    <TestMethod>
    Public Overrides Sub NullTemplate_NoThrow()
        Dim control As ISeparator = RxApi.Separator(Sub(b) b.FromTemplate(Nothing))
    End Sub

    <TestMethod>
    Public Overrides Sub NullConfiguration_NoThrow()
        Dim control As ISeparator = RxApi.Separator(Nothing)
    End Sub

    <TestMethod>
    Public Overrides Sub ProducesLegalRibbonX()
        Assert.That.ValidRibbonXIsProduced(ButtonTests.BuildButton(), BuildSeparator(), ButtonTests.BuildButtonII())
    End Sub

    <TestMethod>
    Public Overrides Sub ContainsNoNullValuesByDefault()
        Assert.That.NoPropertiesAreNull(RxApi.Separator())
    End Sub

    <TestMethod>
    Public Overrides Sub PropertiesAreMappedCorrectly()
        Dim control As ISeparator = BuildReadonlySeparator()

        Assert.IsTrue(control.Visible)
    End Sub

    <TestMethod>
    Public Overrides Sub PropertiesWithoutCallbacksCannotBeModified()
        Assert.That.PropertiesCannotBeModified(BuildReadonlySeparator())
    End Sub

    <TestMethod>
    Public Overrides Sub PropertiesWithCallbacksCanBeModified()
        Dim control As ISeparator = BuildSeparatorII()

        Assert.That.PropertyMayBeModified(control, Function(c) c.Visible, Not control.Visible)
    End Sub

    <TestMethod>
    Public Overrides Sub TemplatePropertiesAreCopiedToNewControl()
        Dim control As ISeparator = BuildReadonlySeparatorII()

        Assert.That.SharedPropertiesAreEqual(control, RxApi.DropDown(Sub(b) b.FromTemplate(control)))
    End Sub

    Public Shared Function BuildReadonlySeparator() As ISeparator
        Return RxApi.Separator(Sub(b) b.Visible().InsertAfter(Excel.About))
    End Function

    Public Shared Function BuildReadonlySeparatorII() As ISeparator
        Return RxApi.Separator(Sub(b) b.Invisible().InsertAfter(Excel.About))
    End Function

    Public Shared Function BuildSeparator() As ISeparator
        Return RxApi.Separator(Sub(b) b.InsertBefore(Excel.About).Visible(AddressOf GetVisibleShared))
    End Function

    Public Shared Function BuildSeparatorII() As ISeparator
        Return RxApi.Separator(Sub(b) b.InsertBefore(Excel.About).Invisible(AddressOf GetVisibleShared))
    End Function
End Class

