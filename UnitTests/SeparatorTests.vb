Imports RibbonX.Controls
Imports RibbonX.Controls.BuiltIn

<TestClass>
Public Class SeparatorTests
    Inherits TestBase

    <TestMethod>
    Public Overrides Sub NullTemplate_NoThrow()
        Dim control As Separator = New Separator(template:=Nothing)
    End Sub

    <TestMethod>
    Public Overrides Sub NullConfiguration_NoThrow()
        Dim control As Separator = New Separator(config:=Nothing)
    End Sub

    <TestMethod>
    Public Overrides Sub ProducesLegalRibbonX()
        Assert.That.ValidRibbonXIsProduced(GroupTests.BuildReadonlyGroup(ButtonTests.BuildButton(), BuildSeparator(), ButtonTests.BuildButtonII()))
    End Sub

    <TestMethod>
    Public Overrides Sub ContainsNoNullValuesByDefault()
        Assert.That.NoPropertiesAreNull(New Separator())
    End Sub

    <TestMethod>
    Public Overrides Sub PropertiesAreMappedCorrectly()
        Dim control As Separator = BuildReadonlySeparator()

        Assert.IsTrue(control.Visible)
    End Sub

    <TestMethod>
    Public Overrides Sub PropertiesWithoutCallbacksCannotBeModified()
        Assert.That.PropertiesCannotBeModified(BuildReadonlySeparator())
    End Sub

    <TestMethod>
    Public Overrides Sub PropertiesWithCallbacksCanBeModified()
        Dim control As Separator = BuildSeparatorII()

        Assert.That.PropertyMayBeModified(control, Function(c) c.Visible, Not control.Visible)
    End Sub

    <TestMethod>
    Public Overrides Sub TemplatePropertiesAreCopiedToNewControl()
        Dim control As Separator = BuildReadonlySeparatorII()

        Assert.That.SharedPropertiesAreEqual(control, New DropDown(template:=control))
    End Sub

    Public Shared Function BuildReadonlySeparator() As Separator
        Return Separator.CreateVisible(Sub(b) b.InsertAfter(Excel.About))
    End Function

    Public Shared Function BuildReadonlySeparatorII() As Separator
        Return Separator.CreateInvisible(Sub(b) b.InsertAfter(Excel.About))
    End Function

    Public Shared Function BuildSeparator() As Separator
        Return New Separator(config:=Sub(b) b.InsertBefore(Excel.About).Visible(AddressOf GetVisibleShared))
    End Function

    Public Shared Function BuildSeparatorII() As Separator
        Return New Separator(config:=Sub(b) b.InsertBefore(Excel.About).Invisible(AddressOf GetVisibleShared))
    End Function
End Class

