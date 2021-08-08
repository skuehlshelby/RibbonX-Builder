Imports RibbonFactory
Imports RibbonFactory.Builders
Imports RibbonFactory.Controls

<TestClass()>
Public Class ButtonTests
    Inherits RibbonTestBase

    <TestMethod()>
    Public Sub Builds_OneButton()
        Dim button As Button = New ButtonBuilder() _
                .WithLabel("My Button") _
                .WithScreenTip("My favorite button.") _
                .InsertAfterMSO(Enums.MSO.Excel.Calculator) _
                .WithKeyTip("Z3") _
                .Disabled(AddressOf GetEnabled) _
                .ThatDoes(AddressOf OnAction, Sub() MsgBox("Hello World!")) _
                .Build()

        Dim errors As RibbonErrorLog = GetErrors(RibbonWithOneTabAndOneGroup(button).RibbonX)

        Assert.IsTrue(errors.None, String.Join(Environment.NewLine, errors.Errors))
    End Sub

    <TestMethod()>
    Public Sub Builds_TwoButtons()
        Dim button As Button = New ButtonBuilder() _
                .WithLabel("My Button") _
                .WithScreenTip("My favorite button.") _
                .InsertAfterMSO(Enums.MSO.Excel.Calculator) _
                .WithKeyTip("Z4") _
                .Disabled(AddressOf GetEnabled) _
                .ThatDoes(AddressOf OnAction, Sub() MsgBox("Hello World!")) _
                .Build()

        Dim buttonTwo As Button = New ButtonBuilder() _
                .WithLabel("Button 2") _
                .WithSuperTip("The Second Button") _
                .InsertAfterMSO(Enums.MSO.Excel.Calculator) _
                .WithKeyTip("Z5") _
                .Disabled(AddressOf GetEnabled) _
                .ThatDoes(AddressOf OnAction, Sub() MsgBox("Hello World!")) _
                .Build()

        Dim errors As RibbonErrorLog = GetErrors(RibbonWithOneTabAndOneGroup(button, buttonTwo).RibbonX)

        Assert.IsTrue(errors.None, String.Join(Environment.NewLine, errors.Errors))
    End Sub

    <TestMethod()>
    Public Sub ActionExecutes()
        Dim button As Button = New ButtonBuilder() _
                .WithLabel("Before Click", AddressOf GetLabel) _
                .ThatDoes(AddressOf OnAction, Sub() button.Label = "After Click") _
                .Build()

        Ribbon = RibbonWithOneTabAndOneGroup(button)

        Dim excel As Stubs.Excel = New Stubs.Excel(ribbonExtensibility:= Me)
        Ribbon.AssignRibbonUI(excel)
        excel.Click(button)

        Assert.AreEqual(button.Label, "After Click")
    End Sub
End Class