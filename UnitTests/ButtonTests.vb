Imports RibbonFactory
Imports RibbonFactory.Builders
Imports RibbonFactory.Controls
Imports RibbonFactory.Utilities

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
                .ThatDoes(Sub() MsgBox("Hello World!"), AddressOf OnAction) _
                .Build()

        Dim errors As RibbonErrorLog = GetErrors(MakeRibbonWithOneTabAndOneGroup(button).RibbonX)

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
                .ThatDoes(Sub() MsgBox("Hello World!"), AddressOf OnAction) _
                .Build()

        Dim buttonTwo As Button = New ButtonBuilder() _
                .WithLabel("Button 2") _
                .WithSuperTip("The Second Button") _
                .InsertAfterMSO(Enums.MSO.Excel.Calculator) _
                .WithKeyTip("Z5") _
                .Disabled(AddressOf GetEnabled) _
                .ThatDoes(Sub() MsgBox("Hello World!"), AddressOf OnAction) _
                .Build()

        Dim errors As RibbonErrorLog = GetErrors(MakeRibbonWithOneTabAndOneGroup(button, buttonTwo).RibbonX)

        Assert.IsTrue(errors.None, String.Join(Environment.NewLine, errors.Errors))
    End Sub

    <TestMethod()>
    Public Sub ActionExecutes()
        Dim button As Button = New ButtonBuilder() _
                .WithLabel("Before Click", AddressOf GetLabel) _
                .ThatDoes(Sub() button.Label = "After Click", AddressOf OnAction) _
                .Build()

        Ribbon = MakeRibbonWithOneTabAndOneGroup(button)

        Dim excel As Stubs.Excel = New Stubs.Excel(ribbonExtensibility:= Me)
        Ribbon.AssignRibbonUI(excel)
        excel.Click(button)

        Assert.AreEqual(button.Label, "After Click")
    End Sub
End Class