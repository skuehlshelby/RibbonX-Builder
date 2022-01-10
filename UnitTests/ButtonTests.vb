Imports RibbonFactory
Imports RibbonFactory.Builders
Imports RibbonFactory.Controls

<TestClass()>
Public Class ButtonTests
    Inherits RibbonTestBase

    Protected Overrides Function CreateRibbon() As Containers.Ribbon
        Dim button As Button = New ButtonBuilder(ControlGenerator.MakeButton()) _
                .WithLabel("Before Click", AddressOf GetLabel) _
                .ThatDoes(Sub() button.Label = "After Click", AddressOf OnAction) _
                .Build()

        Return MakeRibbonWithOneTabAndOneGroup(button)
    End Function

    <TestMethod()>
    Public Sub ActionExecutes()
        Dim button As Button = Ribbon.GetElement(Of Button)(Function(b) True)

        OfficeHost.Click(button)

        Assert.AreEqual(Button.Label, "After Click")
    End Sub

End Class