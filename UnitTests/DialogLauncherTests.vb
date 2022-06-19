Imports RibbonX.Controls

<TestClass()>
Public Class DialogLauncherTests

    <TestMethod()>
    Public Sub ProducesLegalRibbonX()
        Assert.That.ValidRibbonXIsProduced(New DialogLauncher(ButtonTests.BuildButtonII()))
    End Sub

End Class