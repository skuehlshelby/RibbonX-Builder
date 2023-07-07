Imports RibbonX

<TestClass()>
Public Class DialogLauncherTests

    <TestMethod()>
    Public Sub ProducesLegalRibbonX()
        Assert.That.ValidRibbonXIsProduced(RxApi.DialogLauncher(ButtonTests.BuildButtonII()))
    End Sub

End Class