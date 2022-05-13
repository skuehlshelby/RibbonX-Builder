Imports RibbonFactory.Controls

<TestClass()>
Public Class DialogLauncherTests

    <TestMethod()>
    Public Sub ProducesLegalRibbonX()
        Dim dialogLauncher As DialogLauncher = New DialogLauncher(New Button())
        Dim ribbon As Ribbon = New Ribbon(New Tab(New [Group](dialogLauncher)))

        Assert.IsTrue(Ribbon.GetErrors().None)

        Debug.WriteLine(Ribbon.RibbonX)
    End Sub

End Class