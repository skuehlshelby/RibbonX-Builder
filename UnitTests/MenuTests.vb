Imports RibbonFactory
Imports RibbonFactory.Utilities

<TestClass()> Public Class MenuTests
    Inherits RibbonTestBase

    <TestMethod()> 
    Public Sub Builds()
        Dim errors As RibbonErrorLog = MakeRibbonWithOneTabAndOneGroup(GetMenu()).GetErrors()

        Assert.IsTrue(errors.None, String.Join(Environment.NewLine, errors.Errors))
    End Sub

End Class