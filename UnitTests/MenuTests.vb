Imports RibbonFactory.Containers

<TestClass()> Public Class MenuTests
    Inherits RibbonTestBase

    Protected Overrides Function CreateRibbon() As Ribbon
        Return MakeRibbonWithOneTabAndOneGroup(MakeMenu())
    End Function

End Class