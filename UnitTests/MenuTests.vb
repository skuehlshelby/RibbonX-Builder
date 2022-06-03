Imports RibbonX.Controls

<TestClass()> Public Class MenuTests
    Inherits RibbonTestBase

    Protected Overrides Function CreateRibbon() As Ribbon
        Return MakeRibbonWithOneTabAndOneGroup(ControlGenerator.MakeMenu())
    End Function

End Class