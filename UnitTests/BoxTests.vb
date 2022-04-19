Imports RibbonFactory.Controls

<TestClass()>
Public Class BoxTests
    Inherits RibbonTestBase

    Protected Overrides Function CreateRibbon() As Ribbon
        Return MakeRibbonWithOneTabAndOneGroup(ControlGenerator.MakeBox(ControlGenerator.MakeButton(), ControlGenerator.MakeEditBox()))
    End Function

End Class