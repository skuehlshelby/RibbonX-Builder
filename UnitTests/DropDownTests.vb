Imports RibbonFactory

<TestClass()>
Public Class DropDownTests
    Inherits RibbonTestBase

    Protected Overrides Function CreateRibbon() As Containers.Ribbon
        Return MakeRibbonWithOneTabAndOneGroup(ControlGenerator.MakeDropDown())
    End Function

End Class