<TestClass>
Public Class ToggleButtonTests
	Inherits RibbonTestBase

	Protected Overrides Function CreateRibbon() As RibbonFactory.Containers.Ribbon
		Return MakeRibbonWithOneTabAndOneGroup(ControlGenerator.MakeToggleButton())
	End Function

End Class
