<TestClass>
Public Class ToggleButtonTests
	Inherits RibbonTestBase

	Protected Overrides Function CreateRibbon() As RibbonX.Controls.Ribbon
		Return MakeRibbonWithOneTabAndOneGroup(ControlGenerator.MakeToggleButton())
	End Function

End Class
