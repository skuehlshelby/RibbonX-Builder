<TestClass>
Public Class LabelControlTests
	Inherits RibbonTestBase

	Protected Overrides Function CreateRibbon() As RibbonX.Controls.Ribbon
		Return MakeRibbonWithOneTabAndOneGroup(ControlGenerator.MakeLabelControl)
	End Function

End Class
