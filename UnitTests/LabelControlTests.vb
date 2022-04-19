<TestClass>
Public Class LabelControlTests
	Inherits RibbonTestBase

	Protected Overrides Function CreateRibbon() As RibbonFactory.Controls.Ribbon
		Return MakeRibbonWithOneTabAndOneGroup(ControlGenerator.MakeLabelControl)
	End Function

End Class
