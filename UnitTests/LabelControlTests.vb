<TestClass>
Public Class LabelControlTests
	Inherits RibbonTestBase

	Protected Overrides Function CreateRibbon() As RibbonFactory.Containers.Ribbon
		Return MakeRibbonWithOneTabAndOneGroup(ControlGenerator.MakeLabelControl)
	End Function

End Class
