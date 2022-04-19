<TestClass()>
Public Class EditBoxTests
	Inherits RibbonTestBase

	Protected Overrides Function CreateRibbon() As RibbonFactory.Controls.Ribbon
		Return MakeRibbonWithOneTabAndOneGroup(ControlGenerator.MakeEditBox())
	End Function

End Class