<TestClass()>
Public Class EditBoxTests
	Inherits RibbonTestBase

	Protected Overrides Function CreateRibbon() As RibbonFactory.Containers.Ribbon
		Return MakeRibbonWithOneTabAndOneGroup(ControlGenerator.MakeEditBox())
	End Function

End Class