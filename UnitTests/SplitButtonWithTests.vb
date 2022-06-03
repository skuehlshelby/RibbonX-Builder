Imports RibbonX.Controls

<TestClass>
Public Class SplitButtonWithTests
	Inherits RibbonTestBase

	Protected Overrides Function CreateRibbon() As Ribbon
		Return MakeRibbonWithOneTabAndOneGroup(ControlGenerator.MakeSplitButtonWithRegularButtonAndMenu())
	End Function
End Class
