Imports RibbonFactory.Builders
Imports RibbonFactory.Containers

<TestClass>
Public Class SplitButtonWithTests
	Inherits RibbonTestBase

	Protected Overrides Function CreateRibbon() As Ribbon
		Dim sb As SplitButton = New SplitButtonBuilder(ControlGenerator.MakeSplitButtonWithRegularButtonAndMenu()).
			WithButton(ControlGenerator.MakeButton()).
			WithMenu(MakeMenu()).
			Build()

		Return MakeRibbonWithOneTabAndOneGroup(sb)
	End Function
End Class
