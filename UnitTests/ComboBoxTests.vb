<TestClass()>
Public Class ComboBoxTests
	Inherits RibbonTestBase

	Protected Overrides Function CreateRibbon() As RibbonFactory.Containers.Ribbon
		Return MakeRibbonWithOneTabAndOneGroup(ControlGenerator.MakeComboBox())
	End Function
End Class
