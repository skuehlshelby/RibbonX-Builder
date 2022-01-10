Imports RibbonFactory
Imports RibbonFactory.Builders
Imports RibbonFactory.Controls

<TestClass()>
Public Class DynamicUpdating
	Inherits RibbonTestBase

	Private _changeTriggered As Boolean = False

	Protected Overrides Function CreateRibbon() As Containers.Ribbon
		Dim button As Button = New ButtonBuilder(ControlGenerator.MakeButton()).
			Enabled(AddressOf GetEnabled).
			Build()

		Return MakeRibbonWithOneTabAndOneGroup(button)
	End Function

	<TestMethod()>
	Public Sub ChangeTriggersRefresh()
		With Ribbon.GetElement(Function(e As Button) TypeOf e Is Button)
			AddHandler .ValueChanged, AddressOf ChangeTriggered
			.Enabled = False
			.Enabled = True
		End With

		Assert.IsTrue(_changeTriggered)
	End Sub

	Private Sub ChangeTriggered(sender As Object, e As ValueChangedEventArgs)
		_changeTriggered = True
	End Sub

	<TestCleanup()>
	Public Overrides Sub Cleanup()
		MyBase.Cleanup()
		_changeTriggered = False
	End Sub

End Class