Imports NetOffice.OfficeApi
Imports RibbonFactory.Builders
Imports RibbonFactory.Containers
Imports RibbonFactory.Controls

<TestClass()>
Public Class SimpleRibbonTest
    Private _ribbon As Ribbon
    Private _group As Group
    Private _button As Button

    <TestMethod()>
    Public Sub BuildSimple()
        _ribbon = New Ribbon(startFromScratch:= False, [nameSpace]:= "Testing")
        _group = _ribbon.AddTab("TestTab").AddGroup()
        Dim buttonBuilder As ButtonBuilder = New ButtonBuilder()
        buttonBuilder.WithLabel("My Button").WithScreenTip("My favorite button.").ThatDoes(AddressOf OnAction, Sub() MsgBox("Hello World!"))
        _button = buttonBuilder.Build()
        _group.Add(_button)
        dim xml As String = _ribbon.Build()
        _ribbon.ValidateRibbon(xml)
    End Sub

    Private Sub OnAction(control As IRibbonControl)
        _ribbon.AllElements.Single(Function(element) element.ID = control.Id)
    End Sub
End Class