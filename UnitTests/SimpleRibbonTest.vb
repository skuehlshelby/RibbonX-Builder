Imports NetOffice.OfficeApi
Imports RibbonFactory.Builders
Imports RibbonFactory.Containers
Imports RibbonFactory.Controls

<TestClass()>
Public Class SimpleRibbonTest
    Private _ribbon As Ribbon

    <TestMethod()>
    Public Sub BuildSimple()
        _ribbon = New Ribbon(startFromScratch:= False, [nameSpace]:= "Testing")
        _ribbon.Tabs.Add("TestTab").AddGroup().Add(New ButtonBuilder() _
                                                    .WithLabel("My Button") _
                                                    .WithScreenTip("My favorite button.") _
                                                    .ThatDoes(AddressOf OnAction, Sub() MsgBox("Hello World!")) _
                                                    .Build())
        dim xml As String = _ribbon.Build()
    End Sub

    Private Sub OnAction(control As IRibbonControl)
        _ribbon.AllElements.Single(Function(element) element.ID = control.Id)
    End Sub
End Class