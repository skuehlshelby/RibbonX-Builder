Imports RibbonFactory
Imports RibbonFactory.Controls

<TestClass()>
Public Class DynamicUpdating
    Inherits RibbonTestBase

    Private _changeTriggered As Boolean = False

    <TestCleanup()>
    Public Sub Teardown()
        _changeTriggered = False
    End Sub

    <TestMethod()>
    Public Sub ChangeTriggersRefresh()
        Ribbon = MakeRibbonWithOneTabAndOneGroup(MakeButton())

        AssignFakeIRibbonUI()

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
End Class