Imports RibbonX
Imports RibbonX.Controls.BuiltIn
Imports RibbonX.Images.BuiltIn
Imports RibbonX.Testing

<TestClass()>
Public Class IntegrationTests
    Inherits StockRibbonBase

    Private button As IButton
    Private editBox As IEditBox
    Private group As IGroup
    Private tab As ITab

    Protected Overrides Function BuildRibbon() As IRibbon
        button = RxApi.Button(Sub(b) b.
                                    Large(AddressOf GetSize).
                                    WithLabel("My Button", AddressOf GetLabel).
                                    WithSuperTip("More Info", AddressOf GetSuperTip).
                                    WithImage(Common.HappyFace).
                                    ShowImage(AddressOf GetShowImage).
                                    OnClick(AddressOf OnAction, Sub(ab) ab.Do(Sub(c) c.ShowImage = Not c.ShowImage)))

        editBox = RxApi.EditBox(Sub(b) b.
                                          WithImage(Common.FileFind).
                                          WithLabel("My Edit Box", AddressOf GetLabel).
                                          ShowLabel(AddressOf GetShowLabel).
                                          AsWideAs("AAAAAAAAAAAAAAAAAAA").
                                          WithMaximumInputLength(10).
                                          WithText("Text", AddressOf GetText, AddressOf OnChange))

        group = RxApi.Group(Sub(b) b.
                                    WithImage(Common.DollarSign).
                                    WithLabel("My Custom Group", AddressOf GetLabel).
                                    WithSuperTip("More Info", AddressOf GetSuperTip).
                                    WithKeyTip("G", AddressOf GetKeyTip).
                                    WithControls(button, editBox))

        tab = RxApi.Tab(Sub(b) b.
                                InsertAfter(Excel.TabHome).
                                WithLabel("My Custom Tab").
                                WithGroups(group))

        Return RxApi.Ribbon(Sub(b) b.OnLoad(AddressOf OnLoad).WithTabs(tab))
    End Function

    <TestMethod()>
    Public Sub MinimalRibbon()
        Dim host As OfficeHostAppilcation = New OfficeHostAppilcation(Me)

        Debug.WriteLine(Ribbon.RibbonX)

        host.Click(button)

        Assert.IsFalse(button.ShowImage)

        host.Click(button)

        Assert.IsTrue(button.ShowImage)

        host.Type(editBox, "Ta Da")

        Assert.AreEqual(editBox.Text, "Ta Da")
    End Sub
End Class
