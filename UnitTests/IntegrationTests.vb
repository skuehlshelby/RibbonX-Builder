Imports RibbonX
Imports RibbonX.Controls
Imports RibbonX.Controls.BuiltIn
Imports RibbonX.Images.BuiltIn
Imports RibbonX.Testing

<TestClass()>
Public Class IntegrationTests
    Inherits StockRibbonBase

    Private button As Button
    Private editBox As EditBox
    Private group As Group
    Private tab As Tab

    Protected Overrides Function BuildRibbon() As Ribbon
        button = New Button(config:=Sub(b) b.
                                    Large(AddressOf GetSize).
                                    WithLabel("My Button", AddressOf GetLabel).
                                    WithSuperTip("More Info", AddressOf GetSuperTip).
                                    WithImage(Common.HappyFace).
                                    ShowImage(AddressOf GetShowImage).
                                    RouteClickTo(AddressOf OnAction).
                                    OnClick(Sub(c) c.ShowImage = Not c.ShowImage))

        editBox = New EditBox(config:=Sub(b) b.
                                          WithImage(Common.FileFind).
                                          WithLabel("My Edit Box", AddressOf GetLabel).
                                          ShowLabel(AddressOf GetShowLabel).
                                          AsWideAs("AAAAAAAAAAAAAAAAAAA").
                                          WithMaximumInputLength(10).
                                          WithText("Text", AddressOf GetText, AddressOf OnChange))

        group = New Group(config:=Sub(b) b.
                                    WithImage(Common.DollarSign).
                                    WithLabel("My Custom Group", AddressOf GetLabel).
                                    WithSuperTip("More Info", AddressOf GetSuperTip).
                                    WithKeyTip("G", AddressOf GetKeyTip),
                         controls:={button, editBox})

        tab = New Tab(config:=Sub(b) b.
                                InsertAfter(Excel.TabHome).
                                WithLabel("My Custom Tab"),
                      groups:={group})

        Return New Ribbon(configuration:=Sub(b) b.OnLoad(AddressOf OnLoad), tab)
    End Function

    <TestMethod()>
    Public Sub MinimalRibbon()
        Debug.WriteLine(Me.Ribbon.RibbonX)

        Assert.ThrowsException(Of Exception)(Sub() button.Label = String.Empty) 'When ribbonX is generated but IRibbonUI is unassigned, properties are not modifiable

        Dim host As OfficeHostAppilcation = New OfficeHostAppilcation(Me)

        host.Click(button)

        Assert.IsFalse(button.ShowImage)

        host.Click(button)

        Assert.IsTrue(button.ShowImage)

        host.Type(editBox, "Ta Da")

        Assert.AreEqual(editBox.Text, "Ta Da")
    End Sub
End Class
