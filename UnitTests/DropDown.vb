
Imports RibbonFactory.Builders

<TestClass()>
Public Class DropDown
    Inherits RibbonTestBase

    <TestMethod()>
    Public Sub Builds()
        Try
            RedirectConsoleOutput()

            Dim dropDown As RibbonFactory.Controls.DropDown =
                    New DropDownBuilder().WithLabel("Dropdown One").WithScreenTip(
                        "A super cool dropdown. Tell your friends.", AddressOf GetScreenTip).WithMaximumInputLength(50).HideImage().GetItemIdFrom(
                            AddressOf GetItemID).GetItemCountFrom(AddressOf GetItemCount).GetItemImageFrom(
                                AddressOf GetItemImage).GetItemLabelFrom(AddressOf GetItemLabel).GetItemScreenTipFrom(
                                    AddressOf GetItemScreenTip).GetItemSuperTipFrom(AddressOf GetItemSuperTip).Enabled(
                                        AddressOf GetEnabled).Build()

            Ribbon = RibbonWithOneTabAndOneGroup()

            ribbon.Tabs.First().Groups.First().AddControl(dropDown)

            Dim unused As String = ribbon.Build()
        Finally
            RestoreOriginalConsoleOutput()
        End Try

        Assert.AreEqual(String.Empty, GetContentsOfRedirectedConsoleOutput())
    End Sub

    <TestMethod()>
    Public Sub BuildWithAsManyAttributesAsPossible()
        Try
            RedirectConsoleOutput()

            Dim dropDown As RibbonFactory.Controls.DropDown =
                    New DropDownBuilder().WithLabel("Dropdown One").WithScreenTip(
                        "A super cool dropdown. Tell your friends.", AddressOf GetScreenTip).WithMaximumInputLength(50).HideImage().GetItemIdFrom(
                            AddressOf GetItemID).GetItemCountFrom(AddressOf GetItemCount).GetItemImageFrom(
                                AddressOf GetItemImage).GetItemLabelFrom(AddressOf GetItemLabel).GetItemScreenTipFrom(
                                    AddressOf GetItemScreenTip).GetItemSuperTipFrom(AddressOf GetItemSuperTip).Enabled(
                                        AddressOf GetEnabled).Build()

            Ribbon = RibbonWithOneTabAndOneGroup()

            ribbon.Tabs.First().Groups.First().AddControl(dropDown)

            Dim unused As String = ribbon.Build()
        Finally
            RestoreOriginalConsoleOutput()
        End Try

        Assert.AreEqual(String.Empty, GetContentsOfRedirectedConsoleOutput())
    End Sub
End Class