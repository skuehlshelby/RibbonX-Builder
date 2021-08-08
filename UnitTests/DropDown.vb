
Imports RibbonFactory
Imports RibbonFactory.Builders

<TestClass()>
Public Class DropDown
    Inherits RibbonTestBase

    <TestMethod()>
    Public Sub Builds()
        Dim dropDown As Controls.DropDown =
                New DropDownBuilder().WithLabel("Dropdown One").WithScreenTip(
                    "A super cool dropdown. Tell your friends.", AddressOf GetScreenTip).WithSize(New String("l"c, 50)).HideImage().GetItemIdFrom(
                        AddressOf GetItemID).GetItemCountFrom(AddressOf GetItemCount).GetItemImageFrom(
                            AddressOf GetItemImage).GetItemLabelFrom(AddressOf GetItemLabel).GetItemScreenTipFrom(
                                AddressOf GetItemScreenTip).GetItemSuperTipFrom(AddressOf GetItemSuperTip).Enabled(
                                    AddressOf GetEnabled).Build()


        Dim errors As RibbonErrorLog = GetErrors(RibbonWithOneTabAndOneGroup(dropDown).RibbonX)

        Assert.IsTrue(errors.None, String.Join(Environment.NewLine, errors.Errors))
    End Sub

    <TestMethod()>
    Public Sub BuildWithAsManyAttributesAsPossible()
        Dim dropDown As Controls.DropDown =
                New DropDownBuilder().WithLabel("Dropdown One").WithScreenTip(
                    "A super cool dropdown. Tell your friends.", AddressOf GetScreenTip).WithSize(New String("l"c, 50)).HideImage().GetItemIdFrom(
                        AddressOf GetItemID).GetItemCountFrom(AddressOf GetItemCount).GetItemImageFrom(
                            AddressOf GetItemImage).GetItemLabelFrom(AddressOf GetItemLabel).GetItemScreenTipFrom(
                                AddressOf GetItemScreenTip).GetItemSuperTipFrom(AddressOf GetItemSuperTip).Enabled(
                                    AddressOf GetEnabled).Build()

        Ribbon = RibbonWithOneTabAndOneGroup(dropDown)

        Dim errors As RibbonErrorLog = GetErrors(Ribbon.RibbonX)

        Assert.IsTrue(errors.None, String.Join(Environment.NewLine, errors.Errors))
        Console.Write(Ribbon.RibbonX)
    End Sub
End Class