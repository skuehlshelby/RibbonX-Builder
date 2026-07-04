Imports RibbonX
Imports RibbonX.ComTypes.Microsoft.Office.Core
Imports Rx = RibbonX.RibbonXBuilder

''' <summary>
''' Exercises the item-related callbacks (getItemCount, getItemID, getItemLabel,
''' getItemScreenTip, getItemSuperTip, getItemImage) through the implementations
''' supplied by <see cref="CustomRibbonBase"/>. This is the path a consumer takes
''' when routing item callbacks to the built-in base methods (as ExampleRibbon does).
''' It flows through Ribbon.GetContainer / GetContainerItem, which previously threw
''' InvalidCastException because the item containers did not implement
''' IReadOnlyCollection(Of IItem).
''' </summary>
<TestClass()>
Public Class ItemCallbackTests
    Inherits CustomRibbonBase

    Private dropdown As IDropDown
    Private item1 As IItem
    Private item2 As IItem
    Private tab As ITab

    Protected Overrides Function BuildRibbon() As IRibbon
        item1 = Rx.Item(Sub(b) b.WithLabel("Item1"))
        item2 = Rx.Item(Sub(b) b.WithLabel("Item2"))

        dropdown = Rx.DropDown(Sub(b) b.
            WithLabel("DD", AddressOf GetLabel).
            GetItemCountFrom(AddressOf GetItemCount).
            GetItemLabelFrom(AddressOf GetItemLabel))

        dropdown.Add(item1, item2)

        Dim group As IGroup = Rx.Group(Sub(b) b.WithLabel("G").WithControls(dropdown))
        tab = Rx.Tab(Sub(b) b.WithLabel("T").WithGroups(group))
        Return Rx.Ribbon(Sub(b) b.OnLoad(AddressOf OnLoad).WithTabs(tab))
    End Function

    Private Function Control() As IRibbonControl
        GetCustomUI("")  ' builds the ribbon and assigns it to the base
        Return New StubControl(dropdown.Id)
    End Function

    <TestMethod>
    Public Sub GetItemCount_ReturnsNumberOfItems()
        Assert.AreEqual(2, GetItemCount(Control()))
    End Sub

    <TestMethod>
    Public Sub GetItemID_ReturnsIdOfItemAtIndex()
        Dim ctl As IRibbonControl = Control()
        Assert.AreEqual(item1.Id, GetItemID(ctl, 0))
        Assert.AreEqual(item2.Id, GetItemID(ctl, 1))
    End Sub

    <TestMethod>
    Public Sub GetItemLabel_ReturnsLabelOfItemAtIndex()
        Dim ctl As IRibbonControl = Control()
        Assert.AreEqual("Item1", GetItemLabel(ctl, 0))
        Assert.AreEqual("Item2", GetItemLabel(ctl, 1))
    End Sub

    <TestMethod>
    Public Sub GetItemScreenTip_DoesNotThrow()
        Dim ctl As IRibbonControl = Control()
        Assert.AreEqual("Item1", GetItemScreenTip(ctl, 0))
    End Sub

    <TestMethod>
    Public Sub GetItemSuperTip_DoesNotThrow()
        Dim ctl As IRibbonControl = Control()
        Assert.AreEqual(String.Empty, GetItemSuperTip(ctl, 0))
    End Sub

    <TestMethod>
    Public Sub GetItemImage_ReturnsPicture()
        Dim ctl As IRibbonControl = Control()
        Assert.IsNotNull(GetItemImage(ctl, 0))
    End Sub

    Private NotInheritable Class StubControl
        Implements IRibbonControl

        Private ReadOnly _id As String

        Public Sub New(id As String)
            _id = id
        End Sub

        Public ReadOnly Property Id As String Implements IRibbonControl.Id
            Get
                Return _id
            End Get
        End Property

        Public ReadOnly Property Context As Object Implements IRibbonControl.Context
            Get
                Return Nothing
            End Get
        End Property

        Public ReadOnly Property Tag As String Implements IRibbonControl.Tag
            Get
                Return Nothing
            End Get
        End Property
    End Class

End Class
