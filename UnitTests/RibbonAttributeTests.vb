Imports RibbonFactory.Enums
Imports RibbonFactory.RibbonAttributes


<TestClass()>
Public Class RibbonAttributeTests
    Inherits RibbonTestBase
    Private _builder As AttributeGroupBuilder

    <TestInitialize()>
    Public Sub Initialize()
        _builder = new AttributeGroupBuilder()
    End Sub

    <TestMethod()>
    Public Sub AttributesAreMutuallyExclusive_SizeAndGetSize()
        _builder.AddSize(ControlSize.large)
        _builder.AddGetSize(ControlSize.large, AddressOf GetSize)

        Assert.AreEqual(1, _builder.Build().Count)
    End Sub

    <TestMethod()>
    Public Sub AttributesAreMutuallyExclusive_LabelAndGetLabel()
        _builder.AddLabel("Label")
        _builder.AddGetLabel("Label", AddressOf GetLabel)

        Assert.AreEqual(1, _builder.Build().Count)
    End Sub

    <TestMethod()>
    Public Sub AttributesAreMutuallyExclusive_ScreenTipAndGetScreenTip()
        _builder.AddScreentip("Screentip")
        _builder.AddGetScreentip("Screentip", AddressOf GetScreenTip)

        Assert.AreEqual(1, _builder.Build().Count)
    End Sub

    <TestMethod()>
    Public Sub AttributesAreMutuallyExclusive_SuperTipAndGetSuperTip()
        _builder.AddSupertip("Super")
        _builder.AddGetSupertip("Super", AddressOf GetSuperTip)

        Assert.AreEqual(1, _builder.Build().Count)
    End Sub

    <TestMethod()>
    Public Sub AttributesAreMutuallyExclusive_DescriptionAndGetDescription()
        _builder.AddDescription("The Thing")
        _builder.AddGetDescription("The Thing", AddressOf GetDescription)

        Assert.AreEqual(1, _builder.Build().Count)
    End Sub

    <TestMethod()>
    Public Sub AttributesAreMutuallyExclusive_IdAndIdQ()
        _builder.AddID("id")
        _builder.AddIdQ("namespace", "id")

        Assert.AreEqual(1, _builder.Build().Count)
    End Sub
End Class