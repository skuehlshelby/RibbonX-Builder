Imports RibbonFactory.Enums
Imports RibbonFactory.RibbonAttributes


<TestClass()>
Public Class RibbonAttributeTests
    Inherits RibbonTestBase

    <TestMethod()>
    Public Sub AttributesAreMutuallyExclusive_SizeAndGetSize()
        Dim cut As AttributeGroup = New AttributeGroup From {
            New GetSize(ControlSize.large, AddressOf GetSize),
            New Size(ControlSize.normal)
        }

        Assert.AreEqual(cut.Count, 1)
    End Sub

    <TestMethod()>
    Public Sub AttributesAreMutuallyExclusive_LabelAndGetLabel()
        Dim cut As AttributeGroup = New AttributeGroup From {
                New GetLabel("My Label", AddressOf GetLabel),
                New Label("My Label")
        }

        Assert.AreEqual(cut.Count, 1)
    End Sub

    <TestMethod()>
    Public Sub AttributesAreMutuallyExclusive_ScreenTipAndGetScreenTip()
        Dim cut As AttributeGroup = New AttributeGroup From {
                New GetScreentip("My Screentip", AddressOf GetScreenTip),
                New Screentip("My Screentip")
                }

        Assert.AreEqual(cut.Count, 1)
    End Sub

    <TestMethod()>
    Public Sub AttributesAreMutuallyExclusive_SuperTipAndGetSuperTip()
        Dim cut As AttributeGroup = New AttributeGroup From {
                New GetSuperTip("My SuperTip", AddressOf GetSuperTip),
                New Supertip("My SuperTip")
                }

        Assert.AreEqual(cut.Count, 1)
    End Sub

    <TestMethod()>
    Public Sub AttributesAreMutuallyExclusive_DescriptionAndGetDescription()
        Dim cut As AttributeGroup = New AttributeGroup From {
                New GetDescription("My Description", AddressOf GetDescription),
                New Description("My Description")
                }

        Assert.AreEqual(cut.Count, 1)
    End Sub

    <TestMethod()>
    Public Sub AttributesAreMutuallyExclusive_IdAndIdQ()
        Dim cut As AttributeGroup = New AttributeGroup From {
                New Id("Id1"),
                New IdQ("Testing", "Id1")
                }

        Assert.AreEqual(cut.Count, 1)
    End Sub
End Class