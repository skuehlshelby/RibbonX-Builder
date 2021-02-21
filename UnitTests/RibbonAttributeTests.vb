Imports System.Text
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports NetOffice.OfficeApi
Imports RibbonFactory
Imports RibbonFactory.Enums
Imports RibbonFactory.RibbonAttributes
Imports RibbonFactory.RibbonAttributes.Categories.Size

<TestClass()> Public Class RibbonAttributeTests

    <TestMethod()> Public Sub TestMethod1()
        Dim CUT As AttributeGroup = New AttributeGroup()

        CUT.Add(New GetSize(ControlSize.large, AddressOf FromControl)).SetValue(ControlSize.normal)

        CUT.Add(New Size(ControlSize.normal))

        Assert.AreEqual(CUT.Count, 1)
    End Sub

    Public Function FromControl(ByVal Control As IRibbonControl) As ControlSize
        Return ControlSize.large
    End Function
End Class