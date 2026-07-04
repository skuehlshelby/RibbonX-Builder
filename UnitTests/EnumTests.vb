Imports RibbonX.Controls.BuiltIn
Imports RibbonX.Images.BuiltIn
Imports RibbonX.Utilities

<TestClass()>
Public Class EnumTests

    <TestMethod()>
    Public Sub ImageMSOValuesCanBeLoaded()
        Assert.IsTrue(Enumeration.Values(Of Chart).Any())

        Assert.IsTrue(Enumeration.Values(Of Misc).Any())

        Assert.IsTrue(Enumeration.Values(Of Pivot).Any())

        Assert.IsTrue(Enumeration.Values(Of XML).Any())

        Assert.IsTrue(Enumeration.Values(Of Number).Any())

        Assert.IsTrue(Enumeration.Values(Of Mail).Any())

        Assert.IsTrue(Enumeration.Values(Of ThreeDimensional).Any())

        Assert.IsTrue(Enumeration.Values(Of Common).Any())

        Assert.IsTrue(Enumeration.Values(Of Letter).Any())
    End Sub

    <TestMethod()>
    Public Sub IdMSOValuesCanBeLoaded()
        Assert.IsTrue(Enumeration.Values(Of Excel).Any())
        Assert.IsTrue(Enumeration.Values(Of Word).Any())
    End Sub

End Class