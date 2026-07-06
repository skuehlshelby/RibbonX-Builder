Imports System.Drawing
Imports RibbonX.Images
Imports RibbonX.Images.BuiltIn

''' <summary>
''' Equality across the <see cref="RibbonImage"/> variants. Previously every override
''' compared the wrapped value against the outer RibbonImage object, so two images
''' representing the same picture never compared equal.
''' </summary>
<TestClass()>
Public Class RibbonImageEqualityTests

    <TestMethod>
    Public Sub BuiltInImages_WrappingSameImageMso_AreEqual()
        Dim a As RibbonImage = RibbonImage.Create(Common.DollarSign)
        Dim b As RibbonImage = RibbonImage.Create(Common.DollarSign)

        Assert.AreEqual(a, b)
        Assert.AreEqual(a.GetHashCode(), b.GetHashCode())
    End Sub

    <TestMethod>
    Public Sub BuiltInImage_EqualsUnderlyingImageMso()
        Dim a As RibbonImage = RibbonImage.Create(Common.DollarSign)

        Assert.IsTrue(a.Equals(Common.DollarSign))
    End Sub

    <TestMethod>
    Public Sub BuiltInImages_WrappingDifferentImageMso_AreNotEqual()
        ' Note: distinct-valued ImageMSOs are required here. The Common.* set is
        ' all defined with value 1, so those compare equal by design (Enumeration
        ' equality is value-based); Number.One/Two have distinct values.
        Dim a As RibbonImage = RibbonImage.Create(Number.One)
        Dim b As RibbonImage = RibbonImage.Create(Number.Two)

        Assert.AreNotEqual(a, b)
    End Sub

    <TestMethod>
    Public Sub BuiltInImages_FromDifferentCategories_AreNotEqual()
        ' Both carry the underlying value 1, but belong to different ImageMSO
        ' categories; equality is type-aware, so they must not be equal.
        Dim a As RibbonImage = RibbonImage.Create(Common.AttachFile)
        Dim b As RibbonImage = RibbonImage.Create(Number.One)

        Assert.AreNotEqual(a, b)
    End Sub

    <TestMethod>
    Public Sub CachedImages_WithSameId_AreEqual()
        Dim a As RibbonImage = RibbonImage.Create("logo", New Bitmap(1, 1))
        Dim b As RibbonImage = RibbonImage.Create("logo", New Bitmap(2, 2))

        Assert.AreEqual(a, b)
        Assert.AreEqual(a.GetHashCode(), b.GetHashCode())
    End Sub

    <TestMethod>
    Public Sub CachedImages_WithDifferentId_AreNotEqual()
        Dim a As RibbonImage = RibbonImage.Create("logo", New Bitmap(1, 1))
        Dim b As RibbonImage = RibbonImage.Create("banner", New Bitmap(1, 1))

        Assert.AreNotEqual(a, b)
    End Sub

    <TestMethod>
    Public Sub DifferentImageKinds_AreNotEqual()
        Dim builtIn As RibbonImage = RibbonImage.Create(Common.DollarSign)
        Dim cached As RibbonImage = RibbonImage.Create("logo", New Bitmap(1, 1))
        Dim picture As RibbonImage = RibbonImage.Create(New Bitmap(1, 1))

        Assert.AreNotEqual(builtIn, cached)
        Assert.AreNotEqual(builtIn, picture)
        Assert.AreNotEqual(cached, picture)
    End Sub

End Class
