<TestClass>
Public Class GalleryTests
	Inherits RibbonTestBase

	Protected Overrides Function CreateRibbon() As RibbonFactory.Containers.Ribbon
		Return MakeRibbonWithOneTabAndOneGroup(ControlGenerator.MakeGallery())
	End Function

End Class
