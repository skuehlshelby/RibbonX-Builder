<TestClass>
Public Class GalleryTests
	Inherits RibbonTestBase

	Protected Overrides Function CreateRibbon() As RibbonX.Controls.Ribbon
		Return MakeRibbonWithOneTabAndOneGroup(ControlGenerator.MakeGallery())
	End Function

End Class
