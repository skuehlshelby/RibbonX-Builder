Imports RibbonFactory.Builders
Imports stdole

Namespace BuilderInterfaces
    Public Interface ISetChildItemImageSource(Of T As Builder)

        Function GetItemImageFrom(imageSource As FromControlAndIndex(Of IPictureDisp)) As T

    End Interface
End Namespace