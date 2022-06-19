Imports RibbonX.Callbacks
Imports RibbonX.ComTypes.StdOle

Namespace Builders.Internal.Standardization

    Public Interface IGetItemImage(Of T)

        Function GetItemImageFrom(callback As FromControlAndIndex(Of IPictureDisp)) As T

    End Interface

End Namespace