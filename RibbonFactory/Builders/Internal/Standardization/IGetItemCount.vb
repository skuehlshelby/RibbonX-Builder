Imports RibbonX.Callbacks

Namespace Builders.Internal.Standardization

    Public Interface IGetItemCount(Of T)

        Function GetItemCountFrom(callback As FromControl(Of Integer)) As T

    End Interface

End Namespace