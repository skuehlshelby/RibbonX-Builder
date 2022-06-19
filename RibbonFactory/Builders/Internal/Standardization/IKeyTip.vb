Imports RibbonX.Callbacks
Imports RibbonX.SimpleTypes

Namespace Builders.Internal.Standardization

    Public Interface IKeyTip(Of T)

        Function WithKeyTip(keyTip As KeyTip) As T

        Function WithKeyTip(keyTip As KeyTip, callback As FromControl(Of KeyTip)) As T

    End Interface


End Namespace