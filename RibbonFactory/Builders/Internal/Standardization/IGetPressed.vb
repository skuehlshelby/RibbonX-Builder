﻿Imports RibbonX.Callbacks

Namespace Builders.Internal.Standardization

    Public Interface IGetPressed(Of T)

        Function GetPressedFrom(callback As FromControl(Of Boolean)) As T

    End Interface

End Namespace