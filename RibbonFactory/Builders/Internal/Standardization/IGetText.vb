﻿Imports RibbonX.Callbacks

Namespace Builders.Internal.Standardization

    Public Interface IGetText(Of T)

        Function GetTextFrom(callback As FromControl(Of String)) As T

    End Interface

End Namespace