﻿Imports RibbonX.Builders.Internal.Standardization

Namespace Builders

    Public Interface IRibbonBuilder
        Inherits IStartFromScratch(Of IRibbonBuilder)
        Inherits IOnLoad(Of IRibbonBuilder)
        Inherits ILoadImage(Of IRibbonBuilder)
    End Interface

End Namespace
