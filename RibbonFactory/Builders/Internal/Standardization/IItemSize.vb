﻿Namespace Builders.Internal.Standardization

    Public Interface IItemSize(Of T)
        
        Function WithLargeItems() As T

        Function WithNormalSizeItems() As T

    End Interface

End NameSpace