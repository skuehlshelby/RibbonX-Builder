﻿Namespace BuilderInterfaces.API

    Public Interface ITabBuilder
        Inherits IInsert(Of ITabBuilder)
        Inherits IKeyTip(Of ITabBuilder)
        Inherits ILabelOnly(Of ITabBuilder)
        Inherits IVisible(Of ITabBuilder)
    End Interface

End Namespace