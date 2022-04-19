﻿Namespace BuilderInterfaces.API

    Public Interface IMenuSeparatorBuilder
        Inherits IID(Of IMenuSeparatorBuilder)
        Inherits IInsert(Of IMenuSeparatorBuilder)

        Function WithTitle(title As String) As IMenuSeparatorBuilder

        Function WithTitle(title As String, callback As FromControl(Of String)) As IMenuSeparatorBuilder
    End Interface

End Namespace