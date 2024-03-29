﻿Imports RibbonX.Controls.Base
Imports RibbonX.Controls.BuiltIn

Namespace Builders.Internal.Standardization

    Public Interface IInsert(Of T)

        Function InsertBefore(builtInControl As MSO) As T
        Function InsertBefore(qualifiedControl As RibbonElement) As T
        Function InsertAfter(builtInControl As MSO) As T
        Function InsertAfter(qualifiedControl As RibbonElement) As T

    End Interface

End Namespace